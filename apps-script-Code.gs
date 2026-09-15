/**
 * Apps Script 전체 코드 (Code.gs).
 * 편집기의 기존 내용을 전부 지우고 이 파일 내용으로 통째로 교체하세요.
 *
 * [사전 준비]
 * 프로젝트 설정 → 스크립트 속성 → EDIT_PASSWORD = (비밀번호)
 * 비밀번호는 이 파일에 적지 않습니다.
 *
 * [수정 후에는 반드시 재배포]
 * 배포 → 배포 관리 → ✏️ → 버전 "새 버전" → 배포
 */

// 사진을 저장·조회할 드라이브 폴더. 목록 조회와 업로드가 같은 폴더를 쓴다.
var PHOTO_FOLDER_ID = "1jd9_mSxKyazV4OyCUlqItr6NLdrOKItD";

var SHEET_NAME = "시트1";


function isAuthorized(data) {
  var secret = PropertiesService.getScriptProperties().getProperty('EDIT_PASSWORD');
  return !!secret && data && data.password === secret;
}

function unauthorizedResponse() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: 'unauthorized' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME); // 시트 이름 확인 필요

  if (e && e.parameter && e.parameter.action === 'getPhotos') {
    return getDrivePhotos();
  }

  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var data = [];

  // ⭐ 스프레드시트 시간대(예: 아시아/서울) 기준으로 날짜를 문자열 변환하기 위한 값
  var tz = ss.getSpreadsheetTimeZone();

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var obj = { rowIndex: i + 1 }; // 1부터 시작하는 실제 행 번호 저장
    for (var j = 0; j < headers.length; j++) {
      var val = row[j];
      // ⭐ 날짜 셀(Date 객체)은 UTC로 직렬화되면서 하루가 밀리므로,
      //    시트 시간대 기준 'yyyy-MM-dd' 문자열로 미리 변환해서 내려준다.
      if (val instanceof Date) {
        obj[headers[j]] = Utilities.formatDate(val, tz, "yyyy-MM-dd");
      } else {
        obj[headers[j]] = val;
      }
    }
    data.push(obj);
  }

  return jsonResponse(data);
}


/**
 * ⭐ 쓰기 요청은 반드시 스크립트 락 안에서 처리한다.
 *    두 사람이 동시에 저장하면 appendRow / deleteRow 가 겹치면서
 *    엉뚱한 행이 덮이거나 지워질 수 있기 때문이다.
 *    인증·검증처럼 시트를 건드리지 않는 요청은 락을 잡지 않는다.
 */
function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ ok: false, error: '요청 본문을 해석하지 못했습니다.' });
  }

  if (!isAuthorized(data)) return unauthorizedResponse();

  // 비밀번호 모달의 사전 확인 요청. 시트는 건드리지 않는다.
  if (data.action === 'verify') {
    return jsonResponse({ ok: true });
  }

  // 사진 업로드는 드라이브만 쓰므로 시트 락이 필요 없다.
  if (data.action === 'uploadPhoto') {
    return uploadPhoto(data);
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);   // 최대 20초 대기
  } catch (err) {
    return jsonResponse({ ok: false, error: '다른 저장이 진행 중입니다. 잠시 후 다시 시도해 주세요.' });
  }

  try {
    return handleWrite(data);
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}


function handleWrite(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var action = data.action;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (action === 'create') {
    var newRow = headers.map(function(header) {
      return valueForHeader(header, data, null, null);
    });
    sheet.appendRow(newRow);
    return jsonResponse({ ok: true, result: 'success', action: 'create' });
  }

  if (action === 'update') {
    var rowIndex = parseInt(data.rowIndex);
    // ⭐ 기존 행을 한 번에 읽어둔다. 폼이 보내지 않는 열(예: 진행체크)은 그대로 보존.
    var existing = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    var updatedRow = headers.map(function(header, index) {
      return valueForHeader(header, data, existing[index], index);
    });
    sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
    return jsonResponse({ ok: true, result: 'success', action: 'update' });
  }

  // ⭐ 체크리스트는 '진행체크' 한 칸만 건드린다. 나머지 열은 읽지도 쓰지도 않는다.
  if (action === 'updateProgress') {
    var col = headers.indexOf('진행체크') + 1;
    if (col === 0) {
      return jsonResponse({ ok: false, error: "시트 첫 행에 '진행체크' 열을 추가해 주세요." });
    }
    sheet.getRange(parseInt(data.rowIndex), col).setValue(data.progress || '');
    return jsonResponse({ ok: true, result: 'success', action: 'updateProgress' });
  }

  if (action === 'delete') {
    sheet.deleteRow(parseInt(data.rowIndex));
    return jsonResponse({ ok: true, result: 'success', action: 'delete' });
  }

  // ⭐ 알 수 없는 action 일 때도 반드시 JSON 을 돌려준다.
  //    아무것도 반환하지 않으면 Apps Script 가 CORS 헤더 없는 오류 페이지를 내보내
  //    브라우저에서는 원인을 알 수 없는 "Failed to fetch" 로만 보인다.
  return jsonResponse({ ok: false, error: 'unknown action: ' + action });
}


/**
 * 열 이름 → 저장할 값. create 와 update 가 같은 표를 쓴다.
 * 폼이 다루지 않는 열은 기존 값(fallback)을 그대로 돌려줘 덮어쓰지 않는다.
 */
function valueForHeader(header, data, fallback, index) {
  switch (header) {
    case '방문일자': return data.visitDate;
    case '방 이름': return data.name;
    case '주소': return data.address || '';
    case '부동산 이름': return data.agencyName;
    case '평가': return data.status;
    case '보증금(만원)': return data.deposit;
    case '월세+관리비(만원)': return data.rent;
    case '도보거리(분)': return data.walkTime;
    case '구조': return data.structure;
    case '주차': return data.parking;
    case '방향': return data.direction;
    case '채광': return data.lighting;
    case '수압': return data.waterPressure;
    case '콘센트': return data.outlets;
    case '엘리베이터': return data.elevator;
    case '중문': return data.middleDoor;
    case '공동현관비번': return data.buildingLock;
    case '곰팡이': return data.mould;
    case '누수': return data.leak;
    case '벌레': return data.bugs;
    case '옵션': return data.options;
    case '최종점수': return data.score;
    case '사진URL': return data.photoUrl;
    case '메모/특이사항': return data.memo;
    // '진행체크' 를 포함해 폼이 보내지 않는 열은 기존 값 유지
    default: return fallback === null || fallback === undefined ? '' : fallback;
  }
}


/**
 * 📸 현장에서 찍은 사진 업로드.
 * 브라우저가 긴 변 1600px JPEG 로 줄여 base64 로 보내고, 여기서 드라이브 파일로 만든다.
 */
function uploadPhoto(data) {
  try {
    if (!data.dataBase64) return jsonResponse({ ok: false, error: '이미지 데이터가 비어 있습니다.' });

    var name = data.fileName || ('room-' + new Date().getTime() + '.jpg');
    var blob = Utilities.newBlob(
      Utilities.base64Decode(data.dataBase64),
      data.mimeType || 'image/jpeg',
      name
    );

    var file = DriveApp.getFolderById(PHOTO_FOLDER_ID).createFile(blob);

    // 대시보드는 로그인 없이 열리므로 링크가 있으면 볼 수 있게 열어준다.
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      // 조직 정책으로 공유가 막힌 경우에도 업로드 자체는 성공으로 둔다.
    }

    return jsonResponse({
      ok: true,
      id: file.getId(),
      name: file.getName(),
      url: "https://lh3.googleusercontent.com/d/" + file.getId()
    });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}


// 구글 드라이브 사진 가져오기 함수
function getDrivePhotos() {
  try {
    var folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
    var files = folder.getFiles();
    var photos = [];

    while (files.hasNext()) {
      var file = files.next();
      var mimeType = file.getMimeType();
      if (mimeType.indexOf("image/") === 0) {
        photos.push({
          id: file.getId(),
          name: file.getName(),
          url: "https://lh3.googleusercontent.com/d/" + file.getId()
        });
      }
    }

    return jsonResponse(photos);
  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}
