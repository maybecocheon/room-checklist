/**
 * Apps Script 전체 코드 (Code.gs).
 * 편집기의 기존 내용을 전부 지우고 이 파일 내용으로 통째로 교체하세요.
 *
 * [사전 준비]
 * 프로젝트 설정 → 스크립트 속성 → EDIT_PASSWORD = (비밀번호)
 * 비밀번호는 이 파일에 적지 않습니다.
 */

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
  var sheet = ss.getSheetByName("시트1"); // 시트 이름 확인 필요

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


function doPost(e) {
  var data = JSON.parse(e.postData.contents);

  // ⭐ 모든 쓰기 요청은 여기를 먼저 통과해야 한다.
  if (!isAuthorized(data)) return unauthorizedResponse();

  // 비밀번호 모달의 사전 확인 요청. 시트는 건드리지 않는다.
  if (data.action === 'verify') {
    return jsonResponse({ ok: true });
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("시트1");
  var action = data.action;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (action === 'create') {
    var newRow = headers.map(function(header) {
      switch(header) {
        case '방문일자': return data.visitDate;
        case '방 이름': return data.name;
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
        default: return '';
      }
    });
    sheet.appendRow(newRow);
    return jsonResponse({ ok: true, result: 'success', action: 'create' });
  }

  if (action === 'update') {
    var rowIndex = parseInt(data.rowIndex);
    var updatedRow = headers.map(function(header, index) {
      switch(header) {
        case '방문일자': return data.visitDate;
        case '방 이름': return data.name;
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
        default: return sheet.getRange(rowIndex, index + 1).getValue();
      }
    });
    sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
    return jsonResponse({ ok: true, result: 'success', action: 'update' });
  }

  if (action === 'delete') {
    var rowIndex = parseInt(data.rowIndex);
    sheet.deleteRow(rowIndex);
    return jsonResponse({ ok: true, result: 'success', action: 'delete' });
  }

  // ⭐ 알 수 없는 action 일 때도 반드시 JSON 을 돌려준다.
  //    아무것도 반환하지 않으면 Apps Script 가 CORS 헤더 없는 오류 페이지를 내보내
  //    브라우저에서는 원인을 알 수 없는 "Failed to fetch" 로만 보인다.
  return jsonResponse({ ok: false, error: 'unknown action: ' + action });
}


// 구글 드라이브 사진 가져오기 함수
function getDrivePhotos() {
  try {
    var folderId = "1jd9_mSxKyazV4OyCUlqItr6NLdrOKItD";

    var folder = DriveApp.getFolderById(folderId);
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
