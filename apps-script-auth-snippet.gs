/**
 * Apps Script 쪽에 붙여넣을 인증 코드.
 *
 * [사전 준비]
 * Apps Script 편집기 → 좌측 톱니바퀴(프로젝트 설정) → "스크립트 속성" →
 *   속성: EDIT_PASSWORD  /  값: (새로 정한 비밀번호)
 * 비밀번호는 절대 이 파일에 적지 않는다.
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

/**
 * 기존 doPost 의 맨 앞에 아래 3줄을 넣으면 된다.
 *
 * function doPost(e) {
 *   var data = JSON.parse(e.postData.contents);
 *   if (!isAuthorized(data)) return unauthorizedResponse();
 *
 *   ... 기존 create / update / delete 처리 ...
 * }
 *
 * 주의: 기존 코드가 이미 JSON.parse 를 하고 있다면 중복 파싱하지 말고
 *       그 변수를 그대로 isAuthorized 에 넘길 것.
 */
