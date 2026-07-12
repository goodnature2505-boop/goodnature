/*───────────────────────────────────────────────────────────────────────────
  GoodNature 견적·문의 폼  →  Google Sheet + 담당자 메일
  (Google Apps Script Web App — goodnature2505 계정에 배포)

  - goodnature.uk/support 폼이 POST로 접수
  - Cloudflare Turnstile 서버검증(스크립트 속성 TURNSTILE_SECRET 있을 때만)
  - 연락처: 숫자만 저장 + '연락처' 컬럼을 텍스트서식(@)으로 → 앞자리 0 보존
  - 하이픈(010-1234-5678)은 시트 H열 ARRAYFORMULA가 자동생성
  - 담당자 메일 알림(replyTo=문의자 이메일)

  배포: 저장 → 배포 → 새 배포 → 웹 앱
        · 실행 계정 = 나(goodnature2505)
        · 액세스   = 모든 사용자(Anyone)
        → 나오는 /exec URL 을 폼에 연결
  (코드 수정 후엔: 배포 > 배포 관리 > 연필 > 버전 '새 버전' > 배포 — URL 유지)
───────────────────────────────────────────────────────────────────────────*/

var SHEET_ID   = '1DAH1nwYUtxUe-K1m5nYqsWZIVf9YP_YhMV05zCbwRdY';
var SHEET_NAME = 'responses';
var NOTIFY_TO  = 'goodnature@goodnature.uk';
var SUBJECT    = '[GoodNature 홈페이지 문의]';

var HEADERS = ['타임스탬프','회사/성명','연락처','이메일','문의유형','처리대상/물량','문의내용'];

function doPost(e){
  try{
    var p = (e && e.parameter) || {};

    // 허니팟 채워지면 스팸 → 조용히 통과(기록·메일 안함)
    if (p['_honey'] || p['botcheck']) return _ok();

    // Turnstile 서버검증 (스크립트 속성에 secret 있을 때만)
    var secret = PropertiesService.getScriptProperties().getProperty('TURNSTILE_SECRET');
    if (secret && !verifyTurnstile_(secret, p['cf-turnstile-response'])) return _fail('captcha');

    var sh = _sheet_();

    // 폼이 이미 하이픈 포맷하므로 재포맷 안 함. 앞 ' = 텍스트 처리(0 보존 + '+82' 수식오류 방지)
    var row = [
      new Date(),
      p.name  || '',
      "'" + (p.phone || ''),
      p.email || '',
      p.type  || '',
      p.scope || '',
      p.msg   || ''
    ];
    sh.appendRow(row);

    _notify_(row);
    return _ok();
  }catch(err){
    return _fail(String(err));
  }
}

// 간단한 헬스체크(브라우저로 /exec 열면 보임)
function doGet(){
  return ContentService.createTextOutput('GoodNature form endpoint OK');
}

function _sheet_(){
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0){
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function verifyTurnstile_(secret, token){
  if (!token) return false;
  try{
    var res = UrlFetchApp.fetch('https://challenges.cloudflare.com/turnstile/v0/siteverify', {
      method:'post',
      payload:{ secret:secret, response:token },
      muteHttpExceptions:true
    });
    var j = JSON.parse(res.getContentText());
    return !!(j && j.success === true);
  }catch(e){ return false; }
}

function _notify_(row){
  try{
    var body =
      '새 문의가 접수되었습니다.\n\n' +
      '· 회사/성명 : ' + row[1] + '\n' +
      '· 연락처    : ' + String(row[2]).replace(/^'/,'') + '\n' +
      '· 이메일    : ' + row[3] + '\n' +
      '· 문의유형  : ' + row[4] + '\n' +
      '· 처리대상/물량 : ' + row[5] + '\n' +
      '· 문의내용  :\n' + row[6] + '\n\n' +
      '접수시각 : ' + row[0];
    MailApp.sendEmail({
      to: NOTIFY_TO,
      subject: SUBJECT,
      body: body,
      replyTo: (row[3] || NOTIFY_TO)
    });
  }catch(e){ /* 메일 실패해도 시트 접수는 유지 */ }
}

function _ok(){  return ContentService.createTextOutput(JSON.stringify({result:'ok'})).setMimeType(ContentService.MimeType.JSON); }
function _fail(m){ return ContentService.createTextOutput(JSON.stringify({result:'error',error:String(m)})).setMimeType(ContentService.MimeType.JSON); }
