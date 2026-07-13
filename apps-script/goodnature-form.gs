/*───────────────────────────────────────────────────────────────────────────
  GoodNature 견적·문의 폼  →  Google Sheet + 담당자 메일
  (Google Apps Script Web App — goodnature2505 계정에 배포)

  - goodnature.uk/support 폼이 POST로 접수
  - Cloudflare Turnstile 서버검증(스크립트 속성 TURNSTILE_SECRET 있을 때만)
  - 연락처: 폼이 이미 하이픈 포맷 → 시트엔 앞 ' 붙여 텍스트로 저장(앞0 보존 + +82 수식오류 방지)
  - 알림메일: Brevo REST 발송(스크립트속성 BREVO_API_KEY = xkeysib- 키). 보내는=GoodNature<goodnature2505@gmail.com>,
    받는=goodnature@goodnature.uk(도메인, 답장시 자동 From=도메인), replyTo=문의자, 제목 말머리 [GoodNature #0001].
    (BREVO_API_KEY 미설정 시 MailApp 폴백)

  배포: 저장 → 배포 → 새 배포 → 웹 앱
        · 실행 계정 = 나(goodnature2505)
        · 액세스   = 모든 사용자(Anyone)
        → 나오는 /exec URL 을 폼에 연결
  (코드 수정 후엔: 배포 > 배포 관리 > 연필 > 버전 '새 버전' > 배포 — URL 유지)
───────────────────────────────────────────────────────────────────────────*/

var SHEET_ID   = '1DAH1nwYUtxUe-K1m5nYqsWZIVf9YP_YhMV05zCbwRdY';
var SHEET_NAME = 'responses';
var NOTIFY_TO    = 'goodnature@goodnature.uk';   // 알림 받는 주소(도메인). Cloudflare 라우팅→gmail, 답장시 From=이 주소
var SENDER_EMAIL = 'goodnature2505@gmail.com';   // Brevo Senders에 Verified 된 발신 주소
var SEND_AS      = 'goodnature@goodnature.uk';   // (MailApp 폴백용) send-as 별칭 있으면 사용
var SUBJECT      = '홈페이지 문의';  // 실제 제목 = [GoodNature #0001] + 이 문구
// ⚠ Brevo API 키(xkeysib-...)는 코드에 넣지 말 것 → 스크립트 속성 BREVO_API_KEY 에 저장(공개 repo 보호)

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

    // 말머리 티켓번호 = 데이터 행번호(헤더 제외). 4자리 제로패딩(10000+ 자동 5자리)
    var ticket = Utilities.formatString('%04d', Math.max(1, sh.getLastRow() - 1));
    _notify_(row, ticket);
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

function _notify_(row, ticket){
  var subject = '[GoodNature #' + ticket + '] ' + SUBJECT;
  var body =
    '새 문의가 접수되었습니다. (접수번호 #' + ticket + ')\n\n' +
    '· 회사/성명 : ' + row[1] + '\n' +
    '· 연락처    : ' + String(row[2]).replace(/^'/,'') + '\n' +
    '· 이메일    : ' + row[3] + '\n' +
    '· 문의유형  : ' + row[4] + '\n' +
    '· 처리대상/물량 : ' + row[5] + '\n' +
    '· 문의내용  :\n' + row[6] + '\n\n' +
    '접수시각 : ' + row[0];
  // replyTo = 문의자 이메일(형식 유효할 때만), 아니면 받는주소
  var replyEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(row[3])) ? row[3] : NOTIFY_TO;

  // 1) Brevo REST (스크립트속성 BREVO_API_KEY 있으면)
  var key = PropertiesService.getScriptProperties().getProperty('BREVO_API_KEY');
  if (key){
    try{
      UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/email', {
        method:'post',
        contentType:'application/json',
        headers:{ 'api-key': key, 'accept':'application/json' },
        muteHttpExceptions:true,
        payload: JSON.stringify({
          sender:  { name:'GoodNature', email: SENDER_EMAIL },
          to:      [{ email: NOTIFY_TO }],
          replyTo: { email: replyEmail },
          subject: subject,
          textContent: body
        })
      });
      return;
    }catch(e){ /* Brevo 실패 → MailApp 폴백 */ }
  }

  // 2) 폴백: MailApp
  try{
    var opts = { to: NOTIFY_TO, subject: subject, body: body, name:'GoodNature', replyTo: replyEmail };
    try{ if (GmailApp.getAliases().indexOf(SEND_AS) !== -1) opts.from = SEND_AS; }catch(e){}
    MailApp.sendEmail(opts);
  }catch(e){ /* 메일 실패해도 시트 접수는 유지 */ }
}

function _ok(){  return ContentService.createTextOutput(JSON.stringify({result:'ok'})).setMimeType(ContentService.MimeType.JSON); }
function _fail(m){ return ContentService.createTextOutput(JSON.stringify({result:'error',error:String(m)})).setMimeType(ContentService.MimeType.JSON); }
