/**
 * 출시 알림 신청 → 구글 시트 한 줄씩 쌓기
 *
 * 연결하는 법 (한 번만)
 * 1. 구글 시트를 새로 만든다(이름 예: Handoff 출시 알림).
 * 2. 시트 메뉴 [확장 프로그램] → [Apps Script]를 열고, 이 파일 내용을 통째로 붙여 넣고 저장한다.
 * 3. 오른쪽 위 [배포] → [새 배포] → 유형 [웹 앱]
 *    - 다음 사용자로 실행: 나
 *    - 액세스 권한이 있는 사용자: 모든 사용자
 *    → [배포]를 누르고 권한을 허용한다.
 * 4. 나오는 웹 앱 URL(https://script.google.com/macros/s/…/exec)을
 *    index.html의 WAITLIST_ENDPOINT = '' 따옴표 안에 넣는다.
 *
 * 이 파일을 고친 뒤에는 [배포] → [배포 관리] → 연필 → 버전 [새 버전]으로 다시 배포해야 반영된다
 * (URL은 그대로 유지된다).
 */

const SHEET_NAME = 'waitlist';

// =, +, -, @로 시작하는 값은 시트가 수식으로 실행한다 — 앞에 '를 붙여 글자로 저장
function asText(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // 동시에 들어온 신청이 같은 줄을 덮어쓰지 않게
  try {
    const email = String((e && e.parameter && e.parameter.email) || '').trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 254) {
      return ContentService.createTextOutput('invalid');
    }

    const book = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = book.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = book.insertSheet(SHEET_NAME);
      sheet.appendRow(['신청 시각', '이메일', '유입 경로']);
      sheet.setFrozenRows(1);
    }

    // 같은 주소로 두 번 신청하면 한 줄만 남긴다
    const dup = sheet.getRange('B:B').createTextFinder(email).matchEntireCell(true).findNext();
    if (dup) return ContentService.createTextOutput('duplicate');

    const referrer = String(e.parameter.referrer || '').slice(0, 500);
    sheet.appendRow([new Date(), asText(email), asText(referrer)]);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}
