/**
 * CARE ID · DPP 셋업 사전 설문 → Google 스프레드시트 저장 (지정 시트 연결 버전)
 *
 * 아래 SHEET_ID 시트의 첫 번째 탭에 응답을 한 행씩 저장합니다.
 * ※ 이 스크립트 실행 계정(예: mntc.staff@gmail.com)이 해당 시트에 '편집' 권한을 가지고 있어야 합니다.
 *
 * [재배포 방법 — URL 유지]
 * 1) 편집기에서 기존 코드를 모두 지우고 이 내용을 붙여넣고 저장(💾).
 * 2) [배포] → [배포 관리] → 기존 배포의 ✏️(연필) → "버전"을 [새 버전]으로 → [배포].
 *    (이렇게 해야 웹앱 URL이 그대로 유지됩니다.)
 */

const SHEET_ID = '1K2O09x75DPPetz0vLeEnqAKqqKJEl3Tzc28nkJXpDp0';

function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; // 첫 번째 탭 사용
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // 동시 제출 시 행 꼬임 방지
  try {
    const sheet = getSheet_();
    const data = JSON.parse(e.postData.contents);

    // 현재 헤더(첫 행) 읽기
    let headers = sheet.getLastColumn() > 0
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].filter(String)
      : [];

    // 들어온 데이터의 새 key를 헤더에 추가 (문항이 바뀌어도 자동 대응)
    let changed = false;
    Object.keys(data).forEach(k => {
      if (headers.indexOf(k) === -1) { headers.push(k); changed = true; }
    });
    if (changed || sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // 헤더 순서에 맞춰 한 행 추가
    const row = headers.map(h => (data[h] !== undefined && data[h] !== null) ? data[h] : '');
    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput(
    'CARE ID DPP survey endpoint is running.\n저장 시트: https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit'
  );
}
