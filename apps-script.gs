// ============================================================
// PIVOT 진단 저장용 Google Apps Script  (정확한 위치 저장 버전)
// - B열(진단번호)에 수식이 1000행까지 깔려 있어도, 진짜 마지막 데이터 행 다음에 저장한다.
// - B열은 건드리지 않아 기존 진단번호 수식이 그대로 자동 생성된다.
// ------------------------------------------------------------
// [갱신 방법]  ※ 코드를 바꾼 뒤엔 반드시 "새 버전"으로 재배포
// 1) 편집기 코드 전부 지우고 이 내용 붙여넣기 → 저장(Cmd/Ctrl+S)
// 2) "배포" → "배포 관리" → 연필(수정)
// 3) "버전" → "새 버전" 선택  ← 핵심
// 4) 액세스 권한 "모든 사용자" 확인 → "배포"
// ============================================================

var SHEET_ID = '1NKRJfO6V0GtjyDLAUDg4gnH7VE1IOebnGhDfDfAekU4';
var NAME_COL = 5; // E열 = 학생 이름 (이 열 기준으로 마지막 데이터 행을 찾음)

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var tab = data.discipline === 'vocal' ? '보컬' : '댄스';
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(tab);
    if (!sheet) return jsonOut({ ok: false, error: '탭 없음: ' + tab });

    var row = data.row;

    // 진짜 마지막 데이터 행: E열(이름)에 값이 있는 마지막 행
    var names = sheet.getRange(1, NAME_COL, sheet.getMaxRows(), 1).getValues();
    var last = 1;
    for (var i = 0; i < names.length; i++) {
      if (String(names[i][0]).trim() !== '') last = i + 1;
    }
    var target = last + 1;

    // A(학생ID) 기록, C부터 끝까지 기록. B(진단번호)는 시트 수식이 자동 생성하도록 건드리지 않음.
    sheet.getRange(target, 1).setValue(row[0]);
    if (row.length > 2) {
      sheet.getRange(target, 3, 1, row.length - 2).setValues([row.slice(2)]);
    }
    return jsonOut({ ok: true, row: target });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doGet() {
  return jsonOut({ ok: true, msg: 'PIVOT 저장 엔드포인트 정상 작동 중' });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
