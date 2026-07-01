// ============================================================
// PIVOT 진단 저장/수정용 Google Apps Script  (최종본)
// - 새 진단: 행 전체(A~U)를 한 번에 기록, 진단번호는 서버가 행 기준 자동 부여
// - 수정:   진단번호로 기존 행을 찾아 덮어쓰기
// ------------------------------------------------------------
// [갱신 방법]  ※ 코드 바꾼 뒤 반드시 "새 버전"으로 재배포
// 1) 편집기 코드 전부 지우고 이 내용 붙여넣기 → 저장(Cmd/Ctrl+S)
// 2) "배포" → "배포 관리" → 연필(수정) → 버전 "새 버전" → 액세스 "모든 사용자" → 배포
// ============================================================

var SHEET_ID = '1NKRJfO6V0GtjyDLAUDg4gnH7VE1IOebnGhDfDfAekU4';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var tab = data.discipline === 'vocal' ? '보컬' : '댄스';
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(tab);
    if (!sheet) return jsonOut({ ok: false, error: '탭 없음: ' + tab });

    var row = data.row;               // [학생ID, 진단번호, 진단일자, 트레이너, 이름, 학년, 트랙, 장르, 점수들..., 강점, 보완, 최종, 추천클래스, 추천트레이너, 주간일정]
    var maxRows = sheet.getMaxRows();

    // ── 수정: 진단번호(B열)로 기존 행을 찾아 덮어쓰기 ──
    if (data.action === 'update') {
      var nums = sheet.getRange(1, 2, maxRows, 1).getValues();
      var found = 0;
      for (var j = 0; j < nums.length; j++) {
        if (String(nums[j][0]) === String(data.number)) { found = j + 1; break; }
      }
      if (!found) return jsonOut({ ok: false, error: '수정 대상 진단번호 없음: ' + data.number });
      sheet.getRange(found, 1, 1, row.length).setValues([row]);
      return jsonOut({ ok: true, mode: 'update', row: found, cols: row.length });
    }

    // ── 새 진단: 진짜 마지막 데이터 행(E열=이름 기준) 다음에 기록 ──
    var names = sheet.getRange(1, 5, maxRows, 1).getValues();
    var last = 1;
    for (var i = 0; i < names.length; i++) {
      if (String(names[i][0]).trim() !== '') last = i + 1;
    }
    var target = last + 1;

    // 진단번호를 서버가 행 기준으로 부여 (중복 방지, 시트 수식과 동일 규칙: 연도-순번3자리)
    row[1] = new Date().getFullYear() + '-' + ('000' + (target - 1)).slice(-3);

    sheet.getRange(target, 1, 1, row.length).setValues([row]);
    return jsonOut({ ok: true, mode: 'create', row: target, number: row[1], cols: row.length });
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
