// ============================================================
// PIVOT 진단 저장용 Google Apps Script
// 사이트의 "새 진단 작성" 폼에서 보낸 데이터를 시트에 한 줄 추가한다.
// ------------------------------------------------------------
// [설치 방법]
// 1) 구글시트 열기 → 상단 메뉴 "확장 프로그램" → "Apps Script"
// 2) 기존 코드 전부 지우고 이 파일 내용 전체를 붙여넣기 → 저장(디스크 아이콘)
// 3) 오른쪽 위 "배포" → "새 배포"
//      - 유형(톱니바퀴): "웹 앱"
//      - 실행 계정: "나"
//      - 액세스 권한: "모든 사용자"
//    → "배포" → 권한 요청 뜨면 본인 계정으로 승인
// 4) 표시되는 "웹 앱 URL" 복사
// 5) index.html 의  const SAVE_URL = '';  안에 그 URL 붙여넣기 → git push
// ============================================================

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    // 종목에 따라 탭 선택 (탭 이름이 정확히 "댄스" / "보컬" 이어야 함)
    var tab = data.discipline === 'vocal' ? '보컬' : '댄스';
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tab);
    if (!sheet) return jsonOut({ ok: false, error: '탭을 찾을 수 없음: ' + tab });

    // data.row = [학생ID, 진단번호, 진단일자, 트레이너, 이름, 학년, 트랙, 장르, 점수들..., 강점, 보완, 최종평가, 추천클래스, 추천트레이너, 주간일정]
    sheet.appendRow(data.row);
    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
