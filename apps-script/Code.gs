/**
 * 「24시간!」 청소년 정치 성향 설문 — Google Sheets 저장 엔드포인트
 *
 * 사용법
 * 1. 새 Google Sheets 파일을 생성합니다.
 * 2. 메뉴에서 [확장 프로그램 ▸ Apps Script] 를 엽니다.
 * 3. 기본으로 만들어진 Code.gs 의 내용을 모두 지우고, 이 파일의 내용을 그대로 붙여넣습니다.
 * 4. 저장한 뒤, 우측 상단 [배포] ▸ [새 배포] ▸ 유형: 웹 앱 선택.
 *    - 다음 사용자 인증: 본인
 *    - 액세스 권한:       모든 사용자
 *    위 두 옵션을 반드시 위와 같이 설정하고 배포합니다.
 * 5. 배포가 끝나면 표시되는 "웹 앱 URL" 을 복사합니다.
 *    형식: https://script.google.com/macros/s/................/exec
 * 6. 그 URL 을 사이트의 index.html 상단의 WEBHOOK_URL 상수에 붙여넣고 다시 배포하면 끝.
 */

const SHEET_NAME = '응답';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    // 시트가 비어 있으면 헤더부터 작성
    if (sheet.getLastRow() === 0) {
      const headers = [
        '제출시각', '소속학교', '학번',
        'Q1_성평등', 'Q2_성소수자', 'Q3_언론', 'Q4_장애인',
        'Q5_교육', 'Q6_할당제', 'Q7_환경', 'Q8_비정규직',
        'Q9_지원금', 'Q10_기본소득', 'Q11_부동산', 'Q12_최저임금', 'Q13_부의재분배',
        'Q14_전쟁', 'Q15_한미동맹', 'Q16_평화vs국익', 'Q17_대북정책',
        '사회평균', '경제평균', '안보평균',
        '전체평균', '판정카테고리', '본인인식', '2차실험참가',
      ];
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
    }

    const a = body.answers || [];
    const row = [
      new Date(),
      body.school || '',
      body.studentId || '',
      ...a,                          // 17개 점수
      body.bySection?.society ?? '',
      body.bySection?.economy ?? '',
      body.bySection?.security ?? '',
      body.avg ?? '',
      body.category ?? '',
      body.selfIdeology ?? '',
      body.joinSecond == null ? '' : (body.joinSecond ? 'Y' : 'N'),
    ];
    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: '24h survey endpoint alive' }))
    .setMimeType(ContentService.MimeType.JSON);
}
