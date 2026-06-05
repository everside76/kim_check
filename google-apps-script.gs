/**
 * ============================================================
 *  김치옥 MOT 점검 — 구글 스프레드시트 수신 스크립트
 *  웹앱(index.html)에서 보낸 점검 결과를 두 개의 시트에 누적 저장합니다.
 *    1) "점검요약"  : 점검 1건당 1행 (집계·통계용)
 *    2) "항목상세"  : 점검 항목 1건당 1행 (세부 분석용)
 *
 *  [설치 방법]  (자세한 그림 설명은 사용법 README 참고)
 *    1. 구글 스프레드시트 새로 만들기
 *    2. 상단 메뉴 [확장 프로그램] → [Apps Script] 클릭
 *    3. 기존 코드 전부 지우고 이 파일 내용 전체를 붙여넣기
 *    4. [배포] → [새 배포] → 유형 "웹 앱"
 *         - 실행 계정: 나
 *         - 액세스 권한: "모든 사용자"
 *    5. 생성된 웹 앱 URL(.../exec)을 복사해서 앱 "설정" 화면에 입력
 * ============================================================
 */

// 시트 이름 (원하면 바꿔도 됨)
var SHEET_SUMMARY = '점검요약';
var SHEET_DETAIL  = '항목상세';

// 요약 시트 헤더
var HEAD_SUMMARY = [
  '전송일시','점검ID','브랜드','지점명','피점검자','점검자','점검일','요일','점검구분','주문방식',
  '총점','등급','가산점','이행(○)','미흡(△)','불이행(✕)','미응답','10계명위반','총항목수','종합코멘트'
];
// 상세 시트 헤더
var HEAD_DETAIL = [
  '전송일시','점검ID','브랜드','지점명','점검일','No','단계','구분','항목','응답','감점','가산','비고'
];

/**
 * 웹앱 POST 수신 진입점
 */
function doPost(e) {
  var lock = LockService.getScriptLock();   // 동시 전송 충돌 방지
  try {
    lock.waitLock(20000);
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone() || 'Asia/Seoul';
    var stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

    var summary = getOrCreateSheet_(ss, SHEET_SUMMARY, HEAD_SUMMARY);
    var detail  = getOrCreateSheet_(ss, SHEET_DETAIL,  HEAD_DETAIL);

    // --- 중복 전송 방지: 같은 점검ID가 이미 있으면 무시 ---
    if (data.id && idExists_(summary, data.id)) {
      return json_({ result: 'success', duplicated: true, id: data.id });
    }

    // --- 1) 요약 1행 ---
    summary.appendRow([
      stamp,
      data.id || '',
      data.brand || '',
      data.store || '',
      data.target || '',
      data.inspector || '',
      data.date || '',
      data.weekday || '',
      data.inspectType || '',
      data.order || '',
      num_(data.score),
      data.grade || '',
      num_(data.bonus),
      num_(data.ok),
      num_(data.tri),
      num_(data.x),
      num_(data.unanswered),
      num_(data.ruleViolations),
      num_(data.total),
      data.comment || ''
    ]);

    // --- 2) 항목 상세 (응답/미응답 모든 항목) ---
    var rows = [];
    (data.items || []).forEach(function (it) {
      rows.push([
        stamp, data.id || '', data.brand || '', data.store || '', data.date || '',
        it.no || '', it.stage || '', it.type || '',
        it.text || '', it.label || '', num_(it.penalty), num_(it.bonus), it.note || ''
      ]);
    });
    // --- 10계명 위반도 상세 시트에 함께 기록 ---
    (data.rules || []).forEach(function (r) {
      rows.push([
        stamp, data.id || '', data.brand || '', data.store || '', data.date || '',
        '계명' + (r.no || ''), '10계명', r.key || '',
        r.desc || '', '위반', 5, 0, r.action || ''
      ]);
    });
    if (rows.length) {
      detail.getRange(detail.getLastRow() + 1, 1, rows.length, HEAD_DETAIL.length).setValues(rows);
    }

    return json_({ result: 'success', id: data.id, rows: rows.length });

  } catch (err) {
    return json_({ result: 'error', message: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/**
 * 브라우저로 URL 직접 열었을 때 동작 확인용
 */
function doGet() {
  return json_({ result: 'ok', message: '김치옥 MOT 수신 서버가 정상 동작 중입니다.' });
}

/* ---------------- 보조 함수 ---------------- */

function getOrCreateSheet_(ss, name, header) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  // 헤더가 비어 있으면 작성 + 서식
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.getRange(1, 1, 1, header.length)
      .setFontWeight('bold').setBackground('#F5B800').setFontColor('#1f1500');
    sh.setFrozenRows(1);
  }
  return sh;
}

function idExists_(sheet, id) {
  var last = sheet.getLastRow();
  if (last < 2) return false;
  // 요약 시트의 2번째 열(점검ID)에서 검색
  var ids = sheet.getRange(2, 2, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return true;
  }
  return false;
}

function num_(v) {
  var n = Number(v);
  return isNaN(n) ? (v || '') : n;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
