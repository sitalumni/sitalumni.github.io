/**
 * 첨단융합공학부 동문회 온라인 마라톤 — 리더보드 API
 *
 * [배포 방법]
 * 1. 구글 시트 메뉴 → 확장 프로그램 → Apps Script
 * 2. 이 코드 전체를 붙여넣고 저장(Ctrl+S)
 * 3. 오른쪽 위 [배포] → [새 배포]
 *    - 유형: 웹 앱
 *    - 설명: 리더보드 API
 *    - 다음 사용자로 실행: 나(스크립트 소유자)   ← 시트에 직접 접근
 *    - 액세스 권한: 모든 사용자(익명 포함)       ← 페이지가 fetch 할 수 있도록
 * 4. [배포] → 나오는 웹 앱 URL 복사
 * 5. index.html 의 CONFIG.APPS_SCRIPT_URL 에 붙여넣기
 *
 * [시트 열 구조] — 헤더 이름에 아래 키워드가 포함되면 자동 인식 (순서 무관)
 *   학번 · 이름(또는 성함) · 트랙(또는 종목·코스) · 거리 · 기록(또는 시간·페이스·pace)
 *   검증(또는 확인·verif) — 선택. ✓ · y · 예 · 1 이면 검증됨 배지
 *
 * [공개되는 정보] 학번, 이름, 트랙, 거리, 기록, 검증 여부 — 그 외 전부 제외
 */

// 헤더에서 키워드를 포함하는 열 인덱스를 찾는 헬퍼
function findCol(headers, keywords) {
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).toLowerCase();
    for (var k = 0; k < keywords.length; k++) {
      if (h.indexOf(keywords[k]) !== -1) return i;
    }
  }
  return -1;
}

function isVerified(v) {
  var s = String(v || '').trim().toLowerCase();
  return ['✓','y','yes','예','o','true','1','검증','확인'].indexOf(s) !== -1;
}

function normTrack(v) {
  var s = String(v || '').toLowerCase().replace(/\s/g, '');
  if (s.indexOf('10') !== -1) return '10km';
  if (s.indexOf('5') !== -1)  return '5km';
  return '';
}

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var data  = sheet.getDataRange().getValues();

    if (data.length < 2) {
      return jsonResponse({ records: [] });
    }

    var H      = data[0];
    var iId    = findCol(H, ['학번']);
    var iName  = findCol(H, ['이름', '성함']);
    var iTrack = findCol(H, ['트랙', '종목', '코스']);
    var iDist  = findCol(H, ['거리']);
    var iRec   = findCol(H, ['기록', '시간', '페이스', 'pace']);
    var iVer   = findCol(H, ['검증', '확인', 'verif']);

    var records = [];
    for (var r = 1; r < data.length; r++) {
      var row  = data[r];
      var name = String(iName  >= 0 ? row[iName]  : '').trim();
      var track = normTrack(iTrack >= 0 ? row[iTrack] : '');
      if (!name && !track) continue;   // 빈 행 스킵

      var distRaw = iDist >= 0 ? String(row[iDist]).replace(/[^0-9.]/g, '') : '';
      records.push({
        studentId: String(iId  >= 0 ? row[iId]  : '').trim(),
        name:      name,
        track:     track,
        distance:  distRaw ? parseFloat(distRaw) : 0,
        time:      String(iRec >= 0 ? row[iRec] : '').trim(),
        verified:  iVer >= 0 ? isVerified(row[iVer]) : false
        // 사진, 이메일, 타임스탬프 등 나머지 열은 응답에 포함하지 않음
      });
    }

    return jsonResponse({ records: records });

  } catch (err) {
    return jsonResponse({ error: String(err), records: [] });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
