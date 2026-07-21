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
 *   학번 · 이름(또는 성함) · 트랙(또는 종목·코스·track) · 거리 · 기록(총 소요시간, 예 1:24:00 또는 55:30 — 시간·기록·time 키워드)
 *   닉네임(또는 nickname·닉) — 선택. 있으면 이름 대신 표시(가리지 않음)
 *   검증(또는 확인·verif) — 선택. ✓ · y · 예 · 1 이면 검증됨 배지
 *
 * [개인정보 보호] 실명은 서버(여기)에서만 처리하고 절대 반환하지 않습니다.
 *   - 사람(학번+실명)이 한 번이라도 닉네임을 냈으면 → 그 닉네임으로 전체 통일 표시
 *   - 한 번도 안 냈으면 → 실명 가운데를 가린 표시명(홍길동→홍*동)만 반환
 *   - 사람 구분용 key 는 실명 해시라 실명을 역산할 수 없습니다.
 *
 * [공개되는 정보] 학번, 표시명(닉네임/마스킹), 트랙, 거리, 기록, 검증 여부 — 그 외 전부 제외
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

// 이름 가운데 글자 가리기: 홍길동→홍*동, 김철→김*, 남궁민수→남**수
function maskName(name) {
  var n = String(name || '').trim();
  if (n.length <= 1) return n;
  if (n.length === 2) return n.charAt(0) + '*';
  return n.charAt(0) + new Array(n.length - 1).join('*') + n.charAt(n.length - 1);
}

// 실명을 역산할 수 없는 안정적 구분 키(학번+실명의 MD5 앞 12자리)
function personHash(studentId, name) {
  var raw = String(studentId || '') + '' + String(name || '');
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < 6; i++) { var b = (bytes[i] + 256) % 256; hex += ('0' + b.toString(16)).slice(-2); }
  return hex;
}

// 거리 입력 강건 파싱: "5km" · "5 km" · "5.5킬로" · "10K" · "5,5" → 숫자
function parseDistance(v) {
  var s = String(v == null ? '' : v).trim().replace(/,/g, '.');
  var m = s.match(/\d+(\.\d+)?/);
  return m ? parseFloat(m[0]) : 0;
}

// 제출 타임스탬프 → "yy.MM.dd HH:mm" (개인 상세에서 본인 기록 구분용)
function fmtTs(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yy.MM.dd HH:mm');
  return String(v == null ? '' : v).trim();   // 문자열이면 그대로
}

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    // getDisplayValues: 셀에 보이는 '텍스트 그대로' 읽음.
    //  - 기록 "1:24:00"·"55:30" 이 Date/숫자로 자동 변환되는 것 방지
    //  - 학번 "08" 앞자리 0 보존
    var data  = sheet.getDataRange().getDisplayValues();

    if (data.length < 2) {
      return jsonResponse({ records: [] });
    }

    var H      = data[0];
    var iId    = findCol(H, ['학번']);
    var iName  = findCol(H, ['이름', '성함']);
    var iNick  = findCol(H, ['닉네임', 'nickname', '닉']);
    var iTrack = findCol(H, ['트랙', '종목', '코스', 'track']);
    var iDist  = findCol(H, ['거리']);
    var iRec   = findCol(H, ['기록', '소요', '시간', 'time', '페이스', 'pace']);
    var iVer   = findCol(H, ['검증', '확인', 'verif']);
    var iTime  = findCol(H, ['타임스탬프', 'timestamp']);   // 폼 제출 시각(보통 첫 열)

    // 타임스탬프는 실제 Date 값으로 읽어 시간대 안전하게 포맷 (해당 열만 추가로 읽음)
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    var timeVals = iTime >= 0 ? sheet.getRange(1, iTime + 1, data.length, 1).getValues() : null;

    // ── 1차: 행을 파싱하고, 사람(학번+실명)별 닉네임을 모음
    //    같은 사람이 한 번이라도 닉네임을 냈으면 그 닉네임으로 통일(마지막 제출값 우선)
    var parsed = [];
    var nickByPerson = {};   // "학번|실명" → 닉네임
    for (var r = 1; r < data.length; r++) {
      var row  = data[r];
      var name = String(iName  >= 0 ? row[iName]  : '').trim();     // 실명 (반환하지 않음)
      var nick = String(iNick  >= 0 ? row[iNick]  : '').trim();     // 닉네임
      var track = normTrack(iTrack >= 0 ? row[iTrack] : '');
      if (!name && !track) continue;   // 빈 행 스킵

      var studentId = String(iId >= 0 ? row[iId] : '').trim();
      var pid = studentId + '|' + name;   // 사람 구분(서버 내부용, 반환 안 함)
      if (nick) nickByPerson[pid] = nick;

      parsed.push({
        pid:       pid,
        studentId: studentId,
        realName:  name,
        track:     track,
        distance:  iDist >= 0 ? parseDistance(row[iDist]) : 0,
        time:      String(iRec >= 0 ? row[iRec] : '').trim(),
        date:      timeVals ? fmtTs(timeVals[r][0], tz) : '',
        verified:  iVer >= 0 ? isVerified(row[iVer]) : false
      });
    }

    // ── 2차: 사람별 표시명 확정(닉네임 있으면 닉네임, 없으면 마스킹 실명) 후 응답 구성
    var records = parsed.map(function (p) {
      var display = nickByPerson[p.pid] ? nickByPerson[p.pid] : maskName(p.realName);
      return {
        studentId: p.studentId,
        name:      display,                                          // 실명 원본 대신 표시명만
        key:       p.studentId + '|' + personHash(p.studentId, p.realName),  // 실명 역산 불가 구분키
        track:     p.track,
        distance:  p.distance,
        time:      p.time,
        date:      p.date,     // 제출 시각 (개인 상세에서 본인 기록 구분용)
        verified:  p.verified
        // 실명, 사진, 이메일 등 나머지 열은 응답에 포함하지 않음
      };
    });

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
