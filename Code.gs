/**
 * 온라인 출석부 서명 저장소 (Google Apps Script)
 * 구글 시트 > 확장 프로그램 > Apps Script 에 이 코드를 통째로 붙여넣고
 * 아래 ADMIN_KEY 만 바꾼 뒤 "배포 > 새 배포 > 웹 앱"으로 배포하세요.
 *   - 다음 사용자 인증 정보로 실행: 나
 *   - 액세스 권한이 있는 사용자: 모든 사용자
 */
const ADMIN_KEY = '여기에-관리자-비밀번호를-적으세요';
const SHEET_NAME = '서명';
const MAX_SIG_LENGTH = 49000; // 시트 한 칸 최대 50,000자

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['번호', '소속', '성명', '서명(이미지 데이터)', '서명 시각']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function rows_() {
  const sh = sheet_();
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 5).getValues().map(function (r, i) {
    return { row: i + 2, no: Number(r[0]), school: String(r[1]), name: String(r[2]), sig: String(r[3]), time: r[4] };
  });
}

function fmt_(t) {
  return t instanceof Date ? Utilities.formatDate(t, 'Asia/Seoul', 'MM-dd HH:mm') : String(t || '');
}

function safe_(s) {
  s = String(s || '').slice(0, 50);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// GET ?action=status            → 누가 서명했는지(번호·시각)만 공개
// GET ?action=list&key=관리자키  → 서명 이미지까지 전체 (관리자용)
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    const rows = rows_();
    if (p.action === 'list') {
      if (p.key !== ADMIN_KEY) return out_({ ok: false, error: '관리자 비밀번호가 맞지 않습니다.' });
      return out_({ ok: true, items: rows.map(function (r) {
        return { no: r.no, school: r.school, name: r.name, sig: r.sig, time: fmt_(r.time) };
      }) });
    }
    return out_({ ok: true, items: rows.map(function (r) { return { no: r.no, time: fmt_(r.time) }; }) });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

// POST {action:'sign', no, school, name, sig}
// POST {action:'reset', no, key}
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const no = Number(d.no);
    if (!(no >= 1 && no <= 500 && Math.floor(no) === no)) return out_({ ok: false, error: '번호가 올바르지 않습니다.' });

    if (d.action === 'sign') {
      const sig = String(d.sig || '');
      if (sig.indexOf('data:image/png;base64,') !== 0) return out_({ ok: false, error: '서명 이미지가 올바르지 않습니다.' });
      if (sig.length > MAX_SIG_LENGTH) return out_({ ok: false, error: '서명 이미지가 너무 큽니다.' });
      const exists = rows_().some(function (r) { return r.no === no; });
      if (exists) return out_({ ok: false, error: '이미 서명이 등록되어 있습니다. 다시 하려면 담당자에게 초기화를 요청하세요.' });
      sheet_().appendRow([no, safe_(d.school), safe_(d.name), sig, new Date()]);
      return out_({ ok: true });
    }

    if (d.action === 'reset') {
      if (d.key !== ADMIN_KEY) return out_({ ok: false, error: '관리자 비밀번호가 맞지 않습니다.' });
      const sh = sheet_();
      const rows = rows_().filter(function (r) { return r.no === no; }).reverse();
      rows.forEach(function (r) { sh.deleteRow(r.row); });
      return out_({ ok: true, removed: rows.length });
    }

    return out_({ ok: false, error: '알 수 없는 요청입니다.' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
