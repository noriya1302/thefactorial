/**
 * 팩토리얼 모의고사 주문 시트
 * - 사이트에서 주문/신청/취소가 생기면 이 시트에 한 줄씩 쌓이고 메일 알림이 갑니다.
 * - "상태" 칸을 바꾸면 학생 화면의 진행 상황이 바로 바뀝니다.
 * 처음 한 번: 위쪽 메뉴 [팩토리얼] → [1. 시트 준비] → [2. 서비스 계정 키 넣기]
 */
const PROJECT_ID = 'factorial-e1571';
const SHEET_NAME = '주문';
const STEPS = {
  order:  ['주문 접수', '입금 확인', 'PDF 발송'],
  custom: ['신청 접수', '입금 안내', '입금 확인', '제작 중', 'PDF 발송']
};
const CANCEL = '취소';
const HEAD = ['접수 시각', '상태', '주문번호', '종류', '상품', '금액', '입금자명', '받을 이메일',
              '학교·학년', '과목', '범위', '기한', '구성·양식', '로그인 계정', '문서ID'];
const C = {}; HEAD.forEach((h, i) => C[h] = i + 1);

/* ---------- 메뉴 ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('팩토리얼')
    .addItem('1. 시트 준비', 'setup')
    .addItem('2. 서비스 계정 키 넣기', 'saveKey')
    .addItem('연결 확인', 'testConnection')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME, 0);
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#e8f0fc');
  sh.setFrozenRows(1);
  sh.setColumnWidth(C['상품'], 260);
  sh.hideColumns(C['문서ID']);
  ss.setSpreadsheetTimeZone('Asia/Seoul');
  // 상태를 바꾸면 사이트에 반영하는 트리거 (중복 설치 방지)
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'onStatusEdit').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onStatusEdit').forSpreadsheet(ss).onEdit().create();
  SpreadsheetApp.getUi().alert('시트 준비 끝. 이제 [2. 서비스 계정 키 넣기]를 눌러 주세요.');
}

function saveKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('서비스 계정 키', 'Firebase에서 받은 JSON 파일 내용을 통째로 붙여넣어 주세요.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  let j;
  try { j = JSON.parse(r.getResponseText()); } catch (e) { ui.alert('JSON 형식이 아니에요. 파일 내용을 처음 { 부터 마지막 } 까지 전부 복사해 주세요.'); return; }
  if (!j.client_email || !j.private_key) { ui.alert('client_email / private_key 가 없어요. Firebase 서비스 계정 키가 맞는지 확인해 주세요.'); return; }
  PropertiesService.getScriptProperties().setProperties({SA_EMAIL: j.client_email, SA_KEY: j.private_key});
  CacheService.getScriptCache().remove('fs_token');
  testConnection();
}

function testConnection() {
  try {
    const res = UrlFetchApp.fetch(fsUrl('orders') + '?pageSize=1', {headers: {Authorization: 'Bearer ' + token()}, muteHttpExceptions: true});
    SpreadsheetApp.getUi().alert(res.getResponseCode() === 200 ? '연결됐어요. 이제 [배포]만 하면 끝이에요.' : '연결 실패 (' + res.getResponseCode() + ')\n' + res.getContentText().slice(0, 300));
  } catch (e) { SpreadsheetApp.getUi().alert('연결 실패: ' + e.message); }
}

/* ---------- 사이트 → 시트 ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const id = String(body.id || '');
    if (!/^[A-Za-z0-9]{10,40}$/.test(id)) return out('bad id');
    const o = fsGet(id);                 // 사이트가 보낸 내용은 믿지 않고, 실제 주문을 다시 읽음
    if (!o) return out('not found');
    const sh = sheet(); const row = findRow(sh, id);
    if (body.type === 'cancel') {
      if (row && o.cancelled) { sh.getRange(row, C['상태']).setValue(CANCEL); mail('취소', o); }
      return out('ok');
    }
    if (row) return out('dup');
    addRow(sh, id, o); mail(o.kind === 'custom' ? '새 THEFACTORIAL 신청' : '새 주문', o);
    return out('ok');
  } catch (err) { console.error(err); return out('error'); }
  finally { lock.releaseLock(); }
}
function out(t) { return ContentService.createTextOutput(t); }

function addRow(sh, id, o) {
  const a = o.answers || {};
  const goods = o.kind === 'order' ? (o.items || []).map(i => i.name).join(', ') : o.title;
  const vals = new Array(HEAD.length).fill('');
  vals[C['접수 시각'] - 1] = o.createdAt ? new Date(o.createdAt) : new Date();
  vals[C['상태'] - 1] = o.cancelled ? CANCEL : STEPS[o.kind][o.stage || 0];
  vals[C['주문번호'] - 1] = o.no || '';
  vals[C['종류'] - 1] = o.kind === 'custom' ? 'THEFACTORIAL' : '답지';
  vals[C['상품'] - 1] = goods || '';
  vals[C['금액'] - 1] = o.total || 0;
  vals[C['입금자명'] - 1] = o.depositor || '';
  vals[C['받을 이메일'] - 1] = o.email || '';
  vals[C['학교·학년'] - 1] = [a.school, a.grade].filter(Boolean).join(' ');
  vals[C['과목'] - 1] = a.subject || '';
  vals[C['범위'] - 1] = a.range || '';
  vals[C['기한'] - 1] = a.due || '';
  vals[C['구성·양식'] - 1] = [a.plan, a.files, (o.fileNames || []).join(' / ')].filter(Boolean).join(' · ');
  vals[C['로그인 계정'] - 1] = o.account || '';
  vals[C['문서ID'] - 1] = id;
  sh.appendRow(vals);
  const r = sh.getLastRow();
  sh.getRange(r, C['접수 시각']).setNumberFormat('MM/dd HH:mm');
  sh.getRange(r, C['금액']).setNumberFormat('#,##0"원"');
  sh.getRange(r, C['상태']).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STEPS[o.kind].concat(CANCEL), true).setAllowInvalid(false).build());
}

function mail(what, o) {
  const me = Session.getEffectiveUser().getEmail(); if (!me) return;
  const who = o.depositor ? ' · 입금자 ' + o.depositor : '';
  MailApp.sendEmail(me, '[팩토리얼] ' + what + who + ' · ' + Number(o.total || 0).toLocaleString() + '원',
    [what, '주문번호: ' + (o.no || ''), '상품: ' + (o.title || ''), '금액: ' + Number(o.total || 0).toLocaleString() + '원',
     o.depositor ? '입금자명: ' + o.depositor : '', '받을 이메일: ' + (o.email || ''), '', '시트: ' + SpreadsheetApp.getActive().getUrl()]
      .filter(s => s !== '').join('\n'));
}

/* ---------- 시트 → 사이트 ---------- */
function onStatusEdit(e) {
  const r = e.range, sh = r.getSheet();
  if (sh.getName() !== SHEET_NAME || r.getColumn() !== C['상태'] || r.getRow() < 2 || r.getNumRows() > 1) return;
  const id = sh.getRange(r.getRow(), C['문서ID']).getValue();
  const kind = sh.getRange(r.getRow(), C['종류']).getValue() === 'THEFACTORIAL' ? 'custom' : 'order';
  const v = String(e.value || '');
  try {
    if (v === CANCEL) fsPatch(id, {cancelled: {booleanValue: true}});
    else {
      const i = STEPS[kind].indexOf(v); if (i < 0) return;
      fsPatch(id, {stage: {integerValue: String(i)}, cancelled: {booleanValue: false}});
    }
    r.setNote(null).setBackground(null);
  } catch (err) {
    r.setNote('사이트 반영 실패: ' + err.message).setBackground('#fde8e6');
  }
}

/* ---------- Firestore ---------- */
function fsUrl(path) { return 'https://firestore.googleapis.com/v1/projects/' + PROJECT_ID + '/databases/(default)/documents/' + path; }
function fsGet(id) {
  const res = UrlFetchApp.fetch(fsUrl('orders/' + id), {headers: {Authorization: 'Bearer ' + token()}, muteHttpExceptions: true});
  if (res.getResponseCode() !== 200) return null;
  return val({mapValue: JSON.parse(res.getContentText())});
}
function fsPatch(id, fields) {
  const mask = Object.keys(fields).map(k => 'updateMask.fieldPaths=' + k).join('&');
  const res = UrlFetchApp.fetch(fsUrl('orders/' + id) + '?' + mask + '&currentDocument.exists=true', {
    method: 'patch', contentType: 'application/json', payload: JSON.stringify({fields}),
    headers: {Authorization: 'Bearer ' + token()}, muteHttpExceptions: true});
  if (res.getResponseCode() !== 200) throw new Error(res.getResponseCode() + ' ' + res.getContentText().slice(0, 120));
}
function val(v) {
  if ('stringValue' in v) return v.stringValue;
  if ('integerValue' in v) return Number(v.integerValue);
  if ('doubleValue' in v) return v.doubleValue;
  if ('booleanValue' in v) return v.booleanValue;
  if ('timestampValue' in v) return v.timestampValue;
  if ('nullValue' in v) return null;
  if ('arrayValue' in v) return (v.arrayValue.values || []).map(val);
  if ('mapValue' in v) { const o = {}; const f = v.mapValue.fields || {}; Object.keys(f).forEach(k => o[k] = val(f[k])); return o; }
  return null;
}
function token() {
  const cache = CacheService.getScriptCache(); const hit = cache.get('fs_token'); if (hit) return hit;
  const p = PropertiesService.getScriptProperties();
  const email = p.getProperty('SA_EMAIL'), key = p.getProperty('SA_KEY');
  if (!email || !key) throw new Error('서비스 계정 키가 아직 없어요. [팩토리얼] 메뉴에서 넣어 주세요.');
  const now = Math.floor(Date.now() / 1000);
  const b64 = s => Utilities.base64EncodeWebSafe(s).replace(/=+$/, '');
  const head = b64(JSON.stringify({alg: 'RS256', typ: 'JWT'}));
  const claim = b64(JSON.stringify({iss: email, scope: 'https://www.googleapis.com/auth/datastore',
    aud: 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600}));
  const sig = b64(Utilities.computeRsaSha256Signature(head + '.' + claim, key));
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {method: 'post',
    payload: {grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: head + '.' + claim + '.' + sig}});
  const t = JSON.parse(res.getContentText()).access_token;
  cache.put('fs_token', t, 3000);
  return t;
}

function sheet() { return SpreadsheetApp.getActive().getSheetByName(SHEET_NAME); }
function findRow(sh, id) {
  const hit = sh.getRange(1, C['문서ID'], Math.max(sh.getLastRow(), 1), 1).createTextFinder(id).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}
