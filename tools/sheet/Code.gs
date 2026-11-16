/**
 * 팩토리얼 모의고사 주문 시트 + 자동 발송
 * - 사이트에서 주문/신청/취소가 생기면 이 시트에 한 줄씩 쌓이고 메일 알림이 갑니다.
 * - 사이트 관리자 페이지(thefactorial.web.app/?admin)에서 ✓를 누르면 PDF가 자동 발송되고, 이 시트도 자동으로 바뀝니다.
 * - 시트에서 "상태" 칸을 직접 바꿔도 똑같이 동작합니다.
 * - 답지 PDF는 '파일' 탭에 상품별 드라이브 링크를 넣어 두면 됩니다.
 * 처음 한 번: [팩토리얼] → [1. 시트 준비] → [2. 서비스 계정 키 넣기] → [3. 관리자 등록]
 */
const PROJECT_ID = 'factorial-e1571';
const SHEET_NAME = '주문';
const FILE_SHEET = '파일';
const SHOP_NAME = '팩토리얼 모의고사';
const BANK = 'KB국민은행 842401-01-895543 (예금주 이주현(에스피오))';
const STEPS = {
  order:  ['주문 접수', '입금 확인', 'PDF 발송'],
  custom: ['신청 접수', '입금 안내', '입금 확인', '제작 중', 'PDF 발송']
};
const CANCEL = '취소';
const HEAD = ['접수 시각', '상태', '주문번호', '종류', '상품', '금액', '입금자명', '받을 이메일',
              '학교·학년', '과목', '범위', '기한', '구성·양식', '로그인 계정', '문서ID', '휴대폰', '보낸 이름', '보낼 파일', '발송 시각'];
const FILE_HEAD = ['상품 코드', '상품 이름', '파일 링크 (드라이브 파일 또는 폴더, 여러 개면 줄바꿈)'];
const MAX_BYTES = 20 * 1024 * 1024;   // 메일 첨부 한도(25MB)보다 여유 있게
const C = {}; HEAD.forEach((h, i) => C[h] = i + 1);

/* ---------- 메뉴 ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('팩토리얼')
    .addItem('1. 시트 준비', 'setup')
    .addItem('2. 서비스 계정 키 넣기', 'saveKey')
    .addItem('3. 관리자 등록', 'addAdmin')
    .addItem('연결 확인', 'testConnection')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME, 0);
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#e8f0fc');
  sh.setFrozenRows(1);
  sh.setColumnWidth(C['상품'], 260);
  sh.setColumnWidth(C['보낼 파일'], 220);
  sh.hideColumns(C['문서ID']);
  const fs = ss.getSheetByName(FILE_SHEET) || ss.insertSheet(FILE_SHEET);
  fs.getRange(1, 1, 1, FILE_HEAD.length).setValues([FILE_HEAD]).setFontWeight('bold').setBackground('#e8f0fc');
  fs.setFrozenRows(1); fs.setColumnWidth(2, 260); fs.setColumnWidth(3, 420);
  DriveApp.getRootFolder();   // 드라이브 권한을 미리 받아 둠
  ss.setSpreadsheetTimeZone('Asia/Seoul');
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'onStatusEdit').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onStatusEdit').forSpreadsheet(ss).onEdit().create();
  SpreadsheetApp.getUi().alert('시트 준비 끝.');
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

function addAdmin() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('관리자 등록', '사이트 관리자 페이지(thefactorial.web.app/?admin)에 나오는 "내 관리자 코드"를 붙여넣어 주세요.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const uid = r.getResponseText().trim();
  if (!/^[A-Za-z0-9]{10,40}$/.test(uid)) { ui.alert('코드 형식이 아니에요.'); return; }
  try { fsSet('admins/' + uid, {admin: {booleanValue: true}}, false); ui.alert('등록됐어요. 사이트 관리자 페이지를 새로고침하세요.'); }
  catch (e) { ui.alert('등록 실패: ' + e.message); }
}

function testConnection() {
  try {
    const res = UrlFetchApp.fetch(fsUrl('orders') + '?pageSize=1', {headers: {Authorization: 'Bearer ' + token()}, muteHttpExceptions: true});
    SpreadsheetApp.getUi().alert(res.getResponseCode() === 200 ? '연결됐어요.' : '연결 실패 (' + res.getResponseCode() + ')\n' + res.getContentText().slice(0, 300));
  } catch (e) { SpreadsheetApp.getUi().alert('연결 실패: ' + e.message); }
}

/* ---------- 사이트 → 시트 ----------
 * 사이트는 주문 번호만 알려 줌. 여기서 데이터베이스를 직접 다시 읽고, 그 상태에 맞을 때만 움직임.
 * (단계 변경은 데이터베이스 규칙상 관리자만 할 수 있으므로, "발송"도 관리자가 단계를 바꾼 주문에만 일어남) */
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const id = String(body.id || '');
    if (!/^[A-Za-z0-9]{10,40}$/.test(id)) return out('bad id');
    const o = fsGet(id);
    if (!o) return out('not found');
    const sh = sheet(); let row = findRow(sh, id);
    const ensureRow = () => { if (!row) { addRow(sh, id, o); row = findRow(sh, id); } return row; };

    switch (body.type) {
      case 'new':
        if (row) return out('dup');
        addRow(sh, id, o); mail(o.kind === 'custom' ? '새 THEFACTORIAL 신청' : '새 주문', o);
        return out('ok');

      case 'paidAs': {
        const name = String(body.name || '').replace(/[\r\n\t=+@]/g, ' ').trim().slice(0, 20);
        if (!name || o.cancelled) return out('skip');
        sh.getRange(ensureRow(), C['보낸 이름']).setValue(name).setBackground('#fff4c2');
        fsSet('orders/' + id, {paidAs: {stringValue: name}});
        mail('다른 이름으로 입금 · 보낸 이름 ' + name, o);
        return out('ok');
      }
      case 'cancel':
        if (o.cancelled) { setStatus(sh, ensureRow(), CANCEL); mail('취소', o); }
        return out('ok');

      case 'stage':                        // 관리자가 사이트에서 단계를 바꿈 → 시트만 맞춤
        setStatus(sh, ensureRow(), o.cancelled ? CANCEL : STEPS[o.kind][o.stage || 0]);
        return out('ok');

      case 'confirm':                      // 답지: 관리자가 입금 확인 → 자동 발송
        if (o.kind !== 'order' || o.cancelled || (o.stage || 0) < 1 || o.sentAt) return out('skip');
        sendAndMark(sh, ensureRow(), id, o, null);
        return out('ok');

      case 'send':                         // THEFACTORIAL: 관리자가 파일 링크를 넣고 발송
        if (o.kind !== 'custom' || o.cancelled || (o.stage || 0) < 2 || o.sentAt || !o.sendLink) return out('skip');
        sh.getRange(ensureRow(), C['보낼 파일']).setValue(o.sendLink);
        sendAndMark(sh, row, id, o, o.sendLink);
        return out('ok');

      case 'notice':                       // THEFACTORIAL: 입금 안내 메일
        if (o.kind !== 'custom' || o.cancelled || (o.stage || 0) !== 1 || o.noticeAt) return out('skip');
        sendNotice(o);
        fsSet('orders/' + id, {noticeAt: {timestampValue: new Date().toISOString()}});
        setStatus(sh, ensureRow(), STEPS.custom[1]);
        return out('ok');
    }
    return out('unknown');
  } catch (err) { console.error(err); return out('error'); }
  finally { lock.releaseLock(); }
}
function out(t) { return ContentService.createTextOutput(t); }

function sendAndMark(sh, row, id, o, links) {
  try {
    deliver(o, links || rowLinks(sh, row));
    const last = STEPS[o.kind].length - 1;
    fsSet('orders/' + id, {stage: {integerValue: String(last)}, sentAt: {timestampValue: new Date().toISOString()}, sendError: {stringValue: ''}});
    setStatus(sh, row, STEPS[o.kind][last]);
    sh.getRange(row, C['발송 시각']).setValue(new Date()).setNumberFormat('MM/dd HH:mm');
    sh.getRange(row, C['상태']).setNote(null).setBackground(null);
  } catch (err) {
    fsSet('orders/' + id, {sendError: {stringValue: err.message}});
    sh.getRange(row, C['상태']).setNote('발송 실패: ' + err.message).setBackground('#fde8e6');
    mail('발송 실패 · ' + err.message, o);
  }
}

/* ---------- 시트에서 상태를 직접 바꿨을 때 ---------- */
function onStatusEdit(e) {
  const r = e.range, sh = r.getSheet();
  if (sh.getName() !== SHEET_NAME || r.getColumn() !== C['상태'] || r.getRow() < 2 || r.getNumRows() > 1) return;
  const row = r.getRow();
  const id = String(sh.getRange(row, C['문서ID']).getValue());
  const kind = sh.getRange(row, C['종류']).getValue() === 'THEFACTORIAL' ? 'custom' : 'order';
  const v = String(e.value || '');
  const last = STEPS[kind].length - 1;
  try {
    if (v === CANCEL) { fsSet('orders/' + id, {cancelled: {booleanValue: true}}); r.setNote(null).setBackground(null); return; }
    const i = STEPS[kind].indexOf(v); if (i < 0) return;
    const o = fsGet(id); if (!o) throw new Error('주문을 찾을 수 없어요');
    if (i === last) { if (e.oldValue) r.setValue(e.oldValue); sendAndMark(sh, row, id, o, null); return; }
    fsSet('orders/' + id, {stage: {integerValue: String(i)}, cancelled: {booleanValue: false}});
    r.setNote(null).setBackground(null);
    if (kind === 'order' && i === 1 && !o.sentAt) sendAndMark(sh, row, id, o, null);
    if (kind === 'custom' && i === 1 && !o.noticeAt) { sendNotice(o); fsSet('orders/' + id, {noticeAt: {timestampValue: new Date().toISOString()}}); }
  } catch (err) {
    r.setNote('실패: ' + err.message).setBackground('#fde8e6');
  }
}

/* ---------- 시트 줄 ---------- */
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
  vals[C['휴대폰'] - 1] = o.phone ? "'" + o.phone : '';
  vals[C['보낸 이름'] - 1] = o.paidAs || '';
  sh.appendRow(vals);
  if (o.kind === 'order') ensureFileRows(o.items || []);
  const r = sh.getLastRow();
  sh.getRange(r, C['접수 시각']).setNumberFormat('MM/dd HH:mm');
  sh.getRange(r, C['금액']).setNumberFormat('#,##0"원"');
  sh.getRange(r, C['상태']).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STEPS[o.kind].concat(CANCEL), true).setAllowInvalid(false).build());
}
function setStatus(sh, row, v) { if (row) sh.getRange(row, C['상태']).setValue(v); }   // 코드로 바꾸면 onEdit이 다시 돌지 않음
function rowLinks(sh, row) { return row ? String(sh.getRange(row, C['보낼 파일']).getValue() || '').trim() : ''; }

/* ---------- 메일 ---------- */
function mail(what, o) {
  const me = Session.getEffectiveUser().getEmail(); if (!me) return;
  const who = o.depositor ? ' · 입금자 ' + o.depositor : '';
  MailApp.sendEmail(me, '[팩토리얼] ' + what + who + ' · ' + Number(o.total || 0).toLocaleString() + '원',
    [what, '주문번호: ' + (o.no || ''), '상품: ' + (o.title || ''), '금액: ' + Number(o.total || 0).toLocaleString() + '원',
     o.depositor ? '입금자명: ' + o.depositor : '', '받을 이메일: ' + (o.email || ''), '',
     '관리자 페이지: https://thefactorial.web.app/?admin', '시트: ' + SpreadsheetApp.getActive().getUrl()]
      .filter(s => s !== '').join('\n'));
}
function sendNotice(o) {
  if (!o.email) throw new Error('받을 이메일이 없어요');
  MailApp.sendEmail({to: o.email, name: SHOP_NAME,
    subject: '[' + SHOP_NAME + '] THEFACTORIAL 입금 안내 (' + (o.no || '') + ')',
    body: ['안녕하세요, ' + SHOP_NAME + '입니다.', '', '신청하신 내용을 확인했어요. 아래 계좌로 입금해 주시면 바로 제작을 시작해요.', '',
           '· ' + (o.title || ''), '· 입금할 금액: ' + Number(o.total || 0).toLocaleString() + '원', '· 계좌: ' + BANK,
           '· 입금자명: 이름 뒤에 신청번호 ' + (o.no || '') + ' 를 같이 적어 주세요.', '', '문의는 이 메일에 답장해 주세요.'].join('\n')});
}

/* ---------- PDF 발송 ---------- */
function deliver(o, links) {
  if (o.cancelled) throw new Error('취소된 주문이에요');
  if (!o.email) throw new Error('받을 이메일이 없어요');
  const missing = [];
  if (!links) {
    if (o.kind !== 'order') throw new Error('보낼 파일 링크가 없어요');
    const map = fileMap();
    links = (o.items || []).map(it => { const l = map[itemKey(it)]; if (!l) missing.push(it.name || itemKey(it)); return l || ''; }).join('\n');
  }
  if (missing.length) { ensureFileRows(o.items || []); throw new Error("'파일' 탭에 링크가 없어요: " + missing.join(', ')); }
  const blobs = collectFiles(links);
  if (!blobs.length) throw new Error('링크에서 파일을 찾지 못했어요');
  const bytes = blobs.reduce((a, b) => a + b.getBytes().length, 0);
  if (bytes > MAX_BYTES) throw new Error('파일이 너무 커요 (' + Math.round(bytes / 1048576) + 'MB, 20MB까지)');
  const list = o.kind === 'order' ? (o.items || []).map(it => '· ' + it.name).join('\n') : '· ' + o.title;
  MailApp.sendEmail({
    to: o.email, name: SHOP_NAME,
    subject: '[' + SHOP_NAME + '] 주문하신 PDF를 보내드려요 (' + (o.no || '') + ')',
    body: ['안녕하세요, ' + SHOP_NAME + '입니다.', '', '입금이 확인되어 주문하신 파일을 첨부해 보내드려요.', '', list, '',
           '주문번호: ' + (o.no || ''), '파일은 본인만 사용해 주세요. 다른 사람에게 공유하거나 올리면 안 돼요.', '', '문의는 이 메일에 답장해 주세요.'].join('\n'),
    attachments: blobs
  });
}
function itemKey(it) { return it.pid ? it.pid : 'bundle-' + it.bundle; }
function fileMap() {
  const fs = SpreadsheetApp.getActive().getSheetByName(FILE_SHEET); const m = {};
  if (!fs || fs.getLastRow() < 2) return m;
  fs.getRange(2, 1, fs.getLastRow() - 1, 3).getValues().forEach(r => { if (r[0] && String(r[2]).trim()) m[String(r[0])] = String(r[2]).trim(); });
  return m;
}
function ensureFileRows(items) {
  const fs = SpreadsheetApp.getActive().getSheetByName(FILE_SHEET); if (!fs) return;
  const have = fs.getLastRow() > 1 ? fs.getRange(2, 1, fs.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  items.forEach(it => { const k = itemKey(it); if (!have.includes(k)) { fs.appendRow([k, it.name || '', '']); have.push(k); } });
}
function collectFiles(links) {
  const out = [];
  String(links).split(/[\s,]+/).filter(Boolean).forEach(l => {
    const folder = l.match(/folders\/([\w-]{10,})/);
    const file = l.match(/\/d\/([\w-]{10,})/) || l.match(/[?&]id=([\w-]{10,})/) || (/^[\w-]{20,}$/.test(l) ? [0, l] : null);
    if (folder) {
      const it = DriveApp.getFolderById(folder[1]).getFiles(); const fs = [];
      while (it.hasNext()) fs.push(it.next());
      fs.sort((a, b) => a.getName().localeCompare(b.getName())).forEach(f => out.push(f.getBlob()));
    } else if (file) out.push(DriveApp.getFileById(file[1]).getBlob());
  });
  return out;
}

/* ---------- Firestore ---------- */
function fsUrl(path) { return 'https://firestore.googleapis.com/v1/projects/' + PROJECT_ID + '/databases/(default)/documents/' + path; }
function fsGet(id) {
  const res = UrlFetchApp.fetch(fsUrl('orders/' + id), {headers: {Authorization: 'Bearer ' + token()}, muteHttpExceptions: true});
  if (res.getResponseCode() !== 200) return null;
  return val({mapValue: JSON.parse(res.getContentText())});
}
function fsSet(path, fields, mustExist) {
  const mask = Object.keys(fields).map(k => 'updateMask.fieldPaths=' + k).join('&');
  const res = UrlFetchApp.fetch(fsUrl(path) + '?' + mask + (mustExist === false ? '' : '&currentDocument.exists=true'), {
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
