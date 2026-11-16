/**
 * 팩토리얼 모의고사 주문 시트 + 자동 발송
 * - 사이트에서 주문/신청/취소가 생기면 이 시트에 한 줄씩 쌓이고 메일 알림이 갑니다.
 * - 사이트 관리자 페이지(thefactorial.web.app/?admin)에서 ✓를 누르면 PDF가 자동 발송되고, 이 시트도 자동으로 바뀝니다.
 * - 시트에서 "상태" 칸을 직접 바꿔도 똑같이 동작합니다.
 * - 답지 PDF는 '파일' 탭에 상품별 드라이브 링크를 넣어 두면 됩니다.
 * - 텔레그램을 연결하면 알림이 메일 대신 텔레그램으로 오고, 텔레그램 버튼(✓/✗)으로 바로 입금 확인·발송할 수 있습니다.
 * 처음 한 번: [팩토리얼] → [1. 시트 준비] → [2. 서비스 계정 키 넣기] → [3. 관리자 등록] → [4. 텔레그램 연결]
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
    .addItem('4. 텔레그램 연결', 'connectTelegram')
    .addItem('연결 확인', 'testConnection')
    .addItem('워터마크 미리보기', 'testWatermark')
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
  // 실제 처리는 비동기(워터마크) → 이 함수가 끝난 뒤에도 같은 실행 안에서 마저 돌아감
  handlePost(e).catch(err => console.error(err));
  return out('ok');
}
async function handlePost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const id = String(body.id || '');
    if (!/^[A-Za-z0-9]{10,40}$/.test(id)) return out('bad id');
    if (body.type === 'invite') return out(grantInvite(id));
    const o = fsGet(id);
    if (!o) return out('not found');
    const sh = sheet(); let row = findRow(sh, id);
    const ensureRow = () => { if (!row) { addRow(sh, id, o); row = findRow(sh, id); } return row; };

    switch (body.type) {
      case 'new':
        if (row) return out('dup');
        checkPerks(id, o);                 // 무료 답지·친구 쿠폰이 진짜인지 확인하고 차감
        addRow(sh, id, o); row = findRow(sh, id);
        if (o.kind === 'order' && o.freeUse && !o.perkError && (o.total || 0) === 0) {   // 전액 무료 답지 → 바로 발송
          fsSet('orders/' + id, {stage: {integerValue: '1'}}); o.stage = 1;
          try { await sendAndMark(sh, row, id, o, null); mail('무료 답지 주문 · 자동 발송함', o, null); } catch (e) {}
          return out('ok');
        }
        mail((o.kind === 'custom' ? '새 THEFACTORIAL 신청' : '새 주문') + (o.perkError ? ' · ⚠ ' + o.perkError : ''), o, id);
        return out('ok');

      case 'paidAs': {
        const name = String(body.name || '').replace(/[\r\n\t=+@]/g, ' ').trim().slice(0, 20);
        if (!name || o.cancelled) return out('skip');
        sh.getRange(ensureRow(), C['보낸 이름']).setValue(name).setBackground('#fff4c2');
        fsSet('orders/' + id, {paidAs: {stringValue: name}});
        o.paidAs = name; mail('다른 이름으로 입금 · 보낸 이름 ' + name, o, id);
        return out('ok');
      }
      case 'cancel':
        if (o.cancelled) { refundPerks(id, o); setStatus(sh, ensureRow(), CANCEL); mail('취소됨', o, null); }
        return out('ok');

      case 'stage':                        // 관리자가 사이트에서 단계를 바꿈 → 시트만 맞춤
        setStatus(sh, ensureRow(), o.cancelled ? CANCEL : STEPS[o.kind][o.stage || 0]);
        if (o.kind === 'custom' && (o.stage || 0) >= 2 && !o.cancelled) rewardInviter(o);
        return out('ok');

      case 'confirm':                      // 답지: 관리자가 입금 확인 → 자동 발송
        if (o.kind !== 'order' || o.cancelled || (o.stage || 0) < 1 || o.sentAt) return out('skip');
        await sendAndMark(sh, ensureRow(), id, o, null);
        return out('ok');

      case 'send':                         // THEFACTORIAL: 관리자가 파일 링크를 넣고 발송
        if (o.kind !== 'custom' || o.cancelled || (o.stage || 0) < 2 || o.sentAt || !o.sendLink) return out('skip');
        sh.getRange(ensureRow(), C['보낼 파일']).setValue(o.sendLink);
        await sendAndMark(sh, row, id, o, o.sendLink);
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

async function sendAndMark(sh, row, id, o, links) {
  try {
    await deliver(o, links || rowLinks(sh, row));
    const last = STEPS[o.kind].length - 1;
    fsSet('orders/' + id, {stage: {integerValue: String(last)}, sentAt: {timestampValue: new Date().toISOString()}, sendError: {stringValue: ''}});
    setStatus(sh, row, STEPS[o.kind][last]);
    sh.getRange(row, C['발송 시각']).setValue(new Date()).setNumberFormat('MM/dd HH:mm');
    if ((o.total || 0) > 0) rewardInviter(o);
    sh.getRange(row, C['상태']).setNote(null).setBackground(null);
  } catch (err) {
    fsSet('orders/' + id, {sendError: {stringValue: err.message}});
    sh.getRange(row, C['상태']).setNote('발송 실패: ' + err.message).setBackground('#fde8e6');
    mail('발송 실패 · ' + err.message, o, id);
    throw err;
  }
}

/* ---------- 시트에서 상태를 직접 바꿨을 때 ---------- */
function onStatusEdit(e) { statusEdit(e).catch(err => console.error(err)); }
async function statusEdit(e) {
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
    if (i === last) { if (e.oldValue) r.setValue(e.oldValue); await sendAndMark(sh, row, id, o, null); return; }
    fsSet('orders/' + id, {stage: {integerValue: String(i)}, cancelled: {booleanValue: false}});
    r.setNote(null).setBackground(null);
    if (kind === 'order' && i === 1 && !o.sentAt) await sendAndMark(sh, row, id, o, null);
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
function mail(what, o, id) {
  if (tgReady()) { try { tgNotify(what, o, id); return; } catch (e) { console.error('telegram', e); } }
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

/* ---------- 구매자 워터마크 ----------
 * 보내는 PDF 모든 쪽에 주문번호·이메일을 흐리게 넣음 (유출되면 누구 것인지 보임). pdf-lib 사용.
 * 워터마크가 실패해도 판매는 막지 않음: 원본으로 보내고 관리자에게 알림. */
var setTimeout = setTimeout || function (fn) { fn(); return 0; };   // pdf-lib가 쓰는 함수 (Apps Script에는 없음)
const PDFLIB_URL = 'https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js';
function loadPdfLib() { if (typeof PDFLib === 'undefined') (0, eval)(UrlFetchApp.fetch(PDFLIB_URL).getContentText()); }
const ascii_ = t => String(t || '').replace(/[^\x20-\x7E]/g, '?');
async function watermarkAll(blobs, o) {
  const pdfs = blobs.filter(b => /pdf/i.test(b.getContentType() || '') || /\.pdf$/i.test(b.getName() || ''));
  if (!pdfs.length) return blobs;
  try { loadPdfLib(); } catch (e) { mail('워터마크 실패(원본 발송) · ' + e.message, o, null); return blobs; }
  const out = [];
  for (const b of blobs) {
    if (!pdfs.includes(b)) { out.push(b); continue; }
    try { out.push(await watermarkPdf(b, o)); }
    catch (e) { console.error('watermark', e); mail('워터마크 실패(원본 발송) · ' + e.message, o, null); out.push(b); }
  }
  return out;
}
async function watermarkPdf(blob, o) {
  const {PDFDocument, StandardFonts, rgb, degrees} = PDFLib;
  const doc = await PDFDocument.load(Uint8Array.from(blob.getBytes()), {ignoreEncryption: true});
  const bold = await doc.embedFont(StandardFonts.HelveticaBold);
  const reg = await doc.embedFont(StandardFonts.Helvetica);
  const who = ascii_(o.email), no = ascii_(o.no);
  const mark = no + '  ' + who;
  const foot = 'Licensed to ' + who + ' (' + no + ') - personal use only. Do not share, copy or upload.';
  doc.getPages().forEach(p => {
    const {width, height} = p.getSize();
    const size = Math.max(14, Math.min(width, height) / 24);
    const tw = bold.widthOfTextAtSize(mark, size), a = 30 * Math.PI / 180;
    [0.22, 0.5, 0.78].forEach(f => p.drawText(mark, {
      x: width / 2 - (tw / 2) * Math.cos(a), y: height * f - (tw / 2) * Math.sin(a),
      size, font: bold, color: rgb(0.42, 0.48, 0.6), opacity: 0.11, rotate: degrees(30)}));
    p.drawText(foot, {x: 18, y: 10, size: 6.5, font: reg, color: rgb(0.35, 0.38, 0.45), opacity: 0.85});
  });
  doc.setProducer('Factorial');
  const bytes = await doc.save();
  return Utilities.newBlob(Array.from(bytes, x => x > 127 ? x - 256 : x), 'application/pdf', blob.getName());
}
// 메뉴용: 정치 답지에 테스트 워터마크를 넣어 드라이브에 저장
function testWatermark() {
  const say = m => { try { SpreadsheetApp.getUi().alert(m); } catch (_) { Logger.log(m); } };
  const f = DriveApp.searchFiles("title contains '답지' and mimeType = 'application/pdf' and trashed = false");
  if (!f.hasNext()) { say('드라이브에서 답지 PDF를 못 찾았어요.'); return; }
  const src = f.next();
  loadPdfLib();
  watermarkPdf(src.getBlob(), {no: 'F1116-TEST', email: 'student@example.com'})
    .then(b => { const nf = src.getParents().next().createFile(b.setName('[워터마크 미리보기] ' + src.getName())); say('드라이브에 미리보기를 저장했어요:\n' + nf.getUrl()); })
    .catch(e => say('실패: ' + e.message));
}

/* ---------- PDF 발송 ---------- */
async function deliver(o, links) {
  if (o.cancelled) throw new Error('취소된 주문이에요');
  if (!o.email) throw new Error('받을 이메일이 없어요');
  const missing = [];
  if (!links) {
    if (o.kind !== 'order') throw new Error('보낼 파일 링크가 없어요');
    const map = fileMap();
    links = (o.items || []).map(it => { const l = map[itemKey(it)]; if (!l) missing.push(it.name || itemKey(it)); return l || ''; }).join('\n');
  }
  if (missing.length) { ensureFileRows(o.items || []); throw new Error("'파일' 탭에 링크가 없어요: " + missing.join(', ')); }
  const blobs = await watermarkAll(collectFiles(links), o);
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


/* ---------- 텔레그램 ----------
 * 봇이 새 주문을 보내 주고, 버튼을 누르면 1분 안에 처리됩니다(1분마다 확인).
 * 공주님 채팅에서 누른 버튼만 받습니다. */
function tgProps() { return PropertiesService.getScriptProperties(); }
function tgReady() { const p = tgProps(); return !!(p.getProperty('TG_TOKEN') && p.getProperty('TG_CHAT')); }
function tg(method, payload) {
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + tgProps().getProperty('TG_TOKEN') + '/' + method, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload || {}), muteHttpExceptions: true});
  const j = JSON.parse(res.getContentText() || '{}');
  if (!j.ok) throw new Error('텔레그램 ' + method + ': ' + (j.description || res.getResponseCode()));
  return j.result;
}

function connectTelegram() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('텔레그램 연결 (1/2)', '@BotFather 에서 받은 봇 토큰을 붙여넣어 주세요. (예: 123456789:AA...)', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const token = r.getResponseText().trim();
  if (!/^\d+:[\w-]{20,}$/.test(token)) { ui.alert('토큰 형식이 아니에요.'); return; }
  tgProps().setProperty('TG_TOKEN', token); tgProps().deleteProperty('TG_CHAT'); tgProps().deleteProperty('TG_OFFSET');
  let me; try { me = tg('getMe'); tg('deleteWebhook', {}); } catch (e) { ui.alert('토큰이 맞지 않아요: ' + e.message); return; }
  ui.alert('텔레그램 연결 (2/2)', '텔레그램에서 @' + me.username + ' 봇을 열고 [시작] 또는 /start 를 보낸 뒤, 이 창의 [확인]을 눌러 주세요.', ui.ButtonSet.OK);
  const ups = tg('getUpdates', {timeout: 0, allowed_updates: ['message']});
  const m = ups.reverse().find(u => u.message && u.message.chat && u.message.chat.type === 'private');
  if (!m) { ui.alert('봇에서 메시지를 못 찾았어요. /start 를 보낸 뒤 [4. 텔레그램 연결]을 다시 눌러 주세요.'); return; }
  tgProps().setProperty('TG_CHAT', String(m.message.chat.id));
  tgProps().setProperty('TG_OFFSET', String(ups[0].update_id + 1));
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'pollTelegram').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('pollTelegram').timeBased().everyMinutes(1).create();
  tg('sendMessage', {chat_id: m.message.chat.id, text: '연결됐어요. 이제 새 주문이 여기로 와요.\n✓ 버튼을 누르면 1분 안에 입금 확인과 PDF 발송까지 끝나요.'});
  ui.alert('연결됐어요. 텔레그램에 테스트 메시지를 보냈어요.');
}

const won_ = n => Number(n || 0).toLocaleString() + '원';
function tgText(what, o) {
  const lines = ['<b>' + esc_(what) + '</b>', ''];
  if (o.kind === 'custom') {
    const a = o.answers || {};
    lines.push('🎓 ' + esc_([a.school, a.grade].filter(Boolean).join(' ')), '📄 ' + esc_(o.title || ''), '📅 기한: ' + esc_(a.due || '-'));
  } else {
    lines.push('👤 입금자명: <b>' + esc_(o.depositor || '-') + '</b>', '📄 ' + esc_((o.items || []).map(i => i.name).join(', ')));
  }
  lines.push('💰 <b>' + won_(o.total) + '</b>');
  if (o.paidAs) lines.push('🟨 학생이 알린 보낸 이름: <b>' + esc_(o.paidAs) + '</b>');
  lines.push('🔖 ' + esc_(o.no || '') + ' · ' + esc_(o.email || ''));
  return lines.join('\n');
}
function esc_(t) { return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function tgButtons(o, id) {
  if (!id || o.cancelled) return undefined;
  const last = STEPS[o.kind].length - 1, st = o.stage || 0;
  if (st >= last) return undefined;
  const x = {text: '✗ 취소', callback_data: 'x:' + id};
  if (o.kind === 'order') return {inline_keyboard: [[{text: st === 0 ? '✓ 입금 확인 → PDF 발송' : '↻ 다시 보내기', callback_data: 'c:' + id}], [x]]};
  if (st === 0) return {inline_keyboard: [[{text: '✉ 입금 안내 메일 보내기', callback_data: 'n:' + id}], [x]]};
  if (st === 1) return {inline_keyboard: [[{text: '✓ 입금 확인', callback_data: 'p:' + id}], [x]]};
  return undefined;   // 제작·발송은 관리자 페이지에서 (완성 PDF 링크 필요)
}
function tgNotify(what, o, id) {
  tg('sendMessage', {chat_id: tgProps().getProperty('TG_CHAT'), text: tgText(what, o), parse_mode: 'HTML', reply_markup: tgButtons(o, id)});
}

function pollTelegram() { pollTelegramAsync().catch(err => console.error(err)); }
async function pollTelegramAsync() {
  if (!tgReady()) return;
  const lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) return;
  try {
    const p = tgProps(), chat = p.getProperty('TG_CHAT');
    const ups = tg('getUpdates', {offset: Number(p.getProperty('TG_OFFSET') || 0), timeout: 0, allowed_updates: ['callback_query']});
    for (const u of ups) {
      p.setProperty('TG_OFFSET', String(u.update_id + 1));
      const q = u.callback_query; if (!q) continue;
      if (!q.message || String(q.message.chat.id) !== chat) { tg('answerCallbackQuery', {callback_query_id: q.id, text: '권한 없음'}); continue; }
      try { await tgAction(q); } catch (e) { tg('answerCallbackQuery', {callback_query_id: q.id, text: '실패: ' + e.message, show_alert: true}); }
    }
  } finally { lock.releaseLock(); }
}

async function tgAction(q) {
  const [act, id] = String(q.data || '').split(':');
  const msg = {chat_id: q.message.chat.id, message_id: q.message.message_id, parse_mode: 'HTML'};
  const o = fsGet(id); if (!o) throw new Error('주문을 찾을 수 없어요');
  const edit = (head, kb) => tg('editMessageText', Object.assign({}, msg, {text: tgText(head, o), reply_markup: kb}));
  const ask = (question, yes) => edit(question, {inline_keyboard: [[{text: '네', callback_data: yes + ':' + id}, {text: '아니요', callback_data: 'b:' + id}]]});
  const sh = sheet(); let row = findRow(sh, id); if (!row) { addRow(sh, id, o); row = findRow(sh, id); }
  if (o.cancelled && act !== 'b') { edit('이미 취소된 주문', undefined); return tg('answerCallbackQuery', {callback_query_id: q.id}); }

  switch (act) {
    case 'b': edit(o.kind === 'custom' ? 'THEFACTORIAL 신청' : '주문', tgButtons(o, id)); break;
    case 'c': ask("입금자가 '" + (o.depositor || o.no) + "'와 일치하신가요? (네 → 바로 PDF 발송)", 'C'); break;
    case 'C':
      if (o.sentAt) { edit('✅ 이미 발송된 주문', undefined); break; }
      fsSet('orders/' + id, {stage: {integerValue: '1'}, cancelled: {booleanValue: false}});
      o.stage = 1;
      try { await sendAndMark(sh, row, id, o, null); edit('✅ 입금 확인 · PDF 발송 완료 (구매자 워터마크 포함)', undefined); }
      catch (e) { edit('⚠ 발송 실패: ' + e.message, tgButtons(o, id)); }
      break;
    case 'x': ask('이 주문을 취소할까요?', 'X'); break;
    case 'X':
      fsSet('orders/' + id, {cancelled: {booleanValue: true}, cancelledAt: {timestampValue: new Date().toISOString()}});
      refundPerks(id, o); setStatus(sh, row, CANCEL); o.cancelled = true; edit('✗ 취소함', undefined); break;
    case 'n':
      fsSet('orders/' + id, {stage: {integerValue: '1'}});
      o.stage = 1; sendNotice(o); fsSet('orders/' + id, {noticeAt: {timestampValue: new Date().toISOString()}});
      setStatus(sh, row, STEPS.custom[1]); edit('✉ 입금 안내 메일 보냄', tgButtons(o, id)); break;
    case 'p': ask('입금이 확인됐나요?', 'P'); break;
    case 'P':
      fsSet('orders/' + id, {stage: {integerValue: '2'}});
      o.stage = 2; setStatus(sh, row, STEPS.custom[2]); rewardInviter(o);
      edit('✓ 입금 확인 · 완성 PDF는 관리자 페이지 제작 탭에서 보내 주세요', undefined); break;
  }
  tg('answerCallbackQuery', {callback_query_id: q.id});
}


/* ---------- 친구 초대 · 무료 답지 · 쿠폰 ----------
 * 혜택 숫자는 학생이 직접 못 바꾸고(데이터베이스 규칙), 여기서만 늘리고 줄임. */
const KEY_PRICE = 700, BUNDLE_PRICE = 2900, TF_ONE = 15900, TF_THREE = 25900, TF_RUSH = 3000, INVITE_MAX = 5;
function fsDoc(path) {
  const res = UrlFetchApp.fetch(fsUrl(path), {headers: {Authorization: 'Bearer ' + token()}, muteHttpExceptions: true});
  return res.getResponseCode() === 200 ? val({mapValue: JSON.parse(res.getContentText())}) : null;
}
function userByCode(code) {
  const res = UrlFetchApp.fetch('https://firestore.googleapis.com/v1/projects/' + PROJECT_ID + '/databases/(default)/documents:runQuery', {
    method: 'post', contentType: 'application/json', headers: {Authorization: 'Bearer ' + token()}, muteHttpExceptions: true,
    payload: JSON.stringify({structuredQuery: {from: [{collectionId: 'users'}], limit: 1,
      where: {fieldFilter: {field: {fieldPath: 'inviteCode'}, op: 'EQUAL', value: {stringValue: code}}}}})});
  const hit = JSON.parse(res.getContentText() || '[]').find(r => r.document);
  return hit ? hit.document.name.split('/').pop() : null;
}
const int_ = n => ({integerValue: String(Math.max(0, Math.round(n || 0)))});

// 친구가 초대 링크로 가입: 친구에게 무료 답지 1장 + THEFACTORIAL 30% 쿠폰
function grantInvite(uid) {
  const u = fsDoc('users/' + uid); if (!u) return 'no user';
  if (u.inviteGranted || !u.invitedBy || u.invitedBy === u.inviteCode) return 'skip';
  const inviter = userByCode(String(u.invitedBy));
  if (!inviter || inviter === uid) { fsSet('users/' + uid, {inviteGranted: {booleanValue: true}}); return 'bad code'; }
  fsSet('users/' + uid, {freeKeys: int_((u.freeKeys || 0) + 1), tfCoupon: {booleanValue: true}, inviteGranted: {booleanValue: true}, inviterUid: {stringValue: inviter}});
  return 'ok';
}
// 친구가 처음 돈을 내면: 초대한 사람에게 무료 답지 1장 (최대 5명)
function rewardInviter(o) {
  if (!o.uid) return;
  const u = fsDoc('users/' + o.uid); if (!u || !u.inviterUid || u.inviterPaid) return;
  fsSet('users/' + o.uid, {inviterPaid: {booleanValue: true}});
  const inv = fsDoc('users/' + u.inviterUid); if (!inv || (inv.invited || 0) >= INVITE_MAX) return;
  fsSet('users/' + u.inviterUid, {freeKeys: int_((inv.freeKeys || 0) + 1), invited: int_((inv.invited || 0) + 1)});
}
// 새 주문의 무료 답지·쿠폰이 진짜인지 확인하고 차감. 문제가 있으면 perkError를 남김(관리자 확인용)
function checkPerks(id, o) {
  const errs = [];
  if (o.kind === 'order') {
    const items = o.items || [];
    const singles = items.filter(i => i.pid).length;
    const sub = items.reduce((a, i) => a + (i.bundle ? BUNDLE_PRICE : KEY_PRICE), 0);
    const free = Math.max(0, Math.round(o.freeUse || 0));
    if (free) {
      const u = fsDoc('users/' + o.uid) || {};
      if (free > singles || free > (u.freeKeys || 0)) errs.push('무료 답지 사용이 맞지 않아요');
      else { fsSet('users/' + o.uid, {freeKeys: int_((u.freeKeys || 0) - free)}); fsSet('orders/' + id, {freeHeld: int_(free)}); o.freeHeld = free; }
    }
    if ((o.total || 0) !== sub - (errs.length ? 0 : free) * KEY_PRICE) errs.push('금액이 맞지 않아요 (정가 ' + won_(sub - free * KEY_PRICE) + ')');
  } else if (o.kind === 'custom') {
    const a = o.answers || {};
    const base = String(a.plan || '').startsWith('3') ? TF_THREE : TF_ONE;
    const rush = String(a.due || '').startsWith('3일') ? TF_RUSH : 0;
    let expect = base + rush;
    if (o.coupon) {
      const u = fsDoc('users/' + o.uid) || {};
      if (!u.tfCoupon || base !== TF_ONE) errs.push('쿠폰을 쓸 수 없는 계정이에요');
      else { fsSet('users/' + o.uid, {tfCoupon: {booleanValue: false}}); fsSet('orders/' + id, {couponHeld: {booleanValue: true}}); o.couponHeld = true; expect = Math.round(TF_ONE * 0.7 / 10) * 10 + rush; }
    }
    if ((o.total || 0) !== expect) errs.push('금액이 맞지 않아요 (정가 ' + won_(expect) + ')');
  }
  if (errs.length) { o.perkError = errs.join(' · '); fsSet('orders/' + id, {perkError: {stringValue: o.perkError}}); }
}
// 입금 전 취소: 쓴 무료 답지·쿠폰 돌려줌
function refundPerks(id, o) {
  if (o.sentAt || o.perksRefunded || !o.uid) return;
  const u = fsDoc('users/' + o.uid); if (!u) return;
  const f = {};
  if (o.freeHeld) f.freeKeys = int_((u.freeKeys || 0) + o.freeHeld);
  if (o.couponHeld) f.tfCoupon = {booleanValue: true};
  if (Object.keys(f).length) fsSet('users/' + o.uid, f);
  fsSet('orders/' + id, {perksRefunded: {booleanValue: true}});
}
