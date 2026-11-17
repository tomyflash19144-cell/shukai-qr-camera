const APP_TITLE = '集会QR受付システム';
const MEMBERS_SHEET = '会員名簿';
const EVENTS_SHEET = '集会一覧';
const LOG_SHEET = '出席記録';
const STAFF_SHEET = '受付担当者';
const SESSION_TTL = 21600; // 6時間

// 定義された正式なグループ一覧
const VALID_GROUPS = ['東', '西', '大宮西', '北', '見沼/岩槻', 'その他'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, MEMBERS_SHEET, ['会員番号', '氏名', '班・グループ', '状態']);
  ensureSheet_(ss, EVENTS_SHEET, ['集会ID', '集会名', '開催日', '受付状態']);
  ensureSheet_(ss, LOG_SHEET, ['記録日時', '集会ID', '集会名', '会員番号', '氏名', '受付担当']);
  ensureSheet_(ss, STAFF_SHEET, ['担当者ID', '担当者名', 'ログインID', 'PINソルト', 'PINハッシュ', '状態', '登録日時']);
  
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_PIN')) {
    throw new Error('初期設定が未完了です。プロジェクトの設定 > スクリプト プロパティで ADMIN_PIN に6～12桁の数字を設定してから、setupSystem() を再実行してください。');
  }
  return '初期設定・更新が完了しました。既存データは保持されています。';
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9eaf7');
  }
  return sh;
}

// --- 認証・セッション ---

function adminLogin(pin) {
  const saved = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if (!saved || String(pin || '').trim() !== saved) throw new Error('管理者PINが違います。');
  return createSession_({ role: 'admin', name: '管理者' });
}

function changeAdminPin(token, currentPin, newPin) {
  requireSession_(token, ['admin']);
  const props = PropertiesService.getScriptProperties();
  if (String(currentPin || '').trim() !== props.getProperty('ADMIN_PIN')) {
    throw new Error('現在の管理者PINが一致しません。');
  }
  newPin = String(newPin || '').trim();
  if (!/^\d{6,12}$/.test(newPin)) throw new Error('新しい管理者PINは6～12桁の数字にしてください。');
  props.setProperty('ADMIN_PIN', newPin);
  return true;
}

function staffLogin(loginId, pin) {
  const sh = SpreadsheetApp.getActive().getSheetByName(STAFF_SHEET);
  if (!sh || sh.getLastRow() < 2) throw new Error('受付担当者がまだ登録されていません。');
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  const key = String(loginId || '').trim().toUpperCase();
  const row = rows.find(r => String(r[2]).trim().toUpperCase() === key && String(r[5]) === '有効');
  if (!row || !verifyPin_(String(pin || '').trim(), String(row[3]), String(row[4]))) {
    throw new Error('ログインIDまたは暗証番号が違うか、利用停止中です。');
  }
  return createSession_({ role: 'staff', name: String(row[1]), staffId: String(row[0]) });
}

function createSession_(data) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('SESSION_' + token, JSON.stringify(data), SESSION_TTL);
  return { token: token, role: data.role, name: data.name };
}

function requireSession_(token, roles) {
  const raw = CacheService.getScriptCache().get('SESSION_' + String(token || ''));
  if (!raw) throw new Error('SESSION_EXPIRED');
  const session = JSON.parse(raw);
  if (roles && roles.indexOf(session.role) === -1) throw new Error('この操作を行う権限がありません。');
  CacheService.getScriptCache().put('SESSION_' + String(token), raw, SESSION_TTL);
  return session;
}

function logout(token) {
  CacheService.getScriptCache().remove('SESSION_' + String(token || ''));
  return true;
}

// --- 担当者・集会管理 ---

function addStaff(token, name, pin) {
  requireSession_(token, ['admin']);
  name = String(name || '').trim();
  pin = String(pin || '').trim();
  if (!name) throw new Error('担当者名を入力してください。');
  if (!/^\d{6,12}$/.test(pin)) throw new Error('暗証番号は6～12桁の数字にしてください。');
  
  const sh = SpreadsheetApp.getActive().getSheetByName(STAFF_SHEET);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const id = 'S' + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
    const loginId = 'R' + Utilities.getUuid().replace(/-/g, '').slice(0, 7).toUpperCase();
    const salt = Utilities.getUuid();
    sh.appendRow([id, name, loginId, salt, hashPin_(pin, salt), '有効', new Date()]);
    return { name: name, loginId: loginId };
  } finally { lock.releaseLock(); }
}

function listStaff(token) {
  requireSession_(token, ['admin']);
  const sh = SpreadsheetApp.getActive().getSheetByName(STAFF_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues().map(r => ({
    id: String(r[0]), name: String(r[1]), loginId: String(r[2]), status: String(r[5])
  }));
}

function setStaffStatus(token, staffId, status) {
  requireSession_(token, ['admin']);
  const sh = SpreadsheetApp.getActive().getSheetByName(STAFF_SHEET);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  const i = rows.findIndex(r => String(r[0]) === String(staffId));
  if (i < 0) throw new Error('担当者が見つかりません。');
  sh.getRange(i + 2, 6).setValue(status);
  return true;
}

function getEvents(token) {
  requireSession_(token, ['admin', 'staff']);
  const sh = SpreadsheetApp.getActive().getSheetByName(EVENTS_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
    .filter(r => r[0] && String(r[3]) !== '終了')
    .map(r => ({ id: String(r[0]), name: String(r[1]), date: formatDate_(r[2]) }));
}

function addEvent(token, name, date) {
  requireSession_(token, ['admin']);
  name = String(name || '').trim(); date = String(date || '');
  if (!name || !/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('集会名と開催日を入力してください。');
  const sh = SpreadsheetApp.getActive().getSheetByName(EVENTS_SHEET);
  const id = 'E' + Utilities.getUuid().slice(0, 8).toUpperCase();
  sh.appendRow([id, name, new Date(date + 'T00:00:00'), '受付中']);
  return { id: id, name: name, date: date };
}

function closeEvent(token, eventId) {
  requireSession_(token, ['admin']);
  const sh = SpreadsheetApp.getActive().getSheetByName(EVENTS_SHEET);
  if (sh.getLastRow() < 2) return true;
  sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach((r, i) => {
    if (String(r[0]) === String(eventId)) sh.getRange(i + 2, 4).setValue('終了');
  });
  return true;
}

// --- 受付・集計（グループ自動振り分け対応） ---

function normalizeGroup_(g) {
  const val = String(g || '').trim();
  if (VALID_GROUPS.includes(val)) return val;
  // 表記ゆれの吸収
  if (/東/.test(val) && !/大宮西/.test(val)) return '東';
  if (/西/.test(val) && !/大宮西/.test(val)) return '西';
  if (/大宮西/.test(val)) return '大宮西';
  if (/北/.test(val)) return '北';
  if (/見沼|岩槻/.test(val)) return '見沼/岩槻';
  return 'その他';
}

function scanMember(token, eventId, memberId) {
  const session = requireSession_(token, ['admin', 'staff']);
  if (!eventId || !memberId) throw new Error('集会または会員番号が指定されていません。');
  
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActive();
    const members = ss.getSheetByName(MEMBERS_SHEET);
    const events = ss.getSheetByName(EVENTS_SHEET);
    const logs = ss.getSheetByName(LOG_SHEET);
    
    const id = normalizeId_(memberId);
    const memberRows = members.getLastRow() > 1 ? members.getRange(2, 1, members.getLastRow() - 1, 4).getValues() : [];
    const member = memberRows.find(r => normalizeId_(r[0]) === id && String(r[3] || '在籍') !== '退会');
    if (!member) return { ok: false, message: '会員番号「' + id + '」が見つからないか、退会済みです。' };
    
    const eventRows = events.getLastRow() > 1 ? events.getRange(2, 1, events.getLastRow() - 1, 4).getValues() : [];
    const event = eventRows.find(r => String(r[0]) === String(eventId));
    if (!event || String(event[3]) === '終了') throw new Error('この集会は受付中ではありません。');
    
    const logRows = logs.getLastRow() > 1 ? logs.getRange(2, 1, logs.getLastRow() - 1, 6).getValues() : [];
    const duplicate = logRows.find(r => String(r[1]) === String(eventId) && normalizeId_(r[3]) === id);
    if (duplicate) {
      return {
        ok: true, duplicate: true,
        name: String(member[1]),
        group: normalizeGroup_(member[2]),
        memberId: id,
        message: 'すでに受付済みです（' + formatDateTime_(duplicate[0]).slice(11, 16) + ' 登録）'
      };
    }
    
    const now = new Date();
    logs.appendRow([now, String(event[0]), String(event[1]), id, String(member[1]), session.name]);
    return {
      ok: true, duplicate: false,
      name: String(member[1]),
      group: normalizeGroup_(member[2]),
      memberId: id,
      message: '出席を登録しました。'
    };
  } finally { lock.releaseLock(); }
}

function getSummary(token, eventId) {
  requireSession_(token, ['admin', 'staff']);
  const ss = SpreadsheetApp.getActive();
  const ev = ss.getSheetByName(EVENTS_SHEET).getDataRange().getValues().slice(1).find(r => String(r[0]) === String(eventId));
  
  const memSh = ss.getSheetByName(MEMBERS_SHEET);
  const memRows = memSh && memSh.getLastRow() > 1 ? memSh.getRange(2, 1, memSh.getLastRow() - 1, 4).getValues() : [];
  
  const groupMap = {};
  const groupTotal = {};
  
  // 規定の6グループを初期化（0件でもきれいに表示するため）
  VALID_GROUPS.forEach(g => { groupTotal[g] = 0; });

  memRows.forEach(r => {
    if (r[0] && String(r[3] || '在籍') !== '退会') {
      const id = normalizeId_(r[0]);
      const grp = normalizeGroup_(r[2]);
      groupMap[id] = grp;
      groupTotal[grp] = (groupTotal[grp] || 0) + 1;
    }
  });

  const sh = ss.getSheetByName(LOG_SHEET);
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues() : [];
  const selected = rows.filter(r => String(r[1]) === String(eventId));
  
  const groupPresent = {};
  VALID_GROUPS.forEach(g => { groupPresent[g] = 0; });

  selected.forEach(r => {
    const grp = groupMap[normalizeId_(r[3])] || 'その他';
    groupPresent[grp] = (groupPresent[grp] || 0) + 1;
  });

  const groups = VALID_GROUPS.map(g => ({
    name: g,
    present: groupPresent[g] || 0,
    total: groupTotal[g] || 0
  }));

  return {
    eventName: ev ? String(ev[1]) : '',
    count: selected.length,
    groups: groups,
    recent: selected.slice(-15).reverse().map(r => ({
      time: formatDateTime_(r[0]).slice(11),
      id: normalizeId_(r[3]),
      name: String(r[4]),
      group: groupMap[normalizeId_(r[3])] || '',
      staff: String(r[5])
    }))
  };
}

function getMemberCount(token) {
  requireSession_(token, ['admin', 'staff']);
  const sh = SpreadsheetApp.getActive().getSheetByName(MEMBERS_SHEET);
  if (!sh || sh.getLastRow() < 2) return 0;
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().filter(r => r[0] && String(r[3] || '在籍') !== '退会').length;
}

function exportAttendanceCsv(token, eventId) {
  requireSession_(token, ['admin', 'staff']);
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(LOG_SHEET);
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues() : [];
  const selected = rows.filter(r => String(r[1]) === String(eventId));
  
  const lines = [['記録日時', '集会ID', '集会名', '会員番号', '氏名', '受付担当'].join(',')];
  selected.forEach(r => {
    lines.push([
      formatDateTime_(r[0]), r[1],
      '"' + String(r[2]).replace(/"/g, '""') + '"',
      '="' + normalizeId_(r[3]) + '"',
      '"' + String(r[4]).replace(/"/g, '""') + '"',
      '"' + String(r[5]).replace(/"/g, '""') + '"'
    ].join(','));
  });
  return '\uFEFF' + lines.join('\r\n');
}

function hashPin_(pin, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}
function verifyPin_(pin, salt, expected) { return !!salt && !!expected && hashPin_(pin, salt) === expected; }
function normalizeId_(v) { return String(v == null ? '' : v).trim().replace(/\.0$/, '').padStart(4, '0'); }
function formatDate_(v) {
  if (!v) return '';
  const d = v instanceof Date ? v : new Date(v);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}
function formatDateTime_(v) {
  if (!v) return '';
  const d = v instanceof Date ? v : new Date(v);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
}