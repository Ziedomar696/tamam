/**
 * تمام الملاك — مراسي (Google Apps Script)
 * كل قطاع بيسلّم ٤ مرات في اليوم عدد الموجودين (ملاك / مستأجرين / مفوضين) من حسابه.
 *
 * ⚙️ SETUP_CODE: كود بيتطلب مرة واحدة بس علشان تعمل أول حساب "الإدارة".
 *    غيّره لأي كلمة تختارها قبل ما تعمل Deploy.
 */
const SETUP_CODE = 'CHANGE-ME';

const TZ = 'Africa/Cairo';
const TOKEN_DAYS = 30;

const LOG_SHEET = 'سجل التمام';
const EDITS_SHEET = 'سجل التعديلات';
const USERS_SHEET = 'المستخدمين';
const SECTORS_SHEET = 'القطاعات';
const SETTINGS_SHEET = 'الإعدادات';

const L_HEADERS = ['يوم التمام', 'موعد التمام', 'القطاع', 'ملاك', 'مستأجرين', 'مفوضين', 'الإجمالي', 'مشرف القطاع', 'الحساب',
  'وقت التسليم', 'متأخر', 'آخر تعديل', 'ملاحظات', 'المفتاح',
  'ملاك فيلات', 'مستأجرين فيلات', 'مفوضين فيلات', 'ملاك شقق', 'مستأجرين شقق', 'مفوضين شقق'];
// أنواع الوحدات: كل قطاع فيه فيلات أو شقق أو الاتنين، والتمام بيتقسم عليهم
const UTYPES = [{ key: 'v', label: 'فيلات' }, { key: 'a', label: 'شقق' }];
const UCOLS = { v: ['ملاك فيلات', 'مستأجرين فيلات', 'مفوضين فيلات'], a: ['ملاك شقق', 'مستأجرين شقق', 'مفوضين شقق'] };
const LC = {}; L_HEADERS.forEach(function (h, i) { LC[h] = i; });
const E_HEADERS = ['الوقت', 'الحساب', 'يوم التمام', 'موعد التمام', 'القطاع', 'قبل', 'بعد', 'ملاحظات'];
const U_HEADERS = ['اسم المستخدم', 'الاسم', 'الدور', 'القطاعات', 'نشط', 'آخر دخول', 'تاريخ الإنشاء', 'الإيميل', 'salt', 'hash', 'ver'];

const ROLES = { admin: 'الإدارة', manager: 'مدير', guard: 'فرد أمن / مشرف قطاع' };
// المواعيد: h = الساعة (24 = ١٢ منتصف الليل، وبيتحسب على اليوم اللي قبله)
const ROUNDS = [
  { key: '06', label: '6 صباحاً', h: 6 },
  { key: '12', label: '12 ظهراً', h: 12 },
  { key: '18', label: '6 مساءً', h: 18 },
  { key: '24', label: '12 منتصف الليل', h: 24 }
];
const EARLY_MIN = 60;                 // فرد الأمن يقدر يسلّم قبل الميعاد بساعة
const DEFAULT_SECTORS = ['Arezzo', 'Verona', 'Isola', 'Veneto'];

// ================================================================ web app

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('تمام الملاك')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** شغّلها من المحرر بعد أي تحديث: بتجهز الشيت وتنبيه القطاعات اللي ماسلّمتش. */
function setup() {
  const ss = ss_();
  logSheet_(); editsSheet_(); usersSheet_(); sectorsSheet_(); settingsSheet_();
  ss.getSheets().forEach(function (s) {
    if ([LOG_SHEET, EDITS_SHEET, USERS_SHEET, SECTORS_SHEET, SETTINGS_SHEET].indexOf(s.getName()) < 0 && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === 'checkMissing'; }).forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('checkMissing').timeBased().everyMinutes(15).create();
  Logger.log('الشيت جاهز: ' + ss.getUrl());
  Logger.log(needsSetup() ? 'لسه مفيش حساب إدارة: افتح التطبيق واعمل أول حساب بـ SETUP_CODE.' : 'حسابات الإدارة موجودة.');
}

function ss_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActive(); } catch (e) {}
  if (ss) return ss;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  ss = SpreadsheetApp.create('تمام الملاك');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

// ================================================================ sheets

const _sheets = {};
function ensureSheet_(name, headers, init) {
  if (_sheets[name]) return _sheets[name];
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  const style = function (r) { r.setFontWeight('bold').setBackground('#1E5A4C').setFontColor('#FFFFFF'); };
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
    style(sh.getRange(1, 1, 1, headers.length).setValues([headers]));
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    if (init) init(sh);
  } else {
    const w = sh.getLastColumn();
    if (w < headers.length) style(sh.getRange(1, w + 1, 1, headers.length - w).setValues([headers.slice(w)]));
  }
  _sheets[name] = sh;
  return sh;
}
function logSheet_() {
  return ensureSheet_(LOG_SHEET, L_HEADERS, function (sh) {
    sh.getRange(1, LC['ملاك'] + 1, sh.getMaxRows(), 4).setNumberFormat('0');
    sh.getRange(1, LC['ملاك فيلات'] + 1, sh.getMaxRows(), 6).setNumberFormat('0');
    sh.hideColumns(LC['المفتاح'] + 1, 1); sh.hideColumns(LC['الحساب'] + 1, 1);
  });
}
function editsSheet_() { return ensureSheet_(EDITS_SHEET, E_HEADERS); }
function usersSheet_() { return ensureSheet_(USERS_SHEET, U_HEADERS, function (sh) { sh.hideColumns(9, 3); }); }
function sectorsSheet_() {
  return ensureSheet_(SECTORS_SHEET, ['القطاع', 'نشط', 'الوحدات'], function (sh) {
    sh.getRange(2, 1, DEFAULT_SECTORS.length, 3).setValues(DEFAULT_SECTORS.map(function (s) { return [s, 'نعم', 'فيلات وشقق']; }));
  });
}
function settingsSheet_() {
  return ensureSheet_(SETTINGS_SHEET, ['الإعداد', 'القيمة'], function (sh) {
    sh.getRange(2, 1, 2, 2).setValues([['دقايق السماح قبل ما يتحسب متأخر', '30'], ['تنبيه بالإيميل للقطاع اللي ماسلّمش', 'نعم']]);
  });
}

function fmt_(x) { return x instanceof Date ? Utilities.formatDate(x, TZ, 'yyyy-MM-dd HH:mm') : x; }
function values_(sh, width, from) {
  const last = sh.getLastRow();
  const start = Math.max(2, from || 2);
  if (last < start) return [];
  return sh.getRange(start, 1, last - start + 1, width).getValues().map(function (r) { return r.map(fmt_); });
}

function settings_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('settings');
  if (hit) return JSON.parse(hit);
  const out = { grace: 30, alerts: true };
  values_(settingsSheet_(), 2).forEach(function (r) {
    const k = String(r[0]), v = String(r[1]).trim();
    if (/السماح/.test(k)) out.grace = Math.max(0, Math.min(240, Number(v) || 0));
    if (/تنبيه/.test(k)) out.alerts = v !== 'لا';
  });
  cache.put('settings', JSON.stringify(out), 300);
  return out;
}
function saveSettings(token, d) {
  auth_(token, ['admin']);
  const grace = Math.round(Number(d.grace));
  if (!(grace >= 0 && grace <= 240)) throw new Error('دقايق السماح من ٠ لـ ٢٤٠.');
  const sh = settingsSheet_();
  sh.getRange(2, 1, 2, 2).setValues([['دقايق السماح قبل ما يتحسب متأخر', String(grace)], ['تنبيه بالإيميل للقطاع اللي ماسلّمش', d.alerts === false ? 'لا' : 'نعم']]);
  CacheService.getScriptCache().remove('settings');
  return adminData(token);
}

// ================================================================ sectors

function sectors_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('sectors');
  if (hit) return JSON.parse(hit);
  const out = values_(sectorsSheet_(), 3).filter(function (r) { return String(r[0]).trim(); })
    .map(function (r) { return { name: String(r[0]).trim(), active: String(r[1]).trim() !== 'لا', types: typesOf_(r[2]) }; });
  cache.put('sectors', JSON.stringify(out), 300);
  return out;
}
/** "فيلات" / "شقق" / "فيلات وشقق" → ['v'] / ['a'] / ['v','a'] (الفاضي = الاتنين). */
function typesOf_(v) {
  v = String(v || '');
  const t = [];
  if (/فيل/.test(v)) t.push('v');
  if (/شق/.test(v)) t.push('a');
  return t.length ? t : ['v', 'a'];
}
function typesLabel_(t) { return UTYPES.filter(function (u) { return t.indexOf(u.key) >= 0; }).map(function (u) { return u.label; }).join(' و'); }
function sectorTypes_() { const o = {}; sectors_().forEach(function (s) { o[s.name] = s.types; }); return o; }
function activeSectors_() { return sectors_().filter(function (s) { return s.active; }).map(function (s) { return s.name; }); }
function mySectors_(u) {
  const all = activeSectors_();
  if (u.role === 'admin' || u.sectors.indexOf('*') >= 0) return all;
  return all.filter(function (s) { return u.sectors.indexOf(s) >= 0; });
}
function canSee_(u, sector) { return u.role === 'admin' || u.sectors.indexOf('*') >= 0 || u.sectors.indexOf(sector) >= 0; }

function saveSector(token, name, active, types) {
  auth_(token, ['admin']);
  name = clean_(name, 60);
  if (!name) throw new Error('اكتب اسم القطاع.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sectorsSheet_();
    const i = sectors_().map(function (s) { return s.name; }).indexOf(name);
    if (i >= 0) sh.getRange(i + 2, 2).setValue(active === false ? 'لا' : 'نعم');
    else sh.appendRow(["'" + name, 'نعم', typesLabel_(typesOf_(types))]);
    CacheService.getScriptCache().remove('sectors');
  } finally { lock.releaseLock(); }
  return adminData(token);
}
/** نوع وحدات القطاع: ['v'] فيلات، ['a'] شقق، أو الاتنين. */
function saveSectorTypes(token, name, types) {
  auth_(token, ['admin']);
  types = (Array.isArray(types) ? types : []).filter(function (t) { return UCOLS[t]; });
  if (!types.length) throw new Error('القطاع لازم يبقى فيه فيلات أو شقق أو الاتنين.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const i = sectors_().map(function (s) { return s.name; }).indexOf(name);
    if (i < 0) throw new Error('القطاع مش موجود.');
    sectorsSheet_().getRange(i + 2, 3).setValue(typesLabel_(types));
    CacheService.getScriptCache().remove('sectors');
  } finally { lock.releaseLock(); }
  return adminData(token);
}

// ================================================================ rounds & days

function dayKey_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function stamp_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd HH:mm'); }
function pad_(n) { return (n < 10 ? '0' : '') + n; }
/** تاريخ ووقت في توقيت القاهرة → Date */
function cairo_(day, h) {
  const p = day.split('-').map(Number);
  // نحسب فرق توقيت القاهرة في اليوم ده (بيتغير صيف/شتا)
  const guess = new Date(Date.UTC(p[0], p[1] - 1, p[2], h, 0, 0));
  const shown = Utilities.formatDate(guess, TZ, 'yyyy-MM-dd HH:mm').split(/[- :]/).map(Number);
  const asUtc = Date.UTC(shown[0], shown[1] - 1, shown[2], shown[3], shown[4]);
  return new Date(guess.getTime() - (asUtc - guess.getTime()));
}
function shiftDay_(key, n) {
  const p = key.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n));
  return d.getUTCFullYear() + '-' + pad_(d.getUTCMonth() + 1) + '-' + pad_(d.getUTCDate());
}
function roundOf_(key) { return ROUNDS.filter(function (r) { return r.key === key; })[0] || null; }
/** ميعاد التمام الفعلي (١٢ منتصف الليل = أول الساعة صفر من اليوم اللي بعده). */
function schedOf_(day, r) { return r.h === 24 ? cairo_(shiftDay_(day, 1), 0) : cairo_(day, r.h); }

/** الميعاد الحالي لفرد الأمن: آخر ميعاد بدأ (مع سماح ساعة قبله). */
function currentRound_(now) {
  now = now || new Date();
  const today = dayKey_(now);
  let best = null;
  [-1, 0].forEach(function (k) {
    const day = shiftDay_(today, k);
    ROUNDS.forEach(function (r) {
      const t = schedOf_(day, r);
      if (t.getTime() - EARLY_MIN * 60000 <= now.getTime() && (!best || t > best.sched)) best = { day: day, round: r, sched: t };
    });
  });
  const g = settings_().grace;
  return { day: best.day, key: best.round.key, label: best.round.label, at: stamp_(best.sched),
    lateAfter: stamp_(new Date(best.sched.getTime() + g * 60000)), schedMs: best.sched.getTime() };
}

// ================================================================ users & auth

function users_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('users');
  if (hit) return JSON.parse(hit);
  const out = [];
  values_(usersSheet_(), U_HEADERS.length).forEach(function (r, i) {
    if (!r[0]) return;
    out.push({ row: i + 2, username: String(r[0]).trim().toLowerCase(), name: String(r[1]), role: roleKey_(r[2]),
      sectors: String(r[3] || '').split(/\s*[,،]\s*/).map(function (x) { return x.trim(); }).filter(String),
      active: String(r[4]).trim() !== 'لا', lastLogin: String(r[5] || ''), created: String(r[6] || ''), email: String(r[7] || '').trim(),
      salt: String(r[8]), hash: String(r[9]), ver: Number(r[10]) || 1 });
  });
  cache.put('users', JSON.stringify(out), 300);
  return out;
}
function dropCache_(k) { CacheService.getScriptCache().remove(k); }
function findUser_(username) {
  username = String(username || '').trim().toLowerCase();
  return users_().filter(function (u) { return u.username === username; })[0] || null;
}
function roleKey_(v) {
  v = String(v || '').trim();
  for (const k in ROLES) if (k === v || ROLES[k] === v) return k;
  return 'guard';
}
function publicUser_(u) {
  return { username: u.username, name: u.name, role: u.role, roleLabel: ROLES[u.role], sectors: u.role === 'admin' ? ['*'] : u.sectors,
    active: u.active, lastLogin: u.lastLogin, created: u.created, email: u.email };
}
function hash_(salt, pw) {
  let h = String(pw);
  for (let i = 0; i < 50; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + h, Utilities.Charset.UTF_8));
  return h;
}
function secret_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('TOKEN_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('TOKEN_SECRET', s); }
  return s;
}
function sign_(payload) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, secret_())); }
function makeToken_(u) {
  const p = u.username + '|' + (Date.now() + TOKEN_DAYS * 864e5) + '|' + u.ver;
  return Utilities.base64EncodeWebSafe(p) + '.' + sign_(p);
}
function auth_(token, roles) {
  const parts = String(token || '').split('.');
  if (parts.length !== 2) throw new Error('AUTH');
  let p;
  try { p = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString(); } catch (e) { throw new Error('AUTH'); }
  if (sign_(p) !== parts[1]) throw new Error('AUTH');
  const f = p.split('|');
  if (Number(f[1]) < Date.now()) throw new Error('AUTH');
  const u = findUser_(f[0]);
  if (!u || !u.active || String(u.ver) !== f[2]) throw new Error('AUTH');
  if (roles && roles.indexOf(u.role) < 0) throw new Error('مش مسموح لحسابك بالعملية دي.');
  return u;
}

function boot_(u) {
  return { me: publicUser_(u), sectors: mySectors_(u), types: sectorTypes_(), utypes: UTYPES, rounds: ROUNDS.map(function (r) { return { key: r.key, label: r.label }; }),
    current: currentRound_(), grace: settings_().grace };
}
function needsSetup() { return !users_().some(function (u) { return u.role === 'admin' && u.active; }); }

function setupAdmin(code, username, name, password, email) {
  if (SETUP_CODE === 'CHANGE-ME') throw new Error('غيّر SETUP_CODE في أول ملف Code.gs واعمل Deploy بنسخة جديدة الأول.');
  if (String(code || '').trim() !== SETUP_CODE) throw new Error('كود التفعيل غلط.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (!needsSetup()) throw new Error('حساب الإدارة معمول قبل كده. سجّل دخول.');
    writeUser_({ username: username, name: name, role: 'admin', sectors: ['*'], active: true, password: password, email: email }, true);
  } finally { lock.releaseLock(); }
  return login(username, password);
}

function login(username, password) {
  username = String(username || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache(), key = 'fail:' + username;
  const fails = Number(cache.get(key)) || 0;
  if (fails >= 5) throw new Error('محاولات كتير غلط. استنى ١٠ دقايق وجرّب تاني.');
  const u = findUser_(username);
  if (!u || !u.active || hash_(u.salt, password) !== u.hash) {
    cache.put(key, String(fails + 1), 600);
    throw new Error(u && !u.active ? 'الحساب ده متوقف. كلّم الإدارة.' : 'اسم المستخدم أو كلمة السر غلط.');
  }
  cache.remove(key);
  try { usersSheet_().getRange(u.row, 6).setValue("'" + stamp_(new Date())); dropCache_('users'); } catch (e) {}
  const out = boot_(u); out.token = makeToken_(u);
  return out;
}
function whoami(token) { return boot_(auth_(token)); }

function changeMyPassword(token, oldPw, newPw) {
  const u = auth_(token);
  if (hash_(u.salt, oldPw) !== u.hash) throw new Error('كلمة السر الحالية غلط.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { writeUser_({ username: u.username, name: u.name, role: u.role, sectors: u.sectors, active: true, password: newPw, email: u.email }, false); }
  finally { lock.releaseLock(); }
  return makeToken_(findUser_(u.username));
}

function adminData(token) {
  auth_(token, ['admin']);
  return { users: users_().map(publicUser_), roles: ROLES, sectors: sectors_(), settings: settings_() };
}
function saveUser(token, d) {
  auth_(token, ['admin']);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const target = findUser_(d.username);
    if (!d.isNew && target && target.role === 'admin' && (roleKey_(d.role) !== 'admin' || d.active === false)) lastAdminGuard_(target);
    writeUser_(d, !!d.isNew);
  } finally { lock.releaseLock(); }
  return adminData(token);
}
function deleteUser(token, username) {
  const me = auth_(token, ['admin']);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const t = findUser_(username);
    if (!t) throw new Error('المستخدم مش موجود.');
    if (t.username === me.username) throw new Error('مينفعش تمسح حسابك.');
    if (t.role === 'admin') lastAdminGuard_(t);
    usersSheet_().deleteRow(t.row);
    dropCache_('users');
  } finally { lock.releaseLock(); }
  return adminData(token);
}
function lastAdminGuard_(target) {
  const others = users_().filter(function (u) { return u.role === 'admin' && u.active && u.username !== target.username; });
  if (!others.length) throw new Error('لازم يفضل حساب واحد على الأقل للإدارة شغال.');
}
function writeUser_(d, isNew) {
  const username = String(d.username || '').trim().toLowerCase();
  if (!/^[a-z0-9._-]{3,30}$/.test(username)) throw new Error('اسم المستخدم لازم يكون ٣ حروف على الأقل، إنجليزي صغير أو أرقام (مثال: ahmed.v1).');
  const name = clean_(d.name, 80) || username;
  const role = roleKey_(d.role);
  const active = d.active !== false;
  const known = sectors_().map(function (s) { return s.name; });
  let sectors = role === 'admin' ? ['*'] : (Array.isArray(d.sectors) ? d.sectors : []).map(function (s) { return String(s).trim(); })
    .filter(function (s, i, a) { return (s === '*' || known.indexOf(s) >= 0) && a.indexOf(s) === i; });
  if (sectors.indexOf('*') >= 0) sectors = ['*'];
  if (role === 'guard' && sectors.indexOf('*') >= 0) throw new Error('فرد الأمن لازم يبقى على قطاعات محددة.');
  if (role !== 'admin' && !sectors.length) throw new Error('اختار قطاع واحد على الأقل.');
  const email = String(d.email || '').trim().toLowerCase();
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('الإيميل مش مظبوط.');
  const existing = findUser_(username);
  if (isNew && existing) throw new Error('اسم المستخدم ده موجود قبل كده.');
  if (!isNew && !existing) throw new Error('المستخدم مش موجود.');
  const pw = String(d.password || '');
  if ((isNew || pw) && pw.length < 6) throw new Error('كلمة السر لازم تبقى ٦ حروف أو أرقام على الأقل.');
  let salt = existing ? existing.salt : '', hash = existing ? existing.hash : '', ver = existing ? existing.ver : 1;
  if (pw) { salt = Utilities.getUuid(); hash = hash_(salt, pw); if (existing) ver++; }
  if (existing && existing.active && !active) ver++;
  if (existing && existing.role !== role) ver++;
  const row = [username, name, ROLES[role], sectors.join(', '), active ? 'نعم' : 'لا', existing ? existing.lastLogin : '',
    existing ? existing.created : "'" + stamp_(new Date()), d.email === undefined && existing ? existing.email : email, salt, hash, String(ver)];
  const sh = usersSheet_();
  if (existing) sh.getRange(existing.row, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  dropCache_('users');
}

// ================================================================ التمام

function key_(day, round, sector) { return day + '|' + round + '|' + sector; }
function rowOfKey_(sh, k) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const hit = sh.getRange(2, LC['المفتاح'] + 1, last - 1, 1).createTextFinder(k).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}
function rowObj_(r) {
  const n = function (h) { return Math.max(0, Number(r[LC[h]]) || 0); };
  const part = function (t) {
    const c = UCOLS[t].map(function (h) { return Math.max(0, Number(r[LC[h]]) || 0); });
    return { owners: c[0], tenants: c[1], delegates: c[2], total: c[0] + c[1] + c[2] };
  };
  return { day: String(r[LC['يوم التمام']]), round: roundKeyOf_(String(r[LC['موعد التمام']])), sector: String(r[LC['القطاع']]),
    v: part('v'), a: part('a'),
    owners: n('ملاك'), tenants: n('مستأجرين'), delegates: n('مفوضين'), total: n('ملاك') + n('مستأجرين') + n('مفوضين'),
    by: String(r[LC['مشرف القطاع']]), at: String(r[LC['وقت التسليم']]), late: String(r[LC['متأخر']]) === 'نعم',
    edited: String(r[LC['آخر تعديل']] || ''), notes: String(r[LC['ملاحظات']] || '') };
}
function roundKeyOf_(label) { const r = ROUNDS.filter(function (x) { return x.label === label || x.key === label; })[0]; return r ? r.key : label; }
/** كل تسليمات يوم (أو أيام). */
function rowsFor_(days) {
  const set = {}; days.forEach(function (d) { set[d] = 1; });
  const sh = logSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, L_HEADERS.length).getValues().map(function (r) { return r.map(fmt_); })
    .filter(function (r) { return set[String(r[LC['يوم التمام']])]; }).map(rowObj_);
}
function count_(v, label) {
  const s = String(v == null ? '' : v).replace(/[٠-٩]/g, function (d) { return '٠١٢٣٤٥٦٧٨٩'.indexOf(d); }).trim();
  if (!/^\d{1,5}$/.test(s)) throw new Error('اكتب عدد ' + label + ' (رقم، صفر أو أكتر).');
  return Number(s);
}

/** شاشة فرد الأمن: الميعاد الحالي + اللي اتسلّم لقطاعاته النهارده. */
function myState(token) {
  const u = auth_(token);
  const cur = currentRound_();
  const secs = mySectors_(u);
  const rows = rowsFor_([cur.day]).filter(function (r) { return secs.indexOf(r.sector) >= 0; });
  return { current: cur, sectors: secs, types: sectorTypes_(), rows: rows };
}

/**
 * تسليم التمام. فرد الأمن: الميعاد الحالي بس ولقطاعاته بس (ولو سلّم قبل كده بيعدّل، والتعديل بيتسجل).
 * المدير والإدارة: أي يوم وأي ميعاد في قطاعاتهم (للتصحيح).
 */
function submitTamam(token, d) {
  const u = auth_(token);
  const sector = clean_(d.sector, 60);
  if (mySectors_(u).indexOf(sector) < 0) throw new Error('حسابك مش على قطاع ' + sector + '.');
  const cur = currentRound_();
  let day = cur.day, rk = cur.key;
  if (u.role !== 'guard' && d.day && d.round) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d.day) || !roundOf_(d.round)) throw new Error('اليوم أو الميعاد غلط.');
    day = d.day; rk = d.round;
    if (schedOf_(day, roundOf_(rk)).getTime() - EARLY_MIN * 60000 > Date.now()) throw new Error('الميعاد ده لسه ماجاش.');
  } else if (d.round && d.round !== cur.key) {
    throw new Error('ميعاد تمام ' + cur.label + ' هو المفتوح دلوقتي. حدّث الصفحة.');
  }
  const r = roundOf_(rk);
  // الأعداد مقسومة على الفيلات والشقق حسب نوع وحدات القطاع
  const types = sectorTypes_()[sector] || ['v', 'a'];
  const parts = {};
  UTYPES.forEach(function (ut) {
    const src = (d.parts || {})[ut.key];
    if (types.indexOf(ut.key) < 0) { parts[ut.key] = [0, 0, 0]; return; }
    if (!src) throw new Error('اكتب أعداد ال' + ut.label + '.');
    parts[ut.key] = [count_(src.owners, 'ملاك ال' + ut.label), count_(src.tenants, 'مستأجرين ال' + ut.label), count_(src.delegates, 'مفوضين ال' + ut.label)];
  });
  const owners = parts.v[0] + parts.a[0], tenants = parts.v[1] + parts.a[1], delegates = parts.v[2] + parts.a[2];
  const notes = clean_(d.notes, 500);
  const sig = function (pv, pa) { return types.map(function (t) { return (t === 'v' ? 'فيلات ' : 'شقق ') + (t === 'v' ? pv : pa).join('/'); }).join(' · '); };
  const now = new Date(), k = key_(day, rk, sector);
  const late = now.getTime() > schedOf_(day, r).getTime() + settings_().grace * 60000;
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = logSheet_();
    const at = rowOfKey_(sh, k);
    if (at) {
      const old = rowObj_(sh.getRange(at, 1, 1, L_HEADERS.length).getValues()[0].map(fmt_));
      const before = sig([old.v.owners, old.v.tenants, old.v.delegates], [old.a.owners, old.a.tenants, old.a.delegates]), after = sig(parts.v, parts.a);
      if (before === after && old.notes === notes) return { ok: true, same: true, state: myState(token) };
      sh.getRange(at, LC['ملاك'] + 1, 1, 4).setValues([[owners, tenants, delegates, owners + tenants + delegates]]);
      sh.getRange(at, LC['ملاك فيلات'] + 1, 1, 6).setValues([parts.v.concat(parts.a)]);
      sh.getRange(at, LC['آخر تعديل'] + 1, 1, 2).setValues([["'" + u.name + ' ' + stamp_(now), notes]]);
      editsSheet_().appendRow(["'" + stamp_(now), u.username, "'" + day, r.label, sector, before, after, notes]);
      return { ok: true, edited: true, state: myState(token) };
    }
    const row = new Array(L_HEADERS.length).fill('');
    row[LC['يوم التمام']] = "'" + day; row[LC['موعد التمام']] = r.label; row[LC['القطاع']] = "'" + sector;
    row[LC['ملاك']] = owners; row[LC['مستأجرين']] = tenants; row[LC['مفوضين']] = delegates; row[LC['الإجمالي']] = owners + tenants + delegates;
    row[LC['مشرف القطاع']] = u.name; row[LC['الحساب']] = u.username; row[LC['وقت التسليم']] = "'" + stamp_(now);
    row[LC['متأخر']] = late ? 'نعم' : ''; row[LC['ملاحظات']] = notes; row[LC['المفتاح']] = k;
    UTYPES.forEach(function (ut) { UCOLS[ut.key].forEach(function (h, i) { row[LC[h]] = parts[ut.key][i]; }); });
    sh.appendRow(row);
    return { ok: true, late: late, state: myState(token) };
  } finally { lock.releaseLock(); }
}

/** لوحة المتابعة ليوم معيّن (للمدير والإدارة)، في قطاعاتهم بس. */
function dashboard(token, day) {
  const u = auth_(token, ['admin', 'manager']);
  const cur = currentRound_();
  day = /^\d{4}-\d{2}-\d{2}$/.test(String(day || '')) ? day : cur.day;
  const secs = mySectors_(u);
  const rows = rowsFor_([day]).filter(function (r) { return canSee_(u, r.sector); });
  const now = Date.now(), grace = settings_().grace * 60000;
  const rounds = ROUNDS.map(function (r) {
    const got = rows.filter(function (x) { return x.round === r.key; });
    const sum = function (f) { return got.reduce(function (a, x) { return a + x[f]; }, 0); };
    const sumT = function (t) { return got.reduce(function (a, x) { return a + x[t].total; }, 0); };
    const sched = schedOf_(day, r).getTime();
    const state = now < sched - EARLY_MIN * 60000 ? 'future' : now < sched + grace ? 'open' : 'due';
    return { key: r.key, label: r.label, at: stamp_(new Date(sched)), state: state,
      owners: sum('owners'), tenants: sum('tenants'), delegates: sum('delegates'), total: sum('total'), villas: sumT('v'), apts: sumT('a'),
      done: got.length, missing: state === 'future' ? [] : secs.filter(function (s) { return !got.some(function (x) { return x.sector === s; }); }) };
  });
  const cells = {};
  rows.forEach(function (x) { (cells[x.sector] = cells[x.sector] || {})[x.round] = x; });
  // آخر ٧ أيام: إجمالي كل ميعاد
  const days = []; for (let i = 6; i >= 0; i--) days.push(shiftDay_(day, -i));
  const wk = rowsFor_(days).filter(function (r) { return canSee_(u, r.sector); });
  const week = days.map(function (dd) {
    const o = { day: dd };
    ROUNDS.forEach(function (r) { const g = wk.filter(function (x) { return x.day === dd && x.round === r.key; }); o[r.key] = g.length ? g.reduce(function (a, x) { return a + x.total; }, 0) : null; });
    return o;
  });
  return { day: day, today: cur.day, current: cur, sectors: secs, types: sectorTypes_(), rounds: rounds, cells: cells, week: week };
}

/** تفاصيل خانة: مين سلّم وإمتى، وسجل التعديلات. */
function cellDetail(token, day, round, sector) {
  const u = auth_(token, ['admin', 'manager']);
  if (!canSee_(u, sector)) throw new Error('القطاع ده مش عندك.');
  const r = roundOf_(round);
  const sh = editsSheet_();
  const edits = values_(sh, E_HEADERS.length).filter(function (x) { return String(x[2]) === day && String(x[3]) === (r && r.label) && String(x[4]) === sector; })
    .map(function (x) { return { at: String(x[0]), by: String(x[1]), before: String(x[5]), after: String(x[6]) }; }).reverse();
  return { edits: edits };
}

/** تصدير فترة (CSV). */
function exportRange(token, from, to) {
  const u = auth_(token, ['admin', 'manager']);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to) || from > to) throw new Error('اختار الفترة صح.');
  const days = []; for (let d = from, i = 0; d <= to && i < 400; d = shiftDay_(d, 1), i++) days.push(d);
  const order = {}; ROUNDS.forEach(function (r, i) { order[r.key] = i; });
  const rows = rowsFor_(days).filter(function (r) { return canSee_(u, r.sector); })
    .sort(function (a, b) { return a.day < b.day ? -1 : a.day > b.day ? 1 : order[a.round] - order[b.round] || (a.sector < b.sector ? -1 : 1); });
  return [['يوم التمام', 'موعد التمام', 'القطاع', 'ملاك فيلات', 'مستأجرين فيلات', 'مفوضين فيلات', 'إجمالي الفيلات', 'ملاك شقق', 'مستأجرين شقق', 'مفوضين شقق', 'إجمالي الشقق',
    'ملاك', 'مستأجرين', 'مفوضين', 'الإجمالي', 'مشرف القطاع', 'وقت التسليم', 'متأخر', 'آخر تعديل', 'ملاحظات']]
    .concat(rows.map(function (x) { return [x.day, roundOf_(x.round).label, x.sector, x.v.owners, x.v.tenants, x.v.delegates, x.v.total, x.a.owners, x.a.tenants, x.a.delegates, x.a.total,
      x.owners, x.tenants, x.delegates, x.total, x.by, x.at, x.late ? 'نعم' : '', x.edited, x.notes]; }));
}

// ================================================================ تنبيه اللي ماسلّمش

/** كل ربع ساعة: لو ميعاد عدّى ودقايق السماح خلصت وفيه قطاع ماسلّمش → إيميل مرة واحدة. */
function checkMissing() {
  if (!settings_().alerts) return;
  const cur = currentRound_();
  const now = Date.now();
  if (now < cur.schedMs + settings_().grace * 60000) return;
  const props = PropertiesService.getScriptProperties(), pk = 'ALERT_' + cur.day + '_' + cur.key;
  if (props.getProperty(pk)) return;
  const got = rowsFor_([cur.day]).filter(function (r) { return r.round === cur.key; }).map(function (r) { return r.sector; });
  const missing = activeSectors_().filter(function (s) { return got.indexOf(s) < 0; });
  props.setProperty(pk, '1');
  if (!missing.length) return;
  const to = users_().filter(function (u) { return u.active && u.email && (u.role === 'admin' || (u.role === 'manager' && missing.some(function (s) { return canSee_(u, s); }))); })
    .map(function (u) { return u.email; });
  if (!to.length) { try { to.push(Session.getEffectiveUser().getEmail()); } catch (e) {} }
  if (!to.length || !to[0]) return;
  MailApp.sendEmail({ to: to.join(','), subject: '⚠️ تمام ' + cur.label + ' — ' + missing.length + ' قطاع ماسلّمش',
    htmlBody: '<div dir="rtl" style="font-family:Tahoma,Arial,sans-serif"><h3 style="color:#B3261E">⚠️ تمام ' + esc_(cur.label) + ' (' + cur.day + ')</h3>' +
      '<p>القطاعات اللي لسه ماسلّمتش:</p><ul>' + missing.map(function (s) { return '<li><b>' + esc_(s) + '</b></li>'; }).join('') + '</ul>' +
      '<p>افتح تطبيق تمام الملاك للمتابعة.</p></div>',
    body: 'تمام ' + cur.label + ' — لسه ماسلّمش: ' + missing.join('، ') });
}

// ================================================================ helpers
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function clean_(s, max) { return String(s == null ? '' : s).replace(/^[=+\-@]+/, '').trim().slice(0, max || 300); }
