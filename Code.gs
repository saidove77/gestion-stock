// Google Apps Script — نظام تدبير المخزون المدرسي (الإصدار 4: + التوزيع الجماعي)
// - الجميع يستطيع قراءة البيانات وطباعتها (بدون رقم سري).
// - الحفظ والتعديل والحذف وسجل التدقيق: للمشرف فقط برقمه السري.
// - تُنشأ كلمة سر المشرف وتُغيَّر من داخل التطبيق نفسه، لا حاجة لفتح هذا الملف.
// يجب أن يكون هذا السكربت داخل ملف Google Sheets نفسه: Extensions > Apps Script

const SHEETS = {
  Articles:  ['id','name_ar','name_fr','marche','bon_commande','date','stock','min_stock','max_stock'],
  Schools:   ['id','name_ar'],
  Movements: ['id','type','article_id','school_id','quantity','date','notes','batch_id','doc_no','status','cancel_date','plan_id'],
  Plans:     ['id','title','date','status','school_ids','article_ids','created','approved_date','docs_from','docs_to'],
  PlanLines: ['plan_id','school_id','article_id','qty']
};
// أسماء الأوراق المقبولة بدون حساسية لحالة الأحرف (ARTICLES و SCHOOLS و MOVEMENT تعمل كما هي)
const ALIASES = { Articles: ['articles'], Schools: ['schools'], Movements: ['movements', 'movement'], Plans: ['plans'], PlanLines: ['planlines'] };
const AUDIT_SHEET = 'AuditLog';
const AUDIT_COLS = ['id','ts','role','user','action','entity','entity_id','details'];
const WRITE_ACTIONS = ['article_add','article_edit','article_delete','school_add','school_edit','school_delete','entry','distribution','cancel','plan_create','plan_edit','plan_delete','plan_approve','plan_cancel'];
const LOG_EVENTS = ['print','export_excel','export_pdf','logout'];
const VERSION = 4;
const SESSION_TTL = 21600;   // 6 ساعات (الحد الأقصى في CacheService)
const MAX_FAILS = 5, LOCK_SECONDS = 300;

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function props_() { return PropertiesService.getScriptProperties(); }
function now_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }

/* ---------- كلمة سر المشرف (تُخزَّن مشفَّرة، لا تُخزَّن كنص) ---------- */
function sha_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function sessionVersion_() { return Number(props_().getProperty('SESSION_V')) || 0; }
function hasPassword_() { return !!props_().getProperty('ADMIN_HASH'); }
function setPassword_(pw) {
  const salt = Utilities.getUuid(), p = props_();
  p.setProperty('ADMIN_SALT', salt);
  p.setProperty('ADMIN_HASH', sha_(salt + pw));
  p.setProperty('SESSION_V', String(sessionVersion_() + 1));   // يُبطل كل الجلسات القديمة
}
function checkPassword_(pw) {
  const p = props_(), h = p.getProperty('ADMIN_HASH');
  return !!h && sha_(p.getProperty('ADMIN_SALT') + pw) === h;
}
function newSession_() {
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('s_' + token, JSON.stringify({ v: sessionVersion_() }), SESSION_TTL);
  return token;
}
function isAdminToken_(token) {
  if (!token) return false;
  const v = CacheService.getScriptCache().get('s_' + token);
  return !!v && JSON.parse(v).v === sessionVersion_();
}
// للطوارئ فقط: إن نسيت كلمة السر شغّل هذه الدالة من محرر السكربت، ثم ادخل التطبيق وأنشئ كلمة سر جديدة.
function resetAdminPassword() {
  const p = props_();
  p.deleteProperty('ADMIN_HASH'); p.deleteProperty('ADMIN_SALT');
  p.setProperty('SESSION_V', String(sessionVersion_() + 1));
}

/* ---------- الأوراق ---------- */
function findSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const all = ss.getSheets();
  for (let i = 0; i < all.length; i++) {
    if (ALIASES[name].indexOf(all[i].getName().trim().toLowerCase()) >= 0) return all[i];
  }
  const sh = ss.insertSheet(name);
  sh.appendRow(SHEETS[name]);
  return sh;
}
function readSheet(name) {
  const vals = findSheet(name).getDataRange().getValues();
  if (vals.length < 2) return [];
  const head = vals[0].map(String);
  return vals.slice(1).filter(r => r.some(c => c !== '')).map(r => {
    const o = {};
    head.forEach((h, i) => {
      if (!h) return;
      let v = r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      o[h] = v;
    });
    return o;
  });
}
function writeSheet(name, rows) {
  const sh = findSheet(name);
  const cols = SHEETS[name].slice();
  rows.forEach(r => Object.keys(r).forEach(k => { if (cols.indexOf(k) < 0) cols.push(k); }));
  const data = [cols].concat(rows.map(r => cols.map(c => (r[c] === undefined || r[c] === null) ? '' : r[c])));
  sh.clearContents();
  sh.getRange(1, 1, data.length, cols.length).setValues(data);
}
function allData() {
  const res = {};
  Object.keys(SHEETS).forEach(n => { res[n] = readSheet(n); });
  return res;
}

/* ---------- سجل التدقيق ---------- */
function auditSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(AUDIT_SHEET);
  if (!sh) { sh = ss.insertSheet(AUDIT_SHEET); sh.appendRow(AUDIT_COLS); }
  return sh;
}
function audit(action, entity, entityId, details) {
  auditSheet().appendRow([
    Utilities.getUuid().slice(0, 8), now_(), 'admin', 'المشرف',
    String(action || ''), String(entity || ''), String(entityId || ''), String(details || '').slice(0, 500)
  ]);
}
function readAudit(limit) {
  const vals = auditSheet().getDataRange().getValues();
  if (vals.length < 2) return [];
  const head = vals[0].map(String);
  const rows = vals.slice(1).filter(r => r.some(c => c !== '')).map(r => {
    const o = {};
    head.forEach((h, i) => {
      let v = r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      o[h] = v;
    });
    return o;
  });
  return rows.reverse().slice(0, limit);
}

/* ---------- الدخول وكلمة السر ---------- */
function login(req) {
  const cache = CacheService.getScriptCache();
  if (!hasPassword_()) return { status: 'error', code: 'setup', message: 'لم تُنشأ كلمة سر المشرف بعد' };
  const fails = Number(cache.get('fails') || 0);
  if (fails >= MAX_FAILS) return { status: 'error', code: 'locked', message: 'محاولات خاطئة كثيرة. حاول بعد 5 دقائق.' };
  if (!checkPassword_(String(req.password || ''))) {
    cache.put('fails', String(fails + 1), LOCK_SECONDS);
    Utilities.sleep(800);
    audit('login_fail', 'session', '', 'كلمة سر خاطئة');
    return { status: 'error', code: 'bad', message: 'كلمة السر غير صحيحة' };
  }
  cache.remove('fails');
  audit('login', 'session', '', '');
  return { status: 'success', token: newSession_() };
}
function setup(req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (hasPassword_()) return { status: 'error', code: 'exists', message: 'كلمة السر موجودة بالفعل' };
    const pw = String(req.password || '');
    if (pw.length < 6) return { status: 'error', message: 'كلمة السر 6 أحرف على الأقل' };
    setPassword_(pw);
    audit('password_set', 'session', '', 'إنشاء كلمة سر المشرف');
    return { status: 'success', token: newSession_() };
  } finally { try { lock.releaseLock(); } catch (x) {} }
}
function changePassword(req) {
  const pw = String(req.newPassword || '');
  if (pw.length < 6) return { status: 'error', message: 'كلمة السر الجديدة 6 أحرف على الأقل' };
  if (!checkPassword_(String(req.current || ''))) { Utilities.sleep(800); return { status: 'error', code: 'bad', message: 'كلمة السر الحالية غير صحيحة' }; }
  setPassword_(pw);
  audit('password_change', 'session', '', 'تغيير كلمة سر المشرف');
  return { status: 'success', token: newSession_() };
}

/* ---------- الحفظ (للمشرف فقط) ---------- */
function save(req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheets = req.sheets || {};
    const names = Object.keys(sheets);
    if (!names.length) return { status: 'error', message: 'لا توجد بيانات للحفظ' };
    for (let i = 0; i < names.length; i++) {
      if (!SHEETS[names[i]] || !Array.isArray(sheets[names[i]])) return { status: 'error', message: 'ورقة غير صالحة: ' + names[i] };
    }
    names.forEach(n => writeSheet(n, sheets[n]));
    (Array.isArray(req.audit) ? req.audit : []).slice(0, 50).forEach(a => {
      if (WRITE_ACTIONS.indexOf(a.action) >= 0) audit(a.action, String(a.entity || '').slice(0, 40), String(a.entity_id || '').slice(0, 60), a.details);
    });
    return { status: 'success' };
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------- نقاط الدخول ---------- */
function doGet(e) {
  // افتح رابط التطبيق (…/exec) في المتصفح: إن ظهر "version":4 فالنسخة الجديدة منشورة.
  return out({ status: 'ok', version: VERSION, message: 'استعمل التطبيق للوصول إلى البيانات' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out({ status: 'error', message: 'طلب غير صالح' }); }
  try {
    // مفتوحة للجميع (قراءة فقط)
    if (req.action === 'getData') return out({ status: 'success', version: VERSION, admin: isAdminToken_(req.token), configured: hasPassword_(), data: allData() });
    if (req.action === 'login') return out(login(req));
    if (req.action === 'setup') return out(setup(req));
    // ما بعد هذا للمشرف فقط
    if (!isAdminToken_(req.token)) return out({ status: 'error', code: 'auth', message: 'هذه العملية للمشرف فقط. سجّل الدخول.' });
    if (req.action === 'saveData') return out(save(req));
    if (req.action === 'changePassword') return out(changePassword(req));
    if (req.action === 'getAudit') return out({ status: 'success', data: readAudit(Math.min(Number(req.limit) || 300, 1000)) });
    if (req.action === 'log') {
      if (LOG_EVENTS.indexOf(req.event) >= 0) audit(req.event, String(req.entity || '').slice(0, 40), String(req.entity_id || '').slice(0, 60), req.details);
      return out({ status: 'success' });
    }
    return out({ status: 'error', message: 'إجراء غير معروف' });
  } catch (err) {
    return out({ status: 'error', message: String(err) });
  }
}
