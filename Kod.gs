/* =====================================================================
   MİQ və Sertifikasiya Hazırlığı — Google Apps Script (server)
   My New Vocabulary ilə eyni qayda:
   • hər kəs bir dəfə qeydiyyatdan keçir (ad, soyad, telefon)
   • qeydiyyatdan 5 gün sonra avtomatik bloklanır
   • müəllim panelində "💳 Ödəniş edib" basılanda blok açılır və 6 ay bloklanmır
   • müəllimlər heç vaxt bloklanmır
   QURAŞDIRMA: README.md-dəki addımlara baxın.
   ===================================================================== */
const ADMIN_PASS = 'BURAYA-PANEL-SIFRESI'; /* 🔐 Müəllim panelinin şifrəsi — buraya özünüz yazın */
const TRIAL_DAYS = 5, PAY_MONTHS = 6;
const DAY = 864e5;
const UL = ['Telefon', 'Ad', 'Soyad', 'Qeydiyyat tarixi', 'Son giriş', 'Bloklanıb', 'Blok növü', 'Ödəniş tarixi', 'Müəllim', 'İstisna', 'Cihaz'];
const DF = 'dd.mm.yyyy hh:mm';

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
const ms = x => x instanceof Date ? x.getTime() : (+x || 0);
const isTrue = x => x === true || /^(true|bəli|beli|yes|1)$/i.test(String(x).trim());
const cleanTel = t => String(t || '').replace(/\D/g, '').replace(/^994/, '').replace(/^0/, '').slice(0, 9);
function addMonths(t, n) { const d = new Date(t); d.setMonth(d.getMonth() + n); return d.getTime(); }

/* panel açarı: kodda saxlanmır, ilk dəfə avtomatik yaradılır */
function panelKey() {
  const p = PropertiesService.getScriptProperties();
  let k = p.getProperty('PANEL_KEY');
  if (!k) { k = Utilities.getUuid().replace(/-/g, ''); p.setProperty('PANEL_KEY', k); }
  return k;
}

function usersSheet() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('İstifadəçilər');
  if (!s) { s = ss.insertSheet('İstifadəçilər'); s.appendRow(UL); s.getRange(1, 1, 1, UL.length).setFontWeight('bold'); s.setFrozenRows(1); s.setColumnWidth(1, 110); }
  return s;
}
function allRows(s) {
  const n = s.getLastRow() - 1;
  return n > 0 ? s.getRange(2, 1, n, UL.length).getValues() : [];
}
function findRow(v, tel) { for (let i = 0; i < v.length; i++) if (cleanTel(v[i][0]) === tel) return i; return -1; }
const toObj = r => ({ tel: cleanTel(r[0]), ad: String(r[1]), soyad: String(r[2]), qeyd: ms(r[3]), son: ms(r[4]), blok: isTrue(r[5]), ab: String(r[6] || ''), pay: ms(r[7]), tch: isTrue(r[8]), ex: isTrue(r[9]) });

/* qayda: bloklanıbmı? (lazım olsa sətri yeniləyir) */
function ruleBlocked(r, now) {
  if (isTrue(r[8])) return false;                                   /* müəllim */
  const pay = ms(r[7]);
  if (pay && now < addMonths(pay, PAY_MONTHS)) return isTrue(r[5]) && r[6] === 'əl'; /* ödəniş edib: yalnız əl ilə blok */
  if (isTrue(r[5])) return true;
  if (isTrue(r[9])) return false;                                   /* istisna */
  if (now - ms(r[3]) >= TRIAL_DAYS * DAY) { r[5] = true; r[6] = pay ? 'müddət' : 'avto'; return true; }
  return false;
}
function writeRow(s, i, r) {
  s.getRange(i + 2, 1, 1, UL.length).setValues([r]);
  s.getRange(i + 2, 1).setNumberFormat('@');
  s.getRange(i + 2, 4, 1, 2).setNumberFormat(DF); s.getRange(i + 2, 8).setNumberFormat(DF);
}
const dt = t => t ? new Date(t) : '';

/* şagird: qeydiyyat + giriş yoxlaması */
function ping(o, now) {
  const tel = cleanTel(o.tel); if (tel.length !== 9) return { ok: false, error: 'tel' };
  const s = usersSheet(), v = allRows(s); let i = findRow(v, tel), r;
  if (i < 0) {
    const dev = String(o.dev || '').slice(0, 20);
    const j = dev ? v.findIndex(x => String(x[10]) === dev) : -1;   /* eyni cihaz, yeni nömrə: köhnə hesabın vəziyyəti keçir */
    const same = j >= 0 ? v[j] : null;
    r = ["'" + tel, String(o.ad || '').slice(0, 60), String(o.soyad || '').slice(0, 60), dt(same ? ms(same[3]) : (+o.qeyd && +o.qeyd <= now ? +o.qeyd : now)), new Date(now),
      same ? isTrue(same[5]) : false, same ? String(same[6] || '') : '', same ? same[7] : '', false, false, dev];
    i = v.length; v.push(r);
  } else { r = v[i]; r[4] = new Date(now); if (!r[10] && o.dev) r[10] = String(o.dev).slice(0, 20); r[0] = "'" + tel; }
  const b = ruleBlocked(r, now);
  writeRow(s, i, r);
  dvNote_(tel, o.dev, o.dn, now); /* 📱 cihaz */
  return { ok: true, blok: b, tch: isTrue(r[8]) };
}

/* müəllim əməliyyatları */
function setUser(o, now) {
  const s = usersSheet(), v = allRows(s), i = findRow(v, cleanTel(o.tel));
  if (i < 0) return { ok: false, error: 'tel' };
  const r = v[i]; r[0] = "'" + cleanTel(r[0]);
  if (o.a === 'del') { s.deleteRow(i + 2); return { ok: true }; }
  if (o.a === 'pay') { if (o.v) { r[7] = new Date(now); r[5] = false; r[6] = ''; } else { r[7] = ''; } }
  if (o.a === 'block') { r[5] = !!o.v; r[6] = o.v ? 'əl' : ''; if (!o.v) r[9] = true; }
  if (o.a === 'tch') { r[8] = !!o.v; if (o.v) { r[5] = false; r[6] = ''; } }
  if (o.a === 'ren') { r[1] = String(o.ad || r[1]).slice(0, 60); r[2] = String(o.soyad || r[2]).slice(0, 60); }
  writeRow(s, i, r);
  return { ok: true };
}
function listUsers(now) {
  const s = usersSheet(), v = allRows(s);
  let ch = false;
  v.forEach((r, i) => { const was = isTrue(r[5]); ruleBlocked(r, now); if (isTrue(r[5]) !== was) { r[0] = "'" + cleanTel(r[0]); writeRow(s, i, r); ch = true; } });
  const dm = dvMap_(); /* 📱 */
  return v.filter(r => cleanTel(r[0])).map(toObj).map(u => Object.assign(u, dm[u.tel] || {}));
}


/* ✍️ Təklif və rəylər: "Rəylər" vərəqi. Status: Yeni / Oxundu / Cavablandı / Arxiv.
   "Şagird gizlədib" = şagird öz tərəfində silib (müəllim yenə görür). ID = telefon|vaxt. */
const FL = ['Tarix', 'Telefon', 'Ad', 'Soyad', 'Rəy', 'Status', 'Cavab', 'Cavab tarixi', 'Şagird gizlədib', 'ID'];
function fbSheet() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('Rəylər');
  if (!s) { s = ss.insertSheet('Rəylər'); s.appendRow(FL); s.getRange(1, 1, 1, FL.length).setFontWeight('bold'); s.setFrozenRows(1); s.setColumnWidth(5, 420); s.setColumnWidth(7, 360); }
  return s;
}
function fbAll() {
  const s = fbSheet(), n = s.getLastRow() - 1;
  return { s, v: n > 0 ? s.getRange(2, 1, n, FL.length).getValues() : [] };
}
const fbObj = r => ({ id: String(r[9]), at: ms(r[0]), tel: cleanTel(r[1]), ad: String(r[2]), soyad: String(r[3]), x: String(r[4]), st: String(r[5] || 'Yeni'), rep: String(r[6] || ''), rat: ms(r[7]), hid: isTrue(r[8]) });
function fbRows() { return fbAll().v.filter(r => r[9] !== '').map(fbObj); }
function fbAdd(o, now) {
  const x = String(o.x || '').trim().slice(0, 2000), tel = cleanTel(o.tel);
  if (!x) return { ok: false, error: 'x' };
  const at = +o.at && +o.at <= now ? +o.at : now, id = tel + '|' + at;
  if (fbRows().some(f => f.id === id)) return { ok: true, id };
  const s = fbSheet();
  s.appendRow([new Date(at), "'" + tel, String(o.ad || '').slice(0, 60), String(o.soyad || '').slice(0, 60), x, 'Yeni', '', '', false, "'" + id]);
  s.getRange(s.getLastRow(), 1).setNumberFormat(DF);
  return { ok: true, id };
}
function fbSet(o, now) {
  const { s, v } = fbAll(), i = v.findIndex(r => String(r[9]) === String(o.id));
  if (i < 0) return { ok: false, error: 'id' };
  const row = i + 2;
  if (o.del) { s.deleteRow(row); return { ok: true }; }
  if (typeof o.rep === 'string') {
    const t = o.rep.trim().slice(0, 2000);
    s.getRange(row, 6, 1, 3).setValues([[t ? 'Cavablandı' : 'Oxundu', t, t ? new Date(now) : '']]);
    s.getRange(row, 8).setNumberFormat(DF);
    return { ok: true };
  }
  if (['Yeni', 'Oxundu', 'Cavablandı', 'Arxiv'].indexOf(o.st) >= 0) { s.getRange(row, 6).setValue(o.st); return { ok: true }; }
  return { ok: false, error: 'st' };
}

/* tətbiq ayarları və lüğət (S&E redaktoru): "Ayarlar" vərəqində */
function cfgSheet() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('Ayarlar');
  if (!s) { s = ss.insertSheet('Ayarlar'); s.getRange('A1').setValue('Bu vərəqə toxunmayın: tətbiqin ayarları və lüğəti burada saxlanılır.'); }
  return s;
}
function cfgGet() { const v = cfgSheet().getRange('B1').getValue(); try { return v ? JSON.parse(v) : {}; } catch (e) { return {}; } }
function cfgSet(c) { cfgSheet().getRange('B1').setNumberFormat('@').setValue(JSON.stringify(c)); }
function vocabGet() {
  const s = cfgSheet(), n = +s.getRange('C1').getValue() || 0;
  if (!n) return '';
  return s.getRange(2, 3, n, 1).getValues().map(r => String(r[0])).join('');
}
function vocabSet(txt) {
  const s = cfgSheet(), parts = [];
  for (let i = 0; i < txt.length; i += 45000) parts.push([txt.slice(i, i + 45000)]);
  const old = +s.getRange('C1').getValue() || 0;
  if (old > parts.length) s.getRange(2 + parts.length, 3, old - parts.length, 1).clearContent();
  s.getRange(2, 3, parts.length, 1).setNumberFormat('@').setValues(parts);
  s.getRange('C1').setValue(parts.length);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.a === 'cfg') return out({ ok: true, cfg: cfgGet() });
  if (p.a === 'vocab') { const c = cfgGet(); return out({ ok: true, ver: c.ver || 0, vocab: vocabGet() }); }
  return out({ ok: true, app: 'MİQ və Sertifikasiya Hazırlığı' });
}
function doPost(e) {
  { let o0 = {}; try { o0 = JSON.parse(e.postData.contents || '{}'); } catch (x) {} if (o0.a === 'pget' || o0.a === 'pset') return out(psHandle(o0)); } /* ☁️ irəliləyiş: öz kilidi var */
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    let o = {}; try { o = JSON.parse(e.postData.contents || '{}'); } catch (x) {}
    const now = Date.now();
    if (o.a === 'ping') return out(ping(o, now));
    if (o.a === 'tlogin') { /* 🔐 5 səhv cəhddən sonra 15 dəqiqə bağlanır */
      const c = CacheService.getScriptCache(), n = +(c.get('tl_fail') || 0);
      if (!ADMIN_PASS || ADMIN_PASS === 'BURAYA-PANEL-SIFRESI') return out({ ok: false, error: 'nopass' });
      if (n >= 5) return out({ ok: false, error: 'wait' });
      if (String(o.p || '').trim().toUpperCase() !== ADMIN_PASS.trim().toUpperCase()) { c.put('tl_fail', String(n + 1), 900); return out({ ok: false, error: 'pass' }); }
      c.remove('tl_fail'); return out({ ok: true, key: panelKey() });
    }
    if (o.a === 'meren') {                                   /* şagird öz adını düzəldir (nömrəni yox) */
      const s = usersSheet(), v = allRows(s), i = findRow(v, cleanTel(o.tel));
      if (i < 0) return out({ ok: false, error: 'tel' });
      const r = v[i]; if (r[10] && String(r[10]) !== String(o.dev || '')) return out({ ok: false, error: 'dev' });
      r[0] = "'" + cleanTel(r[0]); r[1] = String(o.ad || r[1]).slice(0, 60); r[2] = String(o.soyad || r[2]).slice(0, 60); writeRow(s, i, r);
      return out({ ok: true });
    }
    if (o.a === 'fb') return out(fbAdd(o, now));            /* şagirdin rəyi */
    if (o.a === 'fbmine') {                                  /* şagird: öz rəyləri və müəllimin cavabları */
      const tel = cleanTel(o.tel); if (!tel) return out({ ok: true, rows: [] });
      return out({ ok: true, rows: fbRows().filter(f => f.tel === tel && !f.hid).map(f => ({ id: f.id, at: f.at, x: f.x, rep: f.rep, rat: f.rat })) });
    }
    if (o.a === 'fbhide') {                                  /* şagird öz tərəfində silir: müəllimdə qalır */
      const { s, v } = fbAll(), i = v.findIndex(r => String(r[9]) === String(o.id) && cleanTel(r[1]) === cleanTel(o.tel));
      if (i < 0) return out({ ok: false, error: 'id' });
      s.getRange(i + 2, 9).setValue(true); return out({ ok: true });
    }
    if (String(o.key || '') !== panelKey()) return out({ ok: false, error: 'key' });
    if (o.a === 'fbs') return out({ ok: true, rows: fbRows() });
    if (o.a === 'fbset') return out(fbSet(o, now));
    if (o.a === 'users') return out({ ok: true, rows: listUsers(now) });
    if (['pay', 'block', 'tch', 'ren', 'del'].indexOf(o.a) >= 0) return out(setUser(o, now));
    if (o.a === 'cfg') { const c = cfgGet(); if (typeof o.link === 'string') c.link = o.link.slice(0, 300);
      if (o.note !== undefined) c.note = o.note ? { on: o.note.on !== false, t: String(o.note.t || '').slice(0, 40), x: String(o.note.x || '').slice(0, 600), at: +o.note.at || now } : null;
      cfgSet(c); return out({ ok: true }); }
    if (o.a === 'vocab') { const t = String(o.vocab || ''); JSON.parse(t); vocabSet(t); const c = cfgGet(); c.ver = now; cfgSet(c); return out({ ok: true, ver: now }); }
    return out({ ok: false, error: 'a' });
  } catch (x) { return out({ ok: false, error: String(x && x.message || x) }); }
  finally { lock.releaseLock(); }
}

/* ☁️ İRƏLİLƏYİŞ — şagirdin işlədiyi materiallar telefon nömrəsinə bağlı "İrəliləyiş" vərəqində saxlanır.
   Şagird başqa telefonda / kompüterdə eyni nömrə ilə açanda hamısı geri gəlir. Bu vərəqi əl ilə dəyişməyin.
   Hər açar ayrıca saxlanır: {v: dəyər, t: serverə yazılma vaxtı}. Bir cihaz köhnə versiya üzərində yazırsa,
   heç nə itməsin deyə iki versiya birləşdirilir. */
const PS_NAME = 'İrəliləyiş', PS_CH = 45000, PS_MAXK = 300, PS_MAXV = 1500000;
function psSheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(PS_NAME);
  if (!s) {
    s = ss.insertSheet(PS_NAME);
    s.getRange(1, 1, 1, 4).setValues([['Nömrə', 'Yenilənib', 'Ölçü (KB)', 'Məlumat (avtomatik — dəyişməyin)']]).setFontWeight('bold');
    s.setFrozenRows(1); s.setColumnWidth(1, 110); s.setColumnWidth(2, 140);
  }
  return s;
}
function psRow_(s, tel) {
  if (s.getLastRow() < 2) return 0;
  const c = s.getRange(2, 1, s.getLastRow() - 1, 1).createTextFinder(tel).matchEntireCell(true).findNext();
  return c ? c.getRow() : 0;
}
function psLoad_(s, r) {
  if (!r) return {};
  const w = s.getLastColumn(); if (w < 4) return {};
  const b64 = s.getRange(r, 4, 1, w - 3).getValues()[0].map(x => String(x || '').replace(/^~/, '')).join('');
  if (!b64) return {};
  try { return JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip')).getDataAsString('UTF-8')) || {}; }
  catch (e) { console.error('ps load: ' + e); return null; }
}
function psSave_(s, r, tel, map) {
  const json = JSON.stringify(map);
  const b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(json, 'application/json', 'p.json')).getBytes());
  const ch = []; for (let i = 0; i < b64.length; i += PS_CH) ch.push('~' + b64.slice(i, i + PS_CH));
  const row = ["'" + tel, new Date(), Math.round(json.length / 1024)].concat(ch);
  const put = r => {
    const old = Math.max(s.getLastColumn() - 3, 0);
    if (old > ch.length) s.getRange(r, 4 + ch.length, 1, old - ch.length).clearContent();
    s.getRange(r, 1, 1, row.length).setValues([row]);
    s.getRange(r, 2).setNumberFormat('dd.MM.yyyy HH:mm');
  };
  if (r) return put(r);
  const lock = LockService.getScriptLock(); lock.waitLock(15000); /* yeni sətir: iki şagird eyni sətrə yazmasın */
  try { put(psRow_(s, tel) || s.getLastRow() + 1); SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
}
const psIsO_ = x => x && typeof x === 'object' && !Array.isArray(x);
function psDeep_(a, b) { /* a üstündür; heç nə itmir: siyahılar birləşir, rəqəmlərdən böyüyü, true qalır */
  if (a === undefined || a === null) return b;
  if (b === undefined || b === null) return a;
  if (Array.isArray(a) && Array.isArray(b)) { const seen = new Set(a.map(x => JSON.stringify(x))), o = a.slice(); b.forEach(x => { const k = JSON.stringify(x); if (!seen.has(k)) { seen.add(k); o.push(x); } }); return o; }
  if (psIsO_(a) && psIsO_(b)) { const o = Object.assign({}, b, a); for (const k in a) if (k in b) o[k] = psDeep_(a[k], b[k]); return o; }
  if (typeof a === 'number' && typeof b === 'number') return Math.max(a, b);
  if (typeof a === 'boolean' && typeof b === 'boolean') return a || b;
  return a;
}
function psMergeRaw_(a, b) { try { return JSON.stringify(psDeep_(JSON.parse(a), JSON.parse(b))); } catch (e) { return a; } }
function psHandle(o) {
  const tel = String(o.tel || '').replace(/\D/g, '');
  if (tel.length !== 9 || !psUserOK_(tel)) return { ok: false, error: 'user' };
  const s = psSheet_(), r = psRow_(s, tel), map = psLoad_(s, r);
  if (map === null) return { ok: false, error: 'read' };
  if (o.a === 'pget') return { ok: true, d: map };
  /* pset: d = {açar: {v: dəyər, b: cihazın son gördüyü vaxt}} */
  const d = psIsO_(o.d) ? o.d : {}, now = Date.now(), res = {};
  Object.keys(d).slice(0, PS_MAXK).forEach(k => {
    const x = d[k]; if (!x || typeof x.v !== 'string' || x.v.length > PS_MAXV) return;
    const cur = map[k];
    if (!cur || !cur.t || (+x.b || 0) >= cur.t) { map[k] = { v: x.v, t: now }; res[k] = now; }
    else if (cur.v === x.v) { res[k] = cur.t; }
    else { map[k] = { v: psMergeRaw_(x.v, cur.v), t: now }; res[k] = -1; } /* başqa cihaz yeniləyib: birləşdirildi */
  });
  if (Object.keys(res).length) psSave_(s, r, tel, map);
  return { ok: true, t: res };
}

function psUserOK_(tel) { tel = cleanTel(tel); return tel.length === 9 && findRow(allRows(usersSheet()), tel) >= 0; }

/* 📱 CİHAZLAR — hər nömrənin hansı cihazlarda açıldığı ("Cihazlar" vərəqi; Users vərəqinə toxunmur).
   Müəllim panelində hər istifadəçinin yanında cihaz sayı və son giriş vaxtı görünür: hesabını başqasına verəni görmək üçün. */
function dvSheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('Cihazlar');
  if (!s) { s = ss.insertSheet('Cihazlar'); s.appendRow(['Telefon', 'Cihaz sayı', 'Son giriş', 'Cihazlar (avtomatik — dəyişməyin)']); s.getRange(1, 1, 1, 4).setFontWeight('bold'); s.setFrozenRows(1); s.setColumnWidth(4, 420); }
  return s;
}
function dvNote_(tel, dev, dn, now) {
  tel = String(tel || '').replace(/\D/g, ''); dev = String(dev || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 24);
  if (!tel || !dev) return;
  try {
    const s = dvSheet_(), n = s.getLastRow();
    let r = 0, L = [];
    if (n > 1) { const c = s.getRange(2, 1, n - 1, 1).createTextFinder(tel).matchEntireCell(true).findNext(); if (c) r = c.getRow(); }
    if (r) { try { L = JSON.parse(s.getRange(r, 4).getValue() || '[]'); } catch (e) { L = []; } }
    if (!Array.isArray(L)) L = [];
    const x = L.find(d => d.i === dev), name = String(dn || '').slice(0, 40);
    if (x) { x.l = now; if (name) x.n = name; } else L.push({ i: dev, n: name, f: now, l: now });
    L.sort((a, b) => (b.l || 0) - (a.l || 0)); L = L.slice(0, 30);
    const row = ["'" + tel, L.length, new Date(now), JSON.stringify(L)];
    if (r) s.getRange(r, 1, 1, 4).setValues([row]); else s.appendRow(row);
    s.getRange(r || s.getLastRow(), 3).setNumberFormat('dd.MM.yyyy HH:mm');
  } catch (e) { console.error('dvNote: ' + e); }
}
function dvMap_() {
  const s = SpreadsheetApp.getActive().getSheetByName('Cihazlar'), m = {};
  if (!s || s.getLastRow() < 2) return m;
  s.getRange(2, 1, s.getLastRow() - 1, 4).getValues().forEach(r => {
    const tel = String(r[0]).replace(/^'/, '').replace(/\D/g, ''); if (!tel) return;
    let L = []; try { L = JSON.parse(r[3] || '[]'); } catch (e) {}
    if (!Array.isArray(L)) L = [];
    m[tel] = { dv: L.length, dvl: L.slice(0, 12).map(d => ({ n: d.n || '', f: d.f || 0, l: d.l || 0 })) };
  });
  return m;
}
