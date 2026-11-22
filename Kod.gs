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
  return v.filter(r => cleanTel(r[0])).map(toObj);
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
    if (String(o.key || '') !== panelKey()) return out({ ok: false, error: 'key' });
    if (o.a === 'users') return out({ ok: true, rows: listUsers(now) });
    if (['pay', 'block', 'tch', 'ren', 'del'].indexOf(o.a) >= 0) return out(setUser(o, now));
    if (o.a === 'cfg') { const c = cfgGet(); if (typeof o.link === 'string') c.link = o.link.slice(0, 300); cfgSet(c); return out({ ok: true }); }
    if (o.a === 'vocab') { const t = String(o.vocab || ''); JSON.parse(t); vocabSet(t); const c = cfgGet(); c.ver = now; cfgSet(c); return out({ ok: true, ver: now }); }
    return out({ ok: false, error: 'a' });
  } catch (x) { return out({ ok: false, error: String(x && x.message || x) }); }
  finally { lock.releaseLock(); }
}
