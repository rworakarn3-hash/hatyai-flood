/**
 * ระบบรายงานน้ำท่วมหาดใหญ่โดยภาคประชาชน — Google Apps Script backend
 * พัฒนาระบบโดย นพ.วรกานต์ รุ่งประเสริฐ
 *
 * ไฟล์ในโปรเจกต์: Code.gs (ไฟล์นี้), Index.html, appsscript.json
 * ครั้งแรก: เลือกฟังก์ชัน setup แล้วกด Run (อนุญาตสิทธิ์) → Deploy → New deployment → Web app
 *
 * ข้อมูลเก็บที่: Google Sheet (สร้างอัตโนมัติ ชื่อ "HatYai Flood DB") และรูปเต็มใน Google Drive โฟลเดอร์ "HatYai Flood Photos"
 * ตั้งค่าได้ที่ Project Settings → Script properties:
 *   ADMIN_PIN        รหัสผู้ดูแล (ค่าเริ่มต้น 1516)
 *   TMD_UID/TMD_UKEY คีย์ TMD Open Data ของคุณเอง (ไม่ใส่ = ใช้คีย์สาธารณะ api/api12345)
 *   TMD_NWP_TOKEN    (ไม่บังคับ) token ของ TMD NWP API → ใช้อุณหภูมิ/ความชื้นพยากรณ์รายชั่วโมงตามพิกัดหาดใหญ่
 *   SHARE_URL        (ไม่บังคับ) ลิงก์ที่จะพิมพ์บนการ์ดแชร์/QR เช่น โดเมนสั้น ไม่ใส่ = ใช้ลิงก์ /exec ของเว็บแอปนี้
 */

const SITE = 'https://hatyaicityclimate.org';
const COLS = ['reports', 'votes', 'photos', 'photofull', 'shelters', 'shelterEdits', 'site'];
const ADMIN_ONLY = ['shelters', 'shelterEdits', 'site'];
const LIVE_WINDOW_H = 14;            // reports / votes / photos sent to the page (page shows 12 h)
const KEEP_DAYS = 30;                // cleanup keeps this much history in the sheet
const HAT_YAI = {lat: 7.005, lon: 100.47};
const BOUNDS = {s: 6.6, n: 7.35, w: 100.2, e: 100.8};

/* ===================== web app ===================== */
function doGet(e) {
  // data for the fast static copy of the page (GitHub Pages): one small script with cameras, weather and reports
  if (e && e.parameter && e.parameter.api === 'boot') {
    let pulled = 'null'; try { pulled = pull(); } catch (x) {}
    const js = buildDataBlock_() + '\nvar BOOT_WX = ' + safeJson_(getWeather_()) + ';\nvar BOOT_PULL = ' + safeJson_(pulled) + ';\n';
    return ContentService.createTextOutput(js).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  const t = HtmlService.createTemplateFromFile('Index');
  t.dataBlock = buildDataBlock_();
  t.wxJson = safeJson_(getWeather_());
  return t.evaluate()
    .setTitle('ระบบรายงานน้ำท่วมหาดใหญ่โดยภาคประชาชน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** API for the static copy of the page: POST {fn, args} as text/plain (no CORS preflight). */
function doPost(e) {
  const API = {pull: pull, getDoc: getDoc, write: write, adminLogin: adminLogin};
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}'), fn = API[req.fn];
    if (!fn) throw new Error('INVALID: unknown call');
    out = JSON.stringify({ok: true, r: fn.apply(null, Array.isArray(req.args) ? req.args : [])});
  } catch (x) { out = JSON.stringify({ok: false, err: String((x && x.message) || x)}); }
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the editor: creates the sheet, photo folder, secrets and a daily cleanup trigger. */
function setup() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SHEET_ID')) {
    const ss = SpreadsheetApp.create('HatYai Flood DB');
    const sh = ss.getSheets()[0]; sh.setName('docs');
    sh.getRange(1, 1, 1, 5).setValues([['col', 'id', 'json', 'at', 'updated']]); sh.setFrozenRows(1);
    p.setProperty('SHEET_ID', ss.getId());
  }
  if (!p.getProperty('PHOTO_FOLDER')) p.setProperty('PHOTO_FOLDER', DriveApp.createFolder('HatYai Flood Photos').getId());
  if (!p.getProperty('SECRET')) p.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  if (!p.getProperty('ADMIN_PIN')) p.setProperty('ADMIN_PIN', '1516');
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'cleanup')) ScriptApp.newTrigger('cleanup').timeBased().everyDays(1).atHour(3).create();
  // keeps camera + weather data ready in the background so opening the page never waits for outside websites
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'warm')) ScriptApp.newTrigger('warm').timeBased().everyMinutes(5).create();
  warm();
  Logger.log('พร้อมใช้งาน · Sheet: ' + SpreadsheetApp.openById(p.getProperty('SHEET_ID')).getUrl());
}

/* ===================== storage (one row per document) ===================== */
function sheet_() { const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID'); if (!id) throw new Error('ยังไม่ได้รัน setup()'); return SpreadsheetApp.openById(id).getSheetByName('docs'); }
function readAll_() {
  const sh = sheet_(), n = sh.getLastRow() - 1; if (n < 1) return [];
  return sh.getRange(2, 1, n, 5).getValues().map((r, i) => ({row: i + 2, col: r[0], id: String(r[1]), json: r[2], at: +r[3] || 0}));
}
function findRow_(rows, col, id) { return rows.find(r => r.col === col && r.id === id); }

/** Everything the page needs, as one JSON string: {col: {id: data}} (full-size photos excluded). */
function pull() {
  const cached = getBig_('pull'); if (cached) return cached;
  const since = Date.now() - LIVE_WINDOW_H * 3600e3, out = {};
  COLS.forEach(c => { out[c] = {}; });
  readAll_().forEach(r => {
    if (r.col === 'photofull') return;
    if (['reports', 'votes', 'photos'].indexOf(r.col) >= 0 && r.at < since) return;
    try { out[r.col][r.id] = JSON.parse(r.json); } catch (e) {}
  });
  const s = JSON.stringify(out); putBig_('pull', s, 120); return s;
}

/** Only full-size photos are fetched one at a time (from Drive). */
function getDoc(col, id) {
  if (col !== 'photofull') return 'null';
  const r = findRow_(readAll_(), col, String(id)); if (!r) return 'null';
  const d = JSON.parse(r.json);
  try { const b = DriveApp.getFileById(d.fileId).getBlob(); return JSON.stringify({r: d.r, at: d.at, img: 'data:image/jpeg;base64,' + Utilities.base64Encode(b.getBytes())}); }
  catch (e) { return 'null'; }
}

function write(payload) {
  let q; try { q = JSON.parse(payload); } catch (e) { return err_('INVALID request'); }
  const op = q.op, col = q.col, cid = String(q.cid || '').slice(0, 40), admin = verifyToken_(q.tok);
  let id = q.id == null ? null : String(q.id), body = q.body || {};
  if (COLS.indexOf(col) < 0 || ['add', 'set', 'update', 'delete'].indexOf(op) < 0) return err_('INVALID collection/op');
  if (id && !/^[A-Za-z0-9_\-]{1,120}$/.test(id)) return err_('INVALID id');
  if (!/^u[a-z0-9]{6,40}$/.test(cid)) return err_('INVALID client');
  if (!admin && !rateOk_(cid)) return err_('RATE: ส่งบ่อยเกินไป รอสักครู่แล้วลองใหม่', 'invalid_argument');
  if (ADMIN_ONLY.indexOf(col) >= 0 && !admin) return err_('DENIED: ต้องเข้าสู่โหมดผู้ดูแล');

  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const rows = readAll_(), sh = sheet_(), now = Date.now();
    const cur = id ? findRow_(rows, col, id) : null, curData = cur ? safeParse_(cur.json) : null;
    const owns = d => d && d.by === cid;
    const reportOwner = rid => { const r = findRow_(rows, 'reports', String(rid)); const d = r && safeParse_(r.json); return d && d.by; };

    if (op === 'delete') {
      if (!cur) return ok_({id});
      if (!admin && !owns(curData) && !(col === 'votes' && id.endsWith('__' + cid))) return err_('DENIED: ลบได้เฉพาะของตัวเองหรือผู้ดูแล');
      if (col === 'photofull' && curData.fileId) { try { DriveApp.getFileById(curData.fileId).setTrashed(true); } catch (e) {} }
      sh.deleteRow(cur.row); clearBig_('pull'); return ok_({id});
    }

    if (op === 'update') { if (!cur) return err_('INVALID: ไม่มีเอกสารนี้'); if (!admin && !owns(curData)) return err_('DENIED'); body = Object.assign({}, curData, body); }
    else if (cur && !admin && !owns(curData) && col !== 'votes') return err_('DENIED: แก้ได้เฉพาะของตัวเอง');

    // per-collection rules
    if (col === 'reports') { try { body = cleanReport_(body, curData, cid, now, admin); } catch (e) { return err_(String(e.message || e)); } }
    if (col === 'votes') { if (!id || !id.endsWith('__' + cid)) return err_('DENIED vote'); body = {r: String(body.r || '').slice(0, 60), u: cid, v: ['still', 'down', 'wrong'].indexOf(body.v) >= 0 ? body.v : 'still', at: now}; }
    if (col === 'photos') { if (!admin && reportOwner(id) !== cid) return err_('DENIED photo');
      if (typeof body.thumb !== 'string' || !/^data:image\/jpeg;base64,/.test(body.thumb) || body.thumb.length > 45000) return err_('INVALID thumb');
      body = {r: id, thumb: body.thumb, at: +body.at || now, by: cid, lat: +body.lat, lon: +body.lon}; }
    if (col === 'photofull') { if (!admin && reportOwner(id) !== cid) return err_('DENIED photo');
      const m = /^data:image\/jpeg;base64,(.+)$/.exec(body.img || ''); if (!m || m[1].length > 400000) return err_('INVALID image');
      const folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('PHOTO_FOLDER'));
      const f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[1]), 'image/jpeg', id + '.jpg'));
      body = {r: id, fileId: f.getId(), at: +body.at || now, by: cid}; }
    if (col === 'shelters') body = {name: str_(body.name, 120), org: str_(body.org, 80), tel: (body.tel || []).slice(0, 5).map(x => str_(x, 30)), note: str_(body.note, 160),
      lat: +body.lat, lon: +body.lon, at: now, by: 'admin'};
    if (col === 'shelterEdits') body = {hidden: !!body.hidden, at: now, by: 'admin'};
    if (col === 'site') body = {text: str_(body.text, 300), src: str_(body.src, 60), on: body.on !== false, at: now, by: 'admin'};
    if (col !== 'reports' && col !== 'votes' && col !== 'photos' && col !== 'photofull' && body.by == null) body.by = cid;

    const json = JSON.stringify(body); if (json.length > 48000) return err_('INVALID: ข้อมูลใหญ่เกินไป');
    if (!id) id = 'r' + now.toString(36) + Math.random().toString(36).slice(2, 7);
    const at = +body.at || now;
    if (cur) sh.getRange(cur.row, 3, 1, 3).setValues([[json, at, new Date()]]);
    else sh.appendRow([col, id, json, at, new Date()]);
    clearBig_('pull');
    return ok_({id});
  } finally { lock.releaseLock(); }
}

function cleanReport_(b, cur, cid, now, admin) {
  const lat = +b.lat, lon = +b.lon, lv = Math.round(+b.level);
  if (!(lat > BOUNDS.s && lat < BOUNDS.n && lon > BOUNDS.w && lon < BOUNDS.e)) throw new Error('INVALID: ตำแหน่งอยู่นอกพื้นที่');
  if (!(lv >= 0 && lv <= 4)) throw new Error('INVALID level');
  return {lat, lon, road: str_(b.road, 80), level: lv, veh: str_(b.veh, 12), trend: ['up', 'flat', 'down', 'unk'].indexOf(b.trend) >= 0 ? b.trend : 'unk',
    tags: (b.tags || []).slice(0, 6).map(x => str_(x, 20)), note: str_(b.note, 100), by: cur ? cur.by : cid, at: cur ? cur.at : now,
    status: b.status === 'hidden' ? 'hidden' : 'ok', photo: !!b.photo, verified: admin ? !!b.verified : !!(cur && cur.verified)};
}

/* ===================== admin ===================== */
function adminLogin(pin) {
  const cache = CacheService.getScriptCache(), fails = +(cache.get('pinfail') || 0);
  if (fails >= 20) return JSON.stringify({err: 'ใส่รหัสผิดหลายครั้ง รอ 10 นาที'});
  const want = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || '1516';
  if (String(pin) !== String(want)) { cache.put('pinfail', String(fails + 1), 600); Utilities.sleep(800); return JSON.stringify({err: 'รหัสไม่ถูกต้อง'}); }
  const exp = Date.now() + 12 * 3600e3;
  return JSON.stringify({token: exp + '.' + sign_(String(exp))});
}
function verifyToken_(t) {
  if (!t || typeof t !== 'string' || t.indexOf('.') < 0) return false;
  const [exp, sig] = t.split('.'); return +exp > Date.now() && sig === sign_(exp);
}
function sign_(s) {
  const key = PropertiesService.getScriptProperties().getProperty('SECRET') || 'unset';
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, key)).replace(/=+$/, '');
}

/* ===================== live cameras (hatyaicityclimate.org) ===================== */
function buildDataBlock_() {
  const live = getCams_();
  const head = '/* cameras: ' + (live ? 'live from ' + SITE + '/api/flood/cams' : 'fallback snapshot 12 Oct 2026') + ' */\n'
    + 'const SNAP_AT = ' + JSON.stringify(live ? live.at : FALLBACK_AT) + ';\n'
    + "const NOTICE = '';\nconst SITE = " + JSON.stringify(SITE) + ';\n'
    + 'const SHARE_URL_OVERRIDE = ' + JSON.stringify(shareUrl_()) + ';\n'
    + '// code, title, lat, lon, lastImage, statusMsg, source, photo, area\n'
    + 'const CAMS = ' + safeJson_(live ? live.cams : FALLBACK_CAMS) + ';\n';
  return head + STATIC_JS;
}
/** Link printed on share cards / QR: Script property SHARE_URL (e.g. a short domain) or this web app's /exec URL. */
function shareUrl_() {
  const p = PropertiesService.getScriptProperties().getProperty('SHARE_URL'); if (p) return p.replace(/\/u\/\d+\//, '/').replace(/[?#].*$/, '');   // never hand out an account-specific /u/N/ link
  try { return ScriptApp.getService().getUrl().replace(/\/u\/\d+\//, '/').replace(/\/dev$/, '/exec'); } catch (e) { return ''; }
}
/** Fast path for page opens: cache → last saved copy → (only if nothing saved yet) fetch now. */
function getCams_() {
  const c = CacheService.getScriptCache().get('cams'); if (c) return JSON.parse(c);
  const last = safeParse_(PropertiesService.getScriptProperties().getProperty('CAMS_LAST'));
  if (last && Date.now() - (last.t || 0) < 30 * 60e3) return last;   // the warm trigger keeps this fresh
  return refreshCams_() || last;
}
function refreshCams_() {
  try {
    const res = UrlFetchApp.fetch(SITE + '/api/flood/cams', {muteHttpExceptions: true, headers: {accept: 'application/json'}});
    if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
    // real street/canal cameras only: skip radar/satellite/weather-map items (code R..) and items without coordinates
    const items = (JSON.parse(res.getContentText()).items || []).filter(x => x && x.enable == 1 && x.code && !/^R\d/.test(x.code) && x.location
      && x.location.latitude != null && x.location.longitude != null && Math.abs(+x.location.latitude - 7.0) < 0.6 && Math.abs(+x.location.longitude - 100.47) < 0.6);
    const cams = items.map(x => [String(x.code), String(x.title || x.code), +x.location.latitude, +x.location.longitude, String(x.atDate || ''), String(x.statusMsg || ''),
      String(x.sponsorName || ''), String(x.name || ''), AREA_BY_CODE[x.code] || 'hatyai']).filter(r => isFinite(r[2]) && isFinite(r[3]) && r[4]);
    if (!cams.length) throw new Error('empty');
    const now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    // cameras published on the same image server but not (yet) in the API, e.g. the sukjai.net/cam list
    EXTRA_CAMS.forEach(e => { if (!cams.some(r => r[7] === e[7] || r[0] === e[0])) cams.push(e.slice(0, 4).concat([now, '', e[6], e[7], e[8]])); });
    const out = {at: now, t: Date.now(), cams}, s = JSON.stringify(out);
    CacheService.getScriptCache().put('cams', s, 600);
    try { PropertiesService.getScriptProperties().setProperty('CAMS_LAST', s); } catch (e) {}
    return out;
  } catch (e) { console.warn('cams: ' + e); return null; }
}

/** Time-driven (every 5 min, installed by setup): refresh outside data so visitors never wait for it. */
function warm() {
  refreshCams_();
  const p = PropertiesService.getScriptProperties(), wl = safeParse_(p.getProperty('WX_LAST'));
  if (!wl || Date.now() - (wl._t || 0) > 25 * 60e3) refreshWeather_();
  try { pull(); } catch (e) {}
}

/* ===================== weather (Thai Meteorological Department) ===================== */
const TH_M = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
const TH_D = ['อา.', 'จ.', 'อ.', 'พ.', 'พฤ.', 'ศ.', 'ส.'];
const COND = ['', 'ท้องฟ้าแจ่มใส', 'มีเมฆบางส่วน', 'เมฆเป็นส่วนมาก', 'มีเมฆมาก', 'ฝนตกเล็กน้อย', 'ฝนปานกลาง', 'ฝนตกหนัก', 'ฝนฟ้าคะนอง', 'อากาศหนาวจัด', 'อากาศหนาว', 'อากาศเย็น', 'อากาศร้อนจัด'];
function getWeather_() {
  const c = CacheService.getScriptCache().get('wx'); if (c) return JSON.parse(c);
  const last = safeParse_(PropertiesService.getScriptProperties().getProperty('WX_LAST'));
  if (last && Date.now() - (last._t || 0) < 3 * 3600e3) return last;   // a slightly older forecast beats a 5-second wait
  return refreshWeather_() || last;
}
function refreshWeather_() {
  const cache = CacheService.getScriptCache();
  const p = PropertiesService.getScriptProperties(), uid = p.getProperty('TMD_UID') || 'api', ukey = p.getProperty('TMD_UKEY') || 'api12345';
  try {
    // both TMD calls in parallel
    const [obs, fc] = UrlFetchApp.fetchAll(['Weather3Hours', 'WeatherForecast7Days'].map(api => ({url: 'https://data.tmd.go.th/api/' + api + '/V2/?uid=' + uid + '&ukey=' + ukey + '&format=json', muteHttpExceptions: true})))
      .map(r => JSON.parse(r.getContentText()));
    const st = [].concat((obs.Stations && obs.Stations.Station) || []);
    const S = n => st.find(s => String(s.WmoStationNumber) === n);
    const num = v => Math.round(parseFloat(v) * 10) / 10;
    const hy = S('48569'); if (!hy) throw new Error('no Hat Yai station');
    const o = hy.Observation, m = /^(\d+)\/(\d+)\/(\d+) (\d+):(\d+)/.exec(o.DateTime);   // MM/DD/YYYY HH:mm
    const obsAt = m ? `${m[4]}:${m[5]} น. ${+m[2]} ${TH_M[+m[1] - 1]} ${(+m[3] + 543) % 100}` : '';
    const wx = {obsAt, src: 'TMD', now: {name: 'สนามบินหาดใหญ่', t: num(o.AirTemperature), rh: Math.round(+o.RelativeHumidity), rain24: num(o.Rainfall24Hr), rain1: num(o.Rainfall)},
      stations: [['48571', 'คอหงส์'], ['48574', 'สะเดา'], ['48568', 'สงขลา']].map(([n, name]) => { const s = S(n); return s ? {name, t: num(s.Observation.AirTemperature), rain24: num(s.Observation.Rainfall24Hr)} : null; }).filter(Boolean),
      days: []};
    const prov = [].concat((fc.Provinces && fc.Provinces.Province) || []).find(x => x.ProvinceNameThai === 'สงขลา');
    if (prov) { const f = prov.SevenDaysForecast, today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd');
      wx.days = f.ForecastDate.map((d, i) => { const [dd, mm, yy] = d.split('/').map(Number); const dt = new Date(yy, mm - 1, dd);
        return {key: '' + yy + String(mm).padStart(2, '0') + String(dd).padStart(2, '0'), d: TH_D[dt.getDay()], date: dd, lo: Math.round(+f.MinimumTemperature[i]), hi: Math.round(+f.MaximumTemperature[i]),
          pop: Math.round(+f.PercentRainCover[i]), desc: String(f.DescriptionThai[i] || '')}; })
        .filter(x => x.key >= today).sort((a, b) => a.key < b.key ? -1 : 1).slice(0, 7); }
    const nwp = p.getProperty('TMD_NWP_TOKEN');
    if (nwp) { try {
      const r = JSON.parse(UrlFetchApp.fetch(`https://data.tmd.go.th/nwpapi/v1/forecast/location/hourly/at?lat=${HAT_YAI.lat}&lon=${HAT_YAI.lon}&fields=tc,rh,cond,rain&duration=1`,
        {headers: {accept: 'application/json', authorization: 'Bearer ' + nwp}, muteHttpExceptions: true}).getContentText());
      const d = r.WeatherForcasts[0].forecasts[0].data;
      wx.now = Object.assign(wx.now, {name: 'หาดใหญ่ (พยากรณ์ NWP รายชั่วโมง)', t: num(d.tc), rh: Math.round(d.rh)});
      if (wx.days[0] && COND[d.cond]) wx.days[0].desc = COND[d.cond];
    } catch (e) { console.warn('nwp: ' + e); } }
    wx._t = Date.now(); cache.put('wx', JSON.stringify(wx), 1800); p.setProperty('WX_LAST', JSON.stringify(wx));
    return wx;
  } catch (e) { console.warn('weather: ' + e); return null; }
}

/* ===================== housekeeping ===================== */
function cleanup() {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = sheet_(), rows = readAll_(), old = Date.now() - KEEP_DAYS * 86400e3;
    rows.filter(r => ['reports', 'votes', 'photos', 'photofull'].indexOf(r.col) >= 0 && r.at && r.at < old).reverse().forEach(r => {
      if (r.col === 'photofull') { const d = safeParse_(r.json); if (d && d.fileId) try { DriveApp.getFileById(d.fileId).setTrashed(true); } catch (e) {} }
      sh.deleteRow(r.row);
    });
    clearBig_('pull');
  } finally { lock.releaseLock(); }
}

/* ===================== helpers ===================== */
function ok_(o) { return JSON.stringify(Object.assign({ok: true}, o)); }
function err_(m, code) { return JSON.stringify({err: m, code: code || (/^DENIED|^INVALID/.test(m) ? 'invalid_argument' : 'unavailable')}); }
function str_(v, n) { return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, n); }
function safeParse_(s) { try { return JSON.parse(s); } catch (e) { return null; } }
function safeJson_(o) { return JSON.stringify(o).replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029'); }
function rateOk_(cid) { const c = CacheService.getScriptCache(), k = 'rl_' + cid, n = +(c.get(k) || 0); if (n >= 40) return false; c.put(k, String(n + 1), 600); return true; }
function putBig_(key, s, ttl) { const c = CacheService.getScriptCache(), size = 90000, parts = {}; let n = 0;
  for (let i = 0; i < s.length; i += size) parts[key + '_' + n++] = s.slice(i, i + size);
  if (n > 90) return; parts[key + '_n'] = String(n); try { c.putAll(parts, ttl); } catch (e) {} }
function getBig_(key) { const c = CacheService.getScriptCache(), n = +(c.get(key + '_n') || 0); if (!n) return null;
  const keys = []; for (let i = 0; i < n; i++) keys.push(key + '_' + i); const got = c.getAll(keys); let s = '';
  for (const k of keys) { if (got[k] == null) return null; s += got[k]; } return s; }
function clearBig_(key) { CacheService.getScriptCache().remove(key + '_n'); }

/* ===================== static map data ===================== */
const FALLBACK_AT = '2026-10-12 11:36:43';
const FALLBACK_CAMS = [["S1", "สะพานม่วงก็อง", 6.823193, 100.438272, "2026-10-12 11:36:25", "", "SCCCRN", "muangkong", "hatyai"], ["S2", "สะพานบางศาลา", 6.931207, 100.439536, "2026-10-12 11:35:01", "", "SCCCRN", "bangsala", "hatyai"], ["S3", "ควนขี้แรด ต.พะตง", 6.824277, 100.542231, "2026-10-12 11:35:39", "", "SCCCRN", "khuankhirat", "patong"], ["S4", "ท้ายวัดควนเนียง ต.พะตง", 6.834999, 100.498797, "2026-10-12 11:35:50", "", "SCCCRN", "patongkn", "patong"], ["S5", "ป้อม ตร.ตลาดทุ่งลุง", 6.855114, 100.470618, "2026-10-12 11:35:38", "", "SCCCRN", "tunglung", "patong"], ["S6", "สี่แยกคลองหวะ", 6.975333, 100.479899, "2026-10-12 11:36:13", "", "SCCCRN", "klongwha", "hatyai"], ["S7", "สะพานประชาอุทิศ (คลองหวะ)", 6.987275, 100.474085, "2026-10-12 11:34:27", "", "กรมชลประทาน", "klongwha2", "hatyai"], ["S8", "สะพานข้างที่ว่าการ อ.หาดใหญ่", 7.002231, 100.455775, "2026-08-21 09:00:07", "กล้องขัดข้อง", "สำนักงานทรัพยากรน้ำภาค 8", "hatyainai", "hatyai"], ["S9", "ถนนราษฎร์ยินดี (เหล็กใต้ 30 เมตร)", 7.006898, 100.480973, "2026-10-12 11:35:13", "", "SCCCRN", "road30m", "hatyai"], ["S10", "คลองอิตำ ต.ทุ่งตำเสา", 6.975383, 100.367861, "2026-09-30 00:19:49", "", "SCCCRN", "klongetum", "kuanlang"], ["S11", "คลองนนท์ ต.ควนลัง", 6.983757, 100.419825, "2026-10-12 11:35:30", "", "SCCCRN", "klongnon", "kuanlang"], ["S12", "บางหัก", 7.011591, 100.453763, "2026-10-01 19:51:46", "ไฟฟ้าขัดข้อง", "SCCCRN", "banghak", "hatyai"], ["S13", "อุโมงค์กัลยาณมิตร", 7.017221, 100.466135, "2026-10-12 11:34:56", "", "SCCCRN", "kalyanamit", "hatyai"], ["S14", "บ้านควนหิน ต.พะวง", 7.100897, 100.56769, "2026-10-12 11:34:57", "", "SCCCRN", "kuanhin", "pawong"], ["I1", "ประตูระบายน้ำ คลองอู่ตะเภา", 6.988241, 100.461817, "2026-04-10 20:36:31", "ไฟฟ้าขัดข้อง", "กรมชลประทาน", "utapao", "r1"], ["I2", "ประตูระบายน้ำ คลอง ร.1", 6.985938, 100.458965, "2026-10-12 11:36:06", "", "กรมชลประทาน", "nhakaun", "r1"], ["I3", "ประตูระบายน้ำคลองชลประทาน", 7.008567, 100.44212, "2026-10-12 11:36:20", "", "กรมชลประทาน", "khlongs1r1", "r1"], ["I4", "ประตูระบายน้ำบ้านบางหยี (หน้า)", 7.127922, 100.433471, "2026-10-12 11:36:33", "", "กรมชลประทาน", "bangyheeus", "r1"], ["I5", "ประตูระบายน้ำบ้านบางหยี (หลัง)", 7.12786, 100.43395, "2026-10-12 11:35:55", "", "กรมชลประทาน", "bangyheeds", "r1"], ["K01", "อ่างเก็บน้ำแก้มลิงคลองเรียน", 6.995112, 100.512112, "2026-10-12 11:36:43", "", "เทศบาลเมืองคอหงส์", "khgamling", "korhong"], ["K02", "ต้นคลองเรียน–คลอง ร.6", 6.994829, 100.501188, "2026-10-12 11:36:39", "", "เทศบาลเมืองคอหงส์", "khklongr6", "korhong"], ["K03", "ถนนปลักธง (ป้อม อปพร.)", 6.986162, 100.494535, "2026-03-27 03:10:40", "ไฟฟ้าขัดข้อง", "เทศบาลเมืองคอหงส์", "khsoi1plugtong", "korhong"], ["K04", "ท้ายซอยถนนปลักธง", 6.982545, 100.498096, "2026-10-12 11:35:52", "", "เทศบาลเมืองคอหงส์", "khtaisoi", "korhong"], ["K05", "ปากซอยนายหลง ปลักธง 1", 6.973453, 100.502006, "2026-10-12 11:36:28", "", "เทศบาลเมืองคอหงส์", "khnailhong", "korhong"], ["K06", "สะพานซอยถวิลบูรณะ", 6.972897, 100.496858, "2026-10-12 11:36:16", "", "เทศบาลเมืองคอหงส์", "khtwinburana", "korhong"], ["K07", "ถนนพลพิชัย (หน้า ม.หาดใหญ่)", 6.98162, 100.469922, "2026-10-12 11:36:41", "", "เทศบาลเมืองคอหงส์", "khpholpichaihu", "korhong"], ["K08", "ริมคลอง ร.5 ท้ายซอยกาญจนวนิช 6", 7.032156, 100.490662, "2026-03-27 03:10:39", "ไฟฟ้าขัดข้อง", "เทศบาลเมืองคอหงส์", "khkarnjanavanij6", "korhong"], ["K09", "ริมคลอง ร.5 ท้ายซอย ม.ปริญญา", 7.041927, 100.493368, "2026-10-12 11:36:04", "", "เทศบาลเมืองคอหงส์", "khparinyavillage", "korhong"], ["K10", "ถนนอนุสรณ์ปู่คง", 7.044717, 100.500498, "2026-10-12 11:36:30", "", "เทศบาลเมืองคอหงส์", "khpookong", "korhong"], ["DW01", "สะพานพะตงทุ่งลาน", 6.856416, 100.464361, "2026-10-12 11:36:36", "", "สำนักงานทรัพยากรน้ำภาค 8", "tunglan", "patong"], ["S16", "แยกถนนธรรมนูญวิถี", 7.004864, 100.492375, "2026-10-12 17:48:45", "", "hatyaicityclimate", "thammanoonvithi", "hatyai"], ["X01", "ห้าแยกเกาะยอ (แยกน้ำกระจาย)", 7.1313, 100.57846, "2026-10-12 11:36:43", "", "hatyaicityclimate · sukjai.net/cam", "yakkoyo", "pawong"]];
const AREA_BY_CODE = {"S1": "hatyai", "S2": "hatyai", "S3": "patong", "S4": "patong", "S5": "patong", "S6": "hatyai", "S7": "hatyai", "S8": "hatyai", "S9": "hatyai", "S10": "kuanlang", "S11": "kuanlang", "S12": "hatyai", "S13": "hatyai", "S14": "pawong", "I1": "r1", "I2": "r1", "I3": "r1", "I4": "r1", "I5": "r1", "K01": "korhong", "K02": "korhong", "K03": "korhong", "K04": "korhong", "K05": "korhong", "K06": "korhong", "K07": "korhong", "K08": "korhong", "K09": "korhong", "K10": "korhong", "DW01": "patong", "S16": "hatyai", "X01": "pawong"};
// [code, title, lat, lon, -, -, source, image name, area] — added when the live API does not list them
const EXTRA_CAMS = [["X01", "ห้าแยกเกาะยอ (แยกน้ำกระจาย)", 7.1313, 100.57846, "2026-10-12 11:36:43", "", "hatyaicityclimate · sukjai.net/cam", "yakkoyo", "pawong"]];
const STATIC_JS = "const AREAS = {hatyai:'เทศบาลนครหาดใหญ่ (ศูนย์กลางเมือง)', r1:'คลอง ร.1 และประตูระบายน้ำหลัก', korhong:'เทศบาลเมืองคอหงส์ (แก้มลิง & คลอง ร.5–6)', patong:'พะตง / ทุ่งลุง (ตอนบน)', kuanlang:'กลุ่มสองคลองควนลัง', pawong:'พะวง / ควนหิน'};\n// [upstream, downstream] from /api/flood/cams/flow\nconst EDGES = [['S5','S1'],['S1','S2'],['I1','S8'],['K03','S7'],['S8','S13'],['K02','S9'],['K01','K02'],['S3','S4'],['S4','S5'],['K07','I2'],['S7','I2'],['S5','S2'],['I2','I3'],['S11','I3'],['I3','I4'],['I4','I5'],['K06','S7'],['K05','K04'],['K04','K03'],['K02','K03'],['K08','K09'],['K10','K09'],['S8','S12'],['I2','I1'],['S6','K07'],['S10','S11'],['S2','I2']];\nconst ROUTES = [\n {name:'คลองอู่ตะเภา · สายหลักเข้าเมือง', note:'จากพะตง–ทุ่งลุง ผ่านม่วงก็อง บางศาลา ประตูระบายน้ำ เข้าตัวเมือง', seq:['S3','S4','S5','S1','S2','I2','I1','S8','S13']},\n {name:'ผันน้ำเลี่ยงเมือง · คลอง ร.1', note:'ประตูระบายน้ำผันน้ำเลี่ยงเมืองออกทะเลสาบสงขลาที่บ้านบางหยี', seq:['S2','I2','I3','I4','I5']},\n {name:'กลุ่มสองคลองควนลัง', note:'คลองอิตำ–คลองนนท์ ไหลลงคลองชลประทาน ร.1', seq:['S10','S11','I3','I4','I5']},\n {name:'คลองหวะ', note:'สี่แยกคลองหวะ–ถนนพลพิชัย ลงประตูระบายน้ำ ร.1', seq:['S6','K07','I2']},\n {name:'คอหงส์ · แก้มลิงคลองเรียน', note:'อ่างแก้มลิง → คลอง ร.6 → ปลักธง → คลองหวะ', seq:['K01','K02','K03','S7','I2']},\n {name:'คอหงส์ · ซอยถวิลบูรณะ', note:'ลงคลองหวะที่สะพานประชาอุทิศ', seq:['K06','S7','I2']},\n {name:'คอหงส์ · ปลักธง', note:'ซอยนายหลง → ท้ายซอยปลักธง → ป้อม อปพร.', seq:['K05','K04','K03','S7']},\n {name:'คลอง ร.6 → ถนน 30 เมตร', note:'น้ำจากคลองเรียนเข้าเขตเมืองฝั่งตะวันออก', seq:['K01','K02','S9']},\n {name:'คลอง ร.5', note:'ปู่คง และซอยกาญจนวนิช 6 รวมกันที่ ม.ปริญญา', seq:['K10','K09']},\n {name:'ปลายน้ำในเมือง', note:'ท้ายสะพานที่ว่าการอำเภอ → บางหัก', seq:['S8','S12']}\n];\nconst CALLOUTS = [\n {title:'น้ำจากสะเดา', sub:'คลองอู่ตะเภาตอนบน ไหลเหนือขึ้นสู่ตัวเมือง', at:[6.80,100.441], off:[-62,60], color:'var(--blue)'},\n {title:'น้ำจากพะตง', sub:'ควนขี้แรด–ควนเนียง → ทุ่งลุง', at:[6.828,100.522], off:[30,26], color:'var(--orange)'},\n {title:'น้ำจากควนลัง', sub:'คลองอิตำ–คลองนนท์ → คลอง ร.1 → ทะเลสาบ', at:[6.979,100.39], off:[-230,24], color:'var(--lilac)'},\n {title:'น้ำจากคอหงส์', sub:'แก้มลิงคลองเรียน → คลองหวะ / คลองเตย', at:[6.99,100.517], off:[40,-6], color:'var(--mint)'},\n {title:'น้ำออกสู่ทะเลสาบสงขลา', sub:'ผ่านประตูระบายน้ำบ้านบางหยี', at:[7.128,100.434], off:[-60,-80], color:'var(--blue)', out:true}\n];";
