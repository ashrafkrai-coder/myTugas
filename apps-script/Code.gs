/**
 * myTugas API (Google Apps Script Web App)
 *
 * Pasang:
 * 1. Buka projek Apps Script anda (Extensions > Apps Script dari Google Sheet "myTugas Database").
 *    Sheet dibuka melalui SHEET_ID di bawah, jadi projek mana pun boleh digunakan.
 * 2. Tampal kod ini ke dalam Code.gs dan simpan.
 * 3. Deploy > New deployment > Type: Web app
 *      Execute as: Me
 *      Who has access: Anyone
 *    Jika sudah pernah deploy: Deploy > Manage deployments > Edit > Version: New version.
 * 4. URL Web app (berakhir /exec) diletak dalam auth.js (ENDPOINT).
 *
 * Semua permintaan ialah POST dengan body JSON { action, ... }:
 *   senarai_guru            -> guru aktif yang belum ada akaun (untuk borang daftar)
 *   register {idGuru, username, password}
 *   login    {username, password}
 *   data     {session}      -> semua data aplikasi
 *   laporan  {session, tarikh, kod, kehadiran, kebersihan, isu, tindakan, catatan}
 *
 * Akaun disimpan dalam tab "Akaun". Untuk menyekat guru, tukar Status kepada "Tidak Aktif".
 * Untuk set semula kata laluan, padam baris akaun guru itu dan minta guru mendaftar semula.
 */

// ID Google Sheet "myTugas Database - SMK Taman Medan 2026"
const SHEET_ID = '1W7IcjJ6pkG05t2Vhj28FUe2bMyZKdzsS56NBPOMkal0';
const TZ = 'Asia/Kuala_Lumpur';
const HARI = ['Ahad', 'Isnin', 'Selasa', 'Rabu', 'Khamis', 'Jumaat', 'Sabtu'];

const AKAUN_TAB = 'Akaun';
const AKAUN_HEAD = ['Username', 'ID Guru', 'Nama Guru', 'Salt', 'Hash', 'Status', 'Dicipta Pada', 'Login Terakhir'];
const SESSION_DAYS = 30;
const HASH_ROUNDS = 1000;
const MAX_FAILS = 5;
const FAIL_WINDOW_SECONDS = 900;
const USERNAME_RE = /^[a-z0-9._]{3,30}$/;

let ss_;
function ss() { return ss_ || (ss_ = SpreadsheetApp.openById(SHEET_ID)); }

/* ---------- Pintu masuk ---------- */

function doGet() {
  return json_({ ok: true, app: 'myTugas API' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'Permintaan tidak sah.' });
  }
  try {
    switch (body.action) {
      case 'senarai_guru':
        return json_({ ok: true, guru: guruBelumDaftar_() });
      case 'register':
        return json_(withLock_(() => register_(body)));
      case 'login':
        return json_(login_(body));
      case 'data': {
        const me = requireSession_(body);
        return json_({ ok: true, saya: me.nama, data: readAll_() });
      }
      case 'laporan': {
        const me = requireSession_(body);
        return json_(withLock_(() => addLaporan_(body, me)));
      }
      default:
        return json_({ ok: false, error: 'Tindakan tidak dikenali.' });
    }
  } catch (err) {
    return json_({ ok: false, error: err.message || String(err), code: err.code || '' });
  }
}

/* ---------- Baca data aplikasi ---------- */

function guruAktif_() {
  return rows_('Guru')
    .filter(r => str_(r['ID Guru']) && str_(r['Nama Paparan']) && str_(r['Status'] || 'Aktif') === 'Aktif')
    .map(r => ({ id: str_(r['ID Guru']), nama: str_(r['Nama Paparan']) }));
}

function readAll_() {
  const guru = guruAktif_();

  const lokasi = rows_('Lokasi')
    .filter(r => r['Kod Lokasi'])
    .sort((a, b) => (Number(a['Susunan']) || 99) - (Number(b['Susunan']) || 99))
    .map(r => ({
      kod: str_(r['Kod Lokasi']),
      nama: str_(r['Nama Lokasi']),
      desc: str_(r['Keterangan']),
      ikon: str_(r['Ikon'])
    }));

  const jadual = rows_('Jadual')
    .filter(r => r['Hari'] && r['Kod Lokasi'] && r['Nama Guru'])
    .map(r => ({ hari: str_(r['Hari']), kod: str_(r['Kod Lokasi']), guru: str_(r['Nama Guru']) }));

  // Hanya pertukaran yang diluluskan, dari 7 hari lepas hingga 60 hari akan datang
  const now = new Date();
  const from = new Date(now.getTime() - 7 * 864e5);
  const to = new Date(now.getTime() + 60 * 864e5);
  const pertukaran = rows_('Pertukaran')
    .filter(r => r['Tarikh'] instanceof Date && r['Status'] === 'Diluluskan'
      && r['Tarikh'] >= from && r['Tarikh'] <= to)
    .map(r => ({
      tarikh: Utilities.formatDate(r['Tarikh'], TZ, 'yyyy-MM-dd'),
      kod: str_(r['Kod Lokasi']),
      asal: str_(r['Guru Asal']),
      ganti: str_(r['Guru Ganti'])
    }));

  const tetapan = {};
  rows_('Tetapan').forEach(r => { if (r['Kunci']) tetapan[str_(r['Kunci'])] = str_(r['Nilai']); });

  const catatan = rows_('Catatan')
    .filter(r => r['Catatan'] && r['Aktif'] !== false)
    .sort((a, b) => (Number(a['Susunan']) || 99) - (Number(b['Susunan']) || 99))
    .map(r => str_(r['Catatan']));

  return {
    guru: guru.map(g => ({ nama: g.nama })),
    lokasi, jadual, pertukaran, tetapan, catatan,
    dikemaskini: Utilities.formatDate(now, TZ, "yyyy-MM-dd'T'HH:mm:ssXXX")
  };
}

/* ---------- Laporan ---------- */

function addLaporan_(b, me) {
  const kod = str_(b.kod);
  const lokasiSah = rows_('Lokasi').some(r => str_(r['Kod Lokasi']) === kod);
  if (!lokasiSah) return { ok: false, error: 'Lokasi tidak sah.' };

  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(str_(b.tarikh));
  if (!m) return { ok: false, error: 'Tarikh tidak sah.' };
  const tarikh = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  const pilih = (v, list) => list.indexOf(str_(v)) >= 0 ? str_(v) : '';
  ss().getSheetByName('Laporan').appendRow([
    new Date(),
    tarikh,
    HARI[tarikh.getDay()],
    me.nama, // nama sentiasa daripada akaun yang log masuk
    kod,
    pilih(b.kehadiran, ['Hadir', 'Diganti', 'Tidak Hadir']),
    pilih(b.kebersihan, ['Baik', 'Sederhana', 'Kurang Memuaskan']),
    safe_(b.isu),
    safe_(b.tindakan),
    safe_(b.catatan)
  ]);
  return { ok: true };
}

/* ---------- Akaun ---------- */

function akaunSheet_() {
  let sh = ss().getSheetByName(AKAUN_TAB);
  if (!sh) {
    sh = ss().insertSheet(AKAUN_TAB);
    sh.appendRow(AKAUN_HEAD);
    sh.getRange('A:B').setNumberFormat('@');
  }
  return sh;
}

function akaunSemua_(sh) {
  return sh.getDataRange().getValues().slice(1).map((r, i) => ({
    row: i + 2,
    username: str_(r[0]).toLowerCase(),
    id: str_(r[1]),
    salt: str_(r[3]),
    hash: str_(r[4]),
    status: str_(r[5]).toLowerCase()
  }));
}

function guruBelumDaftar_() {
  const ada = new Set(akaunSemua_(akaunSheet_()).map(a => a.id));
  return guruAktif_().filter(g => !ada.has(g.id));
}

function register_(b) {
  const username = str_(b.username).toLowerCase();
  const password = String(b.password || '');
  if (!USERNAME_RE.test(username)) throw authError_('Username 3-30 aksara: huruf kecil, nombor, titik atau garis bawah sahaja.');
  if (password.length < 8 || password.length > 72) throw authError_('Kata laluan mesti 8-72 aksara.');

  const guru = guruAktif_().find(g => g.id === str_(b.idGuru));
  if (!guru) throw authError_('Sila pilih nama guru daripada senarai.');

  const sh = akaunSheet_();
  const akaun = akaunSemua_(sh);
  if (akaun.some(a => a.id === guru.id)) throw authError_('Guru ini sudah mempunyai akaun. Hubungi pentadbir jika terlupa kata laluan.');
  if (akaun.some(a => a.username === username)) throw authError_('Username sudah digunakan. Pilih username lain.');

  const salt = Utilities.getUuid();
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  sh.appendRow([username, guru.id, guru.nama, salt, hashPassword_(salt, password), 'Aktif', now, now]);
  return Object.assign({ ok: true }, newSession_(username, guru.id, guru.nama));
}

function login_(b) {
  const username = str_(b.username).toLowerCase();
  const failKey = 'fail_' + username;
  if (failCount_(failKey) >= MAX_FAILS) throw authError_('Terlalu banyak percubaan gagal. Cuba lagi dalam 15 minit.');

  const sh = akaunSheet_();
  const akaun = USERNAME_RE.test(username) ? akaunSemua_(sh).find(a => a.username === username) : null;
  // Hash juga untuk username yang tiada supaya masa respons tidak mendedahkan username yang wujud
  const candidate = hashPassword_(akaun ? akaun.salt : 'x', String(b.password || ''));
  if (!akaun || !safeEqual_(candidate, akaun.hash)) {
    addFail_(failKey);
    throw authError_('Username atau kata laluan tidak sah.');
  }
  const guru = guruAktif_().find(g => g.id === akaun.id);
  if (akaun.status !== 'aktif' || !guru) throw authError_('Akaun ini tidak aktif. Hubungi pentadbir.');

  CacheService.getScriptCache().remove(failKey);
  sh.getRange(akaun.row, 8).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'));
  return Object.assign({ ok: true }, newSession_(akaun.username, akaun.id, guru.nama));
}

/* ---------- Sesi (token bertandatangan HMAC, tahan SESSION_DAYS hari) ---------- */

function secret_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('TOKEN_SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('TOKEN_SECRET', s);
  }
  return s;
}

function sign_(payload) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, secret_()));
}

function newSession_(username, id, nama) {
  const exp = Date.now() + SESSION_DAYS * 864e5;
  const payload = Utilities.base64EncodeWebSafe(JSON.stringify({ u: username, id: id, exp: exp }));
  return { session: payload + '.' + sign_(payload), username: username, nama: nama, expiresIn: SESSION_DAYS * 86400 };
}

// Pulangkan {username, id, nama} untuk sesi sah; jika tidak, ralat dengan code SESSION
function requireSession_(b) {
  const expired = authError_('Sesi tamat. Sila log masuk semula.', 'SESSION');
  const parts = str_(b.session).split('.');
  if (parts.length !== 2 || !safeEqual_(sign_(parts[0]), parts[1])) throw expired;
  let info;
  try {
    info = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
  } catch (err) {
    throw expired;
  }
  if (!info.exp || Date.now() > info.exp) throw expired;

  // Semak status akaun setiap 5 minit supaya akaun yang disekat berkesan dengan cepat
  const cache = CacheService.getScriptCache();
  const key = 'act_' + info.u;
  let nama = cache.get(key);
  if (!nama) {
    const akaun = akaunSemua_(akaunSheet_()).find(a => a.username === info.u);
    const guru = akaun && guruAktif_().find(g => g.id === akaun.id);
    if (!akaun || akaun.status !== 'aktif' || akaun.id !== info.id || !guru) throw expired;
    nama = guru.nama;
    cache.put(key, nama, 300);
  }
  return { username: info.u, id: info.id, nama: nama };
}

/* ---------- Bantuan ---------- */

function authError_(message, code) {
  const e = new Error(message);
  if (code) e.code = code;
  return e;
}

function hashPassword_(salt, password) {
  const saltBytes = Utilities.newBlob(salt).getBytes();
  let digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + password);
  for (let i = 0; i < HASH_ROUNDS; i++) {
    digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, digest.concat(saltBytes));
  }
  return digest.map(x => ('0' + (x & 255).toString(16)).slice(-2)).join('');
}

function safeEqual_(a, b) {
  a = String(a); b = String(b);
  let diff = a.length ^ b.length;
  for (let i = 0; i < Math.max(a.length, b.length); i++) {
    diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  }
  return diff === 0;
}

function failCount_(key) {
  return Number(CacheService.getScriptCache().get(key) || 0);
}

function addFail_(key) {
  CacheService.getScriptCache().put(key, String(failCount_(key) + 1), FAIL_WINDOW_SECONDS);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function rows_(name) {
  const sh = ss().getSheetByName(name);
  if (!sh) throw new Error('Tab "' + name + '" tidak dijumpai dalam Google Sheet');
  const values = sh.getDataRange().getValues();
  const head = values.shift().map(h => String(h).trim());
  return values.map(row => {
    const o = {};
    head.forEach((h, i) => { o[h] = row[i]; });
    return o;
  });
}

function str_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

// Elak teks pengguna dibaca sebagai formula dan hadkan panjang
function safe_(v) {
  const s = str_(v).slice(0, 1000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
