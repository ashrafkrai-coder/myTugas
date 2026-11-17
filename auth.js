(function () {
  // URL Web App Apps Script (berakhir dengan /exec)
  const ENDPOINT = "https://script.google.com/macros/s/AKfycbyIfUQ3w8J4uh39UcbkwsU_6xxi90ukZR9saSxtxnqK27N9oWvU5Fu4kLXNdQbYpohb/exec";
  const KEY = "mytugas_auth";

  function readState() {
    try { return JSON.parse(localStorage.getItem(KEY) || "null"); } catch (e) { return null; }
  }
  function writeState(state) {
    try { localStorage.setItem(KEY, JSON.stringify(state)); } catch (e) {}
  }
  function clearState() {
    try { localStorage.removeItem(KEY); } catch (e) {}
  }

  // Sesi dikeluarkan dan disahkan oleh pelayan; di sini hanya untuk tentukan sama ada halaman login perlu dipaparkan
  function isAuthenticated() {
    const s = readState();
    return !!(s && s.session && s.expiresAt && Date.now() < s.expiresAt);
  }

  function currentUser() {
    const s = readState();
    return s && isAuthenticated() ? { username: s.username, nama: s.nama } : null;
  }

  function saveSession(r) {
    writeState({
      session: r.session,
      username: r.username,
      nama: r.nama,
      expiresAt: Date.now() + Math.max(60, (r.expiresIn || 21600) - 60) * 1000
    });
  }

  function setNama(nama) {
    const s = readState();
    if (s && nama && s.nama !== nama) { s.nama = nama; writeState(s); }
  }

  function goToLogin() {
    if (!/login\.html$/.test(location.pathname)) location.replace("login.html");
  }

  async function call(action, data) {
    const s = readState() || {};
    const body = Object.assign({}, data, { action: action, session: s.session });
    let res;
    try {
      // text/plain mengelakkan CORS preflight yang tidak disokong Apps Script
      res = await fetch(ENDPOINT, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(body),
        cache: "no-store"
      });
    } catch (e) {
      const err = new Error("Tiada sambungan ke pelayan. Cuba semula.");
      err.code = "NETWORK";
      throw err;
    }
    if (!res.ok) throw new Error("Pelayan tidak dapat dihubungi (HTTP " + res.status + ").");
    let json;
    try { json = await res.json(); } catch (e) { throw new Error("Respons pelayan tidak sah."); }
    if (!json.ok) {
      if (json.code === "SESSION") { clearState(); goToLogin(); }
      const err = new Error(json.error || "Ralat pelayan.");
      err.code = json.code || "";
      throw err;
    }
    return json;
  }

  function signOut() {
    clearState();
    location.replace("login.html");
  }

  window.myTugasAuth = { isAuthenticated, currentUser, saveSession, setNama, call, signOut, endpoint: ENDPOINT };
})();
