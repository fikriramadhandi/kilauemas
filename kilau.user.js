// ==UserScript==
// @name         KILAU – Balas Otomatis (Outlook Web)
// @namespace    kilau.bbg.bsi
// @version      1.3.0
// @description  Membalas (Reply all) email "[DAS] - FILE MULTIPOSTING" per hari dengan file FIX + Berita Acara + isi email dari paket KILAU.
// @match        https://outlook.office.com/*
// @match        https://outlook.office365.com/*
// @match        https://outlook.cloud.microsoft/*
// @updateURL    https://NAMA-DOMAIN.vercel.app/kilau.user.js
// @downloadURL  https://NAMA-DOMAIN.vercel.app/kilau.user.js
// @grant        GM_xmlhttpRequest
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @run-at       document-idle
// ==/UserScript==

(function () {
  "use strict";
  if (window.top !== window.self) return;           // jangan jalan di iframe
  if (window.__KILAU_OWA__) return;
  window.__KILAU_OWA__ = true;

  // ------------------------------------------------------------------
  // Pengaturan
  // ------------------------------------------------------------------
  const DONE_KEY = "kilau-owa-done-v1";
  const CFG_KEY = "kilau-owa-tracker-v1";
  const W = (typeof unsafeWindow !== "undefined") ? unsafeWindow : window;
  const T = { short: 8000, long: 25000 };            // batas tunggu (ms)

  // Selector Outlook Web (tampilan bahasa Inggris). Beberapa alternatif
  // dicoba berurutan karena Microsoft kadang mengganti tampilannya.
  const SEL = {
    search: ['#topSearchInput', 'input[aria-label="Search"]', 'input[placeholder="Search"]', 'input[type="search"]'],
    listItem: ['[role="listbox"] [role="option"]', 'div[data-convid]', '[role="option"]'],
    readingPane: ['#ReadingPaneContainerId', '[aria-label="Reading Pane"]', '[aria-label="Reading pane"]', '[role="main"]'],
    replyAll: ['button[aria-label="Reply all"]', '[role="button"][aria-label="Reply all"]', 'button[title="Reply all"]', '[role="menuitem"][aria-label="Reply all"]'],
    editor: ['div[contenteditable="true"][aria-label^="Message body"]', 'div[role="textbox"][contenteditable="true"][aria-multiline="true"]', 'div[contenteditable="true"][role="textbox"]'],
    send: ['button[aria-label="Send"]', 'button[title^="Send"]', 'button[name="Send"]'],
    attach: ['button[aria-label="Attach file"]', 'button[aria-label^="Attach"]', 'button[title^="Attach"]', '[role="menuitem"][aria-label^="Attach"]'],
    browse: ['[role="menuitem"][aria-label*="Browse this computer"]', '[role="menuitem"][aria-label*="Upload from this device"]'],
  };

  // ------------------------------------------------------------------
  // Utilitas
  // ------------------------------------------------------------------
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const visible = el => !!el && el.getClientRects().length > 0 && getComputedStyle(el).visibility !== "hidden";
  const qAll = (list, root = document) => { for (const s of list) { const f = [...root.querySelectorAll(s)].filter(visible); if (f.length) return f; } return []; };
  const q = (list, root) => qAll(list, root)[0] || null;
  const byText = (re, root = document, sel = 'button,[role="menuitem"],[role="button"]') => [...root.querySelectorAll(sel)].filter(e => visible(e) && re.test((e.getAttribute("aria-label") || "") + " " + e.textContent));
  async function waitFor(fn, ms = T.short, step = 250) {
    const end = Date.now() + ms;
    while (Date.now() < end) { if (state.stop) throw new Error("Dihentikan"); const v = fn(); if (v) return v; await sleep(step); }
    return null;
  }
  const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const toHtml = t => `<div>${t.split(/\r?\n/).map(l => l ? esc(l) : "<br>").join("</div><div>")}</div>`;
  function setInputValue(el, v) {
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value").set;
    setter.call(el, v);
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
  }
  function key(el, k) {
    const o = { key: k, code: k, keyCode: k === "Enter" ? 13 : 27, which: k === "Enter" ? 13 : 27, bubbles: true, cancelable: true };
    el.dispatchEvent(new KeyboardEvent("keydown", o));
    el.dispatchEvent(new KeyboardEvent("keypress", o));
    el.dispatchEvent(new KeyboardEvent("keyup", o));
  }
  const doneMap = () => { try { return JSON.parse(localStorage.getItem(DONE_KEY) || "{}") || {}; } catch (_) { return {}; } };
  const markDone = (k, v) => { const m = doneMap(); m[k] = v; try { localStorage.setItem(DONE_KEY, JSON.stringify(m)); } catch (_) {} };

  // ------------------------------------------------------------------
  // State & panel (Shadow DOM supaya tidak bentrok dengan gaya Outlook)
  // ------------------------------------------------------------------
  const state = { items: [], ba: null, running: false, stop: false, waitResolve: null };

  const host = document.createElement("div");
  host.style.cssText = "position:fixed;right:16px;bottom:16px;z-index:2147483647";
  document.documentElement.appendChild(host);
  setInterval(() => { if (!document.documentElement.contains(host)) document.documentElement.appendChild(host); }, 2000);
  const root = host.attachShadow({ mode: "open" });
  // Dibangun tanpa innerHTML karena Outlook Web memakai Trusted Types
  const h = (tag, attrs = {}, ...kids) => {
    const e = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) { if (k === "text") e.textContent = v; else if (k === "checked") e.checked = v; else e.setAttribute(k, v); }
    kids.forEach(c => e.append(c)); return e;
  };
  const style = document.createElement("style");
  style.textContent = `
    *{box-sizing:border-box;font-family:"Segoe UI",system-ui,sans-serif}
    .box{width:380px;max-height:78vh;display:flex;flex-direction:column;background:#fff;color:#1c1917;border:1px solid #d6d3d1;border-radius:10px;box-shadow:0 8px 28px rgba(0,0,0,.18);overflow:hidden}
    .hd{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:10px 12px;background:#0f766e;color:#fff;cursor:pointer}
    .hd b{font-size:14px} .hd span{font-size:11px;opacity:.85}
    .bd{padding:10px 12px;overflow:auto;font-size:12.5px}
    .min .bd{display:none}
    label.f{display:block;border:1px dashed #0f766e;border-radius:6px;padding:8px;text-align:center;color:#0f766e;font-weight:600;cursor:pointer;margin-bottom:8px}
    label.f input{display:none}
    .row{display:flex;gap:6px;flex-wrap:wrap;margin:6px 0}
    button{font:inherit;font-size:12.5px;font-weight:600;border-radius:6px;padding:6px 10px;border:1px solid #0f766e;background:#0f766e;color:#fff;cursor:pointer}
    button.g{background:#fff;color:#0f766e} button:disabled{opacity:.45;cursor:not-allowed}
    select{font:inherit;font-size:12.5px;padding:5px;border:1px solid #d6d3d1;border-radius:6px;width:100%}
    .chk{display:flex;gap:6px;align-items:center;margin:6px 0}
    table{width:100%;border-collapse:collapse;margin:6px 0} td{padding:3px 4px;border-bottom:1px solid #eee;vertical-align:top}
    .ok{color:#15803d;font-weight:600}.warn{color:#b45309;font-weight:600}.err{color:#b91c1c;font-weight:600}.mut{color:#78716c}
    .log{background:#fafaf9;border:1px solid #eee;border-radius:6px;padding:6px;height:140px;overflow:auto;font-family:Consolas,monospace;font-size:11px;white-space:pre-wrap}
    details{border:1px solid #e7e5e4;border-radius:6px;padding:6px 8px;margin-bottom:8px} summary{cursor:pointer;font-weight:600;color:#0f766e}
    details input{width:100%;font:inherit;font-size:12px;padding:5px;border:1px solid #d6d3d1;border-radius:6px;margin-top:5px}
    .pause{background:#fef3c7;border:1px solid #f59e0b;border-radius:6px;padding:8px;margin:6px 0;display:none}`;
  const sel = h("select", { id: "mode" },
    h("option", { value: "draft", text: "Simpan sebagai DRAFT (aman, cek dulu di Drafts)" }),
    h("option", { value: "send", text: "KIRIM langsung" }));
  const chk = (id, text) => { const i = h("input", { type: "checkbox", id, checked: true }); return h("label", { class: "chk" }, i, text); };
  const box = h("div", { class: "box min", id: "box" },
    h("div", { class: "hd", id: "hd" }, h("b", { text: "KILAU – Balas Otomatis" }), h("span", { id: "sum", text: "klik untuk buka" })),
    h("div", { class: "bd" },
      h("label", { class: "f" }, 'Pilih file: semua isi folder "Balasan" (.xlsx + .txt) dan PDF Berita Acara',
        h("input", { type: "file", id: "files", multiple: "", accept: ".xlsx,.txt,.pdf" })),
      h("details", { id: "trk" }, h("summary", { text: "Tracker Google Sheets (opsional)" }),
        h("input", { id: "turl", placeholder: "URL Web App (…/exec)" }),
        h("input", { id: "ttok", placeholder: "Token" }),
        h("input", { id: "tnama", placeholder: "Nama lo (tercatat di tracker)" }),
        h("div", { class: "row" }, h("button", { class: "g", id: "tsave", text: "Simpan & tes koneksi" })),
        h("div", { id: "tinfo", class: "mut", text: "Belum terhubung." })),
      h("div", { id: "info", class: "mut", text: "Belum ada file." }),
      h("table", { id: "tbl" }),
      h("div", { class: "row" }, sel),
      chk("slow", " Konfirmasi dulu setiap email sebelum disimpan/dikirim"),
      chk("limit", " Uji coba: proses maksimal 3 email"),
      h("div", { class: "pause", id: "pause" }, h("div", { id: "pmsg" }),
        h("div", { class: "row" }, h("button", { id: "go", text: "Lanjut" }), h("button", { class: "g", id: "skip", text: "Lewati email ini" }))),
      h("div", { class: "row" },
        h("button", { id: "start", text: "Mulai", disabled: "" }), h("button", { class: "g", id: "stop", text: "Stop", disabled: "" }),
        h("button", { class: "g", id: "diag", text: "Diagnosa" }), h("button", { class: "g", id: "reset", text: "Reset status" })),
      h("div", { class: "row" }, h("button", { class: "g", id: "sentAll", text: "Draft sudah saya kirim → tandai Terkirim" })),
      h("div", { class: "log", id: "log" })));
  root.append(style, box);
  const $ = id => root.getElementById(id);
  $("hd").onclick = () => { $("box").classList.toggle("min"); };
  const log = (msg, cls) => { const d = document.createElement("div"); if (cls) d.className = cls; d.textContent = `${new Date().toLocaleTimeString("id-ID")}  ${msg}`; $("log").appendChild(d); $("log").scrollTop = 1e9; };

  // ------------------------------------------------------------------
  // Tracker Google Sheets (status bersama satu tim)
  // ------------------------------------------------------------------
  const cfg = () => { try { return JSON.parse(localStorage.getItem(CFG_KEY) || "{}") || {}; } catch (_) { return {}; } };
  state.remote = {};
  function gm(method, url, body) {
    return new Promise((res, rej) => {
      if (typeof GM_xmlhttpRequest === "undefined") return rej(new Error("GM_xmlhttpRequest tidak tersedia"));
      GM_xmlhttpRequest({ method, url, data: body, headers: { "Content-Type": "text/plain;charset=utf-8" }, timeout: 20000,
        onload: r => { try { res(JSON.parse(r.responseText)); } catch (e) { rej(new Error("Respons tracker tidak valid (cek URL/izin akses Web App)")); } },
        onerror: () => rej(new Error("Tidak bisa menghubungi tracker")), ontimeout: () => rej(new Error("Tracker tidak merespons")) });
    });
  }
  async function pullStatus() {
    const c = cfg(); if (!c.url) return false;
    const r = await gm("GET", `${c.url}?token=${encodeURIComponent(c.token || "")}&t=${Date.now()}`);
    if (!r.ok) throw new Error(r.error === "token" ? "Token tracker salah" : "Tracker menolak permintaan");
    state.remote = r.status || {}; return true;
  }
  async function pushLog(item, status, jumlah, catatan) {
    const c = cfg(); if (!c.url) return;
    try {
      const r = await gm("POST", c.url, JSON.stringify({ token: c.token, rows: [{ tanggal: item.date, program: item.prog, status, jumlah: jumlah ?? "", oleh: c.nama || "", catatan: catatan || "" }] }));
      if (!r.ok) log(`  (tracker: ${r.error})`, "warn");
    } catch (e) { log(`  (tracker gagal dicatat: ${e.message})`, "warn"); }
  }
  const remoteOf = i => state.remote[i.key];
  const blockedByRemote = i => { const r = remoteOf(i), me = (cfg().nama || "").trim().toLowerCase();
    if (!r) return null;
    if (/^terkirim$/i.test(r.status)) return `sudah Terkirim oleh ${r.oleh || "?"}`;
    if (/^draft$/i.test(r.status)) return `sudah Draft oleh ${r.oleh.trim().toLowerCase() === me ? "lo sendiri" : (r.oleh || "?")}`;
    if (/tidak ada email/i.test(r.status)) return "ditandai tidak ada email";
    return null; };

  // ------------------------------------------------------------------
  // Membaca file dari folder Balasan
  // ------------------------------------------------------------------
  const parseName = n => { const m = n.match(/(EMAS30|EMAS50).*?(\d{4}-\d{2}-\d{2})/i); return m ? { prog: m[1].toUpperCase(), date: m[2] } : null; };
  $("files").onchange = async e => {
    const fs = [...e.target.files]; const map = {}; state.ba = null;
    for (const f of fs) {
      if (/\.pdf$/i.test(f.name)) { state.ba = f; continue; }
      const p = parseName(f.name); if (!p) continue;
      const k = `${p.prog}|${p.date}`; map[k] = map[k] || { key: k, ...p };
      if (/\.xlsx$/i.test(f.name)) map[k].fix = f;
      if (/\.txt$/i.test(f.name)) {
        const t = (await f.text()).replace(/^\uFEFF/, "");
        const lines = t.split(/\r?\n/);
        if (lines[0].startsWith("#KILAU")) { map[k].meta = lines[0]; map[k].body = lines.slice(1).join("\n").trim(); }
        else map[k].body = t.trim();
        const m = (map[k].meta || "").match(/baris_asli=(\d+)/); map[k].nAsli = m ? +m[1] : null;
      }
    }
    state.items = Object.values(map).sort((a, b) => a.date === b.date ? a.prog.localeCompare(b.prog) : a.date.localeCompare(b.date));
    renderList();
  };
  function renderList() {
    const done = doneMap();
    const ok = state.items.filter(i => i.fix && i.body).length;
    const info = $("info"); info.replaceChildren(`${state.items.length} hari ditemukan, ${ok} lengkap. Berita acara: `,
      state.ba ? h("span", { class: "ok", text: state.ba.name }) : h("span", { class: "err", text: "belum dipilih (PDF)" }));
    $("tbl").replaceChildren(...state.items.map(i => h("tr", {}, h("td", { text: i.date }), h("td", { text: i.prog }),
      h("td", {}, !i.fix || !i.body ? h("span", { class: "err", text: "file kurang" }) : remoteOf(i) ? h("span", { class: /terkirim/i.test(remoteOf(i).status) ? "ok" : "warn", text: `${remoteOf(i).status} · ${remoteOf(i).oleh}` }) : done[i.key] ? h("span", { class: "ok", text: done[i.key] }) : h("span", { class: "mut", text: "menunggu" })))));
    $("start").disabled = !(ok && state.ba) || state.running;
    $("sum").textContent = state.items.length ? `${ok} hari siap` : "klik untuk buka";
  }
  (() => { const c = cfg(); $("turl").value = c.url || ""; $("ttok").value = c.token || ""; $("tnama").value = c.nama || ""; if (c.url) $("trk").open = false; })();
  $("tsave").onclick = async () => {
    localStorage.setItem(CFG_KEY, JSON.stringify({ url: $("turl").value.trim(), token: $("ttok").value.trim(), nama: $("tnama").value.trim() }));
    $("tinfo").textContent = "Menghubungi tracker…";
    try { await pullStatus(); $("tinfo").textContent = `Terhubung. ${Object.keys(state.remote).length} status tercatat.`; $("tinfo").className = "ok"; renderList(); }
    catch (e) { $("tinfo").textContent = e.message; $("tinfo").className = "err"; }
  };
  $("reset").onclick = () => { if (confirm("Hapus status 'sudah diproses' untuk semua hari?")) { localStorage.removeItem(DONE_KEY); renderList(); log("Status direset."); } };

  // ------------------------------------------------------------------
  // Langkah-langkah di Outlook Web
  // ------------------------------------------------------------------
  async function searchFor(item) {
    const box = q(SEL.search);
    if (!box) throw new Error("Kotak Search tidak ditemukan");
    box.focus(); box.click();
    setInputValue(box, `"Periode Data ${item.date}" ${item.prog} "FILE MULTIPOSTING"`);
    await sleep(300); key(box, "Enter");
    await sleep(1500);
  }
  function itemText(el) { return ((el.getAttribute("aria-label") || "") + " " + el.textContent).replace(/\s+/g, " "); }
  function isOriginal(txt, item) {
    return txt.includes(item.date) && new RegExp(item.prog, "i").test(txt) && /\[DAS\]/i.test(txt)
      && !/\b(RE|Re|FW|Fw|FWD)\s*:\s*\[DAS\]/.test(txt) && !/\[UJI COBA\]/i.test(txt);
  }
  async function openOriginal(item) {
    const el = await waitFor(() => qAll(SEL.listItem).find(e => isOriginal(itemText(e), item)), T.long);
    if (!el) throw new Error("Email DAS tidak ditemukan di hasil pencarian");
    el.scrollIntoView({ block: "center" }); el.click();
    const pane = await waitFor(() => { const p = q(SEL.readingPane); return p && p.innerText.includes(item.date) ? p : null; }, T.long);
    if (!pane) throw new Error("Email tidak terbuka di Reading Pane");
    await sleep(800);
    return pane;
  }
  function checkCount(pane, item) {
    const m = pane.innerText.match(/eligible\s+sebanyak\s*:?\s*(\d+)/i);
    if (m && +m[1] === 0) { const e = new Error("0 transaksi, tidak perlu dibalas"); e.nol = true; throw e; }
    if (m && item.nAsli != null && +m[1] !== item.nAsli) throw new Error(`Jumlah tidak cocok: email DAS ${m[1]} transaksi, file ${item.nAsli} baris`);
    return m ? +m[1] : null;
  }
  async function clickReplyAll(pane) {
    let btn = q(SEL.replyAll, pane) || q(SEL.replyAll) || byText(/^\s*Reply all\s*$/i)[0];
    if (!btn) throw new Error("Tombol Reply all tidak ditemukan");
    const before = new Set(qAll(SEL.editor));
    btn.click();
    const ed = await waitFor(() => qAll(SEL.editor).find(e => !before.has(e)), T.long);
    if (!ed) throw new Error("Jendela balasan tidak terbuka");
    await sleep(1200);
    return ed;
  }
  function composeRoot(ed) {
    let n = ed;
    while (n && n !== document.body) { if (q(SEL.send, n)) return n; n = n.parentElement; }
    return document.body;
  }
  async function insertBody(ed, text) {
    ed.focus();
    const sel = getSelection(), r = document.createRange();
    r.setStart(ed, 0); r.collapse(true); sel.removeAllRanges(); sel.addRange(r);
    const probe = text.split("\n").find(l => l.trim().length > 15) || text.slice(0, 20);
    const html = toHtml(text) + "<div><br></div>";
    try {
      const dt = new DataTransfer(); dt.setData("text/html", html); dt.setData("text/plain", text);
      ed.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
    } catch (_) {}
    if (await waitFor(() => ed.innerText.includes(probe.trim().slice(0, 25)), 2500)) return "paste";
    try { document.execCommand("insertText", false, text + "\n\n"); } catch (_) {}
    if (await waitFor(() => ed.innerText.includes(probe.trim().slice(0, 25)), 2500)) return "insertHTML";
    throw new Error("Isi email gagal ditempel");
  }
  const hasAttachment = (rootEl, name) => { const base = name.replace(/\.[^.]+$/, ""); return rootEl.innerText.includes(base) || [...rootEl.querySelectorAll("[aria-label],[title]")].some(e => ((e.getAttribute("aria-label") || "") + (e.getAttribute("title") || "")).includes(base)); };
  async function attachFiles(ed, files) {
    const cr = composeRoot(ed);
    const allIn = () => files.every(f => hasAttachment(cr, f.name));
    const dt = () => { const d = new DataTransfer(); files.forEach(f => d.items.add(f)); return d; };

    // Cara A: klik "Attach file" → "Browse this computer", lalu cegat jendela pilih file
    const btn = q(SEL.attach, cr) || q(SEL.attach);
    if (btn) {
      const P = W.HTMLInputElement.prototype, orig = P.click;
      let used = false;
      P.click = function () {
        if (this.type === "file" && !used) { used = true; this.files = dt().files; this.dispatchEvent(new Event("change", { bubbles: true })); return; }
        return orig.apply(this, arguments);
      };
      try {
        btn.click(); await sleep(600);
        if (!used) { const b = q(SEL.browse) || byText(/Browse this computer|Upload from this device|This (PC|device|computer)/i)[0]; if (b) { b.click(); await sleep(600); } }
      } finally { P.click = orig; }
      if (used && await waitFor(allIn, T.long)) return "menu";
      if (!used) key(document.activeElement || document.body, "Escape");
    }
    // Cara B: input file yang sudah ada di jendela balasan
    const inp = [...cr.querySelectorAll('input[type="file"]')][0];
    if (inp) { inp.files = dt().files; inp.dispatchEvent(new Event("change", { bubbles: true })); if (await waitFor(allIn, T.long)) return "input"; }
    // Cara C: seret-lepas ke badan email
    const d = dt();
    ["dragenter", "dragover", "drop"].forEach(t => ed.dispatchEvent(new DragEvent(t, { dataTransfer: d, bubbles: true, cancelable: true })));
    if (await waitFor(allIn, T.long)) return "drop";
    throw new Error("Lampiran gagal ditambahkan otomatis");
  }
  async function waitUploads(cr) {
    await waitFor(() => !/Uploading|Attaching|\d+\s*%/.test(cr.innerText), T.long);
    await sleep(1000);
  }
  function pause(msg) {
    $("pmsg").textContent = msg; $("pause").style.display = "block";
    return new Promise(res => { state.waitResolve = res; });
  }
  $("go").onclick = () => { $("pause").style.display = "none"; state.waitResolve && state.waitResolve("go"); };
  $("skip").onclick = () => { $("pause").style.display = "none"; state.waitResolve && state.waitResolve("skip"); };
  async function finish(ed, mode) {
    const cr = composeRoot(ed);
    if (mode === "send") {
      const s = q(SEL.send, cr) || q(SEL.send);
      if (!s) throw new Error("Tombol Send tidak ditemukan");
      s.click();
      if (!await waitFor(() => !document.contains(ed) || !visible(ed), T.long)) throw new Error("Email belum terkirim (jendela balasan masih terbuka)");
      return "terkirim";
    }
    const saved = await waitFor(() => /Draft saved|Saved/i.test(cr.innerText), 20000);
    await sleep(1500);
    return saved ? "draft" : "draft (cek)";
  }

  async function processItem(item, mode, slow) {
    log(`— ${item.date} ${item.prog}: mencari email…`);
    await searchFor(item);
    const pane = await openOriginal(item);
    const n = checkCount(pane, item);
    log(`  email ditemukan${n != null ? `, ${n} transaksi (cocok)` : ""}`);
    const ed = await clickReplyAll(pane);
    log(`  Reply all dibuka, menempel isi email…`);
    log(`  isi email: ${await insertBody(ed, item.body)}`);
    log(`  melampirkan ${item.fix.name} + ${state.ba.name}…`);
    log(`  lampiran: ${await attachFiles(ed, [item.fix, state.ba])}`);
    await waitUploads(composeRoot(ed));
    if (slow) {
      const a = await pause(`${item.date} ${item.prog}: cek isi dan lampiran di jendela balasan. Klik "Lanjut" untuk ${mode === "send" ? "KIRIM" : "simpan draft"}.`);
      if (a === "skip") { log("  dilewati oleh pengguna (balasan dibiarkan terbuka/draft).", "warn"); await pushLog(item, "Dilewati", n, "dilewati oleh pengguna"); return "dilewati"; }
    }
    const res = await finish(ed, mode);
    markDone(item.key, res); log(`  ✓ ${res}`, "ok");
    await pushLog(item, res.startsWith("terkirim") ? "Terkirim" : "Draft", n ?? item.nAsli, res === "draft (cek)" ? "draft belum terkonfirmasi tersimpan" : "");
    return res;
  }

  $("start").onclick = async () => {
    const mode = $("mode").value, slow = $("slow").checked, limit = $("limit").checked ? 3 : Infinity;
    if (mode === "send" && !confirm("Mode KIRIM LANGSUNG: balasan akan terkirim ke DAS/DBO. Lanjutkan?")) return;
    state.running = true; state.stop = false; $("start").disabled = true; $("stop").disabled = false;
    if (cfg().url) { try { await pullStatus(); renderList(); log(`Tracker: ${Object.keys(state.remote).length} status dimuat.`); } catch (e) { log(`Tracker: ${e.message}`, "err"); if (!confirm("Tracker tidak bisa dihubungi, jadi status tim tidak dicek. Tetap lanjut?")) { state.running = false; $("stop").disabled = true; renderList(); return; } } }
    const done = doneMap();
    const todo = [];
    for (const i of state.items) {
      if (!i.fix || !i.body || done[i.key]) continue;
      if (i.nAsli === 0) { log(`Lewati ${i.date} ${i.prog}: 0 transaksi, tidak perlu dibalas.`, "mut"); continue; }
      const b = blockedByRemote(i); if (b) { log(`Lewati ${i.date} ${i.prog}: ${b}.`, "mut"); continue; }
      todo.push(i);
    }
    todo.splice(limit);
    log(`Mulai: ${todo.length} email, mode ${mode === "send" ? "KIRIM" : "DRAFT"}.`);
    let ok = 0, fail = 0;
    for (const item of todo) {
      if (state.stop) break;
      try { await processItem(item, mode, slow); ok++; }
      catch (e) {
        if (e.nol) { log(`  – ${item.date} ${item.prog}: ${e.message}.`, "mut"); markDone(item.key, "0 transaksi"); continue; }
        fail++; log(`  ✗ ${item.date} ${item.prog}: ${e.message}`, "err");
        if (e.message === "Dihentikan") break;
        await pushLog(item, "Gagal", item.nAsli, e.message);
        const a = await pause(`Gagal di ${item.date} ${item.prog}: ${e.message}. Selesaikan manual jika perlu, lalu klik "Lanjut" ke email berikutnya.`);
        if (a === "skip") { /* lanjut juga */ }
      }
      renderList();
    }
    log(`Selesai. Berhasil ${ok}, gagal ${fail}.`, fail ? "warn" : "ok");
    state.running = false; $("stop").disabled = true; renderList();
  };
  $("sentAll").onclick = async () => {
    const done = doneMap();
    const list = state.items.filter(i => /^draft/i.test(done[i.key] || ""));
    if (!list.length) { log("Tidak ada hari berstatus draft di browser ini.", "mut"); return; }
    if (!confirm(`Tandai ${list.length} hari ini sebagai TERKIRIM?\n\n${list.map(i => `${i.date} ${i.prog}`).join("\n")}\n\nPastikan draft-nya memang sudah dikirim dari folder Drafts.`)) return;
    for (const i of list) { markDone(i.key, "terkirim (manual)"); await pushLog(i, "Terkirim", i.nAsli, "dikirim manual dari Drafts"); log(`✓ ${i.date} ${i.prog} ditandai Terkirim.`, "ok"); }
    renderList();
  };
  $("stop").onclick = () => { state.stop = true; state.waitResolve && state.waitResolve("skip"); log("Menghentikan…", "warn"); };

  $("diag").onclick = () => {
    const pane = q(SEL.readingPane);
    const rep = [
      ["Kotak Search", !!q(SEL.search)],
      ["Daftar email", qAll(SEL.listItem).length + " item"],
      ["Reading Pane", !!pane],
      ["Tombol Reply all", !!(q(SEL.replyAll, pane || document) || byText(/^\s*Reply all\s*$/i)[0])],
      ["Editor balasan (jika terbuka)", !!q(SEL.editor)],
      ["Tombol Attach (jika terbuka)", !!q(SEL.attach)],
      ["Tombol Send (jika terbuka)", !!q(SEL.send)],
      ["Halaman", location.host + location.pathname],
    ];
    log("Diagnosa:\n" + rep.map(r => `  ${r[0]}: ${r[1]}`).join("\n"));
  };

  log("KILAU siap. Pilih file dari folder Balasan + PDF BA.");
})();
