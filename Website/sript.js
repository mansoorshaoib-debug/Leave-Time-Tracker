/* =========================================================
   OTHTECH Leave Portal
   - Data comes from Google Sheet via Apps Script (API_URL)
   - Annual leaves: 10 per person
   - Extra leaves: off-day work (sheet) + overtime (every 8h = 1, monthly)
   - Two sections: Leaves · Working Hours (Daily / Weekly / Monthly)
   ========================================================= */

const API_URL = "https://script.google.com/macros/s/AKfycbx-whgnAdu_84LhUL879oUYJ0q5grGgHSHreq4UCMnrCvWmvApuBcU_MRdSV5URX3513Q/exec";
const ANNUAL_LEAVES = 10;

let data = {};      // admin: sab employees (passwords ke baghair)
let adminIssues = [];
const $ = id => document.getElementById(id);

/* ---------------- helpers ---------------- */
const esc = v => String(v ?? "").replace(/[&<>"']/g, c => ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;" }[c]));
const MONTHS = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
const parseDate = d => {
  if(d == null || d === "") return null;
  const s = String(d).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);                 // 2026-02-14
  if(m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})[-\s/]([A-Za-z]+)[-\s/,]*(\d{4})$/);        // 14-Feb-2026, 01-June-2026
  if(m){ const mi = MONTHS.indexOf(m[2].slice(0, 3).toLowerCase()); if(mi > -1) return new Date(+m[3], mi, +m[1]); }
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);                    // 14/02/2026 (day first)
  if(m) return new Date(+m[3], +m[2] - 1, +m[1]);
  const p = new Date(s);
  return Number.isNaN(p.getTime()) ? null : p;
};
const sortByDate = (a, b) => (parseDate(a.date)?.getTime() || 0) - (parseDate(b.date)?.getTime() || 0);
const fmtLong = d => { const p = parseDate(d); return p ? p.toLocaleDateString("en-GB", { weekday:"long", day:"numeric", month:"long", year:"numeric" }) : (d || "—"); };
const initials = name => String(name || "?").trim().split(/\s+/).slice(0, 2).map(w => w[0]).join("").toUpperCase();
const num = v => { const n = parseFloat(v); return Number.isFinite(n) ? n : 0; };

/* Pick the first key that exists on an object (case/format tolerant) */
function pick(obj, keys){
  if(!obj) return undefined;
  const map = {};
  Object.keys(obj).forEach(k => map[k.toLowerCase().replace(/[\s_-]/g, "")] = obj[k]);
  for(const k of keys){
    const v = map[k.toLowerCase().replace(/[\s_-]/g, "")];
    if(v !== undefined && v !== null && v !== "") return v;
  }
  return undefined;
}

const EXTRA_KEYS = ["extraLeaves", "extra", "extraLeave", "extraLeavesList", "offDays", "offDayWork", "offDayWorked", "compLeaves", "compensatoryLeaves", "earnedLeaves", "bonusLeaves"];
const isExtraType = t => /extra|off\s*day|comp|earn|bonus/i.test(String(t || ""));

/* Turn any shape of extra-leave data into { count, items[] }
   Supports: a number (3), a string ("3"), an array of dates,
   an array of {date, reason, days}, or leaves tagged type:"extra". */
function getExtra(emp){
  const raw = pick(emp, EXTRA_KEYS);
  let items = [];
  let count = 0;

  if(Array.isArray(raw)){
    items = raw.map(x => (typeof x === "object" && x !== null)
      ? { date: pick(x, ["date", "offDay", "workedOn", "day"]), reason: pick(x, ["reason", "note", "remarks", "description", "work"]), days: num(pick(x, ["days", "count", "qty", "leaves"])) || 1 }
      : { date: x, reason: "", days: 1 });
    count = items.reduce((s, i) => s + i.days, 0);
  } else if(raw !== undefined){
    count = num(raw);
  }

  // leaves tagged as extra inside the normal list
  (emp.leaves || []).filter(l => isExtraType(l.type)).forEach(l => {
    items.push({ date: l.date, reason: l.reason, days: num(l.days) || 1 });
    count += num(l.days) || 1;
  });

  // explicit count column wins if the sheet sends one
  const explicit = pick(emp, ["extraCount", "extraTotal", "totalExtra", "extraLeavesCount"]);
  if(explicit !== undefined) count = num(explicit);

  items.sort(sortByDate);

  // plain number typed in the Leaves tab "Extra Leaves" column (no date)
  const manual = num(pick(emp, ["extraManual"]));
  if(manual){
    items.push({ date: null, reason: "Added by admin", days: manual, noDate: true });
    count += manual;
  }

  return { count, items };
}

function getUsed(emp){
  return (emp.leaves || []).filter(l => !isExtraType(l.type)).slice().sort(sortByDate);
}

function summary(emp){
  const used = getUsed(emp);
  const usedCount = used.reduce((s, l) => s + (num(l.days) || 1), 0);
  const extra = getExtra(emp);
  const mm = attModel(emp);
  const otRewards = mm.rewards.reduce((t, r) => t + r.days, 0);   // har 8h overtime = +1 leave
  const otCut = mm.cuts.length;                                    // har 8h short = −1 paid leave
  const extraAll = extra.count + otRewards;
  const total = ANNUAL_LEAVES + extraAll - otCut;
  return { used, usedCount, extra, otRewards, otCut, extraAll, total, remaining: total - usedCount };
}

/* ---------------- data ----------------
   Password ab server (Apps Script) par check hota hai.
   Browser ko sirf login karne wale ka data milta hai. */
async function api(body){
  try{
    const r = await fetch(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },   // simple request: CORS preflight nahi
      body: JSON.stringify(body)
    });
    return await r.json();
  }catch(err){
    console.error("API error", err);
    return { ok:false, error:"Could not connect. Check your internet and try again." };
  }
}

/* ---------------- icons ---------------- */
const I = {
  sun: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4.5"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>`,
  moon: `<svg viewBox="0 0 24 24" fill="currentColor"><path d="M20.5 14.2A8.5 8.5 0 0 1 9.8 3.5a8.5 8.5 0 1 0 10.7 10.7z"/></svg>`,
  eye: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/></svg>`,
  eyeOff: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17.9 17.9A10.4 10.4 0 0 1 12 19c-6.4 0-10-7-10-7a18.5 18.5 0 0 1 5.1-5.9M9.9 5.2A9.6 9.6 0 0 1 12 5c6.4 0 10 7 10 7a18.6 18.6 0 0 1-2.2 3.2M14.1 14.1a3 3 0 1 1-4.2-4.2M2 2l20 20"/></svg>`,
  cal: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><rect x="3" y="5" width="18" height="16" rx="3"/><path d="M3 10h18M8 3v4M16 3v4"/></svg>`,
  plus: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>`,
  check: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>`,
  hour: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M6 2h12M6 22h12M7 2v4l5 6-5 6v4M17 2v4l-5 6 5 6v4"/></svg>`,
  users: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><circle cx="9" cy="8" r="3.5"/><path d="M2.5 20a6.5 6.5 0 0 1 13 0M16 4.5a3.5 3.5 0 0 1 0 7M21.5 20a6.5 6.5 0 0 0-4-6"/></svg>`,
  chev: `<svg class="chev" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M9 6l6 6-6 6"/></svg>`,
  search: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>`,
};

/* =========================================================
   LIVE WALLPAPER + THEME
   ========================================================= */
(function setupChrome(){
  const wall = document.createElement("div");
  wall.className = "wallpaper";
  wall.innerHTML = `<div class="blob b1"></div><div class="blob b2"></div><div class="blob b3"></div><div class="blob b4"></div><div class="blob b5"></div>`;
  const grain = document.createElement("div");
  grain.className = "grain";
  document.body.prepend(grain);
  document.body.prepend(wall);

  // parallax: wallpaper follows the mouse / phone tilt (like iOS depth effect)
  const move = (x, y) => {
    wall.style.setProperty("--px", `${x * -28}px`);
    wall.style.setProperty("--py", `${y * -28}px`);
  };
  window.addEventListener("pointermove", e => move(e.clientX / innerWidth - .5, e.clientY / innerHeight - .5), { passive:true });
  window.addEventListener("deviceorientation", e => {
    if(e.gamma == null) return;
    move(Math.max(-1, Math.min(1, e.gamma / 30)) / 2, Math.max(-1, Math.min(1, (e.beta - 45) / 30)) / 2);
  }, { passive:true });

  // theme toggle
  const bar = document.createElement("div");
  bar.className = "topbar";
  bar.innerHTML = `<button class="icon-btn" id="themeBtn" aria-label="Toggle light / dark"></button>`;
  document.body.appendChild(bar);

  const root = document.documentElement;
  const sysDark = matchMedia("(prefers-color-scheme: dark)");
  let saved = null;
  try{ saved = localStorage.getItem("ot-theme"); }catch(e){}
  if(saved) root.dataset.theme = saved;

  const isDark = () => root.dataset.theme ? root.dataset.theme === "dark" : sysDark.matches;
  const paint = () => { $("themeBtn").innerHTML = isDark() ? I.sun : I.moon; };
  paint();
  sysDark.addEventListener?.("change", paint);
  $("themeBtn").onclick = () => {
    root.dataset.theme = isDark() ? "light" : "dark";
    try{ localStorage.setItem("ot-theme", root.dataset.theme); }catch(e){}
    paint();
  };

  // lock-screen clock
  const clock = document.querySelector(".lock-clock");
  if(clock){
    const tick = () => {
      const n = new Date();
      clock.querySelector(".date").textContent = n.toLocaleDateString("en-GB", { weekday:"long", day:"numeric", month:"long" });
      const h = n.getHours() % 12 || 12, mm = String(n.getMinutes()).padStart(2, "0");      // 12 ghante wali ghari
      clock.querySelector(".time").innerHTML = `${h}:${mm}<span class="ampm">${n.getHours() < 12 ? "AM" : "PM"}</span>`;
    };
    tick(); setInterval(tick, 1000);
  }

  // password eye icon
  const t = $("togglePass");
  if(t) t.innerHTML = I.eye;
})();

function togglePassword(){
  const input = $("code"), toggle = $("togglePass");
  const show = input.type === "password";
  input.type = show ? "text" : "password";
  toggle.classList.toggle("active", show);
  toggle.innerHTML = show ? I.eyeOff : I.eye;
}

/* ---------------- UI pieces ---------------- */
const notice = msg => `<div class="notice"><span class="dot"></span>${esc(msg)}</div>`;

async function withLoading(btn, label, fn){
  const old = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = `<span class="spinner"></span>${label}`;
  try{ await fn(); } finally { btn.disabled = false; btn.innerHTML = old; }
}

function calBadge(date, green){
  const p = parseDate(date);
  const m = p ? MONTHS[p.getMonth()] : "—";
  const d = p ? p.getDate() : "?";
  return `<div class="cal${green ? " green" : ""}"><span class="m">${m}</span><span class="d">${d}</span></div>`;
}

function leaveList(items, kind){
  if(!items.length){
    return `<div class="list"><p class="empty-note">${kind === "extra" ? "No extra leaves earned yet." : "No leaves taken yet."}</p></div>`;
  }
  return `<div class="list">${items.map((l, i) => `
    <div class="list-item" style="animation-delay:${i * 40}ms">
      ${l.noDate ? `<div class="cal green"><span class="m">Extra</span><span class="d">+</span></div>` : calBadge(l.date, kind === "extra")}
      <div class="li-body">
        <div class="li-title">${esc(l.reason) || (kind === "extra" ? "Worked on off day" : "Leave")}</div>
        <div class="li-sub">${l.noDate ? "Extra leaves added in sheet" : esc(fmtLong(l.date))}</div>
      </div>
      ${kind === "extra" ? `<span class="badge">+${l.days || 1}</span>` : (num(l.days) > 1 ? `<span class="badge red">${num(l.days)} days</span>` : "")}
    </div>`).join("")}</div>`;
}

function ring(s){
  const R = 52, C = 2 * Math.PI * R;
  const usedPct = s.total ? Math.min(1, s.usedCount / s.total) : 0;
  const extraPct = s.total ? s.extraAll / s.total : 0;
  return `
    <div class="ring">
      <svg viewBox="0 0 120 120">
        <defs>
          <linearGradient id="gUsed" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="var(--blue)"/><stop offset="1" stop-color="var(--purple)"/></linearGradient>
        </defs>
        <circle class="track" cx="60" cy="60" r="${R}" fill="none" stroke-width="12"/>
        <circle class="prog" cx="60" cy="60" r="${R}" fill="none" stroke="var(--green)" stroke-opacity=".35" stroke-width="12" stroke-linecap="round"
          stroke-dasharray="${C}" stroke-dashoffset="${C}" data-target="${C * (1 - Math.min(extraPct, 1 - usedPct))}"
          style="transform:rotate(${usedPct * 360}deg); transform-origin:60px 60px; ${extraPct ? "" : "display:none"}"/>
        <circle class="prog" cx="60" cy="60" r="${R}" fill="none" stroke="${s.remaining < 0 ? "var(--red)" : "url(#gUsed)"}" stroke-width="12" stroke-linecap="round"
          stroke-dasharray="${C}" stroke-dashoffset="${C}" data-target="${C * (1 - usedPct)}"/>
      </svg>
      <div class="center"><div><b class="count-up${s.remaining < 0 ? " neg" : ""}" data-to="${s.remaining}">0</b><small>${s.remaining < 0 ? "Over limit" : "Left"}</small></div></div>
    </div>`;
}

function animateNumbers(scope){
  requestAnimationFrame(() => requestAnimationFrame(() => {
    scope.querySelectorAll(".prog[data-target]").forEach(c => c.style.strokeDashoffset = c.dataset.target);
    scope.querySelectorAll(".bar i[data-w], .pbar i[data-w]").forEach(b => b.style.width = b.dataset.w);
  }));
  scope.querySelectorAll(".count-up").forEach(el => {
    const to = num(el.dataset.to), start = performance.now(), dur = 1100;
    const step = t => {
      const k = Math.min(1, (t - start) / dur), e = 1 - Math.pow(1 - k, 3);
      el.textContent = Math.round(to * e * 10) / 10;
      if(k < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  });
}

/* segmented control: 2 ya zyada tabs */
function segmented(id, ...labels){
  return `
    <div class="segmented" id="${id}" data-active="0" style="--n:${labels.length}">
      <span class="thumb"></span>
      ${labels.map((l, i) => `<button type="button" class="${i ? "" : "on"}" data-i="${i}">${l}</button>`).join("")}
    </div>`;
}
function wireSegmented(id, panes){
  const seg = $(id);
  seg.querySelectorAll("button").forEach(btn => btn.onclick = () => {
    const i = btn.dataset.i;
    seg.dataset.active = i;
    seg.style.setProperty("--i", i);
    seg.querySelectorAll("button").forEach(b => b.classList.toggle("on", b === btn));
    panes.forEach((p, idx) => $(p).hidden = String(idx) !== i);
  });
}

/* =========================================================
   EMPLOYEE LOGIN
   ========================================================= */
function login(){
  const btn = document.querySelector("button.primary");
  withLoading(btn, "Signing in…", async () => {
    const id = $("empId").value.trim();
    const pass = $("code").value.trim();
    if(!id || !pass) return ($("result").innerHTML = notice("Enter your ID and password"));
    const res = await api({ action:"login", id, pass });
    if(!res.ok){ forgetLogin("ot-emp"); return ($("result").innerHTML = notice(res.error || "Login failed")); }
    rememberLogin("ot-emp", { id, pass });
    showEmployee(res.employee);
    startAutoRefresh(id, pass, res.employee);
  });
}

/* ---------- Refresh par sign out na ho ----------
   Login sirf is browser TAB mein yaad rehta hai (sessionStorage): tab band = sign out.
   "Sign out" dabane par bhi mit jata hai. */
function rememberLogin(key, v){ try{ sessionStorage.setItem(key, JSON.stringify(v)); }catch(e){} }
function forgetLogin(key){ try{ sessionStorage.removeItem(key); }catch(e){} }
function savedLogin(key){ try{ return JSON.parse(sessionStorage.getItem(key) || "null"); }catch(e){ return null; } }
function signOut(){ forgetLogin("ot-emp"); forgetLogin("ot-admin"); location.reload(); }
window.addEventListener("DOMContentLoaded", () => {
  const e = savedLogin("ot-emp"), a = savedLogin("ot-admin");
  if(e && $("empId") && $("code")){ $("empId").value = e.id; $("code").value = e.pass; login(); }
  else if(a && $("adminUser") && $("adminPass")){ $("adminUser").value = a.user; $("adminPass").value = a.pass; checkAdmin(); }
});

/* ---------- Auto refresh: har 30 second check, naya Slack data aaye to screen update ----------
   (Slack message -> ~1 minute mein Sheet -> 30 second mein yahan) */
const REFRESH_MS = 30 * 1000;      // har 30 second naya data check
let session = null;
const sigOf = e => JSON.stringify([e.attendance, e.overtime, e.leaves, e.extraLeaves, e.extraManual]);
function startAutoRefresh(id, pass, emp){
  const first = !session;
  session = { id, pass, sig: sigOf(emp) };          // sirf memory mein, kahin save nahi hota
  if(!first) return;
  setInterval(refreshEmployee, REFRESH_MS);
  document.addEventListener("visibilitychange", () => { if(!document.hidden) refreshEmployee(); });
}
async function refreshEmployee(){
  if(!session || document.hidden || session.busy) return;
  session.busy = true;
  try{
    const res = await api({ action:"login", id: session.id, pass: session.pass });
    if(!res.ok || !res.employee) return;
    const sig = sigOf(res.employee);
    if(sig === session.sig) return;                  // kuch naya nahi
    session.sig = sig;
    const tabs = ["mainNav", "hrNav", "otSeg", "lvSeg"].map(id => [id, $(id) ? $(id).dataset.active : "0"]);
    const y = window.scrollY;
    showEmployee(res.employee);
    tabs.forEach(([id, i]) => { if(i && i !== "0" && $(id)) $(id).querySelector(`button[data-i="${i}"]`)?.click(); });
    window.scrollTo(0, y);
  } finally { session.busy = false; }
}

function showEmployee(emp){
  const s = summary(emp), m = attModel(emp);
  document.querySelector(".lock-clock")?.classList.add("compact");
  document.querySelectorAll(".login-only").forEach(el => el.hidden = true);
  document.querySelector(".card").classList.add("wide");

  const first = String(emp.name || "").split(" ")[0] || "there";
  $("result").innerHTML = `
    <div class="dash">
      <div class="dash-head">
        <div>
          <div class="hello">Welcome back</div>
          <h3>${esc(first)}</h3>
        </div>
        <div class="avatar">${esc(initials(emp.name))}</div>
      </div>

      ${otNotice()}
      ${segmented("mainNav", "Leaves", "Working Hours")}
      <div id="vLeaves">${viewLeaves(emp, s, m)}</div>
      <div id="vHours" hidden></div>

      <button class="logout" onclick="signOut()">Sign out</button>
    </div>`;
  mountHours(m, $("vHours"));
  wireSegmented("lvSeg", ["paneUsed", "paneExtra"]);
  const views = ["vLeaves", "vHours"];
  wireSegmented("mainNav", views);
  $("mainNav").querySelectorAll("button").forEach(b => b.addEventListener("click", () => animateNumbers($(views[b.dataset.i]))));
  animateNumbers($("vLeaves"));
}

/* =========================================================
   ADMIN
   ========================================================= */
function checkAdmin(){
  const btn = document.querySelector("button.primary");
  withLoading(btn, "Checking…", async () => {
    const user = $("adminUser").value.trim(), pass = $("adminPass").value.trim();
    const res = await api({ action:"admin", user, pass });
    if(!res.ok){ forgetLogin("ot-admin"); return ($("adminResult").innerHTML = notice(res.error || "Invalid admin credentials")); }
    rememberLogin("ot-admin", { user, pass });
    data = res.employees || {};
    adminIssues = res.issues || [];
    showAdminData();
  });
}

function showAdminData(){
  document.querySelectorAll(".login-only").forEach(el => el.hidden = true);
  document.querySelector(".card").classList.add("wide");
  const box = $("adminResult");
  const ids = Object.keys(data).sort((a, b) => String(data[a].name || a).localeCompare(String(data[b].name || b)));
  if(!ids.length) return (box.innerHTML = `<div class="list"><p class="empty-note">No employee data loaded yet.</p></div>`);

  const all = ids.map(id => ({ id, emp: data[id], s: summary(data[id]) }));
  const totalUsed = all.reduce((t, x) => t + x.s.usedCount, 0);
  const totalExtra = all.reduce((t, x) => t + x.s.extraAll, 0);

  box.innerHTML = `
    <div class="dash">
      <div class="dash-head">
        <div><div class="hello">Admin</div><h3>All employees</h3></div>
        <div class="avatar">OT</div>
      </div>

      <div class="stats-row three">
        <div class="stat"><div class="ico" style="background:var(--blue)">${I.users}</div><span class="num count-up" data-to="${ids.length}">0</span><span class="label">Employees</span></div>
        <div class="stat"><div class="ico" style="background:var(--orange)">${I.hour}</div><span class="num count-up" data-to="${totalUsed}">0</span><span class="label">Leaves taken</span></div>
        <div class="stat extra"><div class="ico" style="background:var(--green)">${I.plus}</div><span class="num">+<span class="count-up" data-to="${totalExtra}">0</span></span><span class="label">Extra earned</span></div>
      </div>

      ${otNotice()}
      ${issuesBlock(adminIssues)}

      <label class="search">${I.search}<input id="empSearch" type="search" placeholder="Search name or ID" autocomplete="off"></label>
      <div id="empList">
        ${all.map(({ id, emp, s }, i) => {
          const pct = s.total ? Math.min(100, s.usedCount / s.total * 100) : 0;
          return `
          <div class="emp" data-q="${esc((emp.name || "") + " " + id).toLowerCase()}" style="animation-delay:${i * 35}ms">
            <div class="emp-head" onclick="this.parentElement.classList.toggle('open')">
              <div class="avatar">${esc(initials(emp.name || id))}</div>
              <div class="emp-meta">
                <div class="emp-name">${esc(emp.name || "Unnamed")}</div>
                ${String(emp.name || "") !== id ? `<div class="emp-id">${esc(id)}</div>` : ""}
                <div class="bar"><i class="${pct >= 80 ? "warn" : ""}" style="width:0" data-w="${pct}%"></i></div>
              </div>
              <div class="emp-count">
                <b class="${s.remaining < 0 ? "neg" : ""}">${s.usedCount}/${s.total}</b>
                ${s.extraAll ? `<small>+${s.extraAll} extra</small>` : ""}
              </div>
              ${I.chev}
            </div>
            <div class="emp-body"><div><div class="inner">
              ${adminAttBlock(emp)}
              <div class="sub-h">Leaves taken</div>
              ${leaveList(s.used, "used")}
              <div class="sub-h">Extra earned (off day + overtime)</div>
              ${s.extra.items.length || s.otRewards || !s.extraAll
                ? leaveList(s.extra.items.concat(attModel(emp).rewards).sort(sortByDate), "extra")
                : `<div class="list"><p class="empty-note">+${s.extraAll} extra leave${s.extraAll > 1 ? "s" : ""}</p></div>`}
            </div></div></div>
          </div>`;
        }).join("")}
      </div>
      <button class="logout" onclick="signOut()">Sign out</button>
    </div>`;

  $("empSearch").oninput = e => {
    const q = e.target.value.trim().toLowerCase();
    box.querySelectorAll(".emp").forEach(el => el.hidden = q && !el.dataset.q.includes(q));
  };
  animateNumbers(box);
}


/* =========================================================
   WORKING HOURS  (Slack check in / check out se — koi nayi entry nahi)
   Server data: emp.attendance = [{date, firstIn, lastOut, work, brk, namaz,
                ot, short, required, status, final, openSince}]  (minutes)
                -> Google Sheet "Daily Summary" tab (Attendance.gs banata hai)

   Hisaab (Attendance.gs ke runMonth jaisa — dono same rakhein):
   - Worked   = check in se check out tak ka waqt (break nahi, namaz kaam mein ginti hai)
   - Required = 8h har Mon–Sat din jis ka check in + check out poora ho
                Sunday weekly off: required 0, saara kaam overtime
   - Din ka overtime = worked − 8h (agar zyada) · short = 8h − worked (agar kam)
   - Check out na ho to din nahi gina jata (admin Raw Log mein theek kare)
   - Balance = Carried in + Overtime − Short (din-ba-din)
     Jis din balance 8h ho: 1 extra leave, −8h. Baqi agle din / mahine.
   - Permanent record Google Sheet mein: "Overtime Months" + "Earned Leaves".
     "Overtime History" mein sirf aakhri 2 haftay.
   ========================================================= */
const ATT = {
  TARGET: 480,                    // 8h
  OFF_DAYS: [0],                  // 0 = Sunday
  PER_LEAVE: 480,                 // 8h overtime = 1 extra leave
  ALLOWED_BREAK: 60,              // Attendance.gs ALLOWED_BREAK_MINUTES (sirf dikhane ke liye)
  OT_START: "2026-10-24",         // overtime / short ISI DIN se (Attendance.gs OVERTIME_START_DATE jaisa)
  NAMAZ_COUNTS_AS_WORK: true,     // Attendance.gs CONFIG ke sath same rakhein (9h office = 8h kaam + 1h break)
  SHORT_REDUCES_OVERTIME: true,   // Attendance.gs CONFIG ke sath same rakhein
  DAY_START_HOUR: 5,
  LIVE_MAX_MIN: 14 * 60
};

const otStartLabel = () => { const d = parseDate(ATT.OT_START); return d ? d.toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric" }) : ATT.OT_START; };
const otNotice = () => `<div class="info-note"><b>Note:</b> Overtime and short time are counted from <b>${otStartLabel()}</b>. Hours worked before this date are shown, but not counted as overtime or short time.</div>`;

const fmtMin = m => {
  m = Math.round(Math.abs(num(m)));
  const h = Math.floor(m / 60), mm = m % 60;
  return h ? `${h}h${mm ? " " + String(mm).padStart(2, "0") + "m" : ""}` : `${mm}m`;
};
const signed = m => (Math.round(m) > 0 ? "+" : Math.round(m) < 0 ? "−" : "") + fmtMin(m);
const tone = m => (Math.round(m) > 0 ? "pos" : Math.round(m) < 0 ? "neg" : "");
const ymd = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
const startOfDay = d => { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; };
const addDays = (d, n) => { const x = new Date(d); x.setDate(x.getDate() + n); return x; };
const weekStart = d => addDays(startOfDay(d), -((d.getDay() + 6) % 7));            // Monday
const workToday = () => startOfDay(new Date(Date.now() - ATT.DAY_START_HOUR * 36e5));  // 5 AM tak pichla din
const dayName = (d, len) => d.toLocaleDateString("en-GB", { weekday: len || "short" });
const shortDate = d => d.toLocaleDateString("en-GB", { day: "numeric", month: "short" });
const longDate = d => d.toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
const monthLabel = d => d.toLocaleDateString("en-GB", { month: "long", year: "numeric" });
const monthShort = d => { const x = MONTHS[d.getMonth()]; return x[0].toUpperCase() + x.slice(1); };
const leavesTxt = n => `${n} extra leave${n === 1 ? "" : "s"}`;

/* "9:05 AM" -> minutes since midnight */
function clockMin(s){
  const m = String(s || "").match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if(!m) return null;
  let h = +m[1] % 12; if(/pm/i.test(m[3])) h += 12;
  return h * 60 + +m[2];
}

/* Ek mahine ka hisaab, din-ba-din — Attendance.gs runMonth() jaisa.
   (Server ka "Overtime Months" / "Earned Leaves" data mile to wohi use hota hai.) */
function runMonth(carryIn, days){
  const red = ATT.SHORT_REDUCES_OVERTIME;
  let bal = carryIn, ot = 0, short = 0;
  const events = [];
  days.forEach(x => {
    if(!x.final) return;
    ot += x.ot; short += x.short;
    bal += x.ot - (red ? x.short : 0);
    while(bal >= ATT.PER_LEAVE){ bal -= ATT.PER_LEAVE; events.push({ date: x.key, after: bal, type: "Earned" }); }
    while(bal <= -ATT.PER_LEAVE){ bal += ATT.PER_LEAVE; events.push({ date: x.key, after: bal, type: "Cut" }); }
  });
  const leaves = events.filter(e => e.type === "Earned").length;
  return { carryIn, ot, short, total: carryIn + ot - (red ? short : 0), leaves, cut: events.length - leaves, carryOut: bal, events };
}

const KIND = {
  done: "Completed", half: "Half day", offWorked: "Sunday · Worked", live: "In progress", incomplete: "No check-out",
  leave: "Leave", off: "Weekly off", absent: "No check-in", notyet: "Not checked in", upcoming: "", nodata: ""
};
const LIVE_TEXT = { "Working now": "Working", "On break": "On break", "At namaz": "At namaz", "Half day": "Half day (checked out)" };

/* Ek din ka record banao (attendance, leave, off, absent…) */
function makeDay(d, r, leave, flags){
  const off = ATT.OFF_DAYS.includes(d.getDay());
  const day = { d, key: ymd(d), off, r, leave, isToday: !!flags.today, preStart: ymd(d) < ATT.OT_START,
    work: 0, brk: 0, namaz: 0, ot: 0, short: 0, required: 0, net: 0, final: false };
  if(r){
    Object.assign(day, { work: r.work, brk: r.brk, namaz: r.namaz, firstIn: r.firstIn, lastOut: r.lastOut, late: r.late || "" });
    day.required = r.required != null ? r.required : (off ? 0 : ATT.TARGET);
    if(r.final){
      day.final = true;
      if(!day.preStart){ day.ot = r.ot; day.short = r.short; day.net = r.ot - r.short; }   // start se pehle: 0
      day.halfDay = !!r.halfDay; day.relief = day.preStart ? 0 : num(r.relief);
      day.kind = off ? "offWorked" : day.halfDay ? "half" : "done";
      if(/final check out/i.test(r.status || "")) day.note = "Last check-out was for a break";
    } else if(flags.today){
      day.kind = "live"; day.halfDay = !!r.halfDay;
      day.liveStatus = LIVE_TEXT[r.status] || r.status || "In progress";
      if(r.openSince && r.status === "Working now"){
        const extra = (Date.now() - r.openSince) / 60000;            // abhi chal raha session
        if(extra > 0 && extra < ATT.LIVE_MAX_MIN) day.work += extra;
      }
    } else day.kind = "incomplete";
  }
  else if(flags.nodata) day.kind = "nodata";
  else if(flags.future) day.kind = off ? "off" : "upcoming";
  else if(leave) day.kind = "leave";
  else if(off) day.kind = "off";
  else if(flags.today){ day.kind = "notyet"; day.required = ATT.TARGET; }
  else day.kind = "absent";
  if(day.kind === "live" || day.kind === "notyet"){
    const diff = day.work - day.required;
    day.remaining = Math.max(-diff, 0); day.extraNow = Math.max(diff, 0);
  }
  return day;
}

/* Sirf poore (final) din gine jate hain — is liye Worked − Required = Overtime − Short hamesha */
function totals(days){
  const t = { required: 0, work: 0, brk: 0, namaz: 0, ot: 0, short: 0, worked: 0, absent: 0, leave: 0, offWorked: 0, incomplete: 0, pre: 0, half: 0, relief: 0 };
  days.forEach(x => {
    if(x.final){
      t.required += x.required; t.work += x.work; t.brk += x.brk; t.namaz += x.namaz;
      t.ot += x.ot; t.short += x.short; t.worked++;
      if(x.preStart) t.pre++;
      if(x.halfDay){ t.half++; t.relief += x.relief || 0; }
    }
    if(x.kind === "absent") t.absent++;
    if(x.kind === "leave") t.leave++;
    if(x.kind === "offWorked") t.offWorked++;
    if(x.kind === "incomplete") t.incomplete++;
  });
  t.net = t.ot - t.short;
  return t;
}

const _models = new WeakMap();
function attModel(emp){
  if(_models.has(emp)) return _models.get(emp);
  const recs = new Map();
  (emp.attendance || []).forEach(r => {
    const d = parseDate(r.date); if(!d) return;
    recs.set(ymd(d), { ...r, work: num(r.work), brk: num(r.brk), namaz: num(r.namaz), ot: num(r.ot), short: num(r.short),
      required: r.required == null || r.required === "" ? null : num(r.required), openSince: num(r.openSince) });
  });
  const leaveMap = new Map();
  getUsed(emp).forEach(l => { const d = parseDate(l.date); if(d) leaveMap.set(ymd(d), l); });

  const today = workToday(), tk = ymd(today);
  const keys = [...recs.keys()].sort();
  const first = keys.length ? parseDate(keys[0]) : null;
  const days = [], byKey = new Map();
  if(first){
    for(let d = new Date(first); d <= today; d = addDays(d, 1)){
      const k = ymd(d);
      const day = makeDay(d, recs.get(k), leaveMap.get(k), { today: k === tk });
      days.push(day); byKey.set(k, day);
    }
  }

  // Mahine ka statement (bank statement jaisa): carry in -> leaves -> carry forward
  const months = [];
  let earned = [];
  const monthOf = key => { const [y, mo] = key.split("-").map(Number); return new Date(y, mo - 1, 1); };
  const srv = emp.overtime && Array.isArray(emp.overtime.months) ? emp.overtime : null;   // Google Sheet (permanent)
  const firstMonth = [first, ...(srv ? srv.months.map(x => monthOf(x.month)) : [])].filter(Boolean).sort((a, b) => a - b)[0];
  if(firstMonth){
    let carry = 0;
    for(let ms = new Date(firstMonth.getFullYear(), firstMonth.getMonth(), 1); ms <= today; ms = new Date(ms.getFullYear(), ms.getMonth() + 1, 1)){
      const me = new Date(ms.getFullYear(), ms.getMonth() + 1, 0), key = ymd(ms).slice(0, 7);
      const inside = days.filter(x => x.d >= ms && x.d <= me);
      const sv = srv && srv.months.find(x => x.month === key);
      const r = sv ? { carryIn: sv.carryIn, ot: sv.ot, short: sv.short, total: sv.total, leaves: sv.leaves, cut: num(sv.cut), carryOut: sv.carryOut, events: [] }
                   : runMonth(carry, inside);
      if(!sv && !inside.length && !carry) continue;
      const current = today <= me;
      months.push({ key, start: ms, end: me, days: inside, t: totals(inside), current,
        status: sv ? sv.status : (current ? "Running" : "Final"), ...r });
      if(!srv) r.events.forEach(e => earned.push({ date: e.date, month: key, used: ATT.PER_LEAVE, after: e.after, final: !current, type: e.type }));
      carry = r.carryOut;
    }
  }
  if(srv) earned = (srv.earned || []).map(e => ({ date: e.date, month: e.month, used: num(e.used) || ATT.PER_LEAVE, after: num(e.after), final: e.status === "Final",
    type: e.type === "Cut" || num(e.leaves) < 0 ? "Cut" : "Earned" }));
  earned.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
  // earned = poori leave history (+1 mili / −1 kati); rewards = sirf +1 wali
  const rewards = earned.filter(e => e.type !== "Cut").map(e => ({ date: e.date, reason: `From overtime (8h) — ${monthLabel(monthOf(e.month))}`, days: 1 }));
  const cuts = earned.filter(e => e.type === "Cut");

  // Overtime history: sirf aakhri 2 haftay (server "Overtime History" tab, warna yahin se)
  const histFrom = addDays(weekStart(today), -7);
  const otHistory = srv && Array.isArray(srv.history)
    ? srv.history.map(h => ({ d: parseDate(h.date), work: num(h.work), required: num(h.required), ot: num(h.ot), short: num(h.short) })).filter(h => h.d)
    : days.filter(x => x.final && !x.preStart && x.d >= histFrom).map(x => ({ d: x.d, work: x.work, required: x.required, ot: x.ot, short: x.short }));

  const m = { days, byKey, recs, leaveMap, months, earned, rewards, cuts, otHistory, first, today,
    balance: months.length ? months[months.length - 1].carryOut : 0,
    todayDay: byKey.get(tk) || makeDay(today, null, leaveMap.get(tk), { today: true }) };
  _models.set(emp, m);
  return m;
}

/* Kisi bhi date ka din */
function dayAt(m, d){
  const k = ymd(d);
  if(m.byKey.has(k)) return m.byKey.get(k);
  if(d > m.today) return makeDay(d, null, m.leaveMap.get(k), { future: true });
  if(k === ymd(m.today)) return m.todayDay;
  return makeDay(d, null, m.leaveMap.get(k), { nodata: true });
}

/* ---------- shared UI ---------- */
function sumPanel(t){
  const extra = [
    t.pre ? `${t.pre} before ${otStartLabel()} (overtime not counted)` : "",
    t.half ? `<span class="pos">Half day relief: ${t.half} day${t.half > 1 ? "s" : ""}</span>` : "",
    t.leave ? `${t.leave} leave` : "",
    t.incomplete ? `<span class="neg">${t.incomplete} day${t.incomplete > 1 ? "s" : ""} not counted (no check-out)</span>` : ""
  ].filter(Boolean).join(" · ");
  return `
    <div class="sum4">
      <div><small>Required</small><b>${fmtMin(t.required)}</b></div>
      <div><small>Worked</small><b>${fmtMin(t.work)}</b></div>
      <div><small>Overtime</small><b class="pos">${t.ot ? "+" + fmtMin(t.ot) : "0m"}</b></div>
      <div><small>Short</small><b class="${t.short ? "neg" : ""}">${t.short ? "−" + fmtMin(t.short) : "0m"}</b></div>
    </div>
    ${breakNamaz(t.brk, t.namaz, t.worked, false)}
    <div class="sum-sub">${t.worked} day${t.worked === 1 ? "" : "s"} counted${extra ? " · " + extra : ""}</div>`;
}

/* Break aur Namaz — alag saaf boxes */
function breakNamaz(brk, namaz, days, single){
  const avg = v => !single && days ? `<small class="bn-sub">avg ${fmtMin(v / days)} / day</small>` : "";
  const over = !single ? ""
    : brk > ATT.ALLOWED_BREAK ? `<small class="bn-sub neg">${fmtMin(brk - ATT.ALLOWED_BREAK)} over ${fmtMin(ATT.ALLOWED_BREAK)} allowed</small>`
    : `<small class="bn-sub">${fmtMin(ATT.ALLOWED_BREAK)} allowed</small>`;
  return `
    <div class="bn2">
      <div><span class="bn-dot brk"></span><small>${single ? "Break" : "Total break"}</small><b>${fmtMin(brk)}</b>${over || avg(brk)}</div>
      <div><span class="bn-dot nmz"></span><small>${single ? "Namaz" : "Total namaz"}</small><b>${fmtMin(namaz)}</b>${avg(namaz)}</div>
    </div>`;
}

/* Bank-statement style: kahan se aaya, kahan gaya */
function statement(mo, next){
  const prevName = monthShort(new Date(mo.start.getFullYear(), mo.start.getMonth() - 1, 1));
  return `
    <div class="bal">
      <div><span>Carried in from ${prevName}</span><b class="${tone(mo.carryIn)}">${signed(mo.carryIn)}</b></div>
      <div><span>Overtime earned</span><b class="pos">+${fmtMin(mo.ot)}</b></div>
      ${ATT.SHORT_REDUCES_OVERTIME ? `<div><span>Short time</span><b class="${mo.short ? "neg" : ""}">${mo.short ? "−" + fmtMin(mo.short) : "0m"}</b></div>` : ""}
      <div class="hl"><span>Total overtime</span><b class="${tone(mo.total)}">${signed(mo.total)}</b></div>
      <div><span>Extra leaves (each 8h overtime = +1)</span><b class="${mo.leaves ? "pos" : ""}">${mo.leaves ? `+${mo.leaves} · −${mo.leaves * 8}h` : "0"}</b></div>
      ${mo.cut ? `<div><span>Paid leaves cut (each 8h short = −1)</span><b class="neg">−${mo.cut} · +${mo.cut * 8}h</b></div>` : ""}
      <div class="hl"><span>Carried forward to ${next}</span><b class="${tone(mo.carryOut)}">${signed(mo.carryOut)}</b></div>
    </div>`;
}

function nextLeavePanel(m){
  const pool = m.balance, neg = pool < 0;
  const pct = Math.min(1, Math.abs(pool) / ATT.PER_LEAVE);
  return `
    <div class="apanel">
      <div class="ph"><span>${neg ? "Short time towards losing a paid leave" : "Overtime towards next extra leave"}</span>
        <b class="${tone(pool)}">${neg ? "−" + fmtMin(pool) : fmtMin(pool)} / 8h</b></div>
      <div class="pbar${neg ? " red" : ""}"><i style="width:0" data-w="${Math.round(pct * 100)}%"></i></div>
      <div class="pnote">${neg
        ? `<b class="neg">${fmtMin(ATT.PER_LEAVE + pool)}</b> more short time and <b>1 paid leave</b> will be cut. Overtime fills this back first — save your paid leave.`
        : `<b>${fmtMin(ATT.PER_LEAVE - pool)}</b> more overtime for your next extra leave.`}</div>
    </div>`;
}

function rightPart(day){
  if(day.kind === "live") return day.remaining
    ? `<span class="anet">${fmtMin(day.remaining)} left</span>`
    : `<span class="anet pos">+${fmtMin(day.extraNow)}</span>`;
  if(day.final && day.preStart) return `<span class="anet muted">Not counted</span>`;
  if(day.final && day.halfDay && day.relief) return `<span class="anet pos">Relief</span>`;
  if(day.final) return `<span class="anet ${tone(day.net)}">${day.net ? signed(day.net) : "On target"}</span>`;
  return "";
}

function dayRow(day, opts = {}){
  const d = day.d;
  const status = day.kind === "live" ? day.liveStatus : KIND[day.kind];
  let main, sub = "";
  if(day.r){
    main = `${fmtMin(day.work)}<span class="of"> / ${day.required ? fmtMin(day.required) : "Sunday"}</span>`;
    const time = day.firstIn ? `${esc(day.firstIn)} – ${esc(day.lastOut || "…")}` : "";
    sub = [time, day.brk ? `Break ${fmtMin(day.brk)}` : "", day.namaz ? `Namaz ${fmtMin(day.namaz)}` : ""].filter(Boolean).join(" · ");
    if(day.kind === "incomplete") sub += ` · <span class="neg">Not counted</span>`;
    if(day.halfDay) sub += ` · <span class="pos">Half day — worked ${fmtMin(day.work)}</span>`;
    if(day.late === "Y") sub += ` · <span class="neg">Late</span>`;
    if(day.late === "Informed") sub += ` · <span class="muted">Late (informed)</span>`;
  } else if(day.kind === "leave"){
    main = `<span class="muted-main">Leave</span>`;
    sub = esc(day.leave.reason || "");
  } else if(day.kind === "upcoming" || day.kind === "nodata"){
    main = `<span class="muted-main">—</span>`;
  } else {
    main = `<span class="muted-main${day.kind === "absent" ? " neg" : ""}">${KIND[day.kind]}</span>`;
  }
  return `
    <div class="arow k-${day.kind}${day.isToday ? " is-today" : ""}${opts.click ? " click" : ""}" data-day="${day.key}">
      <div class="adate">${opts.month ? `<small class="mo">${MONTHS[d.getMonth()]}</small>` : ""}<b>${d.getDate()}</b><small>${dayName(d)}</small></div>
      <div class="abody"><div class="amain">${main}</div>${sub ? `<div class="asub">${sub}</div>` : ""}</div>
      <div class="aright">${rightPart(day)}${status && !["absent", "off", "leave", "notyet"].includes(day.kind) ? `<small class="ast">${status}</small>` : ""}</div>
    </div>`;
}

function progressRing(pct, color, centerTop, centerSub){
  const R = 52, C = 2 * Math.PI * R, p = Math.max(0, Math.min(1, pct));
  return `
    <div class="ring sm">
      <svg viewBox="0 0 120 120">
        <circle class="track" cx="60" cy="60" r="${R}" fill="none" stroke-width="11"/>
        <circle class="prog" cx="60" cy="60" r="${R}" fill="none" stroke="${color}" stroke-width="11" stroke-linecap="round"
          stroke-dasharray="${C}" stroke-dashoffset="${C}" data-target="${C * (1 - p)}"/>
      </svg>
      <div class="center"><div><b>${centerTop}</b><small>${centerSub}</small></div></div>
    </div>`;
}

/* Ek din ka hisaab, line by line */
function dayCalc(t){
  if(!t.r){
    const msg = t.kind === "leave" ? "On leave — no hours required."
      : t.kind === "off" ? "Sunday is the weekly off. If you work on Sunday, all hours count as overtime."
      : t.kind === "notyet" ? "No check-in on Slack yet today."
      : t.kind === "upcoming" ? "This day has not come yet."
      : "No check-in found on Slack for this day.";
    return `<p class="pnote">${msg}</p>`;
  }
  const rows = [];
  const a = clockMin(t.firstIn), b = clockMin(t.lastOut);
  if(t.kind !== "live" && a != null && b != null){
    const span = b >= a ? b - a : b + 1440 - a;
    rows.push(`<div><span>From first check-in to last check-out<br><small>${esc(t.firstIn)} → ${esc(t.lastOut)}</small></span><b>${fmtMin(span)}</b></div>`);
    if(t.brk) rows.push(`<div><span>Break (not counted)</span><b>−${fmtMin(t.brk)}</b></div>`);
    if(t.namaz && !ATT.NAMAZ_COUNTS_AS_WORK) rows.push(`<div><span>Namaz (not counted)</span><b>−${fmtMin(t.namaz)}</b></div>`);
    const away = span - Math.round(t.work) - Math.round(t.brk) - (ATT.NAMAZ_COUNTS_AS_WORK ? 0 : Math.round(t.namaz));
    if(away >= 3) rows.push(`<div><span>Other time checked out</span><b>−${fmtMin(away)}</b></div>`);
  } else if(t.kind === "live"){
    rows.push(`<div><span>Checked in at</span><b>${esc(t.firstIn || "—")}</b></div>`);
    if(t.brk) rows.push(`<div><span>Break so far (not counted)</span><b>−${fmtMin(t.brk)}</b></div>`);
    if(t.namaz && !ATT.NAMAZ_COUNTS_AS_WORK) rows.push(`<div><span>Namaz so far (not counted)</span><b>−${fmtMin(t.namaz)}</b></div>`);
  }
  rows.push(`<div class="hl"><span>Worked${t.kind === "live" ? " so far" : ""}${t.namaz && ATT.NAMAZ_COUNTS_AS_WORK ? `<br><small>includes namaz ${fmtMin(t.namaz)}</small>` : ""}</span><b>${fmtMin(t.work)}</b></div>`);
  rows.push(`<div><span>Required${t.off ? " (Sunday — weekly off)" : ""}</span><b>−${fmtMin(t.required)}</b></div>`);
  let end, note = "";
  if(t.final && t.preStart){
    end = `<div class="hl"><span>Overtime / short time</span><b class="muted">Not counted</b></div>`;
    note = `Overtime and short time are counted from ${otStartLabel()}.`;
  } else if(t.final && t.halfDay && t.relief){
    rows.push(`<div><span>Half day relief (short time not counted)</span><b class="pos">+${fmtMin(t.relief)}</b></div>`);
    end = `<div class="hl"><span>Result</span><b>On target (half day)</b></div>`;
    note = `Checked out as half day after ${fmtMin(t.work)} of work. The missing ${fmtMin(t.relief)} is not counted as short time.`;
  } else if(t.final){
    end = `<div class="hl"><span>${t.net > 0 ? "Overtime" : t.net < 0 ? "Short time" : "Result"}</span><b class="${tone(t.net)}">${t.net ? signed(t.net) : "On target"}</b></div>`;
    if(t.note) note = `${t.note} — the day was closed at that check-out.`;
  } else if(t.kind === "live" && t.halfDay){
    end = `<div class="hl"><span>Half day</span><b>Checked out</b></div>`;
    note = "Checked out as half day. Short time will not be counted for today.";
  } else if(t.kind === "live"){
    end = `<div class="hl"><span>${t.remaining ? "Still needed today" : "Extra so far"}</span><b class="${t.remaining ? "" : "pos"}">${t.remaining ? fmtMin(t.remaining) : "+" + fmtMin(t.extraNow)}</b></div>`;
    note = "Today is still running. Overtime is added after the day ends.";
  } else {
    end = `<div class="hl"><span>Counted</span><b class="neg">Not counted</b></div>`;
    note = "There is no check-out on Slack for this day, so it is not counted. Ask the admin to fix it.";
  }
  rows.push(end);
  return `<div class="bal calc">${rows.join("")}</div>${note ? `<p class="pnote">${note}</p>` : ""}`;
}

/* ---------- WORKING HOURS · DAILY ---------- */
function mountDaily(m, el){
  let cur = m.today;
  const render = () => {
    const t = dayAt(m, cur);
    const isToday = ymd(cur) === ymd(m.today);
    const statusText = t.kind === "live" ? t.liveStatus : t.kind === "notyet" ? "Not checked in yet" : (KIND[t.kind] || "No record");
    const req = t.required || 0;
    const pct = req ? t.work / req : (t.work ? 1 : 0);
    const ring = t.off && !t.r
      ? progressRing(0, "var(--green)", "Off", "Sunday")
      : progressRing(pct, pct >= 1 ? "var(--green)" : "var(--blue)", fmtMin(t.work), req ? `of ${fmtMin(req)}` : "Sunday");
    el.innerHTML = `
      <div class="navrow">
        <button type="button" class="nb" data-d="-1" aria-label="Previous day">‹</button>
        <label class="daypick"><b>${isToday ? "Today" : esc(dayName(cur, "long"))}</b><small>${esc(cur.toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric" }))}</small>
          <input type="date" id="dPick" value="${ymd(cur)}" max="${ymd(m.today)}" aria-label="Pick a date"></label>
        <button type="button" class="nb" data-d="1" aria-label="Next day" ${isToday ? "disabled" : ""}>›</button>
      </div>
      <div class="apanel today">
        <div class="t-status k-${t.kind}"><i></i>${esc(statusText)}</div>
        <div class="today-main" style="margin-top:12px">${ring}
          <div class="facts">
            <div><small>Check in</small><b>${esc(t.firstIn || "—")}</b></div>
            <div><small>Last check-out</small><b>${esc(t.lastOut || "—")}</b></div>
            <div><small>Required</small><b>${t.r || t.kind === "notyet" ? (req ? fmtMin(req) : "0h (Sunday)") : "—"}</b></div>
            <div><small>${t.final ? (t.net < 0 ? "Short" : "Overtime") : "Overtime"}</small><b class="${t.preStart ? "muted" : tone(t.net)}">${t.final ? (t.preStart ? "Not counted" : t.net ? signed(t.net) : "0m") : "—"}</b></div>
          </div>
        </div>
        ${t.r ? breakNamaz(t.brk, t.namaz, 1, true) : ""}
      </div>
      <div class="sub-h">How this day was calculated</div>
      ${dayCalc(t)}`;
    el.querySelectorAll(".nb").forEach(bt => bt.onclick = () => { cur = addDays(cur, +bt.dataset.d); if(cur > m.today) cur = m.today; render(); animateNumbers(el); });
    $("dPick").onchange = e => { const d = parseDate(e.target.value); if(d){ cur = d > m.today ? m.today : d; render(); animateNumbers(el); } };
  };
  el.goTo = d => { cur = d; render(); animateNumbers(el); };
  render();
}

/* ---------- WORKING HOURS · WEEKLY ---------- */
function mountWeekly(m, el, openDay){
  let off = 0;
  const render = () => {
    const ws = addDays(weekStart(m.today), off * 7), we = addDays(ws, 6);
    const days = []; for(let i = 0; i < 7; i++) days.push(dayAt(m, addDays(ws, i)));
    const t = totals(days);
    el.innerHTML = `
      <div class="navrow">
        <button type="button" class="nb" data-d="-1" aria-label="Previous week">‹</button>
        <div><b>${off === 0 ? "This week" : off === -1 ? "Last week" : `${shortDate(ws)} – ${shortDate(we)}`}</b><small>Mon ${shortDate(ws)} – Sun ${shortDate(we)} ${we.getFullYear()}</small></div>
        <button type="button" class="nb" data-d="1" aria-label="Next week" ${off >= 0 ? "disabled" : ""}>›</button>
      </div>
      <div class="apanel">${sumPanel(t)}
        <div class="bal">
          ${t.pre ? `<div><span>${t.pre} day${t.pre > 1 ? "s" : ""} before ${otStartLabel()} — not counted</span><b class="muted">—</b></div>`
                  : `<div><span>Worked − Required</span><b>${fmtMin(t.work)} − ${fmtMin(t.required)}</b></div>`}
          ${t.relief ? `<div><span>Half day relief (${t.half} day${t.half > 1 ? "s" : ""})</span><b class="pos">+${fmtMin(t.relief)}</b></div>` : ""}
          <div class="hl"><span>Overtime this week${ATT.SHORT_REDUCES_OVERTIME ? " (after short time)" : ""}</span><b class="${tone(ATT.SHORT_REDUCES_OVERTIME ? t.net : t.ot)}">${signed(ATT.SHORT_REDUCES_OVERTIME ? t.net : t.ot)}</b></div>
        </div>
      </div>
      <div class="alist">${days.map(x => dayRow(x, { click: x.kind !== "upcoming" })).join("")}</div>
      <p class="att-note">Tap a day to see how it was calculated.</p>`;
    el.querySelectorAll(".nb").forEach(bt => bt.onclick = () => { off += +bt.dataset.d; render(); animateNumbers(el); });
    el.querySelectorAll(".arow.click").forEach(r => r.onclick = () => openDay(parseDate(r.dataset.day)));
  };
  render();
}

/* ---------- WORKING HOURS · MONTHLY (statement) ---------- */
function mountMonthly(m, el, openDay){
  if(!m.months.length){
    el.innerHTML = `<div class="alist"><p class="empty-note">Working hours will appear here once Slack check-ins are synced.</p></div>`;
    return;
  }
  let idx = m.months.length - 1;
  const render = () => {
    const mo = m.months[idx];
    const nextName = monthShort(new Date(mo.start.getFullYear(), mo.start.getMonth() + 1, 1));
    const shown = mo.days.filter(x => x.kind !== "nodata" && x.kind !== "upcoming").slice().reverse();
    el.innerHTML = `
      <div class="navrow">
        <button type="button" class="nb" data-d="-1" aria-label="Previous month" ${idx <= 0 ? "disabled" : ""}>‹</button>
        <label class="msel">
          <select id="mPick" aria-label="Choose month">${m.months.map((x, i) => `<option value="${i}" ${i === idx ? "selected" : ""}>${monthLabel(x.start)}</option>`).slice().reverse().join("")}</select>
          <small>${mo.status === "Final" ? "Final · saved" : mo.current ? "This month — so far" : "Closing — final in a few days"}</small>
        </label>
        <button type="button" class="nb" data-d="1" aria-label="Next month" ${idx >= m.months.length - 1 ? "disabled" : ""}>›</button>
      </div>
      <div class="apanel">
        ${sumPanel(mo.t)}
        ${ymd(mo.end) < ATT.OT_START
          ? `<p class="pnote">Overtime and short time are counted from <b>${otStartLabel()}</b>. This month is shown for your record only.</p>`
          : statement(mo, nextName)}
        ${mo.status !== "Final" ? `<p class="pnote">This month is not final yet. It is saved for good a few days after the month ends.</p>` : ""}
      </div>
      <div class="sub-h">Days in ${monthLabel(mo.start)}</div>
      <div class="alist">${shown.map(x => dayRow(x, { click: true })).join("") || `<p class="empty-note">No working days recorded.</p>`}</div>
      <div class="sub-h">All months</div>
      <div class="alist">${m.months.map((x, i) => ({ x, i })).reverse().map(({ x, i }) => `
        <div class="arow click mrow${i === idx ? " is-today" : ""}" data-i="${i}">
          <div class="adate"><small class="mo">${x.start.getFullYear()}</small><b>${monthShort(x.start)}</b><small>${x.current ? "so far" : ""}</small></div>
          <div class="abody"><div class="amain">${fmtMin(x.t.work)}<span class="of"> / ${fmtMin(x.t.required)}</span></div>
            <div class="asub">${ymd(x.end) < ATT.OT_START ? "Before overtime tracking" : `Overtime +${fmtMin(x.ot)}${x.short && ATT.SHORT_REDUCES_OVERTIME ? ` · Short −${fmtMin(x.short)}` : ""} · Carried ${signed(x.carryOut)}`}</div></div>
          <div class="aright">${ymd(x.end) < ATT.OT_START ? `<span class="ast">Not counted</span>` : x.leaves ? `<span class="badge">+${x.leaves} leave${x.leaves > 1 ? "s" : ""}</span>` : `<span class="ast">No leave</span>`}</div>
        </div>`).join("")}</div>`;
    el.querySelectorAll(".nb").forEach(bt => bt.onclick = () => { idx += +bt.dataset.d; render(); });
    $("mPick").onchange = e => { idx = +e.target.value; render(); };
    el.querySelectorAll(".arow.mrow").forEach(r => r.onclick = () => { idx = +r.dataset.i; render(); el.scrollIntoView({ behavior: "smooth", block: "start" }); });
    el.querySelectorAll(".arow.click:not(.mrow)").forEach(r => r.onclick = () => openDay(parseDate(r.dataset.day)));
  };
  render();
}

/* ---------- WORKING HOURS · OVERTIME (2-week history + earned leaves) ---------- */
function otRow(h){
  const net = h.ot - (ATT.SHORT_REDUCES_OVERTIME ? h.short : 0);
  return `
    <div class="arow click" data-day="${ymd(h.d)}">
      <div class="adate"><b>${h.d.getDate()}</b><small>${dayName(h.d)}</small></div>
      <div class="abody"><div class="amain">${fmtMin(h.work)}<span class="of"> / ${h.required ? fmtMin(h.required) : "Sunday"}</span></div>
        <div class="asub">Worked / required</div></div>
      <div class="aright"><span class="anet ${tone(net)}">${net ? signed(net) : "0m"}</span><small class="ast">${h.short && !h.ot ? "Short" : "Overtime"}</small></div>
    </div>`;
}
function earnedRow(e){
  const cut = e.type === "Cut";
  return `
    <div class="list-item">
      ${calBadge(e.date, !cut)}
      <div class="li-body">
        <div class="li-title">${cut ? "−1 paid leave cut" : "+1 extra leave"}</div>
        <div class="li-sub">${fmtMin(e.used)} ${cut ? "short time" : "overtime"} · ${esc(monthLabel(parseDate(e.month + "-01")))}${e.final ? "" : ` · <span class="${cut ? "neg" : "pos"}">this month</span>`}</div>
      </div>
      <span class="badge${cut ? " red" : ""}">${cut ? "−1" : "+1"}</span>
    </div>`;
}
function mountOvertime(m, el, openDay){
  const weeks = [];
  const ws0 = weekStart(m.today);
  [ws0, addDays(ws0, -7)].forEach((ws, i) => {
    const we = addDays(ws, 6);
    const list = m.otHistory.filter(h => h.d >= ws && h.d <= we).sort((a, b) => b.d - a.d);
    if(!list.length) return;
    const t = list.reduce((s, h) => ({ work: s.work + h.work, required: s.required + h.required, ot: s.ot + h.ot, short: s.short + h.short }), { work: 0, required: 0, ot: 0, short: 0 });
    weeks.push({ label: i ? "Last week" : "This week", ws, we, list, t });
  });
  const totalLeaves = m.earned.length;
  el.innerHTML = `
    ${nextLeavePanel(m)}
    ${segmented("otSeg", "Overtime history", `Leave history <span class="pill">${totalLeaves}</span>`)}
    <div id="otHist">
      ${weeks.map(w => {
        const net = w.t.ot - (ATT.SHORT_REDUCES_OVERTIME ? w.t.short : 0);
        return `
        <div class="sub-h">${w.label} · ${shortDate(w.ws)} – ${shortDate(w.we)}</div>
        <div class="alist">${w.list.map(otRow).join("")}
          <div class="arow wtotal">
            <div class="adate"><b>Σ</b><small>Total</small></div>
            <div class="abody"><div class="amain">${fmtMin(w.t.work)}<span class="of"> / ${fmtMin(w.t.required)}</span></div>
              <div class="asub">Overtime +${fmtMin(w.t.ot)}${w.t.short && ATT.SHORT_REDUCES_OVERTIME ? ` · Short −${fmtMin(w.t.short)}` : ""}</div></div>
            <div class="aright"><span class="anet ${tone(net)}">${net ? signed(net) : "0m"}</span><small class="ast">Week total</small></div>
          </div>
        </div>`;
      }).join("") || `<div class="alist"><p class="empty-note">No completed days in the last 2 weeks.</p></div>`}
      <p class="att-note">Only the latest 2 weeks are shown here. Your overtime balance and earned leaves are saved permanently.</p>
    </div>
    <div id="otEarned" hidden>
      <div class="list">${m.earned.slice().reverse().map(earnedRow).join("") || `<p class="empty-note">No extra or cut leaves yet.</p>`}</div>
      <p class="att-note">Balance +8h = 1 extra leave. Balance −8h = 1 paid leave cut. Short time is taken from overtime first. The rest carries forward. This history is saved permanently.</p>
    </div>`;
  wireSegmented("otSeg", ["otHist", "otEarned"]);
  el.querySelectorAll(".arow.click").forEach(r => r.onclick = () => openDay(parseDate(r.dataset.day)));
}

/* ---------- WORKING HOURS (Daily · Weekly · Monthly · Overtime) ---------- */
function mountHours(m, el){
  el.innerHTML = `
    ${segmented("hrNav", "Daily", "Weekly", "Monthly", "Overtime")}
    <div id="hDaily"></div><div id="hWeekly" hidden></div><div id="hMonthly" hidden></div><div id="hOvertime" hidden></div>
    <p class="att-note">From your Slack check-in / check-out. Break (up to 1h) is not counted as work; namaz time counts as work.</p>`;
  const openDay = d => {
    if(!d) return;
    $("hrNav").querySelector('button[data-i="0"]').click();
    $("hDaily").goTo(d);
    el.scrollIntoView({ behavior: "smooth", block: "start" });
  };
  mountDaily(m, $("hDaily"));
  mountWeekly(m, $("hWeekly"), openDay);
  mountMonthly(m, $("hMonthly"), openDay);
  mountOvertime(m, $("hOvertime"), openDay);
  const panes = ["hDaily", "hWeekly", "hMonthly", "hOvertime"];
  wireSegmented("hrNav", panes);
  $("hrNav").querySelectorAll("button").forEach(b => b.addEventListener("click", () => animateNumbers($(panes[b.dataset.i]))));
}

/* ---------- LEAVES ---------- */
function viewLeaves(emp, s, m){
  const extraItems = s.extra.items.concat(m.rewards).sort(sortByDate);
  return `
    <div class="widget">
      ${ring(s)}
      <div class="legend">
        <div class="legend-row"><span class="sw" style="background:linear-gradient(135deg,var(--blue),var(--purple))"></span><span class="lbl">Used</span><span class="val">${s.usedCount}</span></div>
        <div class="legend-row"><span class="sw" style="background:var(--green)"></span><span class="lbl">Extra earned</span><span class="val">${s.extraAll ? "+" + s.extraAll : 0}</span></div>
        <div class="legend-row"><span class="sw" style="background:color-mix(in srgb,var(--text) 25%,transparent)"></span><span class="lbl">Total available</span><span class="val">${s.total}</span></div>
      </div>
    </div>
    <div class="apanel">
      <div class="lv-rows">
        <div><span>Annual leaves</span><b>${ANNUAL_LEAVES}</b></div>
        <div><span>Extra — worked on off day</span><b class="pos">+${s.extra.count}</b></div>
        <div><span>Extra — from overtime</span><b class="pos">+${s.otRewards}</b></div>
        ${s.otCut ? `<div><span>Cut — short time (8h each)</span><b class="neg">−${s.otCut}</b></div>` : ""}
        <div><span>Leaves used</span><b>${s.usedCount ? "−" + s.usedCount : 0}</b></div>
        <div class="tot"><span>${s.remaining < 0 ? "Over limit" : "Leaves available"}</span><b class="${s.remaining < 0 ? "neg" : ""}">${s.remaining}</b></div>
      </div>
    </div>
    ${nextLeavePanel(m)}
    ${segmented("lvSeg", `Leaves taken <span class="pill">${s.used.length}</span>`, `Extra earned <span class="pill">${s.extraAll}</span>`)}
    <div id="paneUsed">${leaveList(s.used, "used")}</div>
    <div id="paneExtra" hidden>${extraItems.length || !s.extraAll
        ? leaveList(extraItems, "extra")
        : `<div class="list"><p class="empty-note">${s.extraAll} extra leave(s) earned.</p></div>`}</div>`;
}

/* Admin: attendance issues (Slack messages jo samajh nahi aaye, check out bhool gaye, etc.) */
function issuesBlock(issues){
  if(!issues || !issues.length) return "";
  return `
    <div class="sub-h">Attendance issues <span class="pill">${issues.length}</span></div>
    <div class="list issues">${issues.map((x, i) => `
      <div class="list-item" style="animation-delay:${Math.min(i, 12) * 30}ms">
        <div class="li-body">
          <div class="li-title">${esc(x.name)} <span class="li-sub">· ${esc(x.date)}</span></div>
          <div class="li-sub">${esc(x.problem)}</div>
          ${x.text ? `<div class="li-sub quote">“${esc(x.text)}”</div>` : ""}
        </div>
        ${x.link ? `<a class="badge ${x.severity === "ERROR" ? "red" : "neutral"}" href="${esc(x.link)}" target="_blank" rel="noopener">Open</a>` : `<span class="badge ${x.severity === "ERROR" ? "red" : "neutral"}">${esc(x.severity)}</span>`}
      </div>`).join("")}</div>
    <p class="att-note">Fix in Google Sheet → "Raw Log" tab (Manual Type / Manual Time), or tick “Resolved” in the "Issues" tab.</p>`;
}

/* Admin: ek employee ka is mahine ka khulasa */
function adminAttBlock(emp){
  const m = attModel(emp);
  if(!m.recs.size) return `<div class="sub-h">Working hours</div><div class="list"><p class="empty-note">No attendance synced yet.</p></div>`;
  const mo = m.months[m.months.length - 1];
  const last = m.days.slice(-7).reverse();
  return `
    <div class="sub-h">Working hours — ${monthLabel(mo.start)} (so far)</div>
    <div class="apanel">${sumPanel(mo.t)}${statement(mo, monthShort(new Date(mo.start.getFullYear(), mo.start.getMonth() + 1, 1)))}</div>
    <div class="sub-h">Last 7 days</div>
    <div class="alist">${last.map(x => dayRow(x)).join("")}</div>`;
}