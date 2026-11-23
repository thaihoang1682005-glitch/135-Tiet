/* ==========================================================================
   OUTSOURCE STUDIO — phần dùng chung cho 2 trang
   ========================================================================== */
(() => {
const C = window.CONFIG;
const $  = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const ICON = (id, st = "") => `<svg class="ico"${st ? ` style="${st}"` : ""}><use href="#${id}"/></svg>`;
const vnd = n => n.toLocaleString("vi-VN") + "đ";

/* ---------- Icon sprite ---------- */
const SPRITE = `<svg width="0" height="0" style="position:absolute" aria-hidden="true">
<symbol id="i-check" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" d="M5 12.5l4.5 4.5L19 7.5"/></symbol>
<symbol id="i-play" viewBox="0 0 24 24"><path fill="currentColor" d="M8 5.5v13a1 1 0 0 0 1.5.86l10.6-6.5a1 1 0 0 0 0-1.72L9.5 4.64A1 1 0 0 0 8 5.5z"/></symbol>
<symbol id="i-x" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" d="M6 6l12 12M18 6L6 18"/></symbol>
<symbol id="i-arrow" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" d="M5 12h14M13 6l6 6-6 6"/></symbol>
<symbol id="i-cam" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><rect x="2.5" y="6" width="13" height="12" rx="2.5"/><path d="M15.5 10.5l6-3.5v10l-6-3.5z"/></g></symbol>
<symbol id="i-cut" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="M8.2 8.2 20 20M8.2 15.8 20 4"/></g></symbol>
<symbol id="i-user" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="8" r="4"/><path d="M4 21c1.5-4 4.5-6 8-6s6.5 2 8 6"/></g></symbol>
<symbol id="i-team" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="9" cy="8" r="3.4"/><path d="M2.5 20c1-3.6 3.6-5.4 6.5-5.4s5.5 1.8 6.5 5.4"/><path d="M16 4.8a3.4 3.4 0 0 1 0 6.4M18 14.8c1.8.6 3 2.3 3.6 5.2"/></g></symbol>
<symbol id="i-clock" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></g></symbol>
<symbol id="i-shield" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round" d="M12 3l8 3v6c0 4.5-3.4 8.2-8 9-4.6-.8-8-4.5-8-9V6z"/><path fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" d="M8.5 12l2.5 2.5 4.5-5"/></symbol>
<symbol id="i-refresh" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" d="M20 11a8 8 0 0 0-14.6-4.5M4 4v3h3M4 13a8 8 0 0 0 14.6 4.5M20 20v-3h-3"/></symbol>
<symbol id="i-wallet" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><rect x="3" y="6" width="18" height="14" rx="3"/><path d="M3 10h18M16 15h2"/><path d="M6 6l9-3 1.5 3"/></g></symbol>
<symbol id="i-drone" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="9" y="10" width="6" height="4.5" rx="1.5"/><path d="M9 11.5 5.5 8.5M15 11.5l3.5-3M9 13l-3.5 3M15 13l3.5 3"/><path d="M2.5 8h6M15.5 8h6M2.5 16h6M15.5 16h6"/></g></symbol>
<symbol id="i-phone" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round" d="M5 4h4l2 5-2.5 1.5a11 11 0 0 0 5 5L15 13l5 2v4a2 2 0 0 1-2 2A16 16 0 0 1 3 6a2 2 0 0 1 2-2z"/></symbol>
<symbol id="i-chat" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round" d="M4 5h16v11H9l-5 4z"/></symbol>
<symbol id="i-mail" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2.5"/><path d="M3.5 6.5 12 13l8.5-6.5"/></g></symbol>
<symbol id="i-music" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M9 18V5l11-2v13"/><circle cx="6.5" cy="18" r="2.5"/><circle cx="17.5" cy="16" r="2.5"/></g></symbol>
<symbol id="i-fb" viewBox="0 0 24 24"><path fill="currentColor" d="M13.5 21v-7.5h2.6l.4-3h-3V8.6c0-.9.3-1.5 1.6-1.5h1.6V4.4c-.3 0-1.2-.1-2.3-.1-2.3 0-3.9 1.4-3.9 4v2.2H8v3h2.5V21z"/></symbol>
<symbol id="i-spark" viewBox="0 0 24 24"><path fill="currentColor" d="M12 2l1.9 6.1L20 10l-6.1 1.9L12 18l-1.9-6.1L4 10l6.1-1.9z"/></symbol>
<symbol id="i-brief" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><rect x="3" y="7" width="18" height="13" rx="2.5"/><path d="M9 7V5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v2M3 12h18"/></g></symbol>
<symbol id="i-lock" viewBox="0 0 24 24"><g fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><rect x="4.5" y="10.5" width="15" height="10" rx="2.5"/><path d="M8 10.5V8a4 4 0 0 1 8 0v2.5"/></g></symbol>
<symbol id="i-folder" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round" d="M3 6.5A1.5 1.5 0 0 1 4.5 5H9l2 2.5h8.5A1.5 1.5 0 0 1 21 9v9.5a1.5 1.5 0 0 1-1.5 1.5h-15A1.5 1.5 0 0 1 3 18.5z"/></symbol>
</svg>`;

/* ---------- Khung trang: thanh đối tượng, header, footer, modal ---------- */
function chrome(page) {
  const isCtv = page.type === "ctv";
  const ct = C.CONTACT;
  document.body.insertAdjacentHTML("afterbegin", SPRITE + `
  <div class="aud"><div class="wrap">
    <nav class="aud-sw" aria-label="Bạn là">
      <a href="index.html" ${!isCtv ? 'aria-current="page"' : ""}>${ICON("i-user")}Thuê làm video</a>
      <a href="cong-tac-vien.html" ${isCtv ? 'aria-current="page"' : ""}>${ICON("i-team")}Nhận việc freelance</a>
    </nav>
    <a class="aud-hot" href="${esc(ct.phone.url)}">${ICON("i-phone")}${esc(ct.phone.value)}</a>
  </div></div>
  <header class="hdr" id="hdr"><div class="wrap">
    <a href="${isCtv ? "cong-tac-vien.html" : "index.html"}" class="logo" aria-label="${esc(C.BRAND)} – trang chủ">
      <img src="assets/logo.svg" alt="" width="38" height="38">
      <span><b>OUTSOURCE</b><small>${isCtv ? "CỘNG TÁC VIÊN" : "STUDIO"}</small></span>
    </a>
    <nav class="nav" id="nav" aria-label="Điều hướng">${page.nav.map(([h, t]) => `<a href="${h}">${t}</a>`).join("")}</nav>
    <div class="hdr-cta"><a href="${page.cta[0]}" class="btn ${isCtv ? "btn-vio" : "btn-pri"} btn-sm">${page.cta[1]}</a></div>
    <button class="burger" id="burger" aria-label="Mở menu" aria-expanded="false" aria-controls="nav"><span></span><span></span><span></span></button>
  </div></header>`);

  document.body.insertAdjacentHTML("beforeend", `
  <footer><div class="wrap">
    <div class="f-grid">
      <div class="f-brand">
        <a href="index.html" class="logo"><img src="assets/logo.svg" alt="" width="38" height="38"><span><b>OUTSOURCE</b><small>STUDIO</small></span></a>
        <p>Quay & dựng video quảng cáo cho shop, hộ kinh doanh và doanh nghiệp nhỏ.</p>
      </div>
      <div><h4>Liên hệ</h4><ul>
        <li>${ICON("i-phone")}<a href="${esc(ct.phone.url)}">${esc(ct.phone.value)}</a></li>
        <li>${ICON("i-chat")}<a href="${esc(ct.zalo.url)}" target="_blank" rel="noopener">Zalo: ${esc(ct.zalo.value)}</a></li>
        <li>${ICON("i-fb")}<a href="${esc(ct.facebook.url)}" target="_blank" rel="noopener">Facebook</a></li>
        <li>${ICON("i-mail")}<a href="${esc(ct.email.url)}" style="word-break:break-all">${esc(ct.email.value)}</a></li>
      </ul></div>
      <div><h4>Khách hàng</h4><ul>
        <li><a href="index.html#video-mau">Video mẫu</a></li>
        <li><a href="index.html#bang-gia">Bảng giá</a></li>
        <li><a href="index.html#dat-dich-vu">Nhận báo giá</a></li>
        ${C.SURVEY_URL ? `<li><a href="${esc(C.SURVEY_URL)}" target="_blank" rel="noopener">Góp ý / khảo sát</a></li>` : ""}
      </ul></div>
      <div><h4>Cộng tác viên</h4><ul>
        <li><a href="cong-tac-vien.html">Quyền lợi</a></li>
        <li><a href="cong-tac-vien.html#quy-trinh">Quy trình nhận job</a></li>
        <li><a href="cong-tac-vien.html#dang-ky">Đăng ký</a></li>
      </ul></div>
    </div>
    <div class="f-bot"><span>© ${new Date().getFullYear()} ${esc(C.BRAND)}</span><span>Phụ trách: Thái · ${esc(ct.phone.value)}</span></div>
  </div></footer>
  <div class="mbar">
    <a class="btn btn-ghost" href="${esc(ct.zalo.url)}" target="_blank" rel="noopener">${ICON("i-chat")} Zalo</a>
    <a class="btn ${isCtv ? "btn-vio" : "btn-pri"}" href="${page.cta[0]}">${page.cta[1]}</a>
  </div>
  <div class="toast" id="toast" role="status" aria-live="polite"></div>`);

  // menu mobile
  const hdr = $("#hdr"), burger = $("#burger");
  burger.addEventListener("click", () => burger.setAttribute("aria-expanded", hdr.classList.toggle("open")));
  $$("#nav a").forEach(a => a.addEventListener("click", () => { hdr.classList.remove("open"); burger.setAttribute("aria-expanded", false); }));
  // đánh dấu mục đang xem
  const spy = new IntersectionObserver(es => es.forEach(e => {
    if (e.isIntersecting) $$("#nav a").forEach(a => a.classList.toggle("on", a.getAttribute("href") === "#" + e.target.id));
  }), { rootMargin: "-45% 0px -50% 0px" });
  $$("#nav a").forEach(a => { const t = a.getAttribute("href").startsWith("#") && $(a.getAttribute("href")); t && spy.observe(t); });
  // hiệu ứng xuất hiện
  const rv = new IntersectionObserver(es => es.forEach(e => { if (e.isIntersecting) { e.target.classList.add("in"); rv.unobserve(e.target); } }), { threshold: .1 });
  $$(".rv").forEach(el => rv.observe(el));
}

/* ---------- Modal ---------- */
let lastFocus = null;
function openModal(m) {
  lastFocus = document.activeElement;
  m.classList.add("on"); document.body.style.overflow = "hidden";
  setTimeout(() => $(".mclose", m)?.focus(), 30);
}
function closeModal(m) {
  m.classList.remove("on"); document.body.style.overflow = "";
  m.dispatchEvent(new Event("closed"));
  lastFocus && lastFocus.focus();
}
document.addEventListener("click", e => {
  const m = e.target.closest(".modal");
  if (m && (e.target === m || e.target.closest("[data-close]"))) closeModal(m);
});
document.addEventListener("keydown", e => {
  const m = $(".modal.on"); if (!m) return;
  if (e.key === "Escape") closeModal(m);
  if (e.key === "Tab") {
    const f = $$("button,a[href],input,select,textarea,iframe", m).filter(x => !x.disabled && x.offsetParent !== null);
    if (!f.length) return;
    if (e.shiftKey && document.activeElement === f[0]) { e.preventDefault(); f.at(-1).focus(); }
    else if (!e.shiftKey && document.activeElement === f.at(-1)) { e.preventDefault(); f[0].focus(); }
  }
});

/* ---------- Toast ---------- */
let tt;
function toast(msg, err) {
  const t = $("#toast"); t.textContent = msg; t.classList.toggle("err", !!err); t.classList.add("on");
  clearTimeout(tt); tt = setTimeout(() => t.classList.remove("on"), 3800);
}

/* ---------- Kiểm tra dữ liệu ---------- */
const PHONE = /^(0|\+?84)(3|5|7|8|9)\d{8}$/;
function check(inp, ok) {
  const f = inp.closest(".fld"); inp.classList.toggle("bad", !ok); f && f.classList.toggle("has-err", !ok); return ok;
}
function validate(scope) {
  let first = null;
  $$("input,textarea", scope).forEach(i => {
    if (i.classList.contains("hp") || i.type === "checkbox" || i.type === "radio" || i.offsetParent === null) return;
    let ok = true; const v = i.value.trim();
    if (i.required && !v) ok = false;
    if (ok && v && i.type === "tel") ok = PHONE.test(v.replace(/[\s.\-]/g, ""));
    if (ok && v && i.type === "email") ok = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v);
    if (ok && v && i.type === "url") ok = /^https?:\/\/\S+$/.test(v);
    if (!check(i, ok) && !first) first = i;
  });
  if (first) first.focus();
  return !first;
}
document.addEventListener("input", e => { const i = e.target; if (i.classList?.contains("bad")) check(i, true); });

/* ---------- Gửi lên Google Sheet ---------- */
function source() {
  const q = new URLSearchParams(location.search);
  try { return q.get("utm_source") || q.get("src") || (document.referrer ? new URL(document.referrer).hostname : "Trực tiếp"); }
  catch { return "Trực tiếp"; }
}
async function send(payload) {
  if (!C.SHEET_ENDPOINT) throw new Error("Chưa cấu hình SHEET_ENDPOINT");
  // text/plain để không phát sinh preflight; đọc phản hồi để chỉ báo thành công khi Sheet đã ghi xong
  const res = await fetch(C.SHEET_ENDPOINT, { method: "POST", redirect: "follow", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(payload) });
  if (!res.ok) throw new Error("HTTP " + res.status);
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Sheet không ghi được dữ liệu");
  return data;
}
/** Trả về kết quả từ Sheet, hoặc null nếu chưa ghi được (form giữ nguyên để gửi lại) */
async function submitWith(btn, fn) {
  const html = btn.innerHTML; btn.disabled = true; btn.innerHTML = "Đang gửi…";
  try { return await fn(); }
  catch (err) { console.error(err); toast("Chưa gửi được. Bạn thử lại hoặc nhắn Zalo 0399 924 348 nhé.", true); return null; }
  finally { btn.disabled = false; btn.innerHTML = html; }
}

window.OS = { $, $$, esc, ICON, vnd, chrome, openModal, closeModal, toast, validate, send, submitWith, source };
})();
