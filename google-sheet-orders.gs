/**
 * Dollar Store Oman – orders, affiliate program, driver page, returns, payouts.
 * Tabs (created automatically): Orders, Affiliates, Returns, Payouts.
 * After pasting: Save, run setupTriggers once (allow access), then Deploy > Manage deployments > edit > New version.
 */
const ALERT_EMAIL = "m.tayyab.r@outlook.com";   // order alerts + monthly payout list
const STAFF_PIN   = "3302";                     // PIN for the driver / returns page (staff.html)
const RATE        = 0.10;                       // affiliate commission: 10% of items total (no delivery)
const AUTO_DAYS   = 3;                          // orders not marked within 3 days become Delivered

const OH = ["Date","Order #","WhatsApp","Name","Area / notes","Items","Qty","Items OMR","Delivery OMR","Total OMR","Language","Status","Affiliate","Returns OMR","Delivered on","Photo","Order photos","Item IDs"];
const SITE = "https://mtayyabr.github.io/catalogue/";
const thumb_ = id => SITE + "img/" + id + "-t.jpg";
const AH = ["Code","Name","WhatsApp","Payout method","Bank details","Status","Joined","Key"];
const RH = ["Date","Order #","Affiliate","Returned OMR","Commission deducted OMR"];
const PH = ["Month","Code","Name","WhatsApp","Payout method","Bank details","Orders delivered","Items OMR","Commission OMR","Returns deducted OMR","To pay OMR","Paid?"];
const C = {date:0,id:1,phone:2,name:3,area:4,items:5,qty:6,sub:7,del:8,total:9,lang:10,status:11,aff:12,ret:13,delOn:14,photo:15,link:16,ids:17};

/* ---------- helpers ---------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tab_(name, head) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) { sh.appendRow(head); sh.setFrozenRows(1); }
  else {
    const cur = sh.getRange(1, 1, 1, head.length).getValues()[0];
    head.forEach((h, i) => { if (cur[i] !== h) sh.getRange(1, i + 1).setValue(h); });
  }
  return sh;
}
function rows_(sh) { const n = sh.getLastRow(); return n < 2 ? [] : sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues(); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function r3_(x) { return Math.round(x * 1000) / 1000; }
function ym_(d) { d = new Date(d); return d.getFullYear() * 100 + d.getMonth() + 1; }
function ymLabel_(ym) { return Utilities.formatDate(new Date(Math.floor(ym / 100), ym % 100 - 1, 15), ss_().getSpreadsheetTimeZone(), "MMMM yyyy"); }
function findOrder_(sh, id) {
  id = String(id || "").trim().toUpperCase(); if (!id || sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]).toUpperCase() === id) return i + 2;
  return 0;
}
function affiliates_() { return rows_(tab_("Affiliates", AH)); }
function withLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }

/* ---------- web endpoints ---------- */
function doPost(e) {
  const o = JSON.parse(e.postData.contents);
  const a = o.action || "order";
  if (a === "order")  return json_(withLock_(() => saveOrder_(o)));
  if (a === "join")   return json_(withLock_(() => join_(o)));
  if (a === "status") return json_(String(o.pin) === STAFF_PIN ? withLock_(() => setStatus_(o.id, o.status)) : {ok:false, err:"pin"});
  if (a === "return") return json_(String(o.pin) === STAFF_PIN ? withLock_(() => addReturn_(o.id, o.amount)) : {ok:false, err:"pin"});
  return json_({ok:false, err:"action"});
}
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === "aff")   return json_(affInfo_(p.code));
  if (p.action === "stats") return json_(stats_(p.code, p.key, p.phone));
  if (p.action === "open")  return json_(String(p.pin) === STAFF_PIN ? openOrders_() : {ok:false, err:"pin"});
  return ContentService.createTextOutput("DS Orders is connected");
}

/* ---------- orders ---------- */
function saveOrder_(o) {
  const sh = tab_("Orders", OH);
  const items = o.items.map(i => `${i.qty} × ${i.name}`).join("\n");
  const qty = o.items.reduce((a, i) => a + i.qty, 0);
  const aff = o.ref && affInfo_(o.ref).ok ? String(o.ref).toUpperCase() : "";
  const ids = o.items.map(i => i.id + (i.size ? "|" + i.size : "") + ":" + i.qty).join(",");
  const first = o.items[0] ? o.items[0].id : "";
  const link = o.link || "";
  sh.appendRow([new Date(o.date), o.id, "+" + o.phone, o.name, o.area, items, qty,
                Number(o.subtotal), Number(o.delivery), Number(o.total), o.lang === "ar" ? "Arabic" : o.lang === "fil" ? "Filipino" : "English", "New", aff, "", "",
                first ? `=IMAGE("${thumb_(first)}")` : "", link ? `=HYPERLINK("${link}","📷 View ${o.items.length} item${o.items.length > 1 ? "s" : ""}")` : "", ids]);
  sh.setRowHeight(sh.getLastRow(), 60);
  const esc = x => String(x == null ? "" : x).replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
  const rowsHtml = o.items.map(i => `<tr><td style="padding:6px 8px 6px 0"><img src="${thumb_(i.id)}" width="64" height="64" style="border-radius:8px;border:1px solid #eee;display:block" alt=""></td>
    <td style="padding:6px 8px;font:14px Arial">${i.qty} × ${esc(i.name)}</td><td style="padding:6px 0;font:bold 14px Arial;text-align:right;white-space:nowrap">${esc(i.amount)} OMR</td></tr>`).join("");
  const html = `<div style="font:14px Arial;color:#222;max-width:520px">
    <h2 style="margin:0 0 4px">New order ${esc(o.id)}${aff ? ` <span style="font-size:13px;color:#0A8A3A">(affiliate ${aff})</span>` : ""}</h2>
    <div style="color:#555;margin-bottom:10px">${esc(o.name || "No name")} · +${esc(o.phone)}${o.area ? " · " + esc(o.area) : ""}</div>
    <table style="border-collapse:collapse;width:100%">${rowsHtml}</table>
    <p style="font:14px Arial;margin:12px 0">Items ${esc(o.subtotal)} OMR · Delivery ${Number(o.delivery) ? esc(o.delivery) + " OMR" : "FREE"} · <b>Total ${esc(o.total)} OMR</b> (cash on delivery)</p>
    <p><a href="https://api.whatsapp.com/send?phone=${esc(o.phone)}" style="background:#1FA855;color:#fff;padding:10px 16px;border-radius:20px;text-decoration:none;font-weight:bold">Reply on WhatsApp</a>
    ${link ? ` &nbsp; <a href="${esc(link)}" style="color:#F2460D;font-weight:bold">📷 Order with photos</a>` : ""}</p></div>`;
  MailApp.sendEmail({to: ALERT_EMAIL, subject: `New order ${o.id} – ${o.total} OMR` + (aff ? ` (affiliate ${aff})` : ""),
    body: o.text + `\n\nReply on WhatsApp: https://api.whatsapp.com/send?phone=${o.phone}`, htmlBody: html});
  return {ok:true};
}
function setStatus_(id, status) {
  if (["Delivered","Refused","Cancelled","New"].indexOf(status) < 0) return {ok:false, err:"status"};
  const sh = tab_("Orders", OH), r = findOrder_(sh, id);
  if (!r) return {ok:false, err:"notfound"};
  sh.getRange(r, C.status + 1).setValue(status);
  sh.getRange(r, C.delOn + 1).setValue(status === "Delivered" ? new Date() : "");
  return {ok:true};
}
function addReturn_(id, amount) {
  amount = r3_(Number(amount));
  const sh = tab_("Orders", OH), r = findOrder_(sh, id);
  if (!r) return {ok:false, err:"notfound"};
  const row = sh.getRange(r, 1, 1, OH.length).getValues()[0];
  const done = Number(row[C.ret]) || 0, max = r3_(Number(row[C.sub]) - done);
  if (!(amount > 0) || amount > max) return {ok:false, err:"amount", max:max};
  sh.getRange(r, C.ret + 1).setValue(r3_(done + amount));
  const aff = String(row[C.aff] || "");
  const ded = aff && row[C.status] === "Delivered" ? r3_(amount * RATE) : 0;
  tab_("Returns", RH).appendRow([new Date(), row[C.id], aff, amount, ded]);
  return {ok:true, deducted:ded};
}
function openOrders_() {
  const sh = tab_("Orders", OH), since = Date.now() - 30 * 864e5;
  const list = rows_(sh).filter(r => r[C.id] && r[C.status] === "New" && new Date(r[C.date]).getTime() > since)
    .map(r => ({id:r[C.id], date:new Date(r[C.date]).toISOString(), name:r[C.name], area:r[C.area], phone:String(r[C.phone]),
                items:r[C.items], total:Number(r[C.total]), ids:String(r[C.ids] || "")}));
  return {ok:true, orders:list.reverse()};
}

/* ---------- affiliates ---------- */
function affInfo_(code) {
  code = String(code || "").toUpperCase(); if (!code) return {ok:false};
  const a = affiliates_().find(r => String(r[0]).toUpperCase() === code);
  return a && a[5] === "Active" ? {ok:true, name:a[1]} : {ok:false};
}
function join_(o) {
  const name = String(o.name || "").trim().slice(0, 40), phone = String(o.phone || "").replace(/\D/g, "");
  const payout = o.payout ? String(o.payout).slice(0, 30) : "Ask when paying", bank = String(o.bank || "").trim().slice(0, 120);
  if (!name || !/^[79]\d{7}$/.test(phone) || o.agree !== true) return {ok:false, err:"form"};
  const sh = tab_("Affiliates", AH), list = rows_(sh);
  const old = list.find(r => String(r[2]).replace(/\D/g, "").slice(-8) === phone);
  if (old) return old[5] === "Blocked" ? {ok:false, err:"blocked"} : {ok:true, code:String(old[0]).toLowerCase(), key:String(old[7]), name:old[1], existing:true};
  const base = (name.toLowerCase().normalize("NFKD").replace(/[^a-z]/g, "").slice(0, 8)) || "ds";
  const used = list.map(r => String(r[0]).toUpperCase());
  let code; do { code = base + (10 + Math.floor(Math.random() * 90)); } while (used.indexOf(code.toUpperCase()) >= 0);
  const key = Utilities.getUuid().replace(/-/g, "").slice(0, 12);
  sh.appendRow([code.toUpperCase(), name, "+968" + phone, payout, bank, "Active", new Date(), key]);
  MailApp.sendEmail(ALERT_EMAIL, `New affiliate: ${name} (${code.toUpperCase()})`,
    `${name} joined the affiliate program.\nWhatsApp: +968 ${phone}\nCode: ${code.toUpperCase()}\n\nTo block: Affiliates tab > Status = Blocked`);
  return {ok:true, code:code, key:key, name:name};
}
/* commission for one affiliate in one month (ym = 202610) */
function monthFor_(code, ym, orders, returns) {
  let n = 0, items = 0;
  orders.forEach(r => { if (String(r[C.aff]).toUpperCase() === code && r[C.status] === "Delivered" && r[C.delOn] && ym_(r[C.delOn]) === ym) { n++; items += Number(r[C.sub]) || 0; } });
  let ded = 0;
  returns.forEach(r => { if (String(r[2]).toUpperCase() === code && ym_(r[0]) === ym) ded += Number(r[4]) || 0; });
  const com = r3_(items * RATE);
  return {ym:ym, label:ymLabel_(ym), orders:n, items:r3_(items), commission:com, returns:r3_(ded), pay:r3_(com - ded)};
}
function stats_(code, key, phone) {
  code = String(code || "").toUpperCase();
  const ph = String(phone || "").replace(/\D/g, "").slice(-8);
  const a = affiliates_().find(r => String(r[0]).toUpperCase() === code &&
    ((key && String(r[7]) === String(key)) || (ph.length === 8 && String(r[2]).replace(/\D/g, "").slice(-8) === ph)));
  if (!a) return {ok:false};
  const orders = rows_(tab_("Orders", OH)).filter(r => String(r[C.aff]).toUpperCase() === code);
  const returns = rows_(tab_("Returns", RH)).filter(r => String(r[2]).toUpperCase() === code);
  const payouts = rows_(tab_("Payouts", PH)).filter(r => String(r[1]).toUpperCase() === code && String(r[11]).trim() !== "");
  let approved = 0, pending = 0;
  orders.forEach(r => { const v = Number(r[C.sub]) || 0; if (r[C.status] === "Delivered") approved += v * RATE; else if (r[C.status] === "New") pending += v * RATE; });
  const ded = returns.reduce((t, r) => t + (Number(r[4]) || 0), 0);
  const paid = payouts.reduce((t, r) => t + (Number(r[10]) || 0), 0);
  const recent = orders.slice(-30).reverse().map(r => ({id:r[C.id], date:new Date(r[C.date]).toISOString(), items:Number(r[C.sub]), status:r[C.status]}));
  return {ok:true, name:a[1], code:code, status:a[5], due:r3_(approved - ded - paid), pending:r3_(pending), paid:r3_(paid), approved:r3_(approved), returns:r3_(ded), recent:recent};
}

/* ---------- automation (installed by setupTriggers) ---------- */
function autoDeliver() {
  withLock_(() => {
    const sh = tab_("Orders", OH), rows = rows_(sh), cut = Date.now() - AUTO_DAYS * 864e5;
    rows.forEach((r, i) => {
      if (r[C.status] === "New" && r[C.date] && new Date(r[C.date]).getTime() < cut) {
        sh.getRange(i + 2, C.status + 1).setValue("Delivered");
        sh.getRange(i + 2, C.delOn + 1).setValue(new Date());
      }
    });
  });
}
function monthlyPayouts() {
  const now = new Date(), ym = ym_(new Date(now.getFullYear(), now.getMonth() - 1, 15));
  const orders = rows_(tab_("Orders", OH)), returns = rows_(tab_("Returns", RH)), ps = tab_("Payouts", PH);
  const lines = [];
  affiliates_().forEach(a => {
    if (a[5] === "Blocked") return;
    const m = monthFor_(String(a[0]).toUpperCase(), ym, orders, returns);
    if (!m.orders && !m.returns) return;
    ps.appendRow([m.label, a[0], a[1], a[2], a[3], a[4], m.orders, m.items, m.commission, m.returns, m.pay, ""]);
    lines.push(`${a[1]} (${a[0]}) ${a[2]} – ${a[3]}${a[4] ? " " + a[4] : ""}: pay ${m.pay.toFixed(3)} OMR  [${m.orders} orders, ${m.items.toFixed(3)} OMR sales, -${m.returns.toFixed(3)} returns]`);
  });
  MailApp.sendEmail(ALERT_EMAIL, `Affiliate payouts – ${ymLabel_(ym)}`,
    lines.length ? `Pay these affiliates for ${ymLabel_(ym)}:\n\n${lines.join("\n")}\n\nMark them paid in the Payouts tab.` : `No affiliate commission for ${ymLabel_(ym)}.`);
}
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("autoDeliver").timeBased().everyDays(1).atHour(2).create();
  ScriptApp.newTrigger("monthlyPayouts").timeBased().onMonthDay(1).atHour(8).create();
  tab_("Orders", OH); tab_("Affiliates", AH); tab_("Returns", RH); tab_("Payouts", PH);
}
