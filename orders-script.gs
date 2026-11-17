// ===== Google Apps Script: استقبال الطلبات (نسخة محمية) =====
// غيّر الباسورد ده لكلمة سر طويلة وعشوائية (16 حرف أو أكتر)، وهتدخل بيها لوحة الإدارة
const KEY = "knvtalabat";
const SHEET = "Orders";

function sh_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const id = props.getProperty("SSID");
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) { ss = SpreadsheetApp.create("KVN Covers - Orders"); props.setProperty("SSID", ss.getId()); }
  let s = ss.getSheetByName(SHEET);
  if (!s) { s = ss.getSheets()[0]; s.setName(SHEET); s.appendRow(["code","date","name","phone","payment","items","total","status","note"]); }
  return s;
}
// شغّل الدالة دي مرة واحدة من المحرر عشان تعمل الجدول وتديله الصلاحيات
function setup() { Logger.log(sh_().getParent().getUrl()); }

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// ---- حماية كلمة السر: بعد 10 محاولات غلط الدخول بيتقفل 15 دقيقة ----
function auth_(k) {
  const c = CacheService.getScriptCache();
  const fails = Number(c.get("fails") || 0);
  if (fails >= 10) return "locked";
  if (String(k) !== KEY) { c.put("fails", String(fails + 1), 900); return "auth"; }
  c.remove("fails");
  return "";
}
// ---- حد للطلبات: 5 طلبات لكل رقم في الساعة، و60 طلب إجمالي كل 10 دقايق ----
function allow_(phone) {
  const c = CacheService.getScriptCache();
  const pn = Number(c.get("p" + phone) || 0), gn = Number(c.get("g") || 0);
  if (pn >= 5 || gn >= 60) return false;
  c.put("p" + phone, String(pn + 1), 3600);
  c.put("g", String(gn + 1), 600);
  return true;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const s = sh_();
    if (d.action === "create") {
      const phone = String(d.phone || "");
      const name = String(d.name || "").trim();
      const items = Array.isArray(d.items) ? d.items : [];
      if (!/^KVN-[A-Z0-9]{6}$/.test(String(d.code))) return out_({ ok: false });
      if (["cash", "instapay"].indexOf(d.payment) < 0) return out_({ ok: false });
      if (!/^01[0125]\d{8}$/.test(phone)) return out_({ ok: false });
      if (name.length < 2 || name.length > 80) return out_({ ok: false });
      if (items.length < 1 || items.length > 30) return out_({ ok: false });
      const total = Number(d.total);
      if (!(total >= 0 && total < 1000000)) return out_({ ok: false });
      if (!allow_(phone)) return out_({ ok: false, error: "limit" });
      const clean = items.map(i => ({ name: String(i.name).slice(0, 120), model: String(i.model || "").slice(0, 60), qty: Math.max(1, Math.min(99, Number(i.qty) || 1)), price: Number(i.price) || 0 }));
      const r = s.getLastRow() + 1;
      const rng = s.getRange(r, 1, 1, 9);
      rng.setNumberFormat("@");
      rng.setValues([[String(d.code), new Date().toISOString(), name, phone, d.payment, JSON.stringify(clean), String(total), "new", ""]]);
      return out_({ ok: true });
    }
    const a = auth_(d.key);
    if (a) return out_({ ok: false, error: a });
    const rows = s.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(d.code)) {
        if (d.action === "status") { s.getRange(i + 1, 8).setValue(String(d.status)); return out_({ ok: true }); }
        if (d.action === "delete") { s.deleteRow(i + 1); return out_({ ok: true }); }
      }
    }
    return out_({ ok: false });
  } catch (err) {
    return out_({ ok: false });
  } finally { lock.releaseLock(); }
}

function doGet(e) {
  const a = auth_(e.parameter.key);
  if (a) return out_({ ok: false, error: a });
  const rows = sh_().getDataRange().getValues().slice(1).map(r => ({
    code: r[0], date: r[1], name: r[2], phone: r[3], payment: r[4],
    items: JSON.parse(r[5] || "[]"), total: Number(r[6]) || 0, status: r[7], note: r[8]
  }));
  return out_({ ok: true, orders: rows });
}
