// ===== Google Apps Script: استقبال الطلبات =====
// غيّر الباسورد ده لحاجة سرية بتاعتك، وهتدخل بيه على صفحة الطلبات في لوحة الإدارة
const KEY = "غيّر-الباسورد-ده";
const SHEET = "Orders";

function sh_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) { s = ss.insertSheet(SHEET); s.appendRow(["code","date","name","phone","payment","items","total","status","note"]); }
  return s;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const s = sh_();
    if (d.action === "create") {
      if (!/^KVN-[A-Z0-9]{6}$/.test(String(d.code))) return out_({ ok: false });
      if (["cash", "instapay"].indexOf(d.payment) < 0) return out_({ ok: false });
      const items = (d.items || []).slice(0, 30).map(i => ({ name: String(i.name).slice(0, 120), model: String(i.model || "").slice(0, 60), qty: Number(i.qty) || 1, price: Number(i.price) || 0 }));
      const r = s.getLastRow() + 1;
      const rng = s.getRange(r, 1, 1, 9);
      rng.setNumberFormat("@");
      rng.setValues([[String(d.code), new Date().toISOString(), String(d.name).slice(0, 80), String(d.phone).slice(0, 20), d.payment, JSON.stringify(items), String(Number(d.total) || 0), "new", ""]]);
      return out_({ ok: true });
    }
    if (d.key !== KEY) return out_({ ok: false, error: "auth" });
    const rows = s.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(d.code)) {
        if (d.action === "status") { s.getRange(i + 1, 8).setValue(String(d.status)); return out_({ ok: true }); }
        if (d.action === "delete") { s.deleteRow(i + 1); return out_({ ok: true }); }
      }
    }
    return out_({ ok: false });
  } finally { lock.releaseLock(); }
}

function doGet(e) {
  if (e.parameter.key !== KEY) return out_({ ok: false, error: "auth" });
  const rows = sh_().getDataRange().getValues().slice(1).map(r => ({
    code: r[0], date: r[1], name: r[2], phone: r[3], payment: r[4],
    items: JSON.parse(r[5] || "[]"), total: Number(r[6]) || 0, status: r[7], note: r[8]
  }));
  return out_({ ok: true, orders: rows });
}
