// Google Sheet > Extensions > Apps Script me paste karo. Deploy > Web app (Execute as: Me, Access: Anyone)
const SHEET = 'Vehicles';
const HEAD = ['id','type','vehicle_no','bags','shipments','datetime','shift','source','destination','connection','status','updated_at'];

function sh_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET);
    s.appendRow(HEAD);
    s.setFrozenRows(1);
    s.getRange(1, 1, s.getMaxRows(), HEAD.length).setNumberFormat('@'); // plain text, no auto date conversion
  }
  return s;
}
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet() {
  const v = sh_().getDataRange().getValues();
  const rows = v.slice(1).filter(r => r[0]).map(r => Object.fromEntries(HEAD.map((h, i) => [h, String(r[i])])));
  return out_({ ok: true, rows });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    const s = sh_();
    const ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => String(r[0]));
    const idx = ids.indexOf(String(b.id)); // 0-based; row = idx+1
    if (b.action === 'delete') {
      if (idx > 0) s.deleteRow(idx + 1);
    } else {
      b.row.updated_at = new Date().toISOString();
      const vals = HEAD.map(h => b.row[h] ?? '');
      if (idx > 0) s.getRange(idx + 1, 1, 1, HEAD.length).setValues([vals]);
      else { s.appendRow(vals); }
    }
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally { lock.releaseLock(); }
}
