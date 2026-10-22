// BWB・JCXAS 確認シートの受け口（GitHub Pages の送信ボタン → ここ → シートに1行＋森川へメール）
const TO = 'mm738966m@gmail.com';

function setup() {
  const sh = getSheet_();
  MailApp.sendEmail(TO, '【確認シート】受け口のテスト', 'BWB・JCXAS 確認シートの受け口ができました。\n回答シート: ' + sh.getParent().getUrl());
}

function getSheet_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('SHEET_ID');
  let ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('BWB・JCXAS 確認シート 回答');
    p.setProperty('SHEET_ID', ss.getId());
    const sh0 = ss.getSheets()[0];
    sh0.appendRow(['受付日時', 'お名前', '回答（全文）']);
    sh0.setFrozenRows(1);
    sh0.setColumnWidth(3, 700);
  }
  return ss.getSheets()[0];
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.website) return out_({ ok: true });
    const name = String(d.name || '').slice(0, 100);
    const text = String(d.text || '').slice(0, 20000);
    if (!text.trim()) return out_({ ok: false, error: 'empty' });
    const safe = v => /^[=+\-@]/.test(v) ? "'" + v : v;
    const sh = getSheet_();
    sh.appendRow([new Date(), safe(name), safe(text)]);
    const r = sh.getLastRow();
    sh.getRange(r, 1).setNumberFormat('yyyy/m/d H:mm');
    sh.getRange(r, 3).setWrap(true);
    MailApp.sendEmail({
      to: TO,
      subject: '【確認シート】' + (name || '名前なし') + 'さんから回答が届きました',
      body: text + '\n\n回答シート: ' + sh.getParent().getUrl()
    });
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
