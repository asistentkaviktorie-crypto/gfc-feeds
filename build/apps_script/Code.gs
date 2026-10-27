/**
 * GFC – віддає роботу GitHub ТІЛЬКИ публічні дані з таблиці «GFC – Об'єкти та реклама».
 * Встановлення: Розширення → Apps Script → вставити цей код → Зберегти →
 * Розгорнути → Нове розгортання → тип «Вебзастосунок» → Виконувати як: Я; Доступ: Усі → Розгорнути.
 * URL розгортання (…/exec) зберегти в GitHub як секрет SHEET_JSON_URL.
 *
 * Приватні колонки вкладки «Обєкти» (Джерело, Договір з забудовником, ПДВ, Статус будівництва, Прайс,
 * Примітки, Остання перевірка тощо) і вкладки «Статистика», «Новий проєкт – чеклист» НЕ віддаються.
 */
var OBJ_COLUMNS = ['Проєкт', 'Ціна €', 'Площа заг., м²', 'Житлова, м²', 'Ділянка, м²', 'Спальні', 'Ванні',
                   'Realting ID', 'Realting статус', 'У фід Realting', 'У фід JamesEdition', 'У фід Prian'];
var PUBLIC_TABS = ['Тексти', 'Переклади', 'Параметри', 'Сайт', 'Фото'];
var YES = ['так', 'yes', 'да', '1', 'true'];

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = {version: 1, generated: new Date().toISOString(), tabs: {}};
  var obj = ss.getSheets()[0];                      // перша вкладка = «Обєкти»
  var v = obj.getDataRange().getValues();
  var hdr = v[0].map(function (h) { return String(h).trim(); });
  var idx = OBJ_COLUMNS.map(function (c) { return hdr.indexOf(c); });
  var params = ss.getSheetByName('Параметри'), site = ss.getSheetByName('Сайт');
  var known = {};
  [params, site].forEach(function (sh) {
    if (!sh) return;
    sh.getDataRange().getValues().slice(1).forEach(function (r) { if (r[0] !== '') known[String(r[0]).trim().toUpperCase()] = 1; });
  });
  var rid = hdr.indexOf('Realting ID');
  var flags = ['У фід Realting', 'У фід JamesEdition', 'У фід Prian'].map(function (c) { return hdr.indexOf(c); });
  var rows = [OBJ_COLUMNS];
  v.slice(1).forEach(function (r) {
    var id = rid >= 0 ? String(r[rid]).trim() : '';
    if (!id) return;
    var flagged = flags.some(function (i) { return i >= 0 && YES.indexOf(String(r[i]).trim().toLowerCase()) >= 0; });
    if (!flagged && !known[id.toUpperCase()]) return;   // чернетки без «так» і без «Параметри» не віддаємо
    rows.push(idx.map(function (i) { return i >= 0 ? r[i] : ''; }));
  });
  out.tabs['Обєкти'] = rows;
  var sent = {};
  rows.slice(1).forEach(function (r) { sent[String(r[7]).trim().toUpperCase()] = 1; });
  PUBLIC_TABS.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var t = sh.getDataRange().getValues();
    if (name === 'Тексти')        // тексти чернеток (без «так») теж не віддаємо
      t = [t[0]].concat(t.slice(1).filter(function (r) { return sent[String(r[0]).trim().toUpperCase()]; }));
    out.tabs[name] = t;
  });
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
