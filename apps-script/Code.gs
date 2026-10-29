/**
 * «Месопотамія» Жадана — таблиця результатів для вчителя.
 * Сайт надсилає результати учня сюди; кожен учень — один рядок, який оновлюється.
 * Встановлення: див. README.md (Розширення → Apps Script → вставити цей код → Розгорнути як веб-застосунок).
 */
const STORIES = ['Марат', 'Ромео', 'Іван', 'Маріо', 'Юра', 'Фома', 'Матвій', 'Боб', 'Лука'];
const RES_HEAD = ['ID', 'Оновлено', 'Учень']
  .concat(STORIES.map(s => s + ', %'))
  .concat(['Гра: перша спроба, %', 'Гра: найкраща, %', 'Спроб гри', 'Есе зараховано', 'Відповідей на запитання']);
const ESSAY_HEAD = ['Ключ', 'Оновлено', 'Учень', 'Новела / група', '№', 'Запитання', 'Відповідь', 'Знаків'];

function doGet() {
  return ContentService.createTextOutput('Таблиця «Месопотамія» готова приймати результати.');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d || !d.id || !d.name) return out('no name');
    const now = new Date();
    const name = safe(d.name).slice(0, 80);
    const st = d.stories || [];
    const row = [String(d.id).slice(0, 40), now, name]
      .concat(STORIES.map((s, i) => (st[i] && typeof st[i].pct === 'number') ? st[i].pct : ''))
      .concat([num(d.game && d.game.first), num(d.game && d.game.best), num(d.game && d.game.tries), num(d.openDone), num(d.answered)]);
    upsert(sheet('Результати', RES_HEAD, true), row);
    const es = sheet('Есе', ESSAY_HEAD, false);
    (d.open || []).slice(0, 120).forEach(o => {
      const a = safe(o.a).slice(0, 20000);
      upsert(es, [d.id + '#' + o.i, now, name, safe(o.g), Number(o.i) + 1, safe(o.q), a, a.length]);
    });
    return out('ok');
  } finally {
    lock.releaseLock();
  }
}

function out(t) { return ContentService.createTextOutput(t); }
function num(v) { return typeof v === 'number' && isFinite(v) ? v : ''; }
// захист від формул у тексті учня
function safe(v) { const s = String(v == null ? '' : v); return /^[=+\-@]/.test(s) ? "'" + s : s; }

function sheet(name, head, isRes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#99cc00').setWrap(true);
    sh.setFrozenRows(1);
    sh.hideColumns(1);
    sh.getRange('B:B').setNumberFormat('dd.MM.yyyy HH:mm');
    if (isRes) {
      sh.setFrozenColumns(3);
      sh.setColumnWidth(3, 200);
      const pct = sh.getRange(2, 4, 1000, STORIES.length + 2);
      sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue('#f4c7c3', SpreadsheetApp.InterpolationType.NUMBER, '0')
        .setGradientMidpointWithValue('#fce8b2', SpreadsheetApp.InterpolationType.NUMBER, '60')
        .setGradientMaxpointWithValue('#b7e1a1', SpreadsheetApp.InterpolationType.NUMBER, '100')
        .setRanges([pct]).build()]);
    } else {
      sh.setColumnWidth(6, 320); sh.setColumnWidth(7, 480);
      sh.getRange('F:G').setWrap(true);
    }
    const first = ss.getSheetByName('Аркуш1') || ss.getSheetByName('Sheet1');
    if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  }
  return sh;
}

function upsert(sh, row) {
  const n = sh.getLastRow();
  if (n > 1) {
    const keys = sh.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === String(row[0])) { sh.getRange(i + 2, 1, 1, row.length).setValues([row]); return; }
    }
  }
  sh.appendRow(row);
}

/** Запусти один раз вручну (▶ Виконати), щоб побачити аркуші й дати дозволи. */
function setup() {
  sheet('Результати', RES_HEAD, true);
  sheet('Есе', ESSAY_HEAD, false);
}
