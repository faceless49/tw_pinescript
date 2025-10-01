// index.js
const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

// ===== Хелперы =====
function parseLevel(val) {
  if (val === undefined || val === null) return 0.0;
  if (typeof val === 'string') {
    const trimmed = val.trim();
    if (trimmed === '' || trimmed === '—' || trimmed === '--') return 0.0;
    const first = trimmed.includes('/') ? trimmed.split('/')[0].trim() : trimmed;
    const num = parseFloat(first.replace(',', '.'));
    return Number.isFinite(num) ? num : 0.0;
  }
  if (typeof val === 'number') return Number.isFinite(val) ? val : 0.0;
  return 0.0;
}
function mapExcelTickerToPine(excelTicker) {
  if (!excelTicker) return excelTicker;
  const map = { USDRUBF: 'USDRUB.P', CNYRUBF: 'CNYRUB.P', GLDRUBF: 'GLDRUB.P' };
  return map[excelTicker] ?? excelTicker;
}

// ===== Простая форма (GET) =====
const GET_PATHS = ['/', '/generate-pine-script', '/api/generate-pine-script'];
app.get(GET_PATHS, (_req, res) => {
  res.set('Content-Type', 'text/html; charset=utf-8');
  res.send(`
    <h1>Excel → Pine</h1>
    <form action="/generate-pine-script" method="post" enctype="multipart/form-data">
      <input type="file" name="file" accept=".xlsx,.xls" required />
      <button type="submit">Сгенерировать Pine</button>
    </form>
    <p>Альтернативный POST-роут: <code>/api/generate-pine-script</code></p>
  `);
});

// ===== Основной обработчик (POST) =====
const POST_PATHS = ['/', '/generate-pine-script', '/api/generate-pine-script'];
app.post(POST_PATHS, upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).send('Файл не загружен (ожидаю поле "file").');

  let workbook;
  try {
    workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
  } catch (e) {
    return res.status(400).send('Не удалось прочитать Excel: ' + e.message);
  }

  // ===== Шапка Pine: Жёсткая привязка к ЦЕНОВОЙ шкале =====
  let pineScript = `//@version=6
// Форматируем как цену и используем правую ценовую шкалу графика
indicator("Multi-Ticker Levels (Goals/Stop/Cancel/Entry)", overlay=true, format=format.price, scale=scale.right)

// Тикер текущего графика без префикса биржи
var string raw_ticker = syminfo.ticker
var string ticker     = str.replace(syminfo.ticker, "MOEX:", "")

// Точность (опционально): число знаков после запятой из syminfo.mintick
float _mt = syminfo.mintick
int _prec = _mt > 0 ? int(math.round(math.log10(1 / _mt))) : 2
// применим precision к форматированию лейблов, если нужно

// Уровни (series float)
var float goal1        = 0.0
var float goal2        = 0.0
var float stop_level   = 0.0
var float cancel_level = 0.0
var float entry_level  = 0.0
`;

  // ===== Генерация условий из всех листов (кроме "Легенда") =====
  workbook.SheetNames
    .filter((sheetName) => sheetName !== 'Легенда')
    .forEach((sheetName) => {
      const ws = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws);

      rows.forEach((row) => {
        const tRaw = row['Тикер'];
        if (!tRaw) return;
        const t = mapExcelTickerToPine(String(tRaw).trim());

        const g1 = parseLevel(row['Цель 1']);
        const g2 = parseLevel(row['Цель 2']);
        const s  = parseLevel(row['Стоп']);
        const c  = parseLevel(row['Отмена']);
        const e  = parseLevel(row['Вход']);

        pineScript += `
if ticker == "${t}"
    goal1        := ${g1}
    goal2        := ${g2}
    stop_level   := ${s}
    cancel_level := ${c}
    entry_level  := ${e}
`;
      });
    });

  // ===== Хвост Pine: ТОЛЬКО plot-линии (ценовые серии) + подписи у последнего бара =====
  pineScript += `
// ---- Серии уровней (привязаны к цене) ----
goal1_series   = goal1        > 0 ? goal1        : na
goal2_series   = goal2        > 0 ? goal2        : na
stop_series    = stop_level   > 0 ? stop_level   : na
cancel_series  = cancel_level > 0 ? cancel_level : na
entry_series   = entry_level  > 0 ? entry_level  : na

// ---- ЛИНИИ уровней (только plot, на ценовой шкале) ----
plot(goal1_series,   title="Цель 1", color=color.red,    linewidth=2, style=plot.style_line, trackprice=true)
plot(goal2_series,   title="Цель 2", color=color.red,    linewidth=2, style=plot.style_line, trackprice=true)
plot(stop_series,    title="Стоп",   color=color.orange, linewidth=2, style=plot.style_line, trackprice=true)
plot(cancel_series,  title="Отмена", color=color.gray,   linewidth=2, style=plot.style_line, trackprice=true)
plot(entry_series,   title="Вход",   color=color.green,  linewidth=2, style=plot.style_line, trackprice=true)

// ---- Подписи у последнего бара (информативно, не влияет на шкалу) ----
showText = input.bool(true, "Показывать подписи (Вход/Стоп/Отмена/Цели) у последнего бара")

var label lbl_goal1   = na
var label lbl_goal2   = na
var label lbl_stop    = na
var label lbl_cancel  = na
var label lbl_entry   = na

if barstate.islast
    if not na(lbl_goal1)
        label.delete(lbl_goal1), lbl_goal1 := na
    if not na(lbl_goal2)
        label.delete(lbl_goal2), lbl_goal2 := na
    if not na(lbl_stop)
        label.delete(lbl_stop),  lbl_stop  := na
    if not na(lbl_cancel)
        label.delete(lbl_cancel),lbl_cancel:= na
    if not na(lbl_entry)
        label.delete(lbl_entry), lbl_entry := na

    if showText
        // форматируем ценник с учётом _prec (опционально)
        f_fmt(p) => str.tostring(p, _prec)
        if not na(goal1_series)
            lbl_goal1  := label.new(bar_index, goal1,        "Цель 1 " + f_fmt(goal1), style=label.style_label_left,  textcolor=color.white, color=color.new(color.red,    20), size=size.tiny)
        if not na(goal2_series)
            lbl_goal2  := label.new(bar_index, goal2,        "Цель 2 " + f_fmt(goal2), style=label.style_label_left,  textcolor=color.white, color=color.new(color.red,    40), size=size.tiny)
        if not na(stop_series)
            lbl_stop   := label.new(bar_index, stop_level,   "Стоп "   + f_fmt(stop_level), style=label.style_label_left,  textcolor=color.white, color=color.new(color.orange, 20), size=size.tiny)
        if not na(cancel_series)
            lbl_cancel := label.new(bar_index, cancel_level, "Отмена " + f_fmt(cancel_level), style=label.style_label_left, textcolor=color.white, color=color.new(color.gray,   20), size=size.tiny)
        if not na(entry_series)
            lbl_entry  := label.new(bar_index, entry_level,  "Вход "   + f_fmt(entry_level), style=label.style_label_left,  textcolor=color.white, color=color.new(color.green,  20), size=size.tiny)
`;

  res.set('Content-Type', 'text/plain; charset=utf-8');
  res.set('Content-Disposition', 'attachment; filename="generated_pine_script.txt"');
  return res.status(200).send(pineScript);
});

// ===== Экспорт для Vercel =====
module.exports = (req, res) => app(req, res);
