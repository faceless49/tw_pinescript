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

  // ===== Шапка Pine =====
  // Важно: scale=scale.none — индикатор НЕ создаёт вторую ценовую шкалу.
  let pineScript = `//@version=6
indicator("Multi-Ticker Levels (Goals/Stop/Cancel/Entry)", overlay=true, scale=scale.none)

// Тикер текущего графика без префикса биржи
var string raw_ticker = syminfo.ticker
var string ticker     = str.replace(syminfo.ticker, "MOEX:", "")

// Уровни
var float goal1        = 0.0
var float goal2        = 0.0
var float stop_level   = 0.0
var float cancel_level = 0.0
var float entry_level  = 0.0

// Флаг одноразовой отрисовки линий
var bool is_drawn = false
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

  // ===== Хвост Pine: фиксированные лучи + плашки на правой шкале (без второй шкалы) =====
  pineScript += `

// ---- Горизонтальные лучи: рисуем один раз ----
if barstate.islast and not is_drawn
    is_drawn := true
    if goal1 > 0
        line.new(bar_index[200], goal1, bar_index, goal1, extend=extend.right, color=color.red,    style=line.style_solid,  width=2)
    if goal2 > 0
        line.new(bar_index[200], goal2, bar_index, goal2, extend=extend.right, color=color.red,    style=line.style_dotted, width=2)
    if stop_level > 0
        line.new(bar_index[200], stop_level, bar_index, stop_level, extend=extend.right, color=color.orange, style=line.style_solid,  width=2)
    if cancel_level > 0
        line.new(bar_index[200], cancel_level, bar_index, cancel_level, extend=extend.right, color=color.gray,   style=line.style_dashed, width=2)
    if entry_level > 0
        line.new(bar_index[200], entry_level, bar_index, entry_level, extend=extend.right, color=color.green,  style=line.style_solid,  width=2)

// ---- Плашки на правой ценовой шкале (trackprice) ----
// Серии только на последнем баре: линия истории не рисуется, видна только плашка на шкале
goal1_series   = goal1        > 0 ? goal1        : na
goal2_series   = goal2        > 0 ? goal2        : na
stop_series    = stop_level   > 0 ? stop_level   : na
cancel_series  = cancel_level > 0 ? cancel_level : na
entry_series   = entry_level  > 0 ? entry_level  : na

plot(goal1_series,   title="Цель 1", color=color.red,    linewidth=1, style=plot.style_linebr, trackprice=true, show_last=1)
plot(goal2_series,   title="Цель 2", color=color.red,    linewidth=1, style=plot.style_linebr, trackprice=true, show_last=1)
plot(stop_series,    title="Стоп",   color=color.orange, linewidth=1, style=plot.style_linebr, trackprice=true, show_last=1)
plot(cancel_series,  title="Отмена", color=color.gray,   linewidth=1, style=plot.style_linebr, trackprice=true, show_last=1)
plot(entry_series,   title="Вход",   color=color.green,  linewidth=1, style=plot.style_linebr, trackprice=true, show_last=1)
`;

  res.set('Content-Type', 'text/plain; charset=utf-8');
  res.set('Content-Disposition', 'attachment; filename="generated_pine_script.txt"');
  return res.status(200).send(pineScript);
});

// ===== Экспорт для Vercel =====
module.exports = (req, res) => app(req, res);
