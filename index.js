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

// Настройки лейблов (не создают шкалу)
showInlineLabels   = input.bool(false, "Показывать лейблы на графике")
labelsOffsetBars   = input.int(15, "Смещение лейблов вправо (в барах)", minval=0, maxval=500)
labelsSizeStr      = input.string("large", "Размер лейблов", options=["tiny","small","normal","large","huge"])

label_size = size.normal
label_size := labelsSizeStr == "tiny"   ? size.tiny   :
              labelsSizeStr == "small"  ? size.small  :
              labelsSizeStr == "normal" ? size.normal :
              labelsSizeStr == "large"  ? size.large  : size.huge

// Держатели лейблов
var label lbl_goal1   = na
var label lbl_goal2   = na
var label lbl_stop    = na
var label lbl_cancel  = na
var label lbl_entry   = na
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

  // ===== Хвост Pine (без plot и без второй шкалы) =====
  pineScript += `

// ---- Горизонтальные лучи: рисуем один раз ----
if barstate.islast and not is_drawn
    is_drawn := true
    if goal1 > 0
        line.new(bar_index[200], goal1, bar_index, goal1, extend=extend.right, color=color.red,   style=line.style_solid,  width=2)
    if goal2 > 0
        line.new(bar_index[200], goal2, bar_index, goal2, extend=extend.right, color=color.red,   style=line.style_dotted, width=2)
    if stop_level > 0
        line.new(bar_index[200], stop_level, bar_index, stop_level, extend=extend.right, color=color.orange, style=line.style_solid,  width=2)
    if cancel_level > 0
        line.new(bar_index[200], cancel_level, bar_index, cancel_level, extend=extend.right, color=color.gray,   style=line.style_dashed, width=2)
    if entry_level > 0
        line.new(bar_index[200], entry_level, bar_index, entry_level, extend=extend.right, color=color.green,  style=line.style_solid,  width=2)

// ---- Плашки на правой шкале (без второй шкалы и без лейблов) ----
// Делаем серии только на последнем баре — график не «мазюкится», но плашка есть.
goal1_series   = goal1        > 0 ? goal1        : na
goal2_series   = goal2        > 0 ? goal2        : na
stop_series    = stop_level   > 0 ? stop_level   : na
cancel_series  = cancel_level > 0 ? cancel_level : na
entry_series   = entry_level  > 0 ? entry_level  : na

// ВАЖНО: индикатор объявлен с scale=scale.none, поэтому вторая шкала не появится.
plot(goal1_series,   title="Цель 1", color=color.red,    linewidth=1, style=plot.style_linebr, trackprice=true, show_last=1)
plot(goal2_series,   title="Цель 2", color=color.red,    linewidth=1, style=plot.style_linebr, trackprice=true, show_last=1)
plot(stop_series,    title="Стоп",   color=color.orange, linewidth=1, style=plot.style_linebr, trackprice=true, show_last=1)
plot(cancel_series,  title="Отмена", color=color.gray,   linewidth=1, style=plot.style_linebr, trackprice=true, show_last=1)
plot(entry_series,   title="Вход",   color=color.green,  linewidth=1, style=plot.style_linebr, trackprice=true, show_last=1)

// ---- Опциональные лейблы у правого края (не создают шкалу) ----
futureMs = int(timeframe.in_seconds()) * 1000 * labelsOffsetBars
xRight   = timenow + futureMs

// Перерисовываем лейблы на каждом баре, чтобы они «держались» правого края
if true
    // Удаляем предыдущие
    if not na(lbl_goal1)
        label.delete(lbl_goal1)
        lbl_goal1 := na
    if not na(lbl_goal2)
        label.delete(lbl_goal2)
        lbl_goal2 := na
    if not na(lbl_stop)
        label.delete(lbl_stop)
        lbl_stop := na
    if not na(lbl_cancel)
        label.delete(lbl_cancel)
        lbl_cancel := na
    if not na(lbl_entry)
        label.delete(lbl_entry)
        lbl_entry := na

    if showInlineLabels
        if goal1 > 0
            lbl_goal1 := label.new(x=xRight, y=goal1, xloc=xloc.bar_time, style=label.style_label_right, text="Цель 1", color=color.new(color.red, 20), textcolor=color.white, size=label_size)
        if goal2 > 0
            lbl_goal2 := label.new(x=xRight, y=goal2, xloc=xloc.bar_time, style=label.style_label_right, text="Цель 2", color=color.new(color.red, 40), textcolor=color.white, size=label_size)
        if stop_level > 0
            lbl_stop  := label.new(x=xRight, y=stop_level, xloc=xloc.bar_time, style=label.style_label_right, text="Стоп", color=color.new(color.orange, 20), textcolor=color.white, size=label_size)
        if cancel_level > 0
            lbl_cancel := label.new(x=xRight, y=cancel_level, xloc=xloc.bar_time, style=label.style_label_right, text="Отмена", color=color.new(color.gray, 20), textcolor=color.white, size=label_size)
        if entry_level > 0
            lbl_entry := label.new(x=xRight, y=entry_level, xloc=xloc.bar_time, style=label.style_label_right, text="Вход", color=color.new(color.green, 20), textcolor=color.white, size=label_size)

// ---- Алерты (опционально, не влияют на шкалы) ----
crossUp(src, level)   => ta.crossover(src, level)
crossDown(src, level) => ta.crossunder(src, level)

entry_cross_up   = (entry_level  > 0) and crossUp(close, entry_level)
entry_cross_down = (entry_level  > 0) and crossDown(close, entry_level)
alertcondition(entry_cross_up,   "Entry Cross Up",   "Цена пересекла Вход снизу вверх")
alertcondition(entry_cross_down, "Entry Cross Down", "Цена пересекла Вход сверху вниз")

stop_hit        = (stop_level   > 0) and close <= stop_level
cancel_hit_down = (cancel_level > 0) and close <= cancel_level
cancel_hit_up   = (cancel_level > 0) and close >= cancel_level

alertcondition(stop_hit,        "Stop Hit",    "Цена <= Стоп")
alertcondition(cancel_hit_down, "Cancel Down", "Цена <= Отмена")
alertcondition(cancel_hit_up,   "Cancel Up",   "Цена >= Отмена")

goal1_reached = (goal1 > 0) and close >= goal1
goal2_reached = (goal2 > 0) and close >= goal2

alertcondition(goal1_reached, "Goal 1 Reached", "Цена достигла Цель 1")
alertcondition(goal2_reached, "Goal 2 Reached", "Цена достигла Цель 2")
`;

  res.set('Content-Type', 'text/plain; charset=utf-8');
  res.set('Content-Disposition', 'attachment; filename="generated_pine_script.txt"');
  return res.status(200).send(pineScript);
});

// ===== Экспорт для Vercel =====
module.exports = (req, res) => app(req, res);
