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
indicator("Excel Levels — Goals / Stop / Cancel / Entry", overlay=true, format=format.price, scale=scale.right)

// === Тикер без префикса MOEX: ===
var string ticker = str.replace(syminfo.ticker, "MOEX:", "")

// === Уровни (series, заполняются по тикеру) ===
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
  pineScript += `// <= 0 не рисуем
goal1_series  = goal1        > 0 ? goal1        : na
goal2_series  = goal2        > 0 ? goal2        : na
stop_series   = stop_level   > 0 ? stop_level   : na
cancel_series = cancel_level > 0 ? cancel_level : na
entry_series  = entry_level  > 0 ? entry_level  : na


// Цвета линий (по типам)
colGoal_line   = input.color(color.red,    "Цвет линии: Цель")
colStop_line   = input.color(color.orange, "Цвет линии: Стоп")
colCancel_line = input.color(color.gray,   "Цвет линии: Отмена")
colEntry_line  = input.color(color.green,  "Цвет линии: Вход")

// Прозрачность линий (0 = непрозрачно, 100 = полностью прозрачно)
lineTransp     = input.int(0, "Прозрачность линий (0-100)", minval=0, maxval=100)

// === Параметры линий ===
lineWidth = input.int(2, "Толщина линий", minval=1, maxval=6)
colGoal_lblBG   = input.color(color.red,    "Цвет фона лейбла: Цель")
colStop_lblBG   = input.color(color.orange, "Цвет фона лейбла: Стоп")
colCancel_lblBG = input.color(color.gray,   "Цвет фона лейбла: Отмена")
colEntry_lblBG  = input.color(color.green,  "Цвет фона лейбла: Вход")

// Прозрачность фона лейблов
labelTransp     = input.int(20, "Прозрачность фона лейбла (0-100)", minval=0, maxval=100)

// Цвет текста лейблов (общий)
labelTextColor  = input.color(color.white, "Цвет текста лейблов")

// === Хэндлы линий (храним один раз) ===
var line ln_goal1   = na
var line ln_goal2   = na
var line ln_stop    = na
var line ln_cancel  = na
var line ln_entry   = na

// Создаёт/обновляет/удаляет горизонтальную линию уровня и ВСЕГДА продлевает вправо
f_upsert_hline(line h, float lvl, color col, int width, int transp) =>
    line _h = h
    if na(lvl)
        if not na(_h)
            line.delete(_h)
        na
    else
        // скорректированный цвет с прозрачностью
        c = color.new(col, transp)
        if na(_h)
            // создаём отрезок и сразу тянем вправо
            _h := line.new(bar_index, lvl, bar_index, lvl, extend=extend.right)
        // на каждом баре фиксируем координаты и продление вправо
        line.set_xy1(_h, bar_index - 1, lvl)
        line.set_xy2(_h, bar_index,     lvl)
        line.set_extend(_h, extend.right)
        line.set_color(_h, c)
        line.set_width(_h, width)
        line.set_style(_h, line.style_solid)
        _h

// Обновляем линии каждый бар (они тянутся вправо до конца графика)
ln_goal1  := f_upsert_hline(ln_goal1,  goal1_series,  colGoal_line,   lineWidth, lineTransp)
ln_goal2  := f_upsert_hline(ln_goal2,  goal2_series,  colGoal_line,   lineWidth, lineTransp)
ln_stop   := f_upsert_hline(ln_stop,   stop_series,   colStop_line,   lineWidth, lineTransp)
ln_cancel := f_upsert_hline(ln_cancel, cancel_series, colCancel_line, lineWidth, lineTransp)
ln_entry  := f_upsert_hline(ln_entry,  entry_series,  colEntry_line,  lineWidth, lineTransp)

// === Лейблы у последнего бара (жёстко по цене) ===
showText      = input.bool(true,  "Показывать подписи у последнего бара")
labelsSizeStr = input.string("large", "Размер лейблов", options=["tiny","small","normal","large","huge"])
label_size = labelsSizeStr == "tiny" ? size.tiny : labelsSizeStr == "small" ? size.small : labelsSizeStr == "normal" ? size.normal : labelsSizeStr == "large" ? size.large : size.huge

var label l_goal1  = na
var label l_goal2  = na
var label l_stop   = na
var label l_cancel = na
var label l_entry  = na

if barstate.islast
    // чистим старые
    if not na(l_goal1)
        label.delete(l_goal1)
        l_goal1 := na
    if not na(l_goal2)
        label.delete(l_goal2)
        l_goal2 := na
    if not na(l_stop)
        label.delete(l_stop)
        l_stop := na
    if not na(l_cancel)
        label.delete(l_cancel)
        l_cancel := na
    if not na(l_entry)
        label.delete(l_entry)
        l_entry := na

    if showText
        if not na(goal1_series)
            l_goal1 := label.new(bar_index, goal1_series, "Цель 1 " + str.tostring(goal1_series), xloc.bar_index, yloc.price, color.new(colGoal_lblBG,   labelTransp), label.style_label_right, labelTextColor, label_size)
        if not na(goal2_series)
            l_goal2 := label.new(bar_index, goal2_series, "Цель 2 " + str.tostring(goal2_series), xloc.bar_index, yloc.price, color.new(colGoal_lblBG,   math.min(100, labelTransp + 20)), label.style_label_left, labelTextColor, label_size)
        if not na(stop_series)
            l_stop  := label.new(bar_index, stop_series,  "Стоп "   + str.tostring(stop_series),  xloc.bar_index, yloc.price, color.new(colStop_lblBG,   labelTransp), label.style_label_left, labelTextColor, label_size)
        if not na(cancel_series)
            l_cancel:= label.new(bar_index, cancel_series,"Отмена " + str.tostring(cancel_series), xloc.bar_index, yloc.price, color.new(colCancel_lblBG, labelTransp), label.style_label_left, labelTextColor, label_size)
        if not na(entry_series)
            l_entry := label.new(bar_index, entry_series, "Вход "   + str.tostring(entry_series),  xloc.bar_index, yloc.price, color.new(colEntry_lblBG,  labelTransp), label.style_label_left, labelTextColor, label_size)
`;

  res.set('Content-Type', 'text/plain; charset=utf-8');
  res.set('Content-Disposition', 'attachment; filename="generated_pine_script.txt"');
  return res.status(200).send(pineScript);
});

// ===== Экспорт для Vercel =====
module.exports = (req, res) => app(req, res);
