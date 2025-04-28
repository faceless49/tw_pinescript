const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const app = express();
const upload = multer({ dest: '/tmp/' }); // Временная папка для файлов
app.use(express.static('public'));
// Создаем POST-endpoint для загрузки файла
app.post('/generate-pine-script', upload.single('file'), (req, res) => {
  console.log('Получен файл', req.file);
  if (!req.file) {
    return res.status(400).send('Файл не загружен.');
  }

  // Чтение Excel-файла
  const workbook = XLSX.readFile(req.file.path);

  // Начало Pine Script
  let pineScript = `//@version=6
indicator("Multi-Ticker Goal Levels", overlay=true)

// Получение тикера графика
var string raw_ticker = syminfo.ticker
var string ticker = str.replace(syminfo.ticker, "MOEX:", "")
var float goal1 = 0.0
var float goal2 = 0.0

// Флаг для предотвращения повторного создания линий и меток
var bool is_drawn = false

// Переменные для хранения меток
var label goal1_label = na
var label goal2_label = na
var label aromat1_label = na
var label aromat2_label = na
`;

  // Обработка всех листов в Excel-файле
  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);

    data.forEach(row => {
      const ticker = row['Тикер'];
      const goal1 = parseFloat(row['Цель 1']);
      let goal2 = row['Цель 2'];

      // Обработка "Цель 2"
      if (typeof goal2 === 'string' && goal2.includes('/')) {
        goal2 = parseFloat(goal2.split('/')[0].trim());
      } else if (goal2 === '—' || goal2 === '--' || !goal2) {
        goal2 = 0.0;
      } else {
        goal2 = parseFloat(goal2);
      }

      if (isNaN(goal1) || !ticker) {
        console.log(`Ошибка в данных для тикера ${ticker}`);
        return;
      }

      pineScript += `
if ticker == "${ticker}"
    goal1 := ${goal1}
    goal2 := ${goal2}
`;
    });
  });

  // Добавление кода для отрисовки
  pineScript += `
// Отрисовка лучей и меток только один раз
if barstate.islast and not is_drawn
    // Устанавливаем флаг, чтобы не повторять отрисовку
    is_drawn := true

    // Горизонтальные лучи
    if not na(goal1) and goal1 > 0
        line.new(bar_index[200], goal1, bar_index, goal1, extend=extend.right, color=color.red, style=line.style_solid, width=2)
        label.delete(aromat1_label[1])
        aromat1_label := label.new(bar_index, goal1, "Цель по Аромат 1", color=color.red, style=label.style_label_down, textcolor=color.white, size=size.large)

    if not na(goal2) and goal2 > 0
        line.new(bar_index[200], goal2, bar_index, goal2, extend=extend.right, color=color.red, style=line.style_solid, width=2)
        label.delete(aromat1_label[1])
        aromat2_label := label.new(bar_index, goal2, "Цель по Аромат 2", color=color.red, style=label.style_label_down, textcolor=color.white, size=size.large)
`;

  // Удаление временного файла
  fs.unlinkSync(req.file.path);

  res.set('Content-Type', 'text/plain');
  res.set('Content-Disposition', 'attachment; filename="pine_script.txt"');
  // Отправка Pine Script клиенту
  res.send(pineScript);
});
module.exports = app;
// Запуск сервера
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Сервер запущен на порту ${PORT}`);
});
