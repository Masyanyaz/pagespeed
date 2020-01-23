const lighthouse = require('lighthouse'); // Тест скорости
const chromeLauncher = require('chrome-launcher'); // хром, который запускает тест скорости
const Excel = require('exceljs'); // Чтение excel файлов
const {promisify} = require('util');
const {writeFile} = require('fs'); // Для записи файлов
const pWriteFile = promisify(writeFile);

const workbook = new Excel.Workbook(); // Создаем новую книгу

// Функция анализа скорости страниц
async function launchChromeAndRunLighthouse(url , opts, config = null) {
  const chrome = await chromeLauncher.launch({chromeFlags: opts.chromeFlags}); // Запуск хрома
  opts.port = chrome.port; // Установка порта для хрома(если надо)
  const {lhr, report} = await lighthouse(url, opts, config); // Результаты анализа в json и html формате
  await chrome.kill(); // Закрываем хром
  await setTimeout(() => console.log('chrome killed'), 1000); // Для предотвращения ошибки. что хром не закрыт
  return {lhr, report}; // Возвращаем json и html формат
}

// Опции для функции
const opts = {
  chromeFlags: ['--show-paint-rects'],
  output: 'html'
};

// Дополнительные настройки(можно задать какие категории анализировать, что анализировать и в для чего(пк, моб))
const config = {
  extends: 'lighthouse:default'
};

// Запуск анализа, чтение файла excel и его запись
function readAndCheckAndWrite(tableName, sheetName, startCol = 3, endCol = 400, array = null) {
  workbook.xlsx.readFile(`./${tableName}.xlsx`) // Чтение файла
          .then(function (data) {
            let worksheet = workbook.getWorksheet(sheetName); // Чтение листа файла
            let arrUrl = []; // Массив все url которые надо проверить
            for (let i = startCol; i <= endCol; i++) {
              let value = worksheet.getCell(`A${i}`).value; // Значение ячейки с url

              if (value === null || (array !== null && !array.includes(i))) continue; // если пусто (и если задан массив и значение не входит в него) перейти дальше
              arrUrl.push(value.text || value) // добавить url в массив
            }

            (async () => {
              let i = 0;
              for (let a of arrUrl) {
                try {
                  const results = await launchChromeAndRunLighthouse(a, opts, config); // Запуск анализа
                  let fileName = a.replace('https://', '').split('/').join('_'); // Имя файла для сохранения отчета
                  let val = 0; // Сдвиг по стобцам
                  let newCell = 0; // Сдвиг по стобцам, при втором анализе

                  // await pWriteFile(`./report/${sheetName}/${fileName}.html`, results.report); // Сохранение отчета в html формате

                  if (worksheet.getRow(array !== null ? array[i] : startCol + i).getCell(3).value) newCell = 5; // Если url проверяется второй раз, то сдвинуть ячейки

                  for (let key in results.lhr.categories) {
                    let resultValue = +results.lhr.categories[key].score * 100; // Значение анализа
                    let color = ''; // Цвет ячейки
                    const cell = worksheet.getRow(array !== null ? array[i] : startCol + i).getCell(3 + val + newCell); // Рабочая ячейка

                    if (resultValue >= 90) color = '008000'; // Если значение больше 90, то цвет зеленый
                    if (resultValue >= 70 && resultValue < 90) color = 'FFFF00'; // Если значение больше 70 и меньше 90, то цвет желтый
                    if (resultValue < 70) color = 'FF0000'; // Если значение меньше 70, то цвет красный

                    cell.value = resultValue; // Присвоение ячейки значения анализа
                    cell.style = JSON.parse(JSON.stringify(cell.style)); // Чтобы работала возможность перезаписи цвета(фича)
                    // Заливка ячейки цветом
                    cell.fill = {
                      type: 'pattern',
                      pattern: 'solid',
                      fgColor: {argb: color}
                    };
                    // Установка границ у ячейки
                    cell.border = {
                      top: {style: 'thin'},
                      left: {style: 'thin'},
                      bottom: {style: 'thin'},
                      right: {style: 'thin'}
                    };

                    val++
                  }
                  await workbook.xlsx.writeFile(`./${tableName}.xlsx`); // Сохраняем изменения в файле
                  i++;
                  console.log(`${i}/${array !== null ? array.length : endCol - startCol + 1} ${fileName} written!`) // Выводим прогресс
                } catch (e) {
                  console.log(e); // Вывод ошибки
                  i++
                }
              }
            })();
          })
}

readAndCheckAndWrite('PageSpeed', 'FR', 3, 3);