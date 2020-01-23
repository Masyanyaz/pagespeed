const Excel = require('exceljs'); // Чтение excel файлов
const gtm = require('gtmetrix'); // gtmetrix
const email = [
  {
    email: 'tigramasya96@yandex.ru',
    apikey: 'cd496fa0f70e3074dfcfc522d2e7653c'
  }, {
    email: 'd878265@urhen.com',
    apikey: 'be309d4fa7b07ad89ab76dc37b466196'
  }, {
    email: 'd881016@urhen.com',
    apikey: '66291a141f096b2cacc8ee5d4aa96c7c'
  }, {
    email: 'd867605@urhen.com',
    apikey: '3a5eb2a5aaafeb9e46f566a52a9045b5'
  }, {
    email: 'd927864@urhen.com',
    apikey: '9780826307af23b21c741d107fb07a8c'
  }, {
    email: 'dbx17629@eveav.com',
    apikey: 'c2cfbbc3b7110c7ed9e1ec04143aaa92'
  }, {
    email: 'tft87009@zzrgg.com',
    apikey: '75f3c9c66bd516e4c314d0d0c28d01b4'
  }, {
    email: 'seb22249@bcaoo.com',
    apikey: '3a8d7024c3819f1e2f35851f601dbb68'
  }, {
    email: 'yms01563@bcaoo.com',
    apikey: 'e6a624006e81d5e6e053d2d7432289e6'
  }, {
    email: 'jmg77802@zzrgg.com',
    apikey: 'ed410f0ac8635d793729b26d843e503d'
  }, {
    email: 'ldi16624@zzrgg.com',
    apikey: '1dac45f6ae96fced72383d551dd7f7e9'
  }, {
    email: 'hwd92499@eveav.com',
    apikey: '63c2b93664a55b7231c091e6bc3ef442'
  }, {
    email: 'vgk33998@eveav.com',
    apikey: '3db32d48e6412a1e981d97d5c2c6f784'
  }, {
    email: 'd886793@urhen.com',
    apikey: 'c17b5833cac5c1b34b75a4d7b9f360c9'
  }];
let countEmail = 0;
let gtmetrix = gtm(email[countEmail]);
let countCredits;

const workbook = new Excel.Workbook(); // Создаем новую книгу

async function credits() {
  try {
    const data = await gtmetrix.account.status();
    return data.api_credits
  } catch (e) {
    console.log(e)
  }
}

// Функция анализа скорости страниц
async function startGtmetrix(url) {
  try {
    console.log('Test started...');
    const data = await gtmetrix.test.create({url});
    const result = await gtmetrix.test.get(data.test_id, 5000);
    const res = await result.results;
    return {
      pagespeed_score: res.pagespeed_score,
      yslow_score: res.yslow_score,
      fully_loaded_time: +(res.fully_loaded_time / 1000).toFixed(2),
      onload_time: +(res.onload_time / 1000).toFixed(2),
      page_bytes: +(res.page_bytes / 1024 / 1024).toFixed(2),
      page_elements: res.page_elements
    }
  } catch (e) {
    throw e
  }
}

// Запуск анализа, чтение файла excel и его запись
function readAndCheckAndWrite(tableName, sheetName, startCol = 3, endCol = 400, array = null) {
  workbook.xlsx.readFile(`./${tableName}.xlsx`) // Чтение файла
          .then(async function (data) {
            do {
              countCredits = await credits();
              if (!countCredits) gtmetrix = gtm(email[countEmail++]);
            } while (!countCredits && countEmail <= email.length);
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
                  while (!countCredits && countEmail <= email.length) {
                    gtmetrix = gtm(email[countEmail++]);
                    countCredits = await credits();
                    console.log(`Замеры закончились! Данные изменены! Email[${countEmail}]: ${email[countEmail].email}`)
                  }
                  if (!countCredits && countEmail > email.length) {
                    console.log(`Все данные закончились. Подождите или добавьте новые!`)
                    return
                  }
                  const results = await startGtmetrix(a); // Запуск анализа
                  let val = 0; // Сдвиг по стобцам
                  let color = ''; // Цвет ячейки

                  for (let key in results) {
                    let resultValue = results[key]; // Значение анализа
                    switch (key) {
                      case 'yslow_score':
                      case 'pagespeed_score':
                        if (resultValue >= 90) color = '008000'; // Если значение больше 90, то цвет зеленый
                        if (resultValue >= 70 && resultValue < 90) color = 'FFFF00'; // Если значение больше 70 и меньше 90, то цвет желтый
                        if (resultValue < 70) color = 'FF0000'; // Если значение меньше 70, то цвет красный
                        break;
                      case 'page_elements':
                        if (resultValue >= 90) color = 'FF0000'; // Если значение больше 90, то цвет красный
                        if (resultValue >= 50 && resultValue < 90) color = 'FFFF00'; // Если значение больше 70 и меньше 90, то цвет желтый
                        if (resultValue < 50) color = '008000'; // Если значение меньше 70, то цвет зеленый
                        break;
                      case 'fully_loaded_time':
                        if (resultValue >= 10) color = 'FF0000'; // Если значение больше 90, то цвет красный
                        if (resultValue >= 7 && resultValue < 10) color = 'FFFF00'; // Если значение больше 70 и меньше 90, то цвет желтый
                        if (resultValue < 7) color = '008000'; // Если значение меньше 70, то цвет зеленый
                        break;
                      case 'onload_time':
                        if (resultValue >= 5) color = 'FF0000'; // Если значение больше 90, то цвет красный
                        if (resultValue >= 3.5 && resultValue < 5) color = 'FFFF00'; // Если значение больше 70 и меньше 90, то цвет желтый
                        if (resultValue < 3.5) color = '008000'; // Если значение меньше 70, то цвет зеленый
                        break;
                      case 'page_bytes':
                        if (resultValue >= 3.2) color = 'FF0000'; // Если значение больше 90, то цвет красный
                        if (resultValue >= 1.5 && resultValue < 3.2) color = 'FFFF00'; // Если значение больше 70 и меньше 90, то цвет желтый
                        if (resultValue < 1.5) color = '008000'; // Если значение меньше 70, то цвет зеленый
                        break;
                    }

                    const cell = worksheet.getRow(array !== null ? array[i] : startCol + i).getCell(13 + val); // Рабочая ячейка

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
                  countCredits--;
                  console.log(`${i}/${array !== null ? array.length : endCol - startCol + 1} ${a} written! Осталось замеров: ${countCredits}`) // Выводим прогресс
                } catch (e) {
                  console.log(e); // Вывод ошибки
                  i++;
                  countCredits--;
                }
              }
            })();
          })
}

readAndCheckAndWrite('PageSpeed', 'FR', 4, 4);