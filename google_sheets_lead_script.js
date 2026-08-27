/**
 * =========================================================================
 * STELLAR RESIDENCES • TULUM GARDENS | EXCEL / GOOGLE SHEETS LEAD COLLECTOR
 * =========================================================================
 * 
 * ИНСТРУКЦИЯ ПО ПОДКЛЮЧЕНИЮ ЗА 2 МИНУТЫ:
 * 1. Создайте новую таблицу на sheets.google.com (назовите её: "Лиды Tulum Gardens 2026").
 * 2. В верхнем меню таблицы нажмите: Расширения -> Apps Script.
 * 3. Удалите весь стандартный код в окне и вставьте этот скрипт.
 * 4. Нажмите синюю кнопку «Развернуть» (Deploy) в правом верхнем углу -> «Новое развертывание» (New deployment).
 * 5. В типе развертывания выберите «Веб-приложение» (Web app).
 * 6. Настройки:
 *    - Описание: "Сбор лидов с сайта"
 *    - Запуск от имени: "Я" (Me)
 *    - У кого есть доступ: "Все" (Anyone) — это важно для приема заявок с сайта!
 * 7. Нажмите «Развернуть» -> Предоставьте доступ своему аккаунту Google.
 * 8. Скопируйте полученный URL веб-приложения (заканчивается на /exec) и пришлите мне — я вставлю его в код сайта!
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getActiveSheet();

    // Создаем заголовки колонок, если таблица пустая
    if (sheet.getLastRow() === 0) {
      var headers = [
        "Дата",
        "Время",
        "Источник / Форма",
        "Имя клиента",
        "Телефон / WhatsApp",
        "Email",
        "Дополнительно (Квиз / Параметры)",
        "UTM Source",
        "UTM Medium",
        "UTM Campaign",
        "UTM Content",
        "UTM Term",
        "Страница / Реферер"
      ];
      sheet.appendRow(headers);
      
      // Оформление шапки таблицы
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground("#0E1210");
      headerRange.setFontColor("#D4AF37");
      headerRange.setFontWeight("bold");
      headerRange.setFontFamily("Arial");
      sheet.setFrozenRows(1);
    }

    var data;
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      data = e.parameter;
    }

    var nextRow = [
      data.date || Utilities.formatDate(new Date(), "GMT+3", "dd.MM.yyyy"),
      data.time || Utilities.formatDate(new Date(), "GMT+3", "HH:mm:ss"),
      data.source || "Форма сайта",
      data.name || "—",
      data.phone || "—",
      data.email || "—",
      data.extra || "—",
      data.utm_source || "—",
      data.utm_medium || "—",
      data.utm_campaign || "—",
      data.utm_content || "—",
      data.utm_term || "—",
      data.referrer || "Прямой переход"
    ];

    sheet.appendRow(nextRow);

    // Авто-выравнивание колонок
    for (var col = 1; col <= nextRow.length; col++) {
      sheet.autoResizeColumn(col);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ "result": "success", "row": sheet.getLastRow() }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "error", "error": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}