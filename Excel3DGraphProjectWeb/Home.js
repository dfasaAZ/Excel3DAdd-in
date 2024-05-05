(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // Функцию инициализации необходимо выполнять при каждой загрузке новой страницы.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Инициализировать механизм уведомлений и скрыть его
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            
            // Если не используется Excel 2016, использовать резервную логику.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("В этом примере показано отображение значения ячеек, выбранных в таблице.");
                $('#button-text').text("Отобразить!");
                $('#button-desc').text("Отобразить выделение");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $("#template-description").text("В этом примере показано выделение наивысшего значения из выбранных в таблице ячеек.");
            $('#button-text').text("Выделить!");
            $('#button-desc').text("Выделение самого большого числа.");
                
            loadSampleData();

            // Добавить обработчик события щелчка кнопкой мыши для выделенной кнопки.
            $('#highlight-button').click(hightlightHighestValue);
        });
    };

    function loadSampleData() {
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Запустить пакетную операцию для объектной модели Excel
        Excel.run(function (ctx) {
            // Создать прокси-объект для переменной
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // Поставить в очередь команду для записи демонстрационных данных в лист
            sheet.getRange("B3:D5").values = values;

            // Запустить команду из очереди и возвратить обещание отобразить выполнение задачи
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    function hightlightHighestValue() {
        // Запустить пакетную операцию для объектной модели Excel
        Excel.run(function (ctx) {
            // Создать прокси-объект для выделенного диапазона и загрузить его свойства
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // Запустить команду из очереди и возвратить обещание отобразить выполнение задачи
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // Найти ячейку для выделения
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // Выделить ячейку
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
        })
        .catch(errorHandler);
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('Выбранный текст:', '"' + result.value + '"');
                } else {
                    showNotification('Ошибка', result.error.message);
                }
            });
    }

    // Вспомогательная функция для обработки ошибок
    function errorHandler(error) {
        // Всегда перехватывайте любые накопленные ошибки, возникающие при выполнении Excel.run
        showNotification("Ошибка", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Вспомогательная функция для отображения уведомлений
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
