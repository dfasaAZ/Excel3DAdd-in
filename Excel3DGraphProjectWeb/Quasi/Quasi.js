var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
(function () {
    "use strict";
    var cellToHighlight;
    var messageBanner;
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        window.Promise = OfficeExtension.Promise;
        loadSampleData();
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");
                $('#highlight-button').click(displaySelectedCells);
                return;
            }
            $('#color-button').on('click', function () {
                findQuasi();
            });
            $('#frequency-button').on('click', function () {
                buildFrequency();
            });
            $('#centers-button').on('click', function () {
                buildCenters();
            });
        });
    };
    /**
     * Находит квазициклы у выделенного графика, вызывает функцию для его раскраски
     *
     * Наносит на лист с метаданными частоты квазициклов
     */
    function findQuasi() {
        /*
        Взять из инпута погрешность
        Вызвать функцию поиска
        Найти частоты
        Записать в ячейки на листе метаданных
        Найти центры
        Записать в таблицу centers+${id}
        */
        let epsilon = document.getElementById("epsilon").value;
        processQuasiCycles(epsilon);
    }
    function buildFrequency() {
        /*
        Найти на листе с метаданными ячейки
        Если они пустые - вызвать showNotification
        Если нет, построить гистограмму
        */
    }
    function buildCenters() {
        /*
       Найти на листе с метаданными таблицу centers+${id}
       Если её нет - вызвать showNotification
       Если есть, построить полноценный 3d график
       */
    }
    Office.onReady(() => __awaiter(this, void 0, void 0, function* () {
        yield Excel.run((context) => __awaiter(this, void 0, void 0, function* () {
        })).catch(errorHandler);
    }));
    function loadSampleData() {
        return __awaiter(this, void 0, void 0, function* () {
            //TODO: Проверка алгоритмов квазицикла, убрать потом
            quasiTest();
            yield loadSettings();
            // Run the Excel operations
            Excel.run(function (context) {
                // Sync the changes to Excel
                return context.sync();
            })
                .catch(errorHandler);
        });
    }
    function loadSettings() {
        return __awaiter(this, void 0, void 0, function* () {
            let angles;
            yield Excel.run((context) => __awaiter(this, void 0, void 0, function* () {
                let selectedGraph;
                let graphName;
                let graphAngles;
                selectedGraph = context.workbook.getActiveChart().load("name");
                yield context.sync().then(function processSelectedGraphs() {
                    let name = selectedGraph.name;
                    if (name != null) {
                        $('#graphName').text("Работа с графиком id" + name);
                        $('#graphId').text(name);
                        graphName = name;
                        $('#pageContent').show();
                    }
                });
                // Loading graph related properties
                graphAngles = context.workbook.tables.getItem("Angles" + graphName).rows.getItemAt(0).load("values");
                yield context.sync().then(function processSelectedGraphs() {
                    angles = graphAngles.values;
                });
            }));
        });
    }
    function processQuasiCycles(epsilon) {
        return __awaiter(this, void 0, void 0, function* () {
            let graphName = document.getElementById('graphId').innerText; // получить с html, загружать как и в home
            let graphData;
            let result;
            yield Excel.run((context) => __awaiter(this, void 0, void 0, function* () {
                // Loading graph related properties
                graphData = context.workbook.tables.getItem("SourceData" + graphName).rows.load("items");
                return context.sync().then(function () {
                    //Здесь пройтись по строкам и запихать их в квазициклы
                    const points = graphData.items.map(row => {
                        return [
                            row.values[0][0],
                            row.values[0][1],
                            row.values[0][2]
                        ];
                    });
                    result = findMaximumNonIntersectingArrays(findQuasiCycles(points, epsilon)); // Массив квазициклов
                    console.log("Quasi-cycles result:", result);
                    return context.sync();
                }).then(context.sync);
            })).catch(errorHandler);
            colorQuasiCyclesInChart(graphName, result);
        });
    }
    function colorQuasiCyclesInChart(id, quasiCycles) {
        return Excel.run((context) => __awaiter(this, void 0, void 0, function* () {
            let charts = context.workbook.worksheets.getActiveWorksheet().load("charts");
            yield context.sync();
            let pointsLoad = charts.charts.getItem(id).series.getItemAt(0).points.load("items");
            // Set color for chart point.
            yield context.sync();
            let points = pointsLoad;
            const colors = generateColors(quasiCycles.length);
            // Color each quasi-cycle
            quasiCycles.forEach((cycle, index) => {
                const color = colors[index];
                cycle.indices.forEach(pointIndex => {
                    // Excel uses 1-based indexing for points
                    const point = points.items[pointIndex];
                    point.set({ markerForegroundColor: color, markerBackgroundColor: color });
                });
            });
            yield context.sync();
        }));
    }
    // Helper function to generate distinct colors
    function generateColors(count) {
        const colors = [];
        for (let i = 0; i < count; i++) {
            // Use HSL color space for even distribution, then convert to HEX
            const hue = (i * 137.508) % 360; // Use golden angle approximation
            const saturation = 70; // Fixed saturation for vibrant colors
            const lightness = 50; // Fixed lightness for medium brightness
            // Convert HSL to RGB
            const chroma = (1 - Math.abs(2 * lightness / 100 - 1)) * saturation / 100;
            const huePrime = hue / 60;
            const x = chroma * (1 - Math.abs((huePrime % 2) - 1));
            let r, g, b;
            if (huePrime >= 0 && huePrime < 1) {
                [r, g, b] = [chroma, x, 0];
            }
            else if (huePrime >= 1 && huePrime < 2) {
                [r, g, b] = [x, chroma, 0];
            }
            else if (huePrime >= 2 && huePrime < 3) {
                [r, g, b] = [0, chroma, x];
            }
            else if (huePrime >= 3 && huePrime < 4) {
                [r, g, b] = [0, x, chroma];
            }
            else if (huePrime >= 4 && huePrime < 5) {
                [r, g, b] = [x, 0, chroma];
            }
            else {
                [r, g, b] = [chroma, 0, x];
            }
            const m = lightness / 100 - chroma / 2;
            r = Math.round((r + m) * 255);
            g = Math.round((g + m) * 255);
            b = Math.round((b + m) * 255);
            // Convert RGB to HEX
            const hex = '#' + [r, g, b].map(x => {
                const hex = x.toString(16);
                return hex.length === 1 ? '0' + hex : hex;
            }).join('');
            colors.push(hex);
        }
        return colors;
    }
    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                showNotification('The selected text is:', '"' + result.value + '"');
            }
            else {
                showNotification('Error', result.error.message);
            }
        });
    }
    //  Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
    /**
     * AI generated quasi-cycle shit
     *
     * Need to double check
     */
    function euclideanDistance(point1, point2) {
        const [x1, y1, z1] = point1;
        const [x2, y2, z2] = point2;
        return Math.sqrt(Math.pow((x2 - x1), 2) + Math.pow((y2 - y1), 2) + Math.pow((z2 - z1), 2));
    }
    function findQuasiCycles(points, epsilon) {
        const quasiCycles = [];
        for (let i = 0; i < points.length; i++) {
            const start = points[i];
            for (let j = i + 2; j < points.length; j++) {
                const end = points[j];
                const distance = euclideanDistance(start, end);
                if (distance <= epsilon) {
                    const indices = Array.from({ length: j - i + 1 }, (_, k) => i + k);
                    quasiCycles.push({ start: start, end: end, indices });
                }
            }
        }
        return quasiCycles;
    }
    function findMaximumNonIntersectingArrays(quasiCycles) {
        const result = [];
        const usedIndices = new Set();
        quasiCycles.forEach(quasiCycle => {
            const isIntersecting = quasiCycle.indices.some(index => usedIndices.has(index));
            if (!isIntersecting) {
                result.push(quasiCycle);
                quasiCycle.indices.forEach(index => usedIndices.add(index));
            }
        });
        return result;
    }
    /**
     * Function to check work of the qusicycle algorithm
     */
    function quasiTest() {
        // Example usage
        // const points: [number, number, number][] = [[1, 2, 3], [4, 5, 6], [7, 8, 9], [1.1, 2.1, 2.9]];
        const points = [
            [0, 0, 0], [1, 0, 0], [2, 0, 0], [3, 0, 0], [1, 0, 0], // Single quasi-cycle
            [5, 5, 0], [6, 5, 0], [7, 5, 0], [8, 5, 0], [9, 5, 0], [10, 5, 0], [5.1, 5.1, 0], // Multiple quasi-cycles
            [10, 10, 0], [11, 10, 0], [12, 10, 0], [13, 10, 0], [14, 10, 0], [10.1, 10.1, 0], // Intersecting quasi-cycles
            [15, 15, 0], [16, 15, 0], [17, 15, 0], [18, 15, 0], [19, 15, 0], [15.1, 15.1, 0], // Negative case (neighbors within epsilon)
        ];
        const epsilon = 0.5;
        const quasiCycles = findQuasiCycles(points, epsilon);
        let result = findMaximumNonIntersectingArrays(quasiCycles);
        console.log(result);
        const inputQuasiCycles = [
            { start: [0, 0, 0], end: [0, 0, 0], indices: [1, 2, 3, 4] },
            { start: [0, 0, 0], end: [0, 0, 0], indices: [3, 4, 5] },
            { start: [0, 0, 0], end: [0, 0, 0], indices: [8, 9, 10, 11] },
            { start: [0, 0, 0], end: [0, 0, 0], indices: [10, 11, 12] },
            { start: [0, 0, 0], end: [0, 0, 0], indices: [12, 13, 14] },
            { start: [0, 0, 0], end: [0, 0, 0], indices: [13, 14, 15] },
        ];
        result = findMaximumNonIntersectingArrays(inputQuasiCycles);
        console.log(result);
    }
})();
//# sourceMappingURL=Quasi.js.map