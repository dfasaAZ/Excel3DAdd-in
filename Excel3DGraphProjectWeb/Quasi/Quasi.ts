(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {


        (window as any).Promise = OfficeExtension.Promise;
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
   async function findQuasi() {
        /*
        Взять из инпута погрешность
        Вызвать функцию поиска
        Найти частоты
        Записать в ячейки на листе метаданных
        Найти центры
        Записать в таблицу centers+${id}
        */
        let epsilon = (<HTMLInputElement>document.getElementById("epsilon")).value;
       await processQuasiCycles(epsilon).then((result) => {
           findFrequency(result);
           findCenters(result);
       });
        

    }
    function buildFrequency(): void {
        /*
       Найти на листе с метаданными таблицу 
       Если её нет - вызвать showNotification
       Если есть, построить гистограмму
       */
        Excel.run(async (context) => {
            let graphName = (document.getElementById('graphId') as HTMLElement).innerText;
            let sheetName = `Graph${graphName}`;
            let tableName = `Frequencies${graphName}`;

            // Get the active sheet
            let sheet = context.workbook.worksheets.getItem(sheetName);

            // Check if the table exists
            let tables = sheet.tables;
            tables.load("items/name");
            await context.sync();

            let table = tables.items.find(t => t.name === tableName);

            if (!table) {
                showNotification("Отсутствуют данные", `Выполните поиск квазициклов перед постоением графика частот`);
                return;
            }

       
            let dataRange = table.getDataBodyRange();
            dataRange.load("values");
            await context.sync();

            let data = dataRange.values[0];

            let chartRange = context.workbook.worksheets.getActiveWorksheet().getRange("L11:U26");
            let chart = context.workbook.worksheets.getActiveWorksheet().charts.add("ColumnClustered", chartRange, "Rows");
            chart.title.text = "график частот"
            chart.name = "Frequency " + graphName;
            chart.legend.visible = false;
            chart.categoryLabelLevel = -1;
            chart.dataLabels.showValue = true;
            chart.dataLabels.showSeriesName = true;
            chart.dataLabels.separator = "\nколво:\n"
            chart.setData(table.getRange());

            let maxValue = Math.max(...data);
            let maxIndex = data.indexOf(maxValue);
            let pointsCollection = chart.series.getItemAt(maxIndex).points;
            let point = pointsCollection.getItemAt(0);

            // Set color for chart point.
            point.format.fill.setSolidColor('red');

            await context.sync();
        }).catch(error => {
            console.error("Error: " + error);
            showNotification("Error", "An error occurred while building the frequency histogram.");
        });
    }
    function buildCenters() {
        /*
       Найти на листе с метаданными таблицу centers+${id} 
       Если её нет - вызвать showNotification
       Если есть, построить полноценный 3d график
       */
        createNewCentersGraph();
    }
    Office.onReady(async () => {

        await Excel.run(async (context) => {
        }).catch(errorHandler);

    });
    /**
     * Находит центры габаритных прямоугольников и записывает на лист графика
     * @param cycles
     */
    function findCenters(cycles: QuasiCycle[]): void {
        Excel.run(async (context) => {
            let graphName = (document.getElementById('graphId') as HTMLElement).innerText;
            let sheetName = `Graph${graphName}`;
            let centersTableName = `Centers${graphName}`;
            let sourceDataTableName = `SourceData${graphName}`;

            // Get the sheet
            let sheet = context.workbook.worksheets.getItem(sheetName);

            // Check if the centers table already exists
            let tables = sheet.tables;
            tables.load("items/name");
            await context.sync();

            let centersTable = tables.items.find(t => t.name === centersTableName);

            if (centersTable) {
                showNotification("Table Exists", `The table "${centersTableName}" already exists.`);
                return;
            }

            // Get the source data table
            let sourceDataTable = tables.items.find(t => t.name === sourceDataTableName);
            if (!sourceDataTable) {
                showNotification("Error", `Source data table "${sourceDataTableName}" not found.`);
                return;
            }

            // Load source data
            let sourceDataRange = sourceDataTable.getDataBodyRange();
            sourceDataRange.load("values");
            await context.sync();

            let sourceData = sourceDataRange.values;

            // Calculate centers
            let centers = cycles.map(cycle => {
                let points = cycle.indices.map(index => sourceData[index]); // Assuming indices are 1-based
                let sumX = 0, sumY = 0, sumZ = 0;
                points.forEach(point => {
                    sumX += point[0];
                    sumY += point[1];
                    sumZ += point[2];
                });
                return [
                    Number((sumX / points.length).toFixed(3)),
                    Number((sumY / points.length).toFixed(3)),
                    Number((sumZ / points.length).toFixed(3))
                ];
            });

            // Create the centers table
            let range = sheet.getRange("H13:J13");
            centersTable = sheet.tables.add(range, true);
            centersTable.name = centersTableName;

            // Set the headers
            let headerRange = centersTable.getHeaderRowRange();
            headerRange.values = [["X", "Y", "Z"]];

            // Set the data
            let dataRange = centersTable.getDataBodyRange();
            centersTable.rows.add(null, centers);

            await context.sync();
        }).catch(error => {
            console.error("Error: " + error);
            showNotification("Error", "An error occurred while creating the centers table.");
        });
    }
    /**
     * Считает длины квазициклов и записывает на лист
     * 
     * @param cycles
     */
    function findFrequency(cycles): void {
        Excel.run(async (context) => {
            let graphName = (document.getElementById('graphId') as HTMLElement).innerText;
            let sheetName = `Graph${graphName}`;
            let tableName = `Frequencies${graphName}`;

            // Get the sheet
            let sheet = context.workbook.worksheets.getItem(sheetName);

            // Check if the table already exists
            let tables = sheet.tables;
            tables.load("items/name");
            await context.sync();

            let table = tables.items.find(t => t.name === tableName);

            if (table) {
                // Table exists, show notification
                showNotification("Таблица существует", `Таблица "${tableName}" уже существует.`);
                return;
            }

            // Create frequency counts
            let frequencyCounts = new Array(10).fill(0);
            cycles.forEach(cycle => {
                let length = cycle.indices.length;
                if (length > 0 && length <= 10) {
                    frequencyCounts[length - 1]++;
                }
            });

            // Create the table
            let range = sheet.getRange("H10:Q11");
            table = sheet.tables.add(range, true);
            table.name = tableName;

            // Set the headers
            let headerRange = table.getHeaderRowRange();
            headerRange.values = [["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"]];

            // Set the data
            let dataRange = table.rows.getItemAt(0).getRange();
            dataRange.values = [frequencyCounts];

            await context.sync();
        }).catch(error => {
            console.error("Error: " + error);
            showNotification("Error", "An error occurred while creating the frequency table.");
        });
    }
  async  function loadSampleData() {


        //TODO: Проверка алгоритмов квазицикла, убрать потом
        quasiTest();

        await loadSettings();

        // Run the Excel operations
        Excel.run(function (context) {

 

            // Sync the changes to Excel
            return context.sync();
        })
            .catch(errorHandler);
    }

    async function loadSettings() {
        let angles;
        await Excel.run(async (context) => {
            let selectedGraph;
            let graphName;
            let graphAngles;
            selectedGraph = context.workbook.getActiveChart().load("name");
            await context.sync().then(function processSelectedGraphs() {
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
            await context.sync().then(function processSelectedGraphs() {
                angles = graphAngles.values;

            });
        });
    }
    /**
     * Функция для обработки квазициклов, находит их на выбранном графике и красит точки
     *
     * @param epsilon Эпсилон окрестность в которой следует искать пересечение
     * @returns массив квазициклов
     */
   async function processQuasiCycles(epsilon) {
        let graphName = document.getElementById('graphId').innerText;// получить с html, загружать как и в home
        let graphData;
       let result;
      
        await Excel.run(async (context) => {
           
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
               result = findMaximumNonIntersectingArrays(findQuasiCycles(points, epsilon));// Массив квазициклов
                
                console.log("Quasi-cycles result:", result);
               

                return context.sync()

            }).then(context.sync);

        }).catch(errorHandler);
       colorQuasiCyclesInChart(graphName, result);
       return result
    }
    function colorQuasiCyclesInChart(id, quasiCycles) {
        return Excel.run(async (context) => {
            let charts = context.workbook.worksheets.getActiveWorksheet().load("charts");
            await context.sync();
            let pointsLoad = charts.charts.getItem(id).series.getItemAt(0).points.load("items");
            // Set color for chart point.
            await context.sync();
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

            await context.sync();
        });
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

            if (huePrime >= 0 && huePrime < 1) { [r, g, b] = [chroma, x, 0]; }
            else if (huePrime >= 1 && huePrime < 2) { [r, g, b] = [x, chroma, 0]; }
            else if (huePrime >= 2 && huePrime < 3) { [r, g, b] = [0, chroma, x]; }
            else if (huePrime >= 3 && huePrime < 4) { [r, g, b] = [0, x, chroma]; }
            else if (huePrime >= 4 && huePrime < 5) { [r, g, b] = [x, 0, chroma]; }
            else { [r, g, b] = [chroma, 0, x]; }

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
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
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
    function euclideanDistance(point1: [number, number, number], point2: [number, number, number]): number {
        const [x1, y1, z1] = point1;
        const [x2, y2, z2] = point2;
        return Math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2 + (z2 - z1) ** 2);
    }

    interface QuasiCycle {
        start: [number, number, number];
        end: [number, number, number];
        indices: number[];
    }

    function findQuasiCycles(points: [number, number, number][], epsilon: number): QuasiCycle[] {
        const quasiCycles: QuasiCycle[] = [];

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

    function findMaximumNonIntersectingArrays(quasiCycles: QuasiCycle[]): QuasiCycle[] {
        const result: QuasiCycle[] = [];
        const usedIndices: Set<number> = new Set();

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
        const points: [number, number, number][] = [
            [0, 0, 0], [1, 0, 0], [2, 0, 0], [3, 0, 0], [1, 0, 0], // Single quasi-cycle
            [5, 5, 0], [6, 5, 0], [7, 5, 0], [8, 5, 0], [9, 5, 0], [10, 5, 0], [5.1, 5.1, 0], // Multiple quasi-cycles
            [10, 10, 0], [11, 10, 0], [12, 10, 0], [13, 10, 0], [14, 10, 0], [10.1, 10.1, 0], // Intersecting quasi-cycles
            [15, 15, 0], [16, 15, 0], [17, 15, 0], [18, 15, 0], [19, 15, 0], [15.1, 15.1, 0], // Negative case (neighbors within epsilon)
        ];
        const epsilon = 0.5;
        const quasiCycles = findQuasiCycles(points, epsilon);
        let result = findMaximumNonIntersectingArrays(quasiCycles)
        console.log(result);

        const inputQuasiCycles: QuasiCycle[] = [
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
    async function createNewCentersGraph() {
        var activeSheetData;
        let graphName = document.getElementById('graphId').innerText;
        await Excel.run(async (context) => {
            var sourceRange = context.workbook.tables.getItem("Centers"+graphName).getDataBodyRange().load("values, rowCount, columnCount");
            activeSheetData = context.workbook.worksheets.getActiveWorksheet().load("name");
            return context.sync().then(function () {
                //Unique graph related code, use it to name every object related to one particular graph
                const id = "_" + window.crypto.randomUUID().substring(0, 5);
                const values = sourceRange.values;
                const _2dValues = convert3DTo2D(values, id);
                //Creating new sheet for graph
                const graphSheet = context.workbook.worksheets.add("Graph" + id);

                // Create a new table for angles values
                const angleTable = graphSheet.tables.add("O1:Q1", true);
                angleTable.name = "Angles" + id; // Set the table name

                angleTable.getHeaderRowRange().values = [["X", "Y", "Z"]]; // Set the header row
                angleTable.rows.add(null, [[255, 1, 50]]); // Set the data rows

                // Create a new table for angles in radians
                const angleRadTable = graphSheet.tables.add("S1:U1", true);
                angleRadTable.name = "AnglesRad" + id; // Set the table name

                angleRadTable.getHeaderRowRange().values = [["X", "Y", "Z"]]; // Set the header row
                angleRadTable.rows.add(null, [[
                    `=(6.28/360)*${angleTable.name}[X]`,
                    `=(6.28/360)*${angleTable.name}[Y]`,
                    `=(6.28/360)*${angleTable.name}[Z]`
                ]]); // Set the data rows
                // Create a new table for coefficients values
                const coeff = [[
                    `=-COS(${angleRadTable.name}[Y])*COS(${angleRadTable.name}[Z])`,
                    `=COS(${angleRadTable.name}[Z])*-SIN(${angleRadTable.name}[Y])*-SIN(${angleRadTable.name}[X])+SIN(${angleRadTable.name}[Z])*COS(${angleRadTable.name}[X])`,
                    `=SIN(${angleRadTable.name}[Z])*COS(${angleRadTable.name}[Y])`,
                    `=-SIN(${angleRadTable.name}[Z])*-SIN(${angleRadTable.name}[Y])*-SIN(${angleRadTable.name}[X])+COS(${angleRadTable.name}[Z])*COS(${angleRadTable.name}[X])`,
                    `=-SIN(${angleRadTable.name}[Y])`,
                    `=COS(${angleRadTable.name}[Y])*-SIN(${angleRadTable.name}[X])`,
                ]];
                const coeffTable = graphSheet.tables.add("H1:M1", true);
                coeffTable.name = "coefficients" + id; // Set the table name

                coeffTable.getHeaderRowRange().values = [["p1", "p2", "q1", "q2", "r1", "r2"]]; // Set the header row
                coeffTable.rows.add(null, coeff); // Set the data rows

                // Create a new table for source values
                const sourceTable = graphSheet.tables.add("A1:C1", true);
                sourceTable.name = "SourceData" + id; // Set the table name

                sourceTable.getHeaderRowRange().values = [["X", "Y", "Z"]]; // Set the header row
                sourceTable.rows.add(null, values); // Set the data rows


                // Create a new table for converted values
                const _2dTable = graphSheet.tables.add("E1:F1", true);
                _2dTable.name = "GraphData" + id; // Set the table name

                _2dTable.getHeaderRowRange().values = [["X", "Y"]]; // Set the header row
                _2dTable.rows.add(null, _2dValues); // Set the data rows

                const chart = context.workbook.worksheets.getActiveWorksheet().charts.add(
                    "XYScatterSmooth",//XYScatterSmoothNoMarkers or XYScatterSmooth
                    _2dTable.getRange(),//Range of table generated from source points
                    Excel.ChartSeriesBy.columns
                );
                //Turn off default elements
                chart.axes.valueAxis.majorGridlines.visible = false;
                chart.axes.categoryAxis.majorGridlines.visible = false;
                chart.axes.valueAxis.visible = false;
                chart.axes.categoryAxis.visible = false;
                // Set chart title
                chart.title.text = "Движение центров";
                chart.name = id


                return context.sync()

            }).then(context.sync);

        }).catch(errorHandler);
        showNotification("Операция завершена", "Успешно построен график на листе " + activeSheetData.name);
    }
   
    function convert3DTo2D(values, id) {
        const convertedValues = [];
        for (const [x, y, z] of values) {
            const x2d = "=SourceData" + id + "[@X]*coefficients" + id + "[p1]+coefficients" + id + "[q1]*SourceData" + id + "[@Y]+coefficients" + id + "[r2]*SourceData" + id + "[@Z]";
            const y2d = "=SourceData" + id + "[@X]*coefficients" + id + "[p2]+coefficients" + id + "[q2]*SourceData" + id + "[@Y]+coefficients" + id + "[r2]*SourceData" + id + "[@Z]";
            convertedValues.push([x2d, y2d]);
        }
        return convertedValues
    }

})();




