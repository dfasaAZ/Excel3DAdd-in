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
        });
    };

    Office.onReady(async () => {

        await Excel.run(async (context) => {
        }).catch(errorHandler);

    });
    function loadSampleData() {


        //TODO: Проверка алгоритмов квазицикла, убрать потом
        quasiTest();

       

        // Run the Excel operations
        Excel.run(function (context) {

 

            // Sync the changes to Excel
            return context.sync();
        })
            .catch(errorHandler);
    }

    async function handleSelectionChanged(event) {
        let angles;
        await Excel.run(async (context) => {
            let selectedGraph;
            let graphName;
            let graphAngles;
            selectedGraph = context.workbook.getActiveChart().load("name");
            await context.sync().then(function processSelectedGraphs() {
                $('#graphName').text(selectedGraph.name);
                graphName = selectedGraph.name;
            });
            // Loading graph related properties
            graphAngles = context.workbook.tables.getItem("Angles" + graphName).rows.getItemAt(0).load("values");
            await context.sync().then(function processSelectedGraphs() {
                angles = graphAngles.values;

            });
        });
    }

    function processQuasiCycles() {
        let graphName;// получить с html, загружать как и в home
        let graphAngles;
        Excel.run(async (context) => {
            // Loading graph related properties
            graphAngles = context.workbook.tables.getItem("SourceData" + graphName).rows.load("items");
            return context.sync().then(function () {

                //Здесь пройтись по строкам и запихать их в квазициклы

                return context.sync()

            }).then(context.sync);

        }).catch(errorHandler);
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
})();




