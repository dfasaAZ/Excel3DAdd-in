// The initialize function must be run each time a new page is loaded.

Office.onReady(() => {
    // If needed, Office.js is ready to be called
});
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();
async function createNewGraph() {
    var activeSheetData;
    await Excel.run(async (context) => {
        var sourceRange = context.workbook.getSelectedRange().load("values, rowCount, columnCount");
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
            chart.onActivated.add(handleSelectionChanged);
            // Set chart title
            chart.title.text = "3D Chart";
            chart.name = id


            return context.sync()

        }).then(context.sync);

    });
    handleSelectionChanged(null);
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
    //messageBanner.showBanner();
    //messageBanner.toggleExpansion();
}
/**
  * Converts array of points [x,y,z] to it's 2d visualization
  * 
  * @param values initial array
  * @returns array of [x,y] coordinates
  */
function convert3DTo2D(values, id) {
    const convertedValues = [];
    for (const [x, y, z] of values) {
        const x2d = "=SourceData" + id + "[@X]*coefficients" + id + "[p1]+coefficients" + id + "[q1]*SourceData" + id + "[@Y]+coefficients" + id + "[r2]*SourceData" + id + "[@Z]";
        const y2d = "=SourceData" + id + "[@X]*coefficients" + id + "[p2]+coefficients" + id + "[q2]*SourceData" + id + "[@Y]+coefficients" + id + "[r2]*SourceData" + id + "[@Z]";
        convertedValues.push([x2d, y2d]);
    }
    return convertedValues
}