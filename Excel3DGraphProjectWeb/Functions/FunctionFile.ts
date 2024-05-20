// The initialize function must be run each time a new page is loaded.

Office.onReady(() => {
    // If needed, Office.js is ready to be called
});
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();
function buildGraph(event) {

    Excel.run(async (context) => {
        var sourceRange = context.workbook.getSelectedRange().load("values, rowCount, columnCount");
        var activeSheetData = context.workbook.worksheets.getActiveWorksheet().load("name");
        return context.sync().then(function () {
            const values = sourceRange.values;
            const _2dValues = convert3DTo2D(values);
            //Creating new sheet for graph
            const graphSheet = context.workbook.worksheets.add("Graph");

            const table = graphSheet.tables.add("A1:B1", true); // Create a new table with headers
            table.name = "GraphData"; // Set the table name

            table.getHeaderRowRange().values = [["X", "Y"]]; // Set the header row
            table.rows.add(null, _2dValues); // Set the data rows

            const chart = context.workbook.worksheets.getActiveWorksheet().charts.add(
                "XYScatterSmooth",//XYScatterSmoothNoMarkers or XYScatterSmooth
                table.getRange(),//Range of table generated from source points
                "Auto",
            );
            //Turn off default elements
            chart.axes.valueAxis.majorGridlines.visible = false;
            chart.axes.categoryAxis.majorGridlines.visible = false;
            chart.axes.valueAxis.visible = false;
            chart.axes.categoryAxis.visible = false;

            // Set chart title
            chart.title.text = "3D Chart";

            showNotification("Operation complete", "Succesfully built chart at " + activeSheetData.name);
            return context.sync()

        }).then(context.sync);

    }).catch(errorHandler);
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
function convert3DTo2D(values) {
    const convertedValues = [];
    const p1 = -0.35;
    const p2 = -0.35;
    const q1 = 1;
    const q2 = 0;
    const r2 = 1;
    for (const [x, y, z] of values) {
        const x2d = p1 * x + q1 * y + r2 * z;
        const y2d = p2 * x + q2 * y + r2 * z;
        convertedValues.push([x2d, y2d]);
    }
    return convertedValues
}