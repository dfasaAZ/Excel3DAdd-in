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

            $("#template-description").text("Select range with three columns and press \"Build\" button");
            $('#button-text').text("Build!");
            $('#button-desc').text("Build new graph");
                
            
            // Add a click event handler for the highlight button.
            $('#highlight-button').click(createNewGraph);
        });
    };

    Office.onReady(async () => {

        await Excel.run(async (context) => {
        }).catch(errorHandler);
     
    });
    function loadSampleData() {
        // Define spiral parameters
        const numPoints = 100; // Number of points in the spiral
        const radius = 10; // Initial radius of the spiral
        const height = 20; // Height of the spiral
        const angle = Math.PI * 2 / numPoints; // Angular increment

        // Create an array to store the spiral coordinates
        const spiralData = [];

        // Generate the spiral coordinates
        for (let i = 0; i < numPoints; i++) {
            const r = radius + i * 0.1; // Increasing radius
            const x = r * Math.cos(i * angle);
            const y = r * Math.sin(i * angle);
            const z = i * height / (numPoints - 1);
            spiralData.push([x, y, z]);
        }

        // Run the Excel operations
        Excel.run(function (context) {
           
            // Create a new table on the active sheet
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const table = sheet.tables.add("A1:C1", true); // Create a new table with headers
            table.name = "SpiralData"; // Set the table name

            // Populate the table with spiral data
            table.getHeaderRowRange().values = [["X", "Y", "Z"]]; // Set the header row
            table.rows.add(null, spiralData); // Set the data rows (SpiralData)


            // Sync the changes to Excel
            return context.sync();
        })
            .catch(errorHandler);
    }

    async function handleSelectionChanged(event) {
        await Excel.run(async (context) => {
            let selectedGraph;
            selectedGraph = context.workbook.getActiveChart().load("name");
            await context.sync().then(function processSelectedGraphs() {
                    $('#graphName').text(selectedGraph.name);
                });

            
            
        });
    }
    
    function createNewGraph() {
        
        Excel.run(async (context) => {
            var sourceRange = context.workbook.getSelectedRange().load("values, rowCount, columnCount");
            var activeSheetData = context.workbook.worksheets.getActiveWorksheet().load("name");
            return context.sync().then(function () {
                const values = sourceRange.values;
                const _2dValues = convert3DTo2D(values);
                //Creating new sheet for graph
                const graphSheet = context.workbook.worksheets.add("Graph");

                // Create a new table for coefficients values
                const p1 = -0.35;
                const p2 = -0.35;
                const q1 = 1;
                const q2 = 0;
                const r2 = 1;
                const coeff = [[p1, p2, q1, q2, r2]];
                const coeffTable = graphSheet.tables.add("H1:L1", true);
                coeffTable.name = "coefficients"; // Set the table name

                coeffTable.getHeaderRowRange().values = [["p1", "p2", "q1", "q2", "r2"]]; // Set the header row
                coeffTable.rows.add(null, coeff); // Set the data rows

                // Create a new table for source values
                const sourceTable = graphSheet.tables.add("A1:C1", true);
                sourceTable.name = "SourceData"; // Set the table name

                sourceTable.getHeaderRowRange().values = [["X", "Y","Z"]]; // Set the header row
                sourceTable.rows.add(null, values); // Set the data rows


                // Create a new table for converted values
                const _2dTable = graphSheet.tables.add("E1:F1", true); 
                _2dTable.name = "GraphData"; // Set the table name
               
                _2dTable.getHeaderRowRange().values = [["X", "Y"]]; // Set the header row
                _2dTable.rows.add(null, _2dValues); // Set the data rows

                const chart = context.workbook.worksheets.getActiveWorksheet().charts.add(
                    "XYScatterSmooth",//XYScatterSmoothNoMarkers or XYScatterSmooth
                    _2dTable.getRange(),//Range of table generated from source points
                    "Auto",
                );
                //Turn off default elements
                chart.axes.valueAxis.majorGridlines.visible = false;
                chart.axes.categoryAxis.majorGridlines.visible = false;
                chart.axes.valueAxis.visible = false;
                chart.axes.categoryAxis.visible = false;
                chart.onActivated.add(handleSelectionChanged);
                // Set chart title
                chart.title.text = "3D Chart";

                showNotification("Operation complete", "Succesfully built chart at " + activeSheetData.name);
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
   * Converts array of points [x,y,z] to it's 2d visualization
   * 
   * @param values initial array
   * @returns array of [x,y] coordinates
   */
    function convert3DTo2D(values) {
        const convertedValues = [];
        for (const [x, y, z] of values) {
            const x2d = "=SourceData[@X]*coefficients[p1]+coefficients[q1]*SourceData[@Y]+coefficients[r2]*SourceData[@Z]";
            const y2d = "=SourceData[@X]*coefficients[p2]+coefficients[q2]*SourceData[@Y]+coefficients[r2]*SourceData[@Z]";
            convertedValues.push([x2d, y2d]);
        }
        return convertedValues
    }

})();


