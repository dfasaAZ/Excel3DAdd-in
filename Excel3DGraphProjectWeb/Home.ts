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
            // Add a click event handler for the button.
            $('#highlight-button').click(createNewGraph);

            $('#sliderX').on('change', function () {
                rotate('X');
            });
            $('#sliderY').on('change', function () {
                rotate('Y');
            });
            $('#sliderZ').on('change', function () {
                rotate('Z');
            });
        });
    };

    Office.onReady(async () => {

        await Excel.run(async (context) => {
        }).catch(errorHandler);
     
    });
    /**
     * 
     * @param axis [x,y,z]
     */
    function rotate(axis: string): void {
        const graphId = document.getElementById("graphName").textContent;
        //Have to cast types because of typescript limitations
        const value = (<HTMLInputElement>document.getElementById("slider" + axis)).value;

        Excel.run(function (context) {

            // Create a new table on the active sheet
            const sheet = context.workbook.worksheets.getItem("Graph" + graphId);
            const table = sheet.tables.getItem("Angles" + graphId);
            switch (axis) {
                case 'X': table.columns.getItem("X").getDataBodyRange().values = [[parseFloat(value)]];
                    break;
                case 'Y': table.columns.getItem("Y").getDataBodyRange().values = [[parseFloat(value)]];
                    break;
                case 'Z': table.columns.getItem("Z").getDataBodyRange().values = [[parseFloat(value)]];
                    break;
                default: throw new Error("Something bad happened\nNo such axis");
            }

            // Sync the changes to Excel
            return context.sync();
        })
            .catch(errorHandler);
        
    }
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
        //Updating sliders
        const sliderXElement = document.getElementById("sliderX") as HTMLInputElement;
        const sliderYElement = document.getElementById("sliderY") as HTMLInputElement;
        const sliderZElement = document.getElementById("sliderZ") as HTMLInputElement;
        sliderXElement.value = angles[0][0];
        sliderYElement.value = angles[0][1];
        sliderZElement.value = angles[0][2];
    }
    
    function createNewGraph() {
        Excel.run(async (context) => {
            var sourceRange = context.workbook.getSelectedRange().load("values, rowCount, columnCount");
            var activeSheetData = context.workbook.worksheets.getActiveWorksheet().load("name");
            return context.sync().then(function () {
                //Unique graph related code, use it to name every object related to one particular graph
                const id = "_"+window.crypto.randomUUID().substring(0,5);
                const values = sourceRange.values;
                const _2dValues = convert3DTo2D(values,id);
                //Creating new sheet for graph
                const graphSheet = context.workbook.worksheets.add("Graph" + id);

                // Create a new table for angles values
                const angleTable = graphSheet.tables.add("O1:Q1", true);
                angleTable.name = "Angles" + id; // Set the table name

                angleTable.getHeaderRowRange().values = [["X", "Y", "Z"]]; // Set the header row
                angleTable.rows.add(null, [[255,1,50]]); // Set the data rows

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
                coeffTable.name = "coefficients"+id; // Set the table name

                coeffTable.getHeaderRowRange().values = [["p1", "p2", "q1", "q2", "r1", "r2"]]; // Set the header row
                coeffTable.rows.add(null, coeff); // Set the data rows

                // Create a new table for source values
                const sourceTable = graphSheet.tables.add("A1:C1", true);
                sourceTable.name = "SourceData"+id; // Set the table name

                sourceTable.getHeaderRowRange().values = [["X", "Y","Z"]]; // Set the header row
                sourceTable.rows.add(null, values); // Set the data rows


                // Create a new table for converted values
                const _2dTable = graphSheet.tables.add("E1:F1", true); 
                _2dTable.name = "GraphData"+id; // Set the table name
               
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
                chart.name = id

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
    function convert3DTo2D(values,id) {
        const convertedValues = [];
        for (const [x, y, z] of values) {
            const x2d = "=SourceData" + id + "[@X]*coefficients" + id + "[p1]+coefficients" + id + "[q1]*SourceData" + id + "[@Y]+coefficients" + id + "[r2]*SourceData" + id +"[@Z]";
            const y2d = "=SourceData" + id + "[@X]*coefficients" + id + "[p2]+coefficients" + id + "[q2]*SourceData" + id + "[@Y]+coefficients" + id + "[r2]*SourceData" + id +"[@Z]";
            convertedValues.push([x2d, y2d]);
        }
        return convertedValues
    }
   
})();




