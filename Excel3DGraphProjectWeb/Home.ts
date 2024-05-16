﻿(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        (window as any).Promise = OfficeExtension.Promise;
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

            $("#template-description").text("This sample highlights the highest value from the cells you have selected in the spreadsheet.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the largest number.");
                
            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(hightlightHighestValue);
        });
    };
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
            table.rows.add(null, spiralData); // Set the data rows


            // Sync the changes to Excel
            return context.sync();
        })
            .catch(errorHandler);
    }
    function convert3Dto2D(range) {
        const p1 = -0.35;
        const p2 = -0.35;
        const q1 = 1;
        const q2 = 0;
        const r2 = 1;

        return Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const values = range.values;
            const convertedValues = [];

            for (const [x, y, z] of values) {
                const x2d = p1 * x + q1 * y + r2 * z;
                const y2d = p2 * x + q2 * y + r2 * z;
                convertedValues.push([x2d, y2d]);
            }

            await context.sync();
            return convertedValues;
        }).catch(errorHandler);
    }
  

    function hightlightHighestValue() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the selected range and load its properties
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // Run the queued-up command, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // Find the cell to highlight
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

                    // Highlight the cell
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
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }

    // Helper function for treating errors
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
})();
