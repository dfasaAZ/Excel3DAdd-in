var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
(function () {
    "use strict";
    var cellToHighlight;
    var messageBanner;
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        window.Promise = OfficeExtension.Promise;
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
        var numPoints = 100; // Number of points in the spiral
        var radius = 10; // Initial radius of the spiral
        var height = 20; // Height of the spiral
        var angle = Math.PI * 2 / numPoints; // Angular increment
        // Create an array to store the spiral coordinates
        var spiralData = [];
        // Generate the spiral coordinates
        for (var i = 0; i < numPoints; i++) {
            var r = radius + i * 0.1; // Increasing radius
            var x = r * Math.cos(i * angle);
            var y = r * Math.sin(i * angle);
            var z = i * height / (numPoints - 1);
            spiralData.push([x, y, z]);
        }
        // Run the Excel operations
        Excel.run(function (context) {
            // Create a new table on the active sheet
            var sheet = context.workbook.worksheets.getActiveWorksheet();
            var table = sheet.tables.add("A1:C1", true); // Create a new table with headers
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
        var _this = this;
        var p1 = -0.35;
        var p2 = -0.35;
        var q1 = 1;
        var q2 = 0;
        var r2 = 1;
        return Excel.run(function (context) { return __awaiter(_this, void 0, void 0, function () {
            var sheet, values, convertedValues, _i, values_1, _a, x, y, z, x2d, y2d;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        sheet = context.workbook.worksheets.getActiveWorksheet();
                        values = range.values;
                        convertedValues = [];
                        for (_i = 0, values_1 = values; _i < values_1.length; _i++) {
                            _a = values_1[_i], x = _a[0], y = _a[1], z = _a[2];
                            x2d = p1 * x + q1 * y + r2 * z;
                            y2d = p2 * x + q2 * y + r2 * z;
                            convertedValues.push([x2d, y2d]);
                        }
                        return [4 /*yield*/, context.sync()];
                    case 1:
                        _b.sent();
                        return [2 /*return*/, convertedValues];
                }
            });
        }); }).catch(errorHandler);
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
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                showNotification('The selected text is:', '"' + result.value + '"');
            }
            else {
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
//# sourceMappingURL=Home.js.map