/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
    See full license at the bottom of this file. */

/// <reference path="../App.js" />

(function () {
    "use strict";

    var range;
    var targetaddr1;
    var targetaddr2;
    var numofadditionalcols = 0;
    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            // If not using Excel 2016, return
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                app.showNotification("Need Office 2016 or greater", "Sorry, this add-in only works with newer versions of Excel.");
                return;
            }
            $('#runAzureML').click(runAzureML);
        });
    };

    function runAzureML() {
        $('.disable-while-sending').prop('disabled', true);
        var ret = RetrieveDataInRange();
    }

    function RetrieveDataInRange() {
        numofadditionalcols = 0;
        Excel.run(function (ctx) {
            range = ctx.workbook.getSelectedRange().load("values");

            range.load('columnCount');
            range.load('rowCount');
            range.load('rowIndex');

            return ctx.sync().then(function () {

                Excel.run(function (ctx) {
                    var resultRange = ctx.workbook.getSelectedRange().getOffsetRange(0, range.columnCount);
                    resultRange.load('address');
                    var cell = ctx.workbook.getSelectedRange().getOffsetRange(0, range.columnCount).getCell(0, 0);
                    cell.load('address');

                    var rangeshiftby1 = ctx.workbook.getSelectedRange().getOffsetRange(0, 1).getLastCell();
                    rangeshiftby1.load('address');

                    var rangeshiftby2 = ctx.workbook.getSelectedRange().getOffsetRange(0, 2).getLastCell();
                    rangeshiftby2.load('address');

                    return ctx.sync().then(function () {
                        //var startaddr = resultRange.address;
                        var startaddr = cell.address.substring(cell.address.indexOf("!") + 1, cell.address.length);
                        var endaddr1 = rangeshiftby1.address.substring(rangeshiftby1.address.indexOf("!") + 1, rangeshiftby1.address.length);
                        var endaddr2 = rangeshiftby2.address.substring(rangeshiftby2.address.indexOf("!") + 1, rangeshiftby2.address.length);

                        targetaddr1 = startaddr + ":" + endaddr1;
                        targetaddr2 = startaddr + ":" + endaddr2;

                        var test = range.columnCount;
                        var colLength = range.columnCount;

                        var columnnames = [];
                        var rowvalues = [];
                        for (var i = 0; i < range.values.length; i++) {
                            var rowvalue = [];
                            for (var j = 0; j < range.values[i].length; j++) {
                                if (i == 0) {
                                    columnnames.push(String(range.values[i][j]));
                                }
                                else {
                                    rowvalue[j] = String(range.values[i][j]);
                                }
                            }
                            if (i > 0) rowvalues.push(rowvalue);
                        }
                        var MLColumn = "";
                        MLColumn = JSON.stringify(columnnames);
                        var MLInput = "";

                        var MLRows = JSON.stringify(rowvalues);
                        var urltext = $("#urltext").val();
                        var apikey = $("#apikey").val();
                        var ret = sendRequest("AzureML/RunAzureML", urltext, apikey, MLColumn, MLRows);

                    });
                }).catch(function (error) {
                    //console.log("Error: " + error);
                    if (error instanceof OfficeExtension.Error) {
                        //console.log("Debug info: " + JSON.stringify(error.debugInfo));
                    }
                });



            });
        }).catch(function (error) {
            //console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                //console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    };

    // Helper method to Send Data to Web Service (which in turn talks to Azure ML Web Service Endpoint)
    function sendRequest(method, urltext, apikey, mlcolumn, mlrows) {
        $.ajax({
            url: '../../api/' + method,
            type: 'GET',
            data: { URLText: urltext, APIKey: apikey, MLColumns: mlcolumn, MLRows: mlrows },
            contentType: 'application/json;charset=utf-8'
        }).done(function (data) {
            //window.location.href = JSON.stringify(data);
            //app.showNotification("Success1", JSON.stringify(data));
            if (data) {
                loadDataAndCreateChart(data);
            }
            //app.showNotification("Success", data );
            //console.log(JSON.stringify(data));
        }).fail(function (status) {
            app.showNotification('Error', JSON.stringify(status));
        }).always(function () {
            $('.disable-while-sending').prop('disabled', false);
        });
    }
    // Load data into the worksheet and then create a chart
    function loadDataAndCreateChart(returneddata) {

        var jsonobj = JSON.parse(returneddata);
        if (jsonobj.Results) {
            results = new Array();
            var col = new Array();
            //Check to see if we are getting Score Lables and Score Probablities- Sometimes only Score Labels are returned.
            for (var j = 0; j < jsonobj.Results.output1.value.ColumnNames.length; j++) {
                if ((jsonobj.Results.output1.value.ColumnNames[j] == "Scored Labels") || (jsonobj.Results.output1.value.ColumnNames[j] == "Scored Probabilities"))
                {
                    col.push(jsonobj.Results.output1.value.ColumnNames[j]);
                    numofadditionalcols++;
                }
            }

            //for (var j = jsonobj.Results.output1.value.ColumnNames.length - 1; j < jsonobj.Results.output1.value.ColumnNames.length; j++) {
            //    col.push(jsonobj.Results.output1.value.ColumnNames[j]);
            //}
            //results.push(jsonobj.Results.output1.value.ColumnNames);
            results.push(col);

            for (var i = 0; i < jsonobj.Results.output1.value.Values.length; i++) {
                var row = new Array();
                for (var k = jsonobj.Results.output1.value.Values[0].length - numofadditionalcols; k < jsonobj.Results.output1.value.Values[0].length; k++) {
                    row.push(jsonobj.Results.output1.value.Values[i][k]);
                }
                results.push(row);
            }
            // Run a batch operation against the Excel object model
            Excel.run(function (ctx) {
                // Create a proxy object for the active worksheet
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();

                //Queue commands to set the report title in the worksheet
                //sheet.getRange("A10").values = "Azure ML Results";
                //sheet.getRange("A10").format.font.name = "Calibri";
                //sheet.getRange("A10").format.font.size = 20;


                //Queue a command to write the sample data to the specified range
                //in the worksheet and bold the header row
                //ctx.workbook.tables.add('Sheet1!A11:Q13', true);
                //var range = sheet.getRange("A11:Q13");
                //range.values = results;
                //sheet.getRange("A11:O11").format.font.bold = true;

                // ctx.workbook.tables.add('Sheet1!P1:Q3', true);

                ////var newrange = ctx.workbook.getSelectedRange().getOffsetRange(2, 2);
                //var oldrange = ctx.workbook.getSelectedRange().load();
                //var newrange = ctx.workbook.getSelectedRange().getOffsetRange(0, 2);
                //var test = newrange.getCell(0, 0);
                ////newrange.columnIndex = oldrange.columnCount;
                ////newrange.rowIndex = oldrange.rowIndex;
                ////newrange.rowCount = oldrange.rowCount;

                //newrange.columnIndex = 15;
                //newrange.rowIndex = 0;
                //newrange.rowCount = 3;
                //newrange.columnCount = 2
                //range.values = results;

                //ctx.workbook.getSelectedRange().getOffsetRange(0, 2).values = 7;



                //WORKING
                var range;
                if (numofadditionalcols == 1) range = sheet.getRange(targetaddr1);
                else range = sheet.getRange(targetaddr2);
                range.values = results;

                //sheet.getRange("A11:O11").format.font.bold = true;




                //Queue a command to add a new chart
                var chart = sheet.charts.add("ColumnClustered", range, "auto");

                //Queue commands to set the properties and format the chart
                chart.setPosition("R16", "AI36");
                chart.title.text = "Predicted Risk Score";
                chart.legend.position = "right"
                chart.legend.format.fill.setSolidColor("white");
                chart.dataLabels.format.font.size = 15;
                chart.dataLabels.format.font.color = "black";
                var points = chart.series.getItemAt(0).points;
                points.getItemAt(0).format.fill.setSolidColor("pink");
                points.getItemAt(1).format.fill.setSolidColor('indigo');

                //Run the queued-up commands, and return a promise to indicate task completion
                return ctx.sync();
            })
              .then(function () {
                  app.showNotification("Success");
                  console.log("Success!");
              })
            .catch(function (error) {
                // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
                app.showNotification("Error: " + error);
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }

    }
})();
/*
AzureMLExcelAddInStarterKit, https://github.com/OfficeDev/Excel-Add-in-JS-QuarterlySalesReport

Copyright (c) Microsoft Corporation

All rights reserved.

MIT License:

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated
documentation files (the "Software"), to deal in the Software without restriction, including without limitation
the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and
to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial
portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT
LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN
NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/