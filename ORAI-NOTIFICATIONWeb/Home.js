(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
         
            $('#btnSend').click(sendSMS);
            loadSampleData();
        });
    };


    function displayErrorMessage(errormsg) {
        $("#lblErrorMsg").text(errormsg);
    }

    function validate() {

        var result = true;


        if ($('#ddlType').val() == "ORAI-T") {

            if ($('#txtFrom').val() == "") {
                displayErrorMessage("Please Enter From");
                result = false;
                return result;
            }
            else if ($('#txtSID').val() == "") {
                displayErrorMessage("Please Enter SID");
                result = false;
                return result;
            }

            if ($('#txtAuthToken').val() == "") {
                displayErrorMessage("Please Enter AuthToken");
                result = false;
                return result;
            }
            //else if ($('#txtTemplate').val() == "") {
            //    displayErrorMessage("Please Enter Template");
            //    result = false;
            //    return result;
            //}

            if ($('#ddlTemplateType').val() == "Media") {

                if ($('#txtMediaUrl').val() == "") {
                    displayErrorMessage("Please Enter MediaUrl");
                    result = false;
                    return result;
                }
            }
            if ($('#ddlTemplateName').val() == "Select Template Name") {
                displayErrorMessage("Please Select Template Name");
                result = false;
                return result;
            }
        }
        else if ($('#ddlType').val() == "ORAI-K") {

            if ($('#txtFrom').val() == "") {
                displayErrorMessage("Please Enter From");
                result = false;
                return result;
            }
            else if ($('#txtSID').val() == "") {
                displayErrorMessage("Please Enter SID");
                result = false;
                return result;
            }


            if ($('#txtAPIKEY').val() == "") {
                result = false;
                displayErrorMessage("Please Enter API Key");
                return result;
            }

            var files = $("#fileUploadInput").get(0).files;

            if ($('#ddlTemplateType').val() == "Media") {
                if (files.length == 0) {
                    result = false;
                    displayErrorMessage("Please Select File");
                    return result;
                }
            }
            if ($('#ddlTemplateName').val() == "Select Template Name") {
                displayErrorMessage("Please Select Template Name");
                result = false;
                return result;
            }

        }
        else if ($('#ddlType').val() == "ORAI-360") {

            if ($('#txtT360APIKEY').val() == "") {
                result = false;
                displayErrorMessage("Please Enter API Key");
                return result;
            }

            if ($('#txtT360Namespace').val() == "") {
                result = false;
                displayErrorMessage("Please Enter Namespace");
                return result;
            }

            if ($('#ddlTemplateType').val() == "Media") {

                if ($('#txtMediaUrl').val() == "") {
                    displayErrorMessage("Please Enter MediaUrl");
                    result = false;
                    return result;
                }
            }

            if ($('#ddlTemplateName').val() == "Select Template Name") {
                displayErrorMessage("Please Select Template Name");
                result = false;
                return result;
            }

        }

        return result;
    }

    function AddLocalStorage() {
        
        var Type = $('#ddlType').val();
        var TemplateType = $('#ddlTemplateType').val();
        var From = $('#txtFrom').val();
        var SID = $('#txtSID').val();
        var AuthToken = $('#txtAuthToken').val();
        var APIKEY = $('#txtAPIKEY').val();
        var TemplateName = $("#ddlTemplateName option:selected").text();
        var Template = $('#txtTemplate').val();
        var MediaUrl = $('#txtMediaUrl').val();

        var T360APIKEY = $('#txtT360APIKEY').val();
        var T360Nampespace = $('#txtT360Namespace').val();
        

        var Params = ""

        if (Type == "ORAI-T") {
            debugger
            if (TemplateType == "Text") {

                localStorage.setItem("TType", Type);
                localStorage.setItem("TTemplateType", TemplateType);
                localStorage.setItem("TFrom", From);
                localStorage.setItem("TSID", SID);
                localStorage.setItem("TAuthToken", AuthToken);
                localStorage.setItem("TTemplateName", TemplateName);
                localStorage.setItem("TTemplate", Template);

            }
            else {
               
                localStorage.setItem("TMType", Type);
                localStorage.setItem("TMTemplateType", TemplateType);
                localStorage.setItem("TMFrom", From);
                localStorage.setItem("TMSID", SID);
                localStorage.setItem("TMAuthToken", AuthToken);
                localStorage.setItem("TMTemplateName", TemplateName);
                localStorage.setItem("TMTemplate", Template);
                localStorage.setItem("TMMediaUrl", MediaUrl);

            }

        }
        else if (Type == "ORAI-K") {

            if (TemplateType == "Text") {

                localStorage.setItem("KType", Type);
                localStorage.setItem("KTemplateType", TemplateType);
                localStorage.setItem("KFrom", From);
                localStorage.setItem("KSID", SID);
                localStorage.setItem("KAPIKEY", APIKEY);
                localStorage.setItem("KTemplateName", TemplateName);
                localStorage.setItem("KTemplate", Template);
                localStorage.setItem("KParams", Params);

            }
            else {

                localStorage.setItem("KMType", Type);
                localStorage.setItem("KMTemplateType", TemplateType);
                localStorage.setItem("KMFrom", From);
                localStorage.setItem("KMSID", SID);
                localStorage.setItem("KMAPIKEY", APIKEY);
                localStorage.setItem("KMTemplateName", TemplateName);
                localStorage.setItem("KMTemplate", Template);
                localStorage.setItem("KMFile", "");
                localStorage.setItem("KMParams", Params);
            }
        }
        else if (Type == "ORAI-360") {
            
            if (TemplateType == "Text") {
                
                localStorage.setItem("T360Type", Type);
                localStorage.setItem("T360TemplateType", TemplateType);
                localStorage.setItem("T360APIKEY", T360APIKEY);
                localStorage.setItem("T360Namespace", T360Nampespace);
                localStorage.setItem("T360TemplateName", TemplateName);
                localStorage.setItem("T360Template", Template);
                localStorage.setItem("T360Params", Params);
            }
            else {
                localStorage.setItem("TM360Type", Type);
                localStorage.setItem("TM360TemplateType", TemplateType);
                localStorage.setItem("TM360APIKEY", T360APIKEY);
                localStorage.setItem("TM360Namespace", T360Nampespace);
                localStorage.setItem("TM360TemplateName", TemplateName);
                localStorage.setItem("TM360Template", Template);
                localStorage.setItem("TM360Params", Params);
                localStorage.setItem("TM360MediaUrl", MediaUrl);    
            }
        }



    }


    function sendSMS() {
        
        var From = $('#txtFrom').val();
        var SID = $('#txtSID').val();
        var AuthToken = $('#txtAuthToken').val();
        var Template = "";
        
        if ($('#ddlTemplateStatus').val() == "Not Approved") {
            Template = $("#txtNotApprovedTemplate").val();
        }
        else {
            Template = $('#txtTemplate').val();
        }

        var TemplateStatus = "";
        var Type = $('#ddlType').val();

        if (Type == "ORAI-K") {
            if ($('#ddlTemplateStatus').val() == "Not Approved") {
                TemplateStatus = "Not Approved";
            } else {
                TemplateStatus = "Approved";
            }
        }

        var TemplateType = $('#ddlTemplateType').val();
        var MediaUrl = $('#txtMediaUrl').val();

        var modified_Media_url = "";
        if (Type == "ORAI-T" || Type == "ORAI-360") {
            modified_Media_url = MediaUrl;
        }
        else if (Type == "ORAI-K") {
            modified_Media_url = MediaUrl.replace(/\\/g, "/");
        }
       
        var APIKEY = $('#txtAPIKEY').val();

        var T360APIKEY = $('#txtT360APIKEY').val();
        var T360Namespace = $('#txtT360Namespace').val();

        var TemplateName = $("#ddlTemplateName option:selected").text();
        var Params = "";//$('#txtParams').val();

        var result = validate();
        if (result == true) {

            AddLocalStorage();

            Excel.run(function (context) {


                var selectionRange = context.workbook.getSelectedRange();

                selectionRange.format.fill.clear();
                selectionRange.load("values");
                return context.sync()
                    .then(function () {

                        var rowCount = selectionRange.values.length;
                        if (rowCount > 1) {
                            displayErrorMessage("")
                            var columnCount = selectionRange.values[0].length;
                            let PhoneNumber = "";
                            try {
                                for (var row = 1; row < rowCount; row++) {
                                    for (var column = 0; column < columnCount; column++) {
                                        //  if (selectionRange.values[row][column] > 50)
                                        //{
                                        //selectionRange.getCell(row, column).format.fill.color = "yellow";
                                        PhoneNumber = selectionRange.values[row][0];
                                        //selectionRange.getCell(row, column).format.fill.color = "yellow";
                                        //let PhoneNumber = selectionRange.values[row][column];

                                        if (Type == "ORAI-K" || Type == "ORAI-360") {

                                            if (selectionRange.values[0][column] == "Param1") {
                                                Params = selectionRange.values[row][2];

                                            }
                                            else if (selectionRange.values[0][column] == "Param2") {
                                                Params = Params + "," + selectionRange.values[row][3];
                                            }
                                            else if (selectionRange.values[0][column] == "Param3") {
                                                Params = Params + "," + selectionRange.values[row][4];
                                            }
                                            else if (selectionRange.values[0][column] == "Param4") {
                                                Params = Params + "," + selectionRange.values[row][5];
                                            }
                                            if (selectionRange.values[0][column] == "Param5") {
                                                Params = Params + "," + selectionRange.values[row][6];

                                            }
                                            else if (selectionRange.values[0][column] == "Param6") {
                                                Params = Params + "," + selectionRange.values[row][7];
                                            }
                                            else if (selectionRange.values[0][column] == "Param7") {
                                                Params = Params + "," + selectionRange.values[row][8];
                                            }
                                            else if (selectionRange.values[0][column] == "Param8") {
                                                Params = Params + "," + selectionRange.values[row][9];
                                            }
                                            if (selectionRange.values[0][column] == "Param9") {
                                                Params = Params + "," + selectionRange.values[row][10];

                                            }
                                            else if (selectionRange.values[0][column] == "Param10") {
                                                Params = Params + "," + selectionRange.values[row][11];
                                            }
                                            else if (selectionRange.values[0][column] == "Param11") {
                                                Params = Params + "," + selectionRange.values[row][12];
                                            }
                                            else if (selectionRange.values[0][column] == "Param12") {
                                                Params = Params + "," + selectionRange.values[row][13];
                                            }
                                            if (selectionRange.values[0][column] == "Param13") {
                                                Params = Params + "," + selectionRange.values[row][14];

                                            }
                                            else if (selectionRange.values[0][column] == "Param14") {
                                                Params = Params + "," + selectionRange.values[row][15];
                                            }
                                            else if (selectionRange.values[0][column] == "Param15") {
                                                Params = Params + "," + selectionRange.values[row][16];
                                            }
                                            else if (selectionRange.values[0][column] == "Param16") {
                                                Params = Params + "," + selectionRange.values[row][17];
                                            }
                                        }

                                        else if (Type == "ORAI-T"){

                                            if (selectionRange.values[0][column] == "Param1") {
                                                Template = Template.replace("{{1}}", selectionRange.values[row][2]);
                                            }
                                            else if (selectionRange.values[0][column] == "Param2") {
                                                Template = Template.replace("{{2}}", selectionRange.values[row][3]);
                                            }
                                            else if (selectionRange.values[0][column] == "Param3") {
                                                Template = Template.replace("{{3}}", selectionRange.values[row][4]);
                                            }
                                            else if (selectionRange.values[0][column] == "Param4") {
                                                Template = Template.replace("{{4}}", selectionRange.values[row][5]);
                                            }
                                            else if (selectionRange.values[0][column] == "Param5") {
                                                Template = Template.replace("{{5}}", selectionRange.values[row][6]);
                                            }
                                            else if (selectionRange.values[0][column] == "Param6") {
                                                Template = Template.replace("{{6}}", selectionRange.values[row][7]);
                                            }
                                            else if (selectionRange.values[0][column] == "Param7") {
                                                Template = Template.replace("{{7}}", selectionRange.values[row][8]);
                                            }
                                            else if (selectionRange.values[0][column] == "Param8") {
                                                Template = Template.replace("{{8}}", selectionRange.values[row][9]);
                                            }
                                            if (selectionRange.values[0][column] == "Param9") {
                                                Template = Template.replace("{{9}}", selectionRange.values[row][10]);

                                            }
                                            else if (selectionRange.values[0][column] == "Param10") {
                                                Template = Template.replace("{{10}}", selectionRange.values[row][11]);
                                            }
                                            else if (selectionRange.values[0][column] == "Param11") {
                                                Template = Template.replace("{{11}}", selectionRange.values[row][12]);
                                            }
                                            else if (selectionRange.values[0][column] == "Param12") {
                                                Template = Template.replace("{{12}}", selectionRange.values[row][13]);
                                            }
                                            if (selectionRange.values[0][column] == "Param13") {
                                                Template = Template.replace("{{13}}", selectionRange.values[row][14]);

                                            }
                                            else if (selectionRange.values[0][column] == "Param14") {
                                                Template = Template.replace("{{14}}", selectionRange.values[row][15]);
                                            }
                                            else if (selectionRange.values[0][column] == "Param15") {
                                                Template = Template.replace("{{15}}", selectionRange.values[row][16]);
                                            }
                                            else if (selectionRange.values[0][column] == "Param16") {
                                                Template = Template.replace("{{16}}", selectionRange.values[row][17]);
                                            }
                                        }
                                    }

                                    var files = $("#fileUploadInput").get(0).files;
                                    if (Type == "ORAI-K") {
                                        if (TemplateName == "Media") {
                                            if (files.length != 0) {

                                            }
                                        }

                                    }
                                    //for (var i = 0; i < files.length; i++) {
                                    var fileData = new FormData();
                                    fileData.append("fileUploadInput", files[0]);

                                    fileData.append("PhoneNumber", PhoneNumber);
                                    fileData.append("From", From);
                                    fileData.append("Body", Template);
                                    fileData.append("SID", SID);
                                    fileData.append("AuthToken", AuthToken);
                                    fileData.append("MediaUrl", modified_Media_url);
                                    fileData.append("Type", Type);
                                    fileData.append("APIKEY", APIKEY);
                                    fileData.append("T360APIKEY", T360APIKEY);
                                    fileData.append("T360Namespace", T360Namespace );
                                    fileData.append("Params", Params);
                                    fileData.append("TemplateName", TemplateName);
                                    fileData.append("TemplateType", TemplateType);
                                    fileData.append("TemplateStatus", TemplateStatus);
                                    //  }
                                    var aaa;
                                    $.ajax({
                                        type: "POST",
                                        //url: "https://excelutiapi.azurewebsites.net/LoadExcel/SendAddInNotification",
                                        url: "https://e2ewebservice20190528111726.azurewebsites.net/LoadExcel/SendAddInNotification",
                                        //url: "https://localhost:44351/LoadExcel/SendAddInNotification"
                                        dataType: "json",
                                        contentType: false, // Not to set any content header
                                        processData: false, // Not to process data
                                        data: fileData,
                                        async: false,
                                        success: function (response, status, xhr) {
                                            var selectionRangeZ = context.workbook.getSelectedRange();
                                            //selectionRange.getOffsetRange(0, 0).values = "Phone Number";
                                            //selectionRange.getOffsetRange(row, 0).values = response.d;
                                            //selectionRange.values[row][1] = response.d;
                                            // context.sync();
                                            //selectionRange.values[row][columnCount] = response.d;

                                            var sheet = context.workbook.worksheets.getItem("Sheet1");

                                            //var data = [
                                            //    [response.d],
                                            //];

                                            var row_count = row + 1;
                                            var range = sheet.getRange("B" + row_count);
                                            range.values = response;
                                            range.format.autofitColumns();


                                            if (Type == "ORAI-K") {
                                                if ($('#ddlTemplateStatus').val() == "Not Approved") {
                                                    Template = $("#txtNotApprovedTemplate").val();
                                                }
                                            }
                                            else {
                                                Template = $('#txtTemplate').val();
                                            }
                                            return context.sync();
                                        },
                                        error: function (response, status, error) {
                                            var a = response.d;
                                            selectionRange.values[row][1] = response;
                                        }
                                    });
                                }
                            }
                            catch (err) {
                                var a = err.message;
                            }
                        }
                        else {
                            displayErrorMessage("Please Select Excel Data");
                        }
                    }).then(
                        context.sync
                    );



                //var range = context.workbook.getSelectedRange();
                //range.values = "Hello World";
                //

                //var newSh = context.workbook.worksheets.add();
                //new Sh.name = "New Sheet"
                //newSh.activate();

                //var actSh = context.workbook.worksheets.getItem("Sheet1");
                //var rng = actSh.getRange(K10);
                //rng.values = 'Value from js'

                //var selCell = context.workbook.getActiveCell();
                //selCell.values = 'getActiveCell';
                //selCell.getOffsetRange(1, 1).values = "getoffestrange example"




                // Grab the sheet
                //const sheet = context.workbook.worksheets.getActiveWorksheet();

                //// Grab a range using Get Range
                //let salesRng = sheet.getRange("B2:D5");

                //// load the values property & the text properties
                //salesRng.load(["values", "text", "formulas", "formulasR1C1"]);
                ////context.sync();

                ////// let's grab the values
                //let myValues = salesRng.values;

                //// access the first row
                //console.log(myValues[0]);

                ////// access the first column of the first row
                //console.log(myValues[0][0]);

                //// loop through the entire array
                //// let's start with the rows
                //myValues.forEach(function (row, index) {

                //    // print each row
                //    console.log(row);

                //    // then each column in that row
                //    row.forEach(function (col, index2) {

                //        // print each column
                //        console.log(col);

                //    });
                //});

                //// stringify the values:
                //// the first para is the values
                //// the second para is what you want missing values to be
                //// the third is how many spaces you want
                //console.log(JSON.stringify(salesRng.values, null, 1));
                //console.log(JSON.stringify(salesRng.text, null, 1));
                //console.log(JSON.stringify(salesRng.formulas, null, 1));



                //for (i = 0; i < array.length; i++) {
                //    col = i + 6;
                //    worksheet.getCell('E' + col).value = array[i];
                //}
                //var count = context.workbook.worksheets.items.length;

                return context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error)
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
            })
        }
    }

    function loadSampleData() {
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the active sheet
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // Queue a command to write the sample data to the worksheet
            //sheet.getRange("B3:D5").values = values;
            var selectionRange = ctx.workbook.getSelectedRange();
            selectionRange.getOffsetRange(0, 0).values = "Phone Number";
            selectionRange.getOffsetRange(0, 1).values = "Status";
            selectionRange.getOffsetRange(0, 2).values = "Param1";
            selectionRange.getOffsetRange(0, 3).values = "Param2";
            selectionRange.getOffsetRange(0, 4).values = "Param3";
            //selectionRange.getOffsetRange(0, 5).values = "Param4";
            //selectionRange.getOffsetRange(0, 6).values = "Param5";
            //selectionRange.getOffsetRange(0, 7).values = "Param6";
            //selectionRange.getOffsetRange(0, 8).values = "Param7";
            //selectionRange.getOffsetRange(0, 9).values = "Param8";
            //selectionRange.getOffsetRange(1, 0).values = "919881821002";
            //selectionRange.getOffsetRange(1, 2).values = "A";
            //selectionRange.getOffsetRange(1, 3).values = "B";
            //selectionRange.getOffsetRange(1, 4).values = "C";
            //selectionRange.getOffsetRange(1, 5).values = "D";
            //selectionRange.getOffsetRange(1, 6).values = "E";
            //selectionRange.getOffsetRange(1, 6).values = "F";
            //selectionRange.getOffsetRange(1, 7).values = "H";
            //selectionRange.getOffsetRange(1, 8).values = "G";
            //selectionRange.getOffsetRange(1, 9).values = "L";
            setCommonData("TFrom", "TSID", "TAuthToken", "", "ORAI-T");
            setTextMediaDataUsingTypeAndTemplateType("Text", "ORAI-T");
            //$("#ddlTemplateName option:selected").text(localStorage.getItem("TTemplateName"));
            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(errorHandler);
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
