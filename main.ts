// Author: Aliafriend

function AddForm(){
    var form = HtmlService.createHtmlOutputFromFile("SheetsTools").setTitle("Sheets Discord Tools");
    SpreadsheetApp.getUi().showSidebar(form);
}

function onOpen(){
    let menu = SpreadsheetApp.getUi().createAddonMenu();
    menu.addItem('Sheets Discord Tools', 'AddForm');
    menu.addToUi();
}

function localediff(input) {
    const updatedText1 = input.replace(/\{([^{}]*)\}/g, (match, group1) => {
        return '{' + group1.replace(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/g, '\\') + '}';
    });
    return updatedText1.replace(/,(?=(?:[^"]*"[^"]*")*[^"]*$)(?![^{]*\})/g, ';');
}

function reverseLocalediff(input) {
    let restoredText1 = input.replace(/\{([^{}]*)\}/g, (match, group1) => {
        return "{" + group1.replace(/\\/g, ",") + "}";});
    return restoredText1.replace(/;(?![^{]*\})/g, ',');
}

function setTimestamp() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var cell = sheet.getActiveCell();
    var now = new Date();
    var formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM//dd/yyyy HH:mm:ss");
    cell.setValue(formattedDate);
}

function cleanRange() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var range = sheet.getActiveRange();
    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
        for (var j = 0; j < values[i].length; j++) {
            var cellValue = values[i][j];
            if (typeof cellValue === 'string') {
                values[i][j] = cleanAndTrim(cellValue);
            }
        }
    }
    range.setValues(values);
}

function cleanAndTrim(text) {
    var cleanedText = text.replace(/[\x00-\x1F\x7F-\x9F]/g, "");
    return cleanedText.trim();
}


function unpivot(value) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const range = sheet.getActiveRange(); // Get the selected range

    if (!range) {
        Logger.log("No range selected.");
        return;
    }

    const values = range.getValues();
    const headers = values[0]; // First row is assumed to be headers
    const data = values.slice(1); // Remaining rows are data

    const unpivotedData = [];

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        for (let j = value; j < headers.length; j++) { // Start from 1 to skip the first column (assumed ID)
            unpivotedData.push([row[0], headers[j], row[j]]);
        }
    }

    // Write the unpivoted data to a new sheet or the same sheet
    const newSheet = ss.insertSheet("Unpivoted Data"); // Create a new sheet
    newSheet.getRange(1, 1, unpivotedData.length, unpivotedData[0].length).setValues(unpivotedData);
}