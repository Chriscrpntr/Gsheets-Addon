// Set Up Section ------------------------------------------------------------------

function onHomepage(e) {
    return createMainCard();
}

function onOpen(){
    SpreadsheetApp.getUi().createAddonMenu()
        .addItem('Sheets Discord Tools','showSidebar')
        .addItem('Consolidate Data', 'openConsolidateDataDialog')
        .addToUi();
}

function createMainCard() {
    let card = CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("Sheets Discord Tools"))
        .addSection(createToolsSection())
        .addSection(createConsolidateSection());
    return card.build();
}

function createConsolidateSection() {
    let consolidateSection = CardService.newCardSection()
        .setHeader("Consolidate Data");
    let consolidateButton = CardService.newTextButton()
        .setText("Consolidate Tool")
        .setOnClickAction(CardService.newAction().setFunctionName("openConsolidateDataDialog"));
    consolidateSection.addWidget(CardService.newButtonSet().addButton(consolidateButton));

    return consolidateSection;
}

function createToolsSection() {
    let section = CardService.newCardSection()
        .setHeader("Spreadsheet Tools");
    let sidebarbutton = CardService.newTextButton()
        .setText("Show Sidebar")
        .setOnClickAction(CardService.newAction().setFunctionName("showSidebar"));
    section.addWidget(CardService.newButtonSet().addButton(sidebarbutton));

    return section;
}


function openConsolidateDataDialog() {
    var html = HtmlService.createHtmlOutputFromFile('ConsolidateData')
        .setWidth(800)
        .setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(html, 'Consolidate Data');
}

function showSidebar(){
    var form = HtmlService.createHtmlOutputFromFile("SheetsTools").setTitle("Sheets Discord Tools");
    SpreadsheetApp.getUi().showSidebar(form);
}

// Locale Conversion Section ------------------------------------------------------------------

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

// Timestamp Section ------------------------------------------------------------------

function setTimestamp() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var cell = sheet.getActiveCell();
    var now = new Date();
    var formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM//dd/yyyy HH:mm:ss");
    cell.setValue(formattedDate);
}

// Clean Range Section ------------------------------------------------------------------

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

// Unpivot Section ------------------------------------------------------------------

function unpivot(input) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const range = sheet.getActiveRange();

    const values = range.getValues();
    if (!values || values.length < 2 || values[0].length < 2) {
        SpreadsheetApp.getUi().alert("Selected range must have at least 2 rows and 2 columns.");
        return;
    }

    const headers = values[0];
    const data = values.slice(1);
    const unpivotedData = [];

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const idValues = row.slice(0, input);

        for (let j = input; j < headers.length; j++) {
            unpivotedData.push([...idValues, headers[j], row[j]]);
        }
    }

    const newSheet = ss.insertSheet("Unpivoted Data");
    newSheet.getRange(1, 1, unpivotedData.length, unpivotedData[0].length).setValues(unpivotedData);
}

// Convert to Numeric Section ------------------------------------------------------------------

function convertToNumeric() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var range = sheet.getActiveRange();
    var values = range.getValues();

    for (var i = 0; i < values.length; i++) {
        for (var j = 0; j < values[i].length; j++) {
            var cellValue = values[i][j];

            if (typeof cellValue === 'string') {
                var numValue = tryParseNumeric(cellValue.trim());
                if (numValue !== null) {
                    values[i][j] = numValue;
                }
            }
        }
    }
    range.setValues(values);
}

function tryParseNumeric(str) {
    if (/^[-+]?(\d+(\.\d*)?|\.\d+)$/.test(str)) {
        var num = Number(str);
        if (!isNaN(num))
            return num;
    }
    return null;
}

// Letify Section ------------------------------------------------------------------

function letifyFormula() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const cell = sheet.getActiveCell();
    const formula = cell.getFormula();
    if (!formula) return "";

    const regex = /([^",({]+!)?([A-Za-z]+[0-9]*:[A-Za-z]+[0-9]*)(?=(?:[^"]*"[^"]*")*[^"]*$)|([^",({]+!)?([A-Za-z]+[0-9]+)(?=(?:[^"]*"[^"]*")*[^"]*$)/g;

    let match;
    let ranges = [];
    while ((match = regex.exec(formula)) !== null) {
        ranges.push(match[0]);
    }

    ranges = [...new Set(ranges)];
    Logger.log(ranges)

    let letVariables = "";
    let variableCounter = 1;
    let newFormula = formula;

    for (const range of ranges) {
        const variableName = "variable" + variableCounter;
        letVariables += `\n${variableName},${range},`;
        const escapedRange = range.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
        const regexReplace = new RegExp("\\b" + escapedRange + "\\b", "g");
        newFormula = newFormula.replace(regexReplace, variableName);
        variableCounter++;
    }

    if (letVariables) {
        letVariables = letVariables.slice(0, -1);
        if (newFormula.startsWith("=")) {
            newFormula = newFormula.substring(1);
        }
        cell.setValue(`=LET(${letVariables},\n${newFormula})`);
        return `=LET(${letVariables},\n${newFormula})`;
    } else {
        return formula;
    }
}

// Regex Section ------------------------------------------------------------------

function regexCell(str, pattern) {
    const regex = new RegExp(pattern, "g");
    const matches = [];
    let match;

    while ((match = regex.exec(str)) !== null) {
        matches.push(match[0]);
    }

    if (matches.length > 0) {
        return matches;
    } else {
        return "No Matches";
    }
}

function outputRegex(regexoutputValues){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const activeCell = sheet.getActiveCell();

    if (!activeCell) {
        Logger.log("No cell is selected.");
        return;
    }

    const startRow = activeCell.getRow();
    const startColumn = activeCell.getColumn();

    const splitvalues = regexoutputValues.split(",").map(s => s.trim());

    for (let i = 0; i < splitvalues.length; i++) {
        sheet.getRange(startRow + i, startColumn).setValue(splitvalues[i]);
    }
}

// Cropping Section ------------------------------------------------------------------

function cropSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var selection = sheet.getActiveRange();

    var startRow = selection.getRow();
    var startColumn = selection.getColumn();
    var numRows = selection.getNumRows();
    var numCols = selection.getNumColumns();
    var maxRows = sheet.getMaxRows();
    var maxCols = sheet.getMaxColumns();

    if (maxRows > startRow + numRows - 1) {
        sheet.deleteRows(startRow + numRows, maxRows - (startRow + numRows - 1));
    }

    if (maxCols > startColumn + numCols - 1) {
        sheet.deleteColumns(startColumn + numCols, maxCols - (startColumn + numCols - 1));
    }
}

// Consolidation Section ------------------------------------------------------------------

function getSheetNames() {
    return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(sheet => sheet.getName());
}

function consolidateData(formObject) {
    var selectedSheets = formObject.selectedSheets || [];
    var rangeInputs = formObject.rangeInputs || {};

    var dataToConsolidate = [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    for (var i = 0; i < selectedSheets.length; i++) {
        var sheetName = selectedSheets[i];
        var sheet = ss.getSheetByName(sheetName);

        if (sheet) {
            var rangeInput = rangeInputs.hasOwnProperty(sheetName) ? rangeInputs[sheetName] : null;

            if (rangeInput) {
                try {
                    var range = sheet.getRange(rangeInput);
                    var numRows = range.getNumRows();
                    var numCols = range.getNumColumns();
                    var values = range.getValues();

                    for (var j = 0; j < numRows; j++) {
                        var rowData = values[j];
                        var isEmptyRow = true;

                        for (var k = 0; k < numCols; k++) {
                            if (rowData[k] !== "") {
                                isEmptyRow = false;
                                break;
                            }
                        }

                        if (!isEmptyRow) {
                            rowData.push(sheetName);
                            dataToConsolidate.push(rowData);
                        }
                    }
                } catch (e) {
                    return "Error: Invalid range for " + sheetName + ": " + e.message;
                }
            } else {
                return "Error: No range specified for " + sheetName;
            }
        }
    }

    var newSheet = ss.insertSheet("Consolidated Data");
    if (dataToConsolidate.length > 0) {
        newSheet.getRange(1, 1, dataToConsolidate.length, dataToConsolidate[0].length).setValues(dataToConsolidate);
    } else {
        newSheet.getRange(1, 1).setValue("No data to consolidate.");
    }

    return "OK";
}