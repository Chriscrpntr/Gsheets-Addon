function onHomepage(e) {
    return createMainCard();
}

function createMainCard() {
    let card = CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle("Sheets Discord Tools"))
        .addSection(createToolsSection());
    return card.build();
}

function createToolsSection() {
    let section = CardService.newCardSection()
        .setHeader("Tools");
    let sidebarbutton = CardService.newTextButton()
        .setText("Show Sidebar")
        .setOnClickAction(CardService.newAction().setFunctionName("showSidebar"));
    section.addWidget(CardService.newButtonSet().addButton(sidebarbutton));
    return section;
}


function showSidebar(){
    var form = HtmlService.createHtmlOutputFromFile("SheetsTools").setTitle("Sheets Discord Tools");
    SpreadsheetApp.getUi().showSidebar(form);
}

function onOpen(){
    SpreadsheetApp.getUi().createAddonMenu()
        .addItem('Sheets Discord Tools','showSidebar')
        .addToUi();
    showSidebar();
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
    const activeCell = sheet.getActiveCell(); // Get the currently selected cell

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