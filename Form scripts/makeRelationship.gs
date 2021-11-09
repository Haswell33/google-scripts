var sheetID = '1PIdpiijgl8UiT4ehqguiWjaa4r6-XavvltYXfn6Gbg0'

function addFormulaToNewRow(sheet, parentSheetName, col) {
    const row = sheet.getLastRow();
    const currValue = sheet.getRange(row, col).getValue();
    const relatedRow = findNewValueInParentSheet(currValue, parentSheetName, sheetID);
    Logger.log(currValue, relatedRow);
    sheet.getRange(row, col).setValue('=INDEX(' + parentSheetName + '!' + relatedRow + ':' + relatedRow + ', 1, 3)')
}

function findNewValueInParentSheet(newValue, sheetName) {
    const sheet = SpreadsheetApp.openById(sheetID).getSheetByName(sheetName);
    var data = sheet.getDataRange().getValues();
    for(var i = 0; i < data.length; i++) {
        if(data[i][2] == newValue) {
            return i+1;
        }
    }
}

function makeRelationship() {
    const sheet = SpreadsheetApp.openById(sheetID).getSheetByName('Relacja')
    addFormulaToNewRow(sheet, 'Osoba', 3);
    addFormulaToNewRow(sheet, 'Firma', 4);
}