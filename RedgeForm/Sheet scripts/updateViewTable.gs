function updateSheet(sourceSheetName, col, startColAddress, endColAddrress, keyCol) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Widok')
    const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sourceSheetName)
    var data = sheet.getDataRange().getValues();
    var sourceData = sourceSheet.getDataRange().getValues();
    var row = 1;
    sheet.getRange(row, 1).setValue('=INDEX(Relacja!C:D)')
    sheet.getRange(row, 19).setValue('=INDEX(Relacja!E:F)')
    for(var i = 0; i < data.length; i++) {
        const relatedRow = findRelatedData(sourceData, data[i][keyCol])
        sheet.getRange(row, col).setValue('=INDEX(' + sourceSheetName + '!' + startColAddress + relatedRow + ':' + endColAddrress + relatedRow + ')')
        row++;
    }
}

function findRelatedData(data, currCompanyName) {
    for(var i = 0; i < data.length; i++) {
        if(data[i][2] == currCompanyName) {
            return i+1;
        }
    }
    return 1;
}

function updateViewTable() {
    updateSheet('Osoba', 3, 'D', 'M', 0)
    updateSheet('Firma', 13, 'D', 'I', 1)
}