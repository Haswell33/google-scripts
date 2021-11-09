var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('Relacja');


function readColumnData() {
    var column = 2;
    var lastRow = sheet.getLastRow();
    var columnRange = sheet.getRange(1, column, lastRow);
    var rangeArray = columnRange.getValues();
    // Convert to one dimensional array
    rangeArray = [].concat.apply([], rangeArray);
    return rangeArray;
}

function readRowData() {
    var row = 2;
    var lastColumn = sheet.getLastColumn();
    var rowRange = sheet.getRange(row, 1, 1, lastColumn);
    var rangeArray = rowRange.getValues();
    // Convert to one dimensional array
    rangeArray = [].concat.apply([], rangeArray);
    Logger.log(rangeArray);
    return rangeArray;
}

function findDuplicates(data) {
    var sortedData = data.slice().sort();
    var duplicates = [];
    for (var i = 0; i < sortedData.length - 1; i++) {
        if (sortedData[i + 1] == sortedData[i]) {
            duplicates.push(sortedData[i]);
        }
    }
    return duplicates;
}

function getIndexes(data, duplicates) {
    var column = 2;
    var indexes = [];
    i = -1;
    // Loop through duplicates to find their indexes
    for (var n = 0; n < duplicates.length; n++) {
        while ((i = data.indexOf(duplicates[n], i + 1)) != -1) {
            indexes.push(i);
        }
    }
    return indexes;
}

function highlightColumnDuplicates(indexes) {
    var column = 7;
    sheet.getRange(indexes[indexes.length - 2], column).setBackground('green')
    //for (n = 0; n < indexes.length; n++) {
    //  sheet.getRange(indexes[indexes.length-1] + 1, column).setBackground("yellow");
    //}
}

function columnMain() {
    var data = readColumnData();
    var duplicates = findDuplicates(data);
    var indexes = getIndexes(data, duplicates);
    highlightColumnDuplicates(indexes);
}

//function rowMain() {
//  var data = readRowData();
//  var duplicates = findDuplicates(data);
//  var indexes = getIndexes(data, duplicates);
// highlightRowDuplicates(indexes);
//}