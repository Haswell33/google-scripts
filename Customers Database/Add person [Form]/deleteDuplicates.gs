//it should immediately deletes duplicates, but unfinished
var sheetID = 'google-sheet-id' //main sheet with forms data

function deleteDuplicates() {
  const sheet = SpreadsheetApp.openById(sheetID).getSheetByName('Firma');
  const lastRow = sheet.getLastRow();
  const lastValues = sheet.getRange('C'+lastRow+':C'+lastRow).getValues();
  var name = lastValues[0][0];
  const allNames = sheet.getRange('C2:C').getValues()
  var row;
  var len;
  //name = 'Festival Group Sp. z o.o.'
  for (row = 0, len = allNames.length; row < len - 1; row++) // TRY AND FIND EXISTING NAME
    if (allNames[row] == name) {
      Logger.log('duplicated value found in C%d\nname of duplicated value: %s', row+2, sheet.getRange(row+2, 3).getValue())
      //sheet.getRange('C2').offset(0, 0, row, lastValues.length).setValues([lastValues]); // OVERWRITE OLD DATA
      sheet.deleteRow(lastRow); //delete duplicated row
      break;
  }
}
