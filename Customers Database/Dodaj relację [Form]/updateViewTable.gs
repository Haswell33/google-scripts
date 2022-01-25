// onSubmitEvent()
// Useful in sheet to read-only usage. where are only references from other sheets with data, thanks of this script it updates references on e. g. every edit or smth else
// row/columns to reference should be defined as parameters in getDataFromReferenceSheet()

var sheetID = 'google-sheet-id' //arkusz Formularz

function updateViewTable() {
  const sheet = SpreadsheetApp.openById(sheetID).getSheetByName('google-subsheet-name')
  sheet.getRange(1, 1).setValue('=INDEX(google-subsheet-name!C:D)') // source sheets
  sheet.getRange(1, 21).setValue('=INDEX(google-subsheet-name!E:H)')
  getDataFromReferenceSheet('google-subsheet-name', 3, 12, 'E', 'N', 0, sheet) // source sheets
  getDataFromReferenceSheet('google-subsheet-name', 13, 19, 'D', 'J', 1, sheet)
  getDataFromReferenceSheet('google-subsheet-name', 20, 20, 'L', 'L', 1, sheet) 
}

function getDataFromReferenceSheet(referenceSheetName, startColNumber, endColNumber, startColAddress, endColAddrress, keyCol, sheet) {
  const referenceSheet = SpreadsheetApp.openById(sheetID).getSheetByName(referenceSheetName)
  var data = sheet.getDataRange().getValues()
  var referenceData = referenceSheet.getDataRange().getValues()
  sheet.getRange(1, startColNumber, endColNumber).setValue('=INDEX(' + referenceSheetName + '!' + startColAddress + '1:' + endColAddrress + '1)') // header row
  var row = 2
  for(var i = 1; i < data.length; i++) {
    const relatedRow = findRelatedData(referenceData, data[i][keyCol])
    if (relatedRow == 'REFERENCE_NOT_EXISTS'){
      Logger.log('there is a reference error in row: %d', row)
      sheet.getRange(row, startColNumber).setValue(relatedRow)
      //sheet.getRange(row, 1, 1, sheet.getMaxColumns()).setBackground('#e06666');
    }
    else
      sheet.getRange(row, startColNumber).setValue('=INDEX(' + referenceSheetName + '!' + startColAddress + relatedRow + ':' + endColAddrress + relatedRow + ')')
    row++
  }
}

function findRelatedData(data, currCompanyName) {
  for(var i = 0; i < data.length; i++) {
    if(data[i][2] == currCompanyName)
      return i+1 
  }
  return 'REFERENCE_NOT_EXISTS'
}
