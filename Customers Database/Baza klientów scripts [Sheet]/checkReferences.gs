//TO UPDATE
var sheetID = '1PIdpiijgl8UiT4ehqguiWjaa4r6-XavvltYXfn6Gbg0' //arkusz Formularz

function checkReferences() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Relacja')
  checkReference(sheet, 'Osoba', 1)
  checkReference(sheet, 'Firma', 4);
}

function checkReference(sheet, referenceSheetName, col) {
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(referenceSheetName)
  const lastRow = sheet.getLastRow()
  var row = 2
  for(var i = 0; i < lastRow - 1; i++) {
    const currValue = sheet.getRange(row, col).getValue() // (row, col)
    const currFormula = sheet.getRange(row, col).getFormula()
    if (currFormula.substring(0, 1) != '=') // czyli jak nie ma referencji, tylko string
      addReference(currValue, sourceSheet, referenceSheetName, row, col, sheet)
    row++;
  }
}

function addReference(currValue, sourceSheet, referenceSheetName, row, col, sheet){ 
  const relatedRow = findValueInReferenceSheet(currValue, sourceSheet)
  sheet.getRange(row, col).setValue('=INDEX(' + referenceSheetName + '!' + relatedRow + ':' + relatedRow + ', 1, 3)')
  Logger.log('formula %s created for %s ', sheet.getRange(row, col).getFormula(), currValue);
}

function findValueInReferenceSheet(newValue, sourceSheet) {
  const data = sourceSheet.getDataRange().getValues()
  if (newValue.includes('+'))
    newValue = newValue.replace('+', '.')
  newValue = newValue.replace(' ', '.')
  var regexPattern = new RegExp(newValue)
  try{
    for(var i = 0; i <= data.length; i++) {
      if (regexPattern.test(data[i][2]) == true)
        return i+1;  
    }
  } catch(e){
    if (e instanceof TypeError)
      MailApp.sendEmail('karol.siedlaczek@redge.com', 'REFERENCE_ERROR in Formularz', ('Reference has not been created for ' + newValue))
  }
}


