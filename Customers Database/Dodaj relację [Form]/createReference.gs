// Tworzy relacje z arkuszu Firma oraz Osoba w arkuszu Relacja w kolumnie Osoba oraz Firma przy dodaniu nowej relacji, OnSubmitEvent

var sheetID = '1PIdpiijgl8UiT4ehqguiWjaa4r6-XavvltYXfn6Gbg0' //arkusz Formularz

function createReference() {
  const sheet = SpreadsheetApp.openById(sheetID).getSheetByName('Relacja')
  addReference(sheet, 'Osoba', 3)
  addReference(sheet, 'Firma', 4)
}

function addReference(sheet, referenceSheetName, col) { // Replace common string to formula in order to make reference
  const referenceSheet = SpreadsheetApp.openById(sheetID).getSheetByName(referenceSheetName) 
  const data = referenceSheet.getDataRange().getValues()
  const row = sheet.getLastRow()
  const currValue = sheet.getRange(row, col).getValue()
  const relatedRow = findNewValueInReferenceSheet(currValue, data)
  Logger.log('reference %s created for %s ', sheet.getRange(row, col).getFormula(), currValue)
  sheet.getRange(row, col).setValue('=INDEX(' + referenceSheetName + '!' + relatedRow + ':' + relatedRow + ', 1, 3)')
}

function findNewValueInReferenceSheet(newValue, data) { // Find related value in selected sheet
  if (newValue.includes('+'))
    newValue = newValue.replace('+', '.')
  newValue = newValue.replace(' ', '.')
  var regexPattern = new RegExp(newValue)
  try {
    for(var i = 0; i <= data.length; i++) {
      if (regexPattern.test(data[i][2]))
        return i+1
    }
  } catch(e) {
    if (e instanceof TypeError)  // Jakby sie pojawil blad w utworzeniu referencji, to poleci info do mnie na maila
      MailApp.sendEmail('karol.siedlaczek@redge.com', '[Formularz sheet] REFERENCE_ERROR', ('Reference has not been created for ' + newValue))
  }
}
