// Adds question to Google Form from provided google config sheet. 
// IMPORTANT - in order to import questions u need to also add this question in form in very manual way

function updateAddPersonGoogleForm() {
  const googleFormID = 'source-google-sheet-id'
  const choices = getDataFromSheet('google-subsheet-name') //sheet with defined responses, title row as question name and all below as options in dictionary
  importDataToGoogleForm(googleFormID, choices)
}

function updateAddCompanyGoogleForm() {
  const googleFormID = 'source-google-sheet-id'
  const choices = getDataFromSheet('google-subsheet-name')
  importDataToGoogleForm(googleFormID, choices)
}

function updateAddRelationShipGoogleForm() {
  const googleFormID = 'source-google-sheets-id'
  const choices = getDataFromSheet('google-subsheet-name')
  importDataToGoogleForm(googleFormID, choices)
}

function getDataFromSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
  const [header, ...data] = sheet.getDataRange().getDisplayValues();
  const choices = {}
  header.forEach(function(title, index) {
    choices[title] = data.map(row => row[index]).filter(e => e !== '')
  })
  return choices;
}

function importDataToGoogleForm(googleFormID, choices){
  const googleForm = FormApp.openById(googleFormID)
  const items = googleForm.getItems()
  items.forEach(function(item) {
    const itemTitle = item.getTitle()
    if (itemTitle in choices){
      const itemType = item.getType()
      switch (itemType) {
        case FormApp.ItemType.CHECKBOX:
          item.asCheckboxItem().setChoiceValues(choices[itemTitle])
          break
        case FormApp.ItemType.LIST:
          item.asListItem().setChoicesValues(choices[itemTitle])
          break
        case FormApp.ItemType.MULTIPLE_CHOICE:
          item.asMultipleChoiceItem().setChoiceValues(choices[itemTitle])
          break
        default:
          Logger.log('question ignored: %s', itemTitle)
      }
    }
  })
}
