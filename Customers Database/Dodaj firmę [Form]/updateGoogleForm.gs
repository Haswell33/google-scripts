//Dodaje opcje odpowiedzi synchronizując je z arkuszy Konifguracja (Osoba), Konifguracja (Firma), Konifguracja (Relacje) do odpowiadających im formularzy
//Aby słownik został zaimportowany konieczne jest dodanie tego pytania w fomularzu !!!

function updateGoogleFormRelacja() {
  const googleFormID = '1pMGRJl2RTZVkppE7Ny9EoxWwkCUM-fwc_MV0hs58h2M';
  const choices = getDataFromSheet('Konfiguracja (Relacje)');
  importDataToGoogleForm(googleFormID, choices)
}

function getDataFromSheet(sheetName) {
  const sheet = SpreadsheetApp.openById(sheetID).getSheetByName(sheetName);
  const [header, ...data] = sheet.getDataRange().getDisplayValues();
  const choices = {}
  header.forEach(function(title, index) {
    choices[title] = data.map(row => row[index]).filter(e => e !== '');
  });
  return choices;
}

function importDataToGoogleForm(googleFormID, choices){
  const googleForm = FormApp.openById(googleFormID);
  const items = googleForm.getItems();
  items.forEach(function(item) {
    const itemTitle = item.getTitle();
    if (itemTitle in choices){
      const itemType = item.getType();
      switch (itemType) {
        case FormApp.ItemType.CHECKBOX:
          item.asCheckboxItem().setChoiceValues(choices[itemTitle]);
          break;
        case FormApp.ItemType.LIST:
          item.asListItem().setChoicesValues(choices[itemTitle]);
          break;
        case FormApp.ItemType.MULTIPLE_CHOICE:
          item.asMultipleChoiceItem().setChoiceValues(choices[itemTitle]);
          break;
        default:
          Logger.log('question ignored: %s', itemTitle);
      }
    }
  });
}
