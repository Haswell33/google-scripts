function Osoba() {
  const googleFormID = '1h6S2iTtqLpVMNTGhAM1hrKr4yzNYzh6DiLwy42DjIZw';
  const choices = getDataFromSheet('Konfiguracja (Osoba)');
  importDataToGoogleForm(googleFormID, choices)
}

function Firma() {
  const googleFormID = '1J0NKweKL6C7NOeMHF60UZRGowZe1pI7zRhiDPkkB-sA';
  const choices = getDataFromSheet('Konfiguracja (Firma)');
  importDataToGoogleForm(googleFormID, choices)
}

function Relacje() {
  const googleFormID = '1pMGRJl2RTZVkppE7Ny9EoxWwkCUM-fwc_MV0hs58h2M';
  const choices = getDataFromSheet('Konfiguracja (Relacje)');
  importDataToGoogleForm(googleFormID, choices)
}

function getDataFromSheet(SHEET_NAME) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const [header, ...data] = sheet.getDataRange().getDisplayValues();
  const choices = {}
  header.forEach(function(title, index) {
    choices[title] = data.map(row => row[index]).filter(e => e !== '');
  });
  return choices;
}

function importDataToGoogleForm(GOOGLE_FORM_ID, choices){
  const googleForm = FormApp.openById(GOOGLE_FORM_ID);
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
          Logger.log('Ignore question', itemTitle);
      }
    }
  });
}