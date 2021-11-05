function getDataFromOsoba() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Konfiguracja (Osoba)');
  const [header, ...data] = sheet.getDataRange().getDisplayValues();
  const choices = {}
  header.forEach(function(title, index) {
    choices[title] = data.map(row => row[index]).filter(e => e !== '');
  });
  //debugger;
  return choices;
}

function getDataFromFirma() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Konfiguracja (Firma)');
  const [header, ...data] = sheet.getDataRange().getDisplayValues();
  const choices = {}
  header.forEach(function(title, index) {
    choices[title] = data.map(row => row[index]).filter(e => e !== '');
  });
  //debugger;
  return choices;
}

function getDataFromRelacje() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Konfiguracja (Relacje)');
  const [header, ...data] = sheet.getDataRange().getDisplayValues();
  const choices = {}
  header.forEach(function(title, index) {
    choices[title] = data.map(row => row[index]).filter(e => e !== '');
  });
  //debugger;
  return choices;
}

function importToOsobaForm() {
  const GOOGLE_FORM_ID = '1h6S2iTtqLpVMNTGhAM1hrKr4yzNYzh6DiLwy42DjIZw';
  const googleForm = FormApp.openById(GOOGLE_FORM_ID);
  const items = googleForm.getItems();
  const choices = getDataFromOsoba();
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

function importToFirmaForm() {
  const GOOGLE_FORM_ID = '1J0NKweKL6C7NOeMHF60UZRGowZe1pI7zRhiDPkkB-sA';
  const googleForm = FormApp.openById(GOOGLE_FORM_ID);
  const items = googleForm.getItems();
  const choices = getDataFromFirma();
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

function importToRelacjeForm() {
  const GOOGLE_FORM_ID = '1pMGRJl2RTZVkppE7Ny9EoxWwkCUM-fwc_MV0hs58h2M';
  const googleForm = FormApp.openById(GOOGLE_FORM_ID);
  const items = googleForm.getItems();
  const choices = getDataFromRelacje();
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
