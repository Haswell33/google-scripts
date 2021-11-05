function searchCompanyName(new_value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Firma');
  var data = sheet.getDataRange().getValues();
  for(var i = 0; i<data.length;i++){
    if(data[i][1] == new_value){ //[1] because column B
      Logger.log((i+1))
      return i+1;
    }
  }
}

function addFormulaToNewValue() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Relacja");
  var row_nr = sheet.getLastRow();
  var col_nr = 3;
  var curr_value = sheet.getRange(row_nr, col_nr).getValue();
  firma_row_nr = searchCompanyName(curr_value);
  Logger.log(curr_value, firma_row_nr);
  sheet.getRange(row_nr, col_nr).setValue('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1PIdpiijgl8UiT4ehqguiWjaa4r6-XavvltYXfn6Gbg0", "Firma!B'+firma_row_nr+':B'+firma_row_nr+'")')
}
