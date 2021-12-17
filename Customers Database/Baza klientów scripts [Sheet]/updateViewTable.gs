//updatuje arkusz Widok aktualizując o nowo dodaną relację przy każdej edycji tego arkusza
// all columns 23

function updateViewTable() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Widok')
  sheet.getRange(1, 1).setValue('=INDEX(Relacja!C:D)') //Osoba Firma w kolumnie(row=1, col=1)
  sheet.getRange(1, 21).setValue('=INDEX(Relacja!E:H)') //Nazwa stanowiska, Klasyfikacja roli, Data od, Data do w kolumnie
  getDataFromReferenceSheet('Osoba', 3, 12, 'E', 'N', 0, sheet) // email priv, email firm, telefon priv, telefon firm., jezyk, rola, klasa prez., event notatka, ogola notatka, opiekun
  getDataFromReferenceSheet('Firma', 13, 19, 'D', 'J', 1, sheet) // ulica, kod poczt., miasto, kraj, var, dzialalnosc, status
  getDataFromReferenceSheet('Firma', 20, 20, 'L', 'L', 1, sheet) //Recepcja e-mail
}

function getDataFromReferenceSheet(referenceSheetName, startColNumber, endColNumber, startColAddress, endColAddrress, keyCol, sheet) {
  const referenceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(referenceSheetName)
  var data = sheet.getDataRange().getValues()
  var referenceData = referenceSheet.getDataRange().getValues()
  sheet.getRange(1, startColNumber, endColNumber).setValue('=INDEX(' + referenceSheetName + '!' + startColAddress + '1:' + endColAddrress + '1)') // header row
  var row = 2
  for(var i = 1; i < data.length; i++) {
    const relatedRow = findRelatedData(referenceData, data[i][keyCol])
    if (relatedRow == 'REFERENCE_NOT_EXISTS'){
      Logger.log('there is a reference error in row: %d', row)
      sheet.getRange(row, startColNumber).setValue(relatedRow)
      sheet.getRange(row, 1, 1, sheet.getMaxColumns()).setBackground('#e06666');
    }
    else
      sheet.getRange(row, startColNumber).setValue('=INDEX(' + referenceSheetName + '!' + startColAddress + relatedRow + ':' + endColAddrress + relatedRow + ')')
      sheet.getRange(row, 1, 1, sheet.getMaxColumns()).setBackground('#ffffff');
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
