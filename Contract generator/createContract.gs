const trueString = 'Tak' // cale to temporary
const dataStartCol = 'B'
const dataEndCol = 'AE'
const urlCol = 'AF'
const minHourPerMonthRange = 'A2'
const maxHourPerMonthRange = 'B2'
const salaryTypeGodzinowaRange = 'C2'
const salaryTypeRyczaltRange = 'D2'
const motivatingSystemRange = 'E2'
const benefits1Range = 'F2'
const benefits2Range = 'G2'
const taxpayerVatRange = 'H2'
const startTextRange = ['I2', 'I3', 'I4', 'I5', 'I6', 'I7', 'I8'] //temporary
const workerPositionCol = 'P'

class Contract {
  constructor(workerName, workerCity, workerAddress, workerPesel, workerCompanyName, workerCompanyNIP, workerCompanyCity, workerCompanyAddress,                            
              agentName, agentPhone, agentEmail, maxHourPerMonth, benefits, motivatingSystem, taxpayerVAT, workerPosition, 
              periodOfNotice, salaryType, salaryConstNumeric, salaryConstWords, salaryRatioType, salaryRatioNumeric, salaryRatioWords,  
              salaryRatioMaxYearNumeric, salaryRatioMaxYearWords, startDate,  timeType,  timestamp, minHourPerMonth, responderEmail){
    this.workerName = workerName                                  //  Imie i nazwisko (Pracownik)                      {worker_name}
    this.workerCity = workerCity                                  //  Miejscowosc (Pracownik)                          {worker_city}
    this.workerAddress = workerAddress                            //  Adres (Pracownik)                                {worker_address}  
    this.workerPesel = workerPesel                                //  PESEL                                            {worker_pesel}
    this.workerPosition = workerPosition                          //  Stanowisko                                       {worker_position}
    this.workerCompanyName = workerCompanyName                    //  Nazwa (Firma)                                    {worker_company_name}
    this.workerCompanyNIP = workerCompanyNIP                      //  NIP                                              {worker_company_nip}
    this.workerCompanyCity = workerCompanyCity                    //  Miejscowosc (Firma)                              {worker_company_city}
    this.workerCompanyAddress = workerCompanyAddress              //  Adres (Firma)                                    {worker_company_address}
    this.agentName = agentName                                    //  Imie i nazwisko (Przedstawiciel)                 {agent_name}
    this.agentPhone = agentPhone                                  //  Nr telefonu                                      {agent_phone}
    this.agentEmail = agentEmail                                  //  Adres e-mail                                     {agent_email}
    //this.minHourPerMonth = minHourPerMonth                      //  Min. limit godzin w miesiacu                     {min_hours_per_month}
    this.maxHourPerMonth = maxHourPerMonth                        //  Maks. limit godzin w miesiacu                    {max_hours_per_month}            if "Godzinowa"
    this.minHourPerMonth = minHourPerMonth                        //  Wartosc min. ilosci godzin w miesiacu            {min_hours_per_month_value}      if min_hours_per_month == 'Tak'
    //this.maxHourPerMonthValue = maxHourPerMonthValue            //  Wartosc max. ilosci godzin w miesiacu            {max_hours_per_month_value}      if max_hours_per_month == 'Tak'    
    this.benefits = benefits                                      //  Benefity                                         {benefits1}, {benefits2}
    this.motivatingSystem = motivatingSystem                      //  Dedykowany system motywacyjny                    {motivatingSystem}
    this.taxpayerVAT = taxpayerVAT                                //  Podatnik VAT                                     {taxpayer_vat} 
    this.periodOfNotice =  periodOfNotice                         //  Okres wypowiedzenia                              {period_of_notice}
    this.salaryType = salaryType                                  //  Rodzaj stawki                                    {salary_type}                    {"Ryczałtowa", "Godzinowa"}
    this.salaryConstNumeric = salaryConstNumeric                  //  Wysokośc wynagrodzenia (liczbowo)                {salary_const_numeric}           if "Ryczałtowa"
    this.salaryConstWords = salaryConstWords                      //  Wysokośc wynagrodzenia (słownie)                 {salary_const_words}             if "Ryczałtowa"
    this.salaryRatioType = salaryRatioType                        //  Jednostka iloczynu                               {salary_ratio_type}              if "Godzinowa"
    this.salaryRatioNumeric = salaryRatioNumeric                  //  Wysokość wynagr. za wsk. jdn. czas. (liczbowo)   {salary_ratio_numeric}           if "Godzinowa"
    this.salaryRatioWords = salaryRatioWords                      //  Wysokość wynagr. za wsk. jdn. czas. (słownie)    {salary_ratio_words}             if "Godzinowa"
    this.salaryRatioMaxYearNumeric = salaryRatioMaxYearNumeric    //  Maks. kwota wynagr. netto za 12 mies. (liczbowo) {salary_ratio_max_year_numeric}  if "Godzinowa"
    this.salaryRatioMaxYearWords = salaryRatioMaxYearWords        //  Maks. kwota wynagr. netto za 12 mies. (słownie)  {salary_ratio_max_year_words}    if "Godzinowa"
    this.startDate = startDate.toLocaleDateString("pl-PL")        //  Dzien zawarcia umowy                             {start_date}
    this.timeType = timeType                                      //  Czas zawarcia umowy                              {time_type}                      {"Określony","Nieokreślony"} 
    this.timestamp = timestamp                                    //  Ilosc miesięcy                                   {timestamp}                      if "Określony"  
    this.responderEmail = responderEmail                          //  Osoba która wypełniła formularz                  
  }
}

function onOpen(){
  const sheetUi = getSheetUi()
  const menu = sheetUi.createMenu('Wygeneruj umowę');
  menu.addItem('Wybierz wykonawcę', 'createContactSelect')
  menu.addItem('Generowanie zbiorcze', 'createContactBulk')
  menu.addToUi();
}

function createContractLast() {
  const dataSheet = getDataSheet()
  var lastRow = dataSheet.getLastRow()
  createDocument(lastRow, true)
}

function createContactSelect() {
  const sheetUi = getSheetUi()
  const dataSheet = getDataSheet()
  var data = dataSheet.getDataRange().getValues()
  var result = sheetUi.prompt('Podaje imię i nazwisko');
  result =  result.getResponseText()
  Logger.log('Person to generate contract: ' + result);
  for(var i = 1; i < data.length; i++) {
    if (data[i][1] == result){
      var rowNumber = i + 1
      break
    }
  }
  if (rowNumber == null){
    Logger.log('Did not found selected worker, creation aborted')
    sheetUi.alert('Nie znaleziono wskazanego wykonawcy, generowanie umowy zakończone niepowodzeniem')
  }
  else{
    Logger.log('Found selected worker in row number: ' + rowNumber)
    sheetUi.alert('Rozpoczęto generowanie umowy dla: ' + result)
    var url = createDocument(rowNumber, false)
    sheetUi.alert('Generowanie umowy zakończone:\n' + url)
    Logger.log('Contract from row number ' + i + ' generated, mail has not been sent')
  }
}

function createContactBulk() {
  const dataSheet = getDataSheet()
  //ui = getSheetUi()
  //var htmlDlg = HtmlService.createHtmlOutputFromFile('msg1.html')
  //    .setSandboxMode(HtmlService.SandboxMode.NATIVE)
  //    .setWidth(200)
  //    .setHeight(90);
  //SpreadsheetApp.getUi()
  //    .showModalDialog(htmlDlg, 'Wybierz wykonawcę');;
  //dataSheet = getDataSheet()
  var lastRow = dataSheet.getLastRow()
  for (var i = 2; i <= lastRow; i++){
    createDocument(i, false)
    Logger.log('Contract from row number ' + i + ' generated, mail has not been sent')
  }
}

function createDocument(rowNumber, sendMailBoolean){
  const templateDoc = getTemplateDoc()
  const destDir = getDocumentRepository()
  const dataSheet = getDataSheet()
  const textSheet = getTextSheet()
  var contract = getData(dataSheet.getRange(dataStartCol + rowNumber + ':' + dataEndCol + rowNumber).getValues())
  const newContract = templateDoc.makeCopy(contract.workerName + ' umowa B2B', destDir)
  var contractDoc = DocumentApp.openById(newContract.getId())
  var contractBody = contractDoc.getBody()
  var contractId = contractDoc.getId()
  var styleHeader = {
    [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER,
    [DocumentApp.Attribute.FONT_SIZE]: 12,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.SPACING_AFTER]: 10
  }
  var styleParagraph = {
    [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
    [DocumentApp.Attribute.FONT_SIZE]: 10,
    [DocumentApp.Attribute.BOLD]: false,
    [DocumentApp.Attribute.SPACING_AFTER]: 10
  }
  
  insertWorkerPositionTasks(contract.workerPosition, contractBody) 

  if (contract.minHourPerMonth != null){
    addListItem('{min_hour_per_month}', (textSheet.getRange(minHourPerMonthRange).getValues()).toString(), 1, contractBody)
    contractBody.replaceText('{min_hour_per_month_value}', contract.minHourPerMonth)
  }
  contractBody.replaceText('{min_hour_per_month}', '')    
  if (contract.maxHourPerMonth != null){ // nie istnieje?
    addListItem('{max_hour_per_month}', (textSheet.getRange(minHourPerMonthRange).getValues()).toString(), 1, contractBody)
    contractBody.replaceText('{max_hour_per_month_value}', contract.minHourPerMonth)
  }
  contractBody.replaceText('{max_hour_per_month}', '')
  if (contract.benefits == trueString) {
    addListItem('{benefits1}', (textSheet.getRange(benefits1Range).getValues()).toString(), 1, contractBody)
    addListItem('{benefits2}', (textSheet.getRange(benefits2Range).getValues()).toString(), 1, contractBody)
  }
  contractBody.replaceText('{benefits1}', '')
  contractBody.replaceText('{benefits2}', '')  
  if (contract.motivatingSystem == trueString)
    addListItem('{motivating_system}', (textSheet.getRange(motivatingSystemRange).getValues()).toString(), 1, contractBody)
  contractBody.replaceText('{motivating_system}', '')
  if (contract.salaryType == 'Godzinowa'){
    addListItem('{salary_type}', (textSheet.getRange(salaryTypeGodzinowaRange).getValues()).toString(), 1, contractBody)
    contractBody.replaceText('{salary_ratio_numeric}', contract.salaryRatioNumeric)
    contractBody.replaceText('{salary_ratio_words}', contract.salaryRatioWords)
    contractBody.replaceText('{salary_ratio_type}', contract.salaryRatioType)
    contractBody.replaceText('{salary_ratio_max_year_numeric}', contract.salaryRatioMaxYearNumeric)
    contractBody.replaceText('{salary_ratio_max_year_words}', contract.salaryRatioMaxYearWords)
  }
  else if (contract.salaryType == 'Ryczałtowa'){
    addListItem('{salary_type}', (textSheet.getRange(salaryTypeRyczaltRange).getValues()).toString(), 1, contractBody)
    contractBody.replaceText('{salary_const_numeric}', contract.salaryConstNumeric)
    contractBody.replaceText('{salary_const_words}', contract.salaryConstWords)
  }
  contractBody.replaceText('{salary_type}', '')
  if (contract.taxpayerVAT == trueString)
    addListItem('{taxpayer_VAT}', (textSheet.getRange(taxpayerVatRange).getValues()).toString(), 1, contractBody)
  contractBody.replaceText('{taxpayer_VAT}', '')      
  if (contract.timeType == 'Określony')
    contractBody.replaceText('{time_type}', 'okres ' + contract.timestamp + ' miesięcy')
  else if (contract.timeType == 'Nieokreślony')
    contractBody.replaceText('{time_type}', 'czas nieokreślony') 
  contractDoc.saveAndClose() // becasue API request can be send only to closed document
  customizeDoc(contractId, 1) //send requests via Google Docs API

  contractDoc = DocumentApp.openById(contractId)
  contractBody = contractDoc.getBody()
  contractBody.insertParagraph(0, textSheet.getRange(startTextRange[0]).getValues().toString()).setAttributes(styleHeader)
  contractBody.insertParagraph(1, textSheet.getRange(startTextRange[1]).getValues().toString()).setAttributes(styleParagraph)
  contractBody.insertParagraph(2, textSheet.getRange(startTextRange[2]).getValues().toString()).setAttributes(styleParagraph)
  contractBody.insertParagraph(3, textSheet.getRange(startTextRange[3]).getValues().toString()).setAttributes(styleParagraph)
  contractBody.insertParagraph(4, textSheet.getRange(startTextRange[4]).getValues().toString()).setAttributes(styleParagraph)
  contractBody.insertParagraph(5, textSheet.getRange(startTextRange[5]).getValues().toString()).setAttributes(styleParagraph)
  contractBody.insertParagraph(5, textSheet.getRange(startTextRange[6]).getValues().toString()).setAttributes(styleParagraph)

  insertBoldText(contractBody, textSheet, startTextRange)
  
  contractBody.replaceText('{start_date}', contract.startDate)
  contractBody.replaceText('{worker_name}', contract.workerName);
  contractBody.replaceText('{worker_city}', contract.workerCity);
  contractBody.replaceText('{worker_address}', contract.workerAddress);
  contractBody.replaceText('{worker_pesel}', contract.workerPesel);
  contractBody.replaceText('{worker_company_NIP}', contract.workerCompanyNIP);
  contractBody.replaceText('{worker_company_name}', contract.workerCompanyName)
  contractBody.replaceText('{worker_company_city}', contract.workerCompanyCity)
  contractBody.replaceText('{worker_company_address}', contract.workerCompanyAddress)
  contractBody.replaceText('{agent_name}', contract.agentName)
  contractBody.replaceText('{agent_phone}', contract.agentPhone)
  contractBody.replaceText('{agent_email}', contract.agentEmail)
  contractBody.replaceText('{period_of_notice}', contract.periodOfNotice)

  contractDoc.saveAndClose()
  customizeDoc(contractId, 0) //send requests via Google Docs API
  var url = saveDocOnGoogleDrive(contractDoc, dataSheet, rowNumber) 
  if (sendMailBoolean == true)
    sendMail(contract.responderEmail, 'Umowa B2B - ' + contract.workerName, DriveApp.getFileById(contractId))
  return url
}

// ##############################################################################################################################################################

function addListItem(textToFind, textToInsert, nestingLevel, body){
  var style = {
    [DocumentApp.Attribute.FONT_SIZE]: 10,
    [DocumentApp.Attribute.NESTING_LEVEL]: nestingLevel
  }
  
  var found = body.findText(textToFind)
  if (found) {
    var elem = found.getElement().getParent();
    var index = body.getChildIndex(elem);
    var listItem = body.insertListItem(index + 1, textToInsert).setGlyphType(DocumentApp.GlyphType.NUMBER)
    listItem.setAttributes(style);
   }
}

function saveDocOnGoogleDrive(doc, sheet, rowNumber) {
  const url = doc.getUrl()
  sheet.getRange(urlCol + rowNumber).setValue(url)
  Logger.log('Document has been saved in ' + url)
  return url
}

function sendMail(emailRecipent, subject, file){
  MailApp.sendEmail(emailRecipent, subject, 'Umowa B2B', {
    name: 'Generator umów',
    attachments: [file.getAs(MimeType.PDF)]
  })
  Logger.log('Document has been sent to ' + emailRecipent)
}

function customizeDoc(googleDocId, requestNum) {
  //googleDocId = '1MM9VeJ96cWn0vwly4mbEoWmn9ee52Ur9BEpwDuUs_qU'
  var requests = [Docs.newRequest(), Docs.newRequest()]
  requests[0] = {
    'updateParagraphStyle': {
      'range': {
        'startIndex': 1,
        'endIndex':  8000
      },
      'paragraphStyle': {
        'spaceAbove': {
          'magnitude': 10,
          'unit': 'PT'
        },
        'spaceBelow': {
          'magnitude': 5,
          'unit': 'PT'
        },
      },
      'fields': 'spaceAbove, spaceBelow'
    }
  };
  requests[1] = {
    'createParagraphBullets': {
      'bulletPreset': 'NUMBERED_DECIMAL_NESTED',
      'range': {
        'startIndex': 1,
        'endIndex': 2
      }
    }
  };
  Logger.log('Sent request: ' + requests[requestNum])
  var batchUpdateRequest = Docs.newBatchUpdateDocumentRequest()
  batchUpdateRequest.requests = requests[requestNum]
  Docs.Documents.batchUpdate(batchUpdateRequest, googleDocId)
}

function getNumberInWords(number){
  var number = 444
  var digits = ['', ' jeden', ' dwa', ' trzy', ' cztery', ' pięć', ' sześć', ' siedem', ' osiem', ' dziewięć']
  var tens = ['', ' jedenaście', ' dwanaście', ' trzynaście', ' czternaście', ' piętnaście', ' szesnaście', ' siedemnaście', ' osiemnaście', ' dziewiętnaście']
  var dozens = ['', ' dziesieć', ' dwadzieścia', ' trzydzieści', ' czterdzieści', ' pięćdziesiąt', ' sześćdziesiąt', ' siedemdziesiąt', ' osiemdziesiąt', ' dziewięćdziesiąt']
  var hundreds = ['', ' sto', ' dwieście', ' trzysta', ' czterysta', ' pięćset', ' sześćset', ' siedemset', ' osiemset', ' dziewięćset']
  var groups = [ ['', '', ''], [' tysiąc', ' tysiące', ' tysięcy'], [" milion" ," miliony" ," milionów"], [" miliard"," miliardy"," miliardów"],
                 [" bilion" ," biliony" ," bilionów"], [" biliard"," biliardy"," biliardów"], [" trylion"," tryliony"," trylionów"]];
  Logger.log(number)
  if (typeof number == 'number'){
    const logMessage = 'Number ' + number + ' converted in words: '
    var result = ''
    var char = ''
    if (number == 0)
      result = "zero";
    if (number < 0) {
      char = "minus ";
      number = -number;
    }
  
    var g = 0;
    while (number > 0) {
      var s = Math.floor((number % 1000)/100);
      var n = 0;
      var d = Math.floor((number % 100)/10);
      var j = Math.floor(number % 10);
    
      if (d == 1 && j>0) {
        n = j;
        d = 0;
        j = 0;
      }
    
      var k = 2;
      if (j == 1 && s + d + n == 0)
        k = 0;
      if (j == 2 || j == 3 || j == 4)
        k = 1
      if (s + d + n + j > 0)
        result = hundreds[s] + dozens[d] + tens[n] + digits[j] + groups[g][k] + result;
      g++;
      number = Math.floor(number/1000);
    }
    result = result.substring(1)
    if (result.startsWith('jeden tysiąc'))
      result = result.replace('jeden tysiąc', 'tysiąc')
    Logger.log(logMessage +  char + result)
    return char + result
  }
  else {
    Logger.log('"' + number + '" is not a number')
    return null
  }
}

function insertWorkerPositionTasks(workerPosition, body){
  const configSheet = getConfigSheet()
  const lastCol = configSheet.getLastColumn()
  const lastRow = configSheet.getLastRow()
  
  var formulas = configSheet.getRange('A1:BA2').getFormulas()
  for (var i in formulas) {
    for (var j in formulas[i]) {
      if (formulas[i][j].startsWith('=TRANSPOSE')){ // find location where "Stanowisko" is refered 
        Logger.log(formulas[i][j]);
        var tasksCol = +j + +1
        var tasksRow = +i + +1
        break
      }
    }
  }

  for (var i = tasksCol; i <= lastCol; i++){
    if (configSheet.getRange(getColumnToLetter(i)+tasksRow).getValues() == workerPosition){
      const workerTasks = configSheet.getRange(getColumnToLetter(i) + (+tasksRow + +1) + ':' + getColumnToLetter(i) + lastRow).getValues()
      workerTasks.forEach(function(elem){
        var found = body.findText('{worker_position}')
        if (found) {
          var foundElem = found.getElement().getParent()
          var index = body.getChildIndex(foundElem);
          if (elem != ''){
            Logger.log('Task added: ' + elem.toString())
            body.insertListItem(index+1, elem.toString()).setNestingLevel(2);
            index += 1
          } 
        }
      })
      break
    }
  }
  body.replaceText('{worker_position}', '')
}

function insertBoldText(body, sheet, textRange) { // find and replaces string between <boldStart> and <boldEnd>
  var regexPattern = /<boldStart>\s*(.*?)\s*<boldEnd>/g
  var boldStrings = []
  textRange.forEach(function(elem) {
    var value = sheet.getRange(elem).getValues().toString()
    var matches = []
    var match = regexPattern.exec(value)
    while (match != null) {
      matches.push(match[1]);
      match = regexPattern.exec(value);      
    }
    if (matches != null){
      for (var i = 0; i < matches.length; i++){
        var foundElement = body.findText(matches[i])
        var foundText = foundElement.getElement().asText();
        var start = foundElement.getStartOffset()
        var end = foundElement.getEndOffsetInclusive()
        foundText.setBold(start, end, true)
        boldStrings.push(matches[i])
      }      
    }
  })
  Logger.log('Bolded text in document: ' + boldStrings)
  body.replaceText('<boldStart>', '')
  body.replaceText('<boldEnd>', '')
}

function getColumnToLetter(column){
  var temp, letter = '';
  while (column > 0){
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter
}

function getLetterToColumn(letter){
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++){
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column
}

function getData(rows){
  let contract = new Contract(rows[0][0], rows[0][1], rows[0][2], rows[0][3], rows[0][4], rows[0][5], rows[0][6], rows[0][7], rows[0][8], rows[0][9], rows[0][10], 
                              rows[0][11], rows[0][12], rows[0][13], rows[0][14], rows[0][15], rows[0][16], rows[0][17], rows[0][18], rows[0][19], rows[0][20], 
                              rows[0][21], rows[0][22], rows[0][23], rows[0][24], rows[0][25], rows[0][26], rows[0][27], rows[0][28], rows[0][29])
  Logger.log('Loaded data: \n\t' + rows)
  return contract
}

function getDataSheet() {return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('')} // odpowiedzi z arkusza
function getConfigSheet() {return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('')} // odpowiedzi z arkusza
function getTextSheet() {return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('')} // odpowiedzi z arkusza
function getSheetUi() {return SpreadsheetApp.getUi()} 
function getTempaltedOC() {return DriveApp.getFileById('') } // umowa wzorocowa ze zmiennymi
function getDocumentRepository() {return DriveApp.getFolderById('')} // docelowy folder gdzie wpadają umowy
