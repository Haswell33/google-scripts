const trueString = 'Tak' // cale to temporaryco
const falseString = 'Nie'
const dataStartCol = 'B'
const dataEndCol = 'AA'
const urlCol = 'AB'
const startTextAnnexMarker = '{start_text_annex}'
const startTextContractMarker = '{start_text_contract}'
const annexTitle = 'Aneks do umowy B2B'
const contractTitle = 'Umowa B2B'

class Worker { // Wykonawca
  constructor(name, city, address, pesel){
    this.name = { // Imie i nazwisko (Pracownik)
      listItem: false,
      marker: '{worker_name}',
      value: name
    }                          
    this.city = { // Miejscowosc (Pracownik)
      listItem: false,
      marker: '{worker_city}',
      value: city 
    }                    
    this.address = { // Adres (Pracownik)
      listItem: false,
      marker: '{worker_address}',
      value: address
    }                    
    this.pesel = { // PESEL
      listItem: false,
      marker: '{worker_pesel}',
      value: pesel
    }            
  }
}
class Company { // Firma wykonawcy
  constructor(name, nip, city, address) {
    this.name = { // Nazwa (Firma) 
		  listItem: false,
		  marker: '{company_name}',
      value: name
    }	
    this.nip = { // NIP
		  listItem: false,
		  marker: '{company_nip}',
		  value: nip
	  }
	  this.city = { // Miejscowosc (Firma) 
		  listItem: false,
		  marker: '{company_city}',
		  value: city
	  }
	  this.address = { // Adres (Firma)
		  listItem: false,
		  marker: '{company_address}',
		  value: address
	  }
  }
}

class Agent {
  constructor(name, phone, email){
    this.name = { // Imie i nazwisko (Przedstawiciel) 
		  listItem: false,
		  marker: '{agent_name}',
		  value: name
    }                          
    this.phone = { // Nr telefonu
		  listItem: false,
		  marker: '{agent_phone}',
		  value: phone
    }
    this.email =  { // Adres e-mail
		  listItem: false,
		  marker: '{agent_email}',
		  value: email
    }  
  }
}
                                       
class Contract { // Warunki umowy
  constructor(maxHourPerMonth, benefits, motivatingSystem, taxpayerVat, position, periodOfNotice, salaryType, 
              salaryConstNumeric, salaryRatioNumeric, startDate,  timeType,  timestamp, minHourPerMonth, startDateActive){  
    this.periodOfNotice =  { // Okres wypowiedzenia
		  listItem: false,
		  marker: '{period_of_notice}',
		  value: periodOfNotice
    }
    this.startDate = { // Dzien zawarcia umowy
		  listItem: false,
		  marker: '{start_date}',
		  value: startDate.toLocaleDateString('pl-PL')
    }
	  this.timeType = { // Czas zawarcia umowy
		  listItem: false,
		  marker: '{time_type}',
		  value: timeType
    }
    this.benefits1 = { // Benefity
		  listItem: true,
		  marker: '{benefits1}',
		  value: benefits
    }
    this.benefits2 = {
      listItem: true,
      marker: '{benefits2}',
      value: benefits
    }
    this.motivatingSystem = { // Dedykowany system motywacyjny
		  listItem: true,
		  marker: '{motivating_system}',
		  value: motivatingSystem
    }
	  this.taxpayerVat = { // Podatnik VAT
		  listItem: true,
		  marker: '{taxpayer_vat}',
		  value: taxpayerVat
    }
    this.position = { // Stanowisko
      listItem: true,
      marker: '{worker_position}',
      value: 'position',
      name: position
    } 
    switch(startDateActive){
      case '':
        this.startDateActive = {
          listItem: true,
          marker: '{start_date_active}',
          value: false
        }
        break
      default:
        this.startDateActive = {
          listItem: true,
          marker: '{start_date_active}',
          value: true
        }
        this.startDateActiveValue = {
          listItem: false,
          marker: '{start_date_active_value}',
          value: startDateActive.toLocaleDateString('pl-PL')
        }
        break
    }
    switch(timeType){
      case 'Określony':
        this.timestamp = { // Ilosc miesięcy
			    listItem: false,
			    marker: '{timestamp}',
			    value: 'okres ' + timestamp + ' miesięcy'
		    }
        break
      case 'Nieokreślony':
        this.timestamp = {
          listItem: false,
          marker: '{timestamp}',
          value: 'czas nieokreślony'
        }
    }
    if (salaryType == 'Ryczałtowa'){
      this.salaryTypeRyczalt = { // Rodzaj stawki
		    listItem: true,
		    marker: '{salary_type_ryczalt}',
		    value: true
	    }
      this.salaryTypeGodzinowa = { // Rodzaj stawki
		    listItem: true,
		    marker: '{salary_type_godzinowa}',
		    value: false
	    }
		  this.salaryNumeric = { // Wysokosc wynagrodzenia (liczbowo) {ZL.GR} OR {ZL}
			  listItem: false,
			  marker: '{salary_numeric}',
			  value: salaryConstNumeric
		  }
      this.salaryWordsPLN = { // Wysokośc wynagrodzenia (slownie) zl
        listItem: false,
        marker: '{salary_words_pln}',
        value: getNumberInWords(salaryConstNumeric)
      }
      if ((salaryConstNumeric.toString()).split('.')[1] != null){
        this.salaryWordsPenny = { // Wysokosc wynagrodzenia (liczbowo) gr
          listItem: false,
          marker: '{salary_numeric_penny}',
          value: (salaryConstNumeric.toString()).split('.')[1]
        } 
      }
      else {
        this.salaryWordsPenny = {
          listItem: false,
          marker: '{salary_numeric_penny}',
          value: false
        }
      }   
    }
    else if (salaryType == 'Godzinowa'){
      this.salaryTypeGodzinowa = { // Rodzaj stawki
		    listItem: true,
		    marker: '{salary_type_godzinowa}',
		    value: true
      }
      this.salaryTypeRyczalt = { // Rodzaj stawki
		    listItem: true,
		    marker: '{salary_type_ryczalt}',
		    value: false
      }              
      this.salaryNumeric = { // Wysokość wynagrodzenia za godz. (liczbowo) {ZL.GR} OR {ZL}
        listItem: false,
        marker: '{salary_numeric}',
        value: salaryRatioNumeric
      }   
      this.salaryWordsPLN = {// Wysokość wynagrodzenia za godz. (slownie) zl
        listItem: false,
        marker: '{salary_words_pln}',
        value: getNumberInWords(salaryRatioNumeric)
      }
      if ((salaryRatioNumeric.toString()).split('.')[1] != null)
        this.salaryWordsPenny = { // Wysokosc wynagrodzenia (liczbowo) gr
          listItem: false,
          marker: '{salary_numeric_penny}',
          value: (salaryRatioNumeric.toString()).split('.')[1]
        } 
      if (maxHourPerMonth > 0){
        this.maxHourPerMonth = { // Maksymalna ilość godzin w miesiącu
          listItem: true,
          marker: '{max_hour_per_month}',
          value: true
        }
        this.maxHourPerMonthValue = {
          listItem: false,
          marker: '{max_hour_per_month_value}',
          value: maxHourPerMonth
        }
      }
      else {
        this.maxHourPerMonth = {
          listItem: true,
          marker: '{max_hour_per_month}',
          value: false
        }
      }    
      if (minHourPerMonth > 0){
        this.minHourPerMonth = { // Minimalna ilość godzin w miesiącu
          listItem: true,
          marker: '{min_hour_per_month}',
          value: true
        }
        this.minHourPerMonthValue = {
          listItem: false,
          marker: '{min_hour_per_month_value}',
          value: minHourPerMonth
        }
      }
      else {
        this.minHourPerMonth = {
          listItem: true,
          marker: '{min_hour_per_month}',
          value: false
        }
      }
    }
    this.pointsOrder = {
      listItem: false,
      marker: '{points_order}',
      value: getPointsOrder(benefits, motivatingSystem)
    }                          
  }
}

function onOpen(){
  const sheetUi = getSheetUi()
  const menuContract = sheetUi.createMenu('Wygeneruj umowę')
  menuContract.addItem('Wybierz wykonawcę', 'createContractSelect')
  menuContract.addItem('Generowanie zbiorcze', 'createContractBulk')
  menuContract.addItem('Szablon umowy', 'createContractTemplate')
  menuContract.addToUi()
  const menuAnnex = sheetUi.createMenu('Wygeneruj aneks')
  menuAnnex.addItem('Wybierz wykonawcę', 'createAnnexSelect')
  menuAnnex.addItem('Generowanie zbiorcze', 'createAnnexBulk')
  menuAnnex.addToUi()
}

function createContractLast() { // on FormSubmit
  const dataSheet = getDataSheet()
  const lastRow = dataSheet.getLastRow()
  createDocument(lastRow, true, contractTitle, startTextContractMarker)
}

function createContractSelect() {
  const sheetUi = getSheetUi()
  const results = getPersonFromSheet(sheetUi)
  const rowNumber = results[0]
  const workerName = results[1]
  switch(rowNumber){
    case null:
      console.error('not found + ' + workerName + ', creation aborted')
      sheetUi.alert('Nie znaleziono wskazanego wykonawcy, generowanie umowy zakończone niepowodzeniem')
      break
    default:
      console.info('found + ' + workerName + ' in row number ' + rowNumber)
      sheetUi.alert('Rozpoczęto generowanie umowy dla: ' + result)
      const url = createDocument(rowNumber, true, contractTitle, startTextContractMarker)
      sheetUi.alert('Generowanie umowy zakończone pomyślnie:\n' + url)
      break
  }
}

function createContractTemplate() {
  const sheetUi = getSheetUi()
  var html = HtmlService.createHtmlOutputFromFile('messageFailed.html')
      .setSandboxMode(HtmlService.SandboxMode.NATIVE)
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'My custom dialog');
  //dataSheet = getDataSheet()
  //createDocument(null, false, contractTitle, startTextContractMarker)
  sheetUi.alert('Generowanie szablonu umowy zakończone pomyślnie')
}

function createContractBulk() {
  const dataSheet = getDataSheet()
  const sheetUi = getSheetUi()
  //var htmlDlg = HtmlService.createHtmlOutputFromFile('msg1.html')
  //    .setSandboxMode(HtmlService.SandboxMode.NATIVE)
  //    .setWidth(200)
  //    .setHeight(90);
  //SpreadsheetApp.getUi()
  //    .showModalDialog(htmlDlg, 'Wybierz wykonawcę');;
  //dataSheet = getDataSheet()
  const lastRow = dataSheet.getLastRow()
  for (var i = 2; i <= lastRow; i++)
    createDocument(i, false, contractTitle, startTextContractMarker)
  sheetUi.alert('Zbiorcze generowanie umów zakończone')
}

function createAnnexSelect() {
  const sheetUi = getSheetUi()
  const results = getPersonFromSheet(sheetUi)
  const rowNumber = results[0]
  const workerName = results[1]
  switch(rowNumber){
    case null:
      console.error('not found + ' + workerName + ', creation aborted')
      sheetUi.alert('Nie znaleziono wskazanego wykonawcy, generowanie aneksu zakończone niepowodzeniem')
      break
    default:
      console.info('found + ' + workerName + ' in row number ' + rowNumber)
      sheetUi.alert('Rozpoczęto generowanie aneksu dla: ' + result)
      const url = createDocument(rowNumber, true, annexTitle, startTextAnnexMarker)
      sheetUi.alert('Generowanie aneksu zakończone pomyślnie:\n' + url)
      break
  }
}

function createAnnexBulk() {
  const dataSheet = getDataSheet()
  const sheetUi = getSheetUi()
  const lastRow = dataSheet.getLastRow()
  for (var i = 2; i <= lastRow; i++)
    createDocument(i, false , annexTitle, startTextAnnexMarker)
  sheetUi.alert('Zbiorcze generowanie aneksów zakończone')
}

function createDocument(rowNumber, sendMailBoolean, titleDoc, startTextType){
  rowNumber = 5
  startTextType = '{start_text_contract}'
  titleDoc = 'test'
  const templateDoc = getTemplateDoc()
  const destDir = getDocumentRepository()
  const data = getData(rowNumber)
  const worker = data[0]
  const company = data[1]
  const agent = data[2]
  const contract = data[3]
  const emailRecipent = data[4]
  console.info('START: generate document for ' + worker.name['value'])
  const docNew = templateDoc.makeCopy(worker.name['value'] + ' ' + titleDoc, destDir)
  var doc = DocumentApp.openById(docNew.getId())
  var docBody = doc.getBody()
  var docId = doc.getId()

  addListItems(contract, docBody)
  doc.saveAndClose() // API request can be send only to closed document
  customizeDocAPI(docId, 'createParagraphBullets') //send requests via Google Docs API 
  doc = DocumentApp.openById(docId)
  docBody = doc.getBody()
  insertStartText(startTextType, docBody)
  replaceMarkers(contract, docBody)
  replaceMarkers(worker, docBody)
  replaceMarkers(company, docBody)
  replaceMarkers(agent, docBody) 
  customizeDocApp(/<bold>\s*(.*?)\s*<.bold>/g, 'bold', docBody) // bold text between <bold></bold>
  customizeDocApp(/<header>\s*(.*?)\s*<.header>/g, 'header', docBody) // set as header text between <header></header
  doc.saveAndClose()
  customizeDocAPI(docId, 'updateParagraphStyle') //send requests via Google Docs API 
  var url = saveDocInDir(doc, rowNumber) 
  if (sendMailBoolean == true)
    sendMail(emailRecipent, titleDoc + ' - ' + worker.name, DriveApp.getFileById(docId), url)
  else
    console.info('MAIL: document has not been sent to ' + emailRecipent)
  return url
}

function saveDocInDir(doc, rowNumber) {
  const dataSheet = getDataSheet()
  const url = doc.getUrl()
  dataSheet.getRange(urlCol + rowNumber).setValue(url)
  console.info('SAVE: document has been saved in ' + url)
  return url
}

function sendMail(emailRecipent, title, file, url) {
  MailApp.sendEmail(emailRecipent, title, url, {
    name: 'Generator umów',
    attachments: [file.getAs(MimeType.PDF)]
  })
  console.info('MAIL: document has been sent to ' + emailRecipent)
}

function insertWorkerPositionTasks(workerPosition, body) {
  const configSheet = getConfigSheet()
  const lastCol = configSheet.getLastColumn()
  const lastRow = configSheet.getLastRow()
  var formulas = configSheet.getRange('A1:BA2').getFormulas()
  for (var i in formulas) {
    for (var j in formulas[i]) {
      if (formulas[i][j].startsWith('=TRANSPOSE')){ // find location where "Stanowisko" is refered 
        var tasksCol = +j + +1
        var tasksRow = +i + +1
        break
      }
    }
  }
  for (var i = tasksCol; i <= lastCol; i++){
    var positionName = configSheet.getRange(getColumnToLetter(i)+tasksRow).getValues()
    if (positionName == workerPosition){
      const workerTasks = configSheet.getRange(getColumnToLetter(i) + (+tasksRow + +1) + ':' + getColumnToLetter(i) + lastRow).getValues()
      workerTasks.forEach(function(elem){
        var found = body.findText('{worker_position}')
        if (found) {
          var foundElem = found.getElement().getParent()
          var index = body.getChildIndex(foundElem);
          if (elem != ''){
            console.info('EDIT: task "' + elem.toString() + '" added to document')
            body.insertListItem(index+1, elem.toString()).setNestingLevel(2);
            index += 1
          } 
        }
      })
      break
    }
  }
}

function insertStartText(startTextType, body) {
  const textSheet = getTextSheet()
  var styleParagraph = {
    [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
    [DocumentApp.Attribute.FONT_SIZE]: 10,
    [DocumentApp.Attribute.BOLD]: false,
    [DocumentApp.Attribute.SPACING_AFTER]: 10
  }
  var textLocation = getTextFromSheet(textSheet, startTextType)
  var textToInsert = ((textSheet.getRange(textLocation).getValues()).toString()).split('<eol>')
  for (var i = 0; i < textToInsert.length; i++)
    body.insertParagraph(i, textToInsert[i]).setAttributes(styleParagraph)
  console.info('EDIT: header ' + startTextType + ' added to document')
}

function replaceMarkers(data, body) {
  for (const [key, value] of Object.entries(data)) {
    if (value['listItem'] == false) {
      var found = body.findText(value['marker'])
      if (found) {
          body.replaceText(value['marker'], value['value'])
          console.info('REPLACE: found ' + value['marker'] + ' in document and replaced for "' + value['value'] + '"')
      }
      else
        console.error('REPLACE: not found ' + value['marker'] + ' in document, can not insert "' + value['value'] + '"')
    }
  }
}

function addListItems(data, body) {
  const textSheet = getTextSheet()
  var nestingLevel = 1
  var style = {
    [DocumentApp.Attribute.FONT_SIZE]: 10, 
    [DocumentApp.Attribute.NESTING_LEVEL]: nestingLevel
  }
  for (const [key, value] of Object.entries(data)) {
    if (value['listItem'] == true){
      switch(value['value']){
        case true: // if true, markers are selected to use, so add them
          var textToFind = value['marker']
          var textLocation = getTextFromSheet(textSheet, value['marker'])
          var textToInsert = (textSheet.getRange(textLocation).getValues()).toString()
          var found = body.findText(textToFind)
          if (found) {
            var elem = found.getElement().getParent()
            var index = body.getChildIndex(elem)
            var listItem = body.insertListItem(index + 1, textToInsert)
            listItem.setAttributes(style)
            body.replaceText(value['marker'], '')
            console.info('REPLACE: found ' + value['marker'] + ' in document and replaced for text located in ' + textLocation)
          }
          else
            console.error('REPLACE: not found ' + value['marker'] + ' in document, can not replace for text located in ' + textLocation)
          break
        case false: // if false, markers will be unused, so delete
          body.replaceText(value['marker'], '')
          console.info('DELETE: found ' + value['marker'] + ' and deleted from document')
          break
        case 'position':
          insertWorkerPositionTasks(value['name'], body)
          body.replaceText(value['marker'], '')
          break
        default: 
          console.warn('??? ' + key + '\n marker: ' + value['marker'] + '\n value:' + value['value'])  
          break     
      }
    }
  }
}

function customizeDocApp(regexPattern, operation, body) {
  var regex = regexPattern
  var docText = body.editAsText().getText()
  do {
    var match = regex.exec(docText)
    if (match){
      var foundElement = body.findText(match[1])
      switch(operation){
        case 'bold':  
          var foundText = foundElement.getElement().asText();
          var start = foundElement.getStartOffset()
          var end = foundElement.getEndOffsetInclusive()
          foundText.setBold(start, end, true)
          console.log('EDIT: set bold text "' + match[1] + '" in document')
          break
        case 'header':
          var foundParagraph = foundElement.getElement().getParent().asParagraph()
          var header = {
            [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER,
            [DocumentApp.Attribute.FONT_SIZE]: 12,
            [DocumentApp.Attribute.BOLD]: true,
            [DocumentApp.Attribute.SPACING_AFTER]: 10}
          foundParagraph.setAttributes(header)
          console.log('EDIT: set header "' + match[1] + '" in document')
          break
        default:
          console.error('EDIT: not found provided operation to customize, can not custom ' + match[i])
          break
      }
    } 
  } while (match)
  switch(operation){
    case 'bold':
      body.replaceText('<bold>', '')
      body.replaceText('</bold>', '')
      break
    case 'header':
      body.replaceText('<header>', '')
      body.replaceText('</header>', '')
      break
  }
}

function customizeDocAPI(googleDocId, requestTitle) {
  var request = Docs.newRequest()
  switch(requestTitle){
    case 'updateParagraphStyle':
      var request = {
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
      }
      break
    case 'createParagraphBullets':
      var request = {
        'createParagraphBullets': {
          'bulletPreset': 'NUMBERED_DECIMAL_NESTED',
          'range': {
            'startIndex': 1,
            'endIndex': 2
          }
        }
      }
      break
    default:
      console.error('EDIT: not found provided request title to customize')
      break
  }
  var batchUpdateRequest = Docs.newBatchUpdateDocumentRequest()
  batchUpdateRequest.requests = request
  console.info('EDIT: sent request: ' + batchUpdateRequest)
  Docs.Documents.batchUpdate(batchUpdateRequest, googleDocId)
}

function getPointsOrder(benefits, motivatingSystem) {
  if (motivatingSystem == true || benefits == true)
    return 'punktach 3.1 - 3.3'
  else if (motivatingSystem == false || benefits == false)
    return 'punkcie 3.1'
  else
    return 'punktach 3.1 - 3.2'
}

function getTextFromSheet(sheet, textToFind) {
  var result = sheet.createTextFinder(textToFind).findAll()
  var row
  var col
  result.forEach(function (elem){
    col = getColumnToLetter(elem.getColumn())
    row = elem.getRow()
  })
  return col + (+row + +1) //row below to get text
}

function getPersonFromSheet(sheetUi){
  const dataSheet = getDataSheet()
  var data = dataSheet.getDataRange().getValues()
  var result = sheetUi.prompt('Podaje imię i nazwisko');
  result =  result.getResponseText()
  for(var i = 1; i < data.length; i++)
    if (data[i][1] == result)
      return [i + 1, result]
}

function getColumnToLetter(column) {
  var temp, letter = ''
  while (column > 0){
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26
  }
  return letter
}

function getLetterToColumn(letter) {
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1)
  return column
}

function getNumberInWords(number) {
  var digits = ['', ' jeden', ' dwa', ' trzy', ' cztery', ' pięć', ' sześć', ' siedem', ' osiem', ' dziewięć']
  var tens = ['', ' jedenaście', ' dwanaście', ' trzynaście', ' czternaście', ' piętnaście', ' szesnaście', ' siedemnaście', ' osiemnaście', ' dziewiętnaście']
  var dozens = ['', ' dziesieć', ' dwadzieścia', ' trzydzieści', ' czterdzieści', ' pięćdziesiąt', ' sześćdziesiąt', ' siedemdziesiąt', ' osiemdziesiąt', ' dziewięćdziesiąt']
  var hundreds = ['', ' sto', ' dwieście', ' trzysta', ' czterysta', ' pięćset', ' sześćset', ' siedemset', ' osiemset', ' dziewięćset']
  var groups = [ ['', '', ''], [' tysiąc', ' tysiące', ' tysięcy'], [" milion" ," miliony" ," milionów"], [" miliard"," miliardy"," miliardów"],
                 [" bilion" ," biliony" ," bilionów"], [" biliard"," biliardy"," biliardów"], [" trylion"," tryliony"," trylionów"]];

  if (typeof number == 'number'){
    var result = ''
    var char = ''
    if (number == 0)
      result = "zero"
    if (number < 0) {
      char = "minus "
      number = -number
    }
  
    var g = 0;
    while (number > 0) {
      var s = Math.floor((number % 1000)/100)
      var n = 0;
      var d = Math.floor((number % 100)/10)
      var j = Math.floor(number % 10)
    
      if (d == 1 && j>0) {
        n = j
        d = 0
        j = 0
      }
    
      var k = 2;
      if (j == 1 && s + d + n == 0)
        k = 0;
      if (j == 2 || j == 3 || j == 4)
        k = 1
      if (s + d + n + j > 0)
        result = hundreds[s] + dozens[d] + tens[n] + digits[j] + groups[g][k] + result
      g++
      number = Math.floor(number/1000);
    }
    result = result.substring(1)
    if (result.startsWith('jeden tysiąc'))
      result = result.replace('jeden tysiąc', 'tysiąc')
    return char + result
  }
  else {
    console.error('can not convert number in words, ' + number + ' is not number')
    return null
  }
}

function getData(rowNumber) {
  switch(rowNumber){
    case null:
      //get testowe dane
      break
    default:
      const dataSheet = getDataSheet()
      var rows = dataSheet.getRange(dataStartCol + rowNumber + ':' + dataEndCol + rowNumber).getValues() 
      break
  }
  for (var i = 0; i < rows[0].length; i++){
    if (rows[0][i] == trueString)
      rows[0][i] = true
    else if (rows[0][i] == falseString)
      rows[0][i] = false
  }
  let worker = new Worker(rows[0][0], rows[0][1], rows[0][2], rows[0][3])
  let company = new Company(rows[0][4], rows[0][5], rows[0][6], rows[0][7])
  let agent = new Agent(rows[0][8], rows[0][9], rows[0][10])
  let contract = new Contract(rows[0][11], rows[0][12], rows[0][13], rows[0][14], rows[0][15], rows[0][16], rows[0][17], 
                              rows[0][18], rows[0][19], rows[0][20], rows[0][21], rows[0][22], rows[0][23], rows[0][24])
  return [worker, company, agent, contract, rows[0][25]]
}

function getDataSheet() {return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Contract Data')} // sheet with data about worker, contract etc.
function getConfigSheet() {return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config')} // sheet with current config
function getTextSheet() {return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Texts')} // sheet with texts to replace
function getSheetUi() {return SpreadsheetApp.getUi()} 
function getTemplateDoc() {return DriveApp.getFileById('doc-id') } // template contract
function getDocumentRepository() {return DriveApp.getFolderById('google-drive-folder-id')} // dest directory for new contracts
