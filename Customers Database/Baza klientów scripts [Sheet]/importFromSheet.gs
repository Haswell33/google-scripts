// import from sheet to google form as responses
function importFromSheet() {
  const googleFormOsobaID = '1h6S2iTtqLpVMNTGhAM1hrKr4yzNYzh6DiLwy42DjIZw';
  const googleFormFirmaID = '1J0NKweKL6C7NOeMHF60UZRGowZe1pI7zRhiDPkkB-sA';
  const googleFormRelacjeID = '1pMGRJl2RTZVkppE7Ny9EoxWwkCUM-fwc_MV0hs58h2M';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tmp_sheet_to_import');
  var data = sheet.getDataRange().getValues();
  for (i in data) {
    if (i != 0){
      sumbitResponsesToGoogleForm(googleFormRelacjeID, data[i])
      Utilities.sleep(8000); // aby inny skrypt nadazyl
    }
  }
}

function sumbitResponsesToGoogleForm(GOOGLE_FORM_ID, data){
  const googleForm = FormApp.openById(GOOGLE_FORM_ID);
  Logger.log(data)
  for (i in data){
    if (data[i] == ""){
      data[i] = 'NO_ANSWER'
    }
  }
  const questions = googleForm.getItems();
  const q0 = questions[0].asMultipleChoiceItem();
  const q1 = questions[1].asMultipleChoiceItem();
  /*const q2 = questions[2].asTextItem();
  const q3 = questions[3].asTextItem();
  const q4 = questions[4].asTextItem();
  const q5 = questions[5].asTextItem();
  const q6 = questions[6].asMultipleChoiceItem();
  const q7 = questions[7].asMultipleChoiceItem();
  const q8 = questions[8].asMultipleChoiceItem();
  const q9 = questions[9].asMultipleChoiceItem();
  const q10 = questions[10].asParagraphTextItem();
  const q11 = questions[11].asParagraphTextItem();*/
  const qR0 = q0.createResponse(data[0]);
  const qR1 = q1.createResponse(data[1]);
  /*const qR2 = q2.createResponse(data[2]);
  const qR3 = q3.createResponse(data[3]);
  const qR4 = q4.createResponse(data[4]);
  const qR5 = q5.createResponse(data[5]);
  const qR6 = q6.createResponse(data[6]);
  const qR7 = q7.createResponse(data[7]);
  const qR8 = q8.createResponse(data[8]);
  const qR9 = q9.createResponse(data[9]);
  const qR10 = q10.createResponse(data[10]);
  const qR11 = q11.createResponse(data[11]);*/
  const FormResponse = googleForm.createResponse();
  FormResponse.withItemResponse(qR0);
  FormResponse.withItemResponse(qR1);
 /* FormResponse.withItemResponse(qR2);
  FormResponse.withItemResponse(qR3);
  FormResponse.withItemResponse(qR4);
  FormResponse.withItemResponse(qR5);
  FormResponse.withItemResponse(qR6);
  FormResponse.withItemResponse(qR7);
  FormResponse.withItemResponse(qR8);
  FormResponse.withItemResponse(qR9);
  FormResponse.withItemResponse(qR10);
  FormResponse.withItemResponse(qR11);*/
  FormResponse.submit();
}
