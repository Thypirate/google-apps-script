
var ui = SpreadsheetApp.getUi();
var EMAIL_SENT = "SIM";
var ESTADO = "FEITO";
var DEST = DriveApp.getFolderById("0B3v25NWgY5XrVHRsUy1LMVFFTmM");//destination for the new document

function generateLetters() {

  var templateID = "1dPyF0xLx5CKyRkIxCHtQ0di1MDhfccT0_SejJiy2GHY";//id of the document template
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = ObjApp.rangeToObjects(ss.getDataRange().getValues());
  for (var i = 0; i < data.length; i++){
    var estado = data[i].estado;
    var emailEnviado = data[i].emailEnviado;
    var email = data[i].email;
    var message = "This is amazing";
    if (estado != ESTADO && emailEnviado != EMAIL_SENT) {
    var subject = "Sending emails from a Spreadsheet";
    MailApp.sendEmail(email, subject, message);
    sheet.getRange(2 + i, 13).setValue(ESTADO);
    sheet.getRange(2 + i, 14).setValue(EMAIL_SENT);
    SpreadsheetApp.flush();

    var date = Utilities.formatDate(new Date(), "GMT-1", "yyyy-MM-dd");//actual date
    var nameWithSpaces = date+"_"+data[i].unidade+"_"+"ESI"+"_"+"SI"+"_"+"carta_acolhimento_estagiarios"+"_"+data[i].nome+"_"+data[i].numero;
    var nameWithUnderscores = nameWithSpaces.toString().replace(/\s/g,"_");//replace spaces with underscores
    var doc = DriveApp.getFileById(templateID).makeCopy(nameWithUnderscores.toLowerCase(), DEST);//make a copy of the document template and place it to the destination folder
    var docID = doc.getId();

    var docActive = DocumentApp.openById(docID);
    var body = docActive.getBody();
    body.replaceText("%CURSO%", data[i].curso);
    body.replaceText("%NOME%", data[i].nome);
    body.replaceText("%ANO%", data[i].ano);
    body.replaceText("%LOCAL%", data[i].local);
    body.replaceText("%CONTACTO%", data[i].responsavel);
    body.replaceText("%FUNCAO%", data[i].funcao);
    body.replaceText("%VARIANTE%", data[i].funcao);
    body.replaceText("%UNIDADE%", data[i].unidade);
    body.replaceText("%DATA%", date);
    docActive.saveAndClose();
    ss.toast("Cartas criadas com sucesso !!!");
    }
  }
}

function onOpen(e) {
  ui.createMenu("Fazer cartas").addItem("Fazer cartas", "generateLetters").addToUi();
}
