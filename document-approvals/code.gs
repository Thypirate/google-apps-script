var ui = DocumentApp.getUi();

function onOpen(){

  ui
  .createMenu('Aprovação')
  .addItem('Pedido para aprovar/rejeitar', 'approveRejectRequest')
  .addItem('Adicionar aprovadores', 'startWorkflow')
  .addItem('Estado da aprovação', 'approvalStatus')
  .addToUi();

}

function startWorkflow() {

  var html = HtmlService.createTemplateFromFile('startWorkflow').evaluate()
  .setTitle('Adicionar aprovadores')
  .setWidth(300)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  ui.showSidebar(html);

}

function approveRejectRequest() {

  var html = HtmlService.createTemplateFromFile('approveRejectRequest').evaluate()
  .setTitle('Pedido para aprovar/rejeitar')
  .setWidth(300)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  ui.showSidebar(html);

}

function approvalStatus() {

  var html = HtmlService.createTemplateFromFile('approverStatus').evaluate()
  .setTitle('Estado da aprovação')
  .setWidth(300)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  ui.showSidebar(html);

}



function addApprovers(email) {

  var db = SpreadsheetApp.openById("1dr2DvUTSw8EBKw1j_Gju365TiCzCa7BUkPfxXx_PzI4");
  var ss = db.getActiveSheet();
  var nextRow = ss.getLastRow() + 1;
  var docID = DocumentApp.getActiveDocument().getId();
    ss.getRange(nextRow, 1).setValue(docID);
    ss.getRange(nextRow, 2).setValue(email);
    ss.getRange(nextRow, 3).setValue("null");
    ss.getRange(nextRow, 4).setValue("false");
    SpreadsheetApp.flush();

}



function getApprovers() {

  var db = SpreadsheetApp.openById("1dr2DvUTSw8EBKw1j_Gju365TiCzCa7BUkPfxXx_PzI4");
  var docID = DocumentApp.getActiveDocument().getId();
  var ss = db.getActiveSheet();
  var approvers = [];
  var data = ObjApp.rangeToObjects(ss.getDataRange().getValues());
  for(var i = 0; i < data.length; i++){
    approvers.push(data[i]);
    Logger.log(approvers[i].email);
  }

  return approvers;

}

function getApprover(email) {

  var db = SpreadsheetApp.openById("1dr2DvUTSw8EBKw1j_Gju365TiCzCa7BUkPfxXx_PzI4");
  var docID = DocumentApp.getActiveDocument().getId();
  var ss = db.getActiveSheet();
  var approvers = [];
  var data = ObjApp.rangeToObjects(ss.getDataRange().getValues());
  for(var i = 0; i < data.length; i++){
    if(email == data[i].email){
       approvers.push(data[i]);
      Logger.log(approvers[i].email);
      break;
    }
  }
  return approvers;
}


function start() {

  var db = SpreadsheetApp.openById("1dr2DvUTSw8EBKw1j_Gju365TiCzCa7BUkPfxXx_PzI4");
  var docID = DocumentApp.getActiveDocument().getId();
  var ss = db.getActiveSheet();
  var data = ObjApp.rangeToObjects(ss.getDataRange().getValues());

  for(var i = 0; i < data.length; i++){
    var email = data[i].email;
    var estado = data[i].estado;
    var docid = data[i].docid;

    if(estado == "null" && docID == docid) {
    var doc = DocumentApp.getActiveDocument();
    doc.addEditor(email);
    var url = doc.getUrl();
      MailApp.sendEmail({
        to: email,
        subject: "Por favor valide o documento",
        htmlBody: "O utilizador foi pedido para validar e aprovar um documento<br>"+
        'Por favor <a href="'+url+'">clique aqui</a> para abrir o documento. '+
        '<br><br>Depois de validar clique no menu de aprovação e selecione a opção de pedido para '+
        'aprovar/rejeitar',
      });
      ss.getRange(2 + i, 4).setValue("Email Sent");
       SpreadsheetApp.flush();

    }
  }
}

function setApproverStatus(email, status){

  var db = SpreadsheetApp.openById("1dr2DvUTSw8EBKw1j_Gju365TiCzCa7BUkPfxXx_PzI4");
  var docID = DocumentApp.getActiveDocument().getId();
  var ss = db.getActiveSheet();
  var data = ObjApp.rangeToObjects(ss.getDataRange().getValues());

  for(var i = 0; i < data.length; i++){

    if(email == data[i].email && docID == data[i].docid){

      ss.getRange(2 + i, 3).setValue(status);
      SpreadsheetApp.flush();
    }
  }
}

function sendReminder(approverEmail){
  MailApp.sendEmail({
    to: approverEmail,
    subject: "Por favor valide o documento",
    htmlBody: "O utilizador foi pedido para validar e aprovar um documento<br>"+
        'Por favor <a href="'+url+'">clique aqui</a> para abrir o documento. '+
        '<br><br>Depois de validar clique no menu de aprovação e selecione a opção de pedido para '+
        'aprovar/rejeitar',
      });
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function getSession() {//get the current user session

  return Session.getEffectiveUser().getEmail();

}
