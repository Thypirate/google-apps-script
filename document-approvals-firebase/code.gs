var ui = DocumentApp.getUi();
var firebaseUrl ="https://spike-40e0b.firebaseio.com/";
var secret = "2pxgw328lpTdQMbhqD0x1slg9Rank9ZC6yGSwZvu";
var base = FirebaseApp.getDatabaseByUrl(firebaseUrl,secret);

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
  var docID = DocumentApp.getActiveDocument().getId();
  var approvers = {"docid": docID, "email": email, "estado": "null", "emailEnviado": "false" };
  var data = base.pushData("approvers", approvers);

}


function getApprovers() {
  var approvers = [];
  var data = base.getData("approvers");
  for(var i in data) {
    approvers.push(data[i]);
    Logger.log(data[i].email);
  }
  return approvers;
}


function getApprovers(email) {
  var approvers = [];
  var data = base.getData("approvers");
  for(var i in data) {
    if(email == data[i].email){
    approvers.push(data[i]);
    Logger.log(data[i].email);
    break;
  }
}
  return approvers;
}


function start() {
  var docID = DocumentApp.getActiveDocument().getId();
  var data = base.getData("approvers");
  for(var i in data) {
    var email = data[i].email;
    var estado = data[i].estado;
    var docid = data[i].docid;

    if(estado == "null" && docID == docid){
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
  }
}
}

function setApproverStatus(email, status){
  var docID = DocumentApp.getActiveDocument().getId();
  var data = base.getData("approvers");
  for(var i in data) {
    if(email == data[i].email && docID == data[i].docid){
      base.setData("approvers/"+i+"/estado", status);
  }
}
}


function sendReminder(approverEmail){
  var doc = DocumentApp.getActiveDocument();
  var url = doc.getUrl();
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
