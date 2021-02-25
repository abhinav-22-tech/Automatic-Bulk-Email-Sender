function gmaildraft(){
  showDialog();
  var subject = PropertiesService.getScriptProperties().getProperty('subject');
  return subject;
}

function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('All Draft')
      .setWidth(417)
      .setHeight(170);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'Select Gmail Draft');
}

function getAllDrafts() {
//  Logger.log("Running");
  var i = 0;
  var d = new Array(10);
  GmailApp.getDrafts().forEach(draft => {
    var message = draft.getMessage();
    PropertiesService.getScriptProperties().setProperty('Draft Mail  '+ i, message.getSubject());
    d[i] = PropertiesService.getScriptProperties().getProperty('Draft Mail  ' + i);
    i++;
//    Logger.log(d[i] + "  D");
 });
 
 return d;
}

function submitDraftSubject(values){
   var value = PropertiesService.getScriptProperties().getProperty('Draft Mail  ' + values)
   PropertiesService.getScriptProperties().setProperty('subject', value);
   
   Logger.log(value);
   return value;
}

function saveEmailTemplate(data){
  Logger.log(data);
  var emailBody = data.emailBody;
//  var html = convertHtmlToPlain(emailBody);
  
  var emailSubject = data.emailSubject;
  PropertiesService.getScriptProperties().setProperty('subject', emailSubject);
  GmailApp.createDraft("", emailSubject, emailBody);
}
