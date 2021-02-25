
// Created By Abhinav Jain 



const RECIPIENT_COL  = "Recipient";
const EMAIL_SENT_COL = "Email Sent";

function onOpen() {

  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Advance")
    .addItem("Create Template", "createTemplate")
    .addItem("Mail Merge","sidebar")
    .addSubMenu(ui.createMenu("Email Validation")
      .addItem("Enter Emails", "Insertnewsheet")
      .addItem("Verify","emailValidation"))
    .addToUi();
}

function sidebar() {

  var html = HtmlService.createHtmlOutputFromFile("join").setTitle("Mail Merge");
  var ui = SpreadsheetApp.getUi();
  SpreadsheetApp.getActiveSpreadsheet().toast("Wait a minute", "Waiting");
  SpreadsheetApp.getActiveSpreadsheet().toast("Almost done");
  ui.showSidebar(html);
  Template();
}

function visualeditor() {
  var html = HtmlService.createHtmlOutputFromFile("visual").setWidth(800)
      .setHeight(562);
  var ui = SpreadsheetApp.getUi();
  SpreadsheetApp.getActiveSpreadsheet().toast("Wait a minute", "Waiting");

  SpreadsheetApp.getActiveSpreadsheet().toast("Almost done");
  ui.showModalDialog(html, "Desgin the content for your email");
}


function createTemplate() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D7').activate();
  spreadsheet.insertSheet(1);
  spreadsheet.getActiveSheet().setName('Summary');
  spreadsheet.insertSheet(2);
  spreadsheet.getActiveSheet().setName('Logs');
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Summary'), true);
  spreadsheet.getActiveRange().setFormula('=UNIQUE(Logs!B1:C)');
  spreadsheet.getRange('C1').activate()
  .setValue('Last Read');
  spreadsheet.getRange('D1').activate()
  .setValue('Number of Reads');
  spreadsheet.getRange('C2').activate()
  .setFormula('=MAXIFS(Logs!$A$2:$A,Logs!$B$2:$B,A2,Logs!$C$2:$C,B2)');
  spreadsheet.getRange('D2').activate()
  .setFormula('=COUNTIFS(Logs!$B$2:$B,A2,Logs!$C$2:$C,B2)');
  spreadsheet.getRange('C2').activate();
  spreadsheet.getActiveSheet().setColumnWidth(3, 148);
  spreadsheet.getActiveSheet().setColumnWidth(2, 178);
  spreadsheet.getActiveSheet().setColumnWidth(1, 125);
  spreadsheet.getActiveSheet().setColumnWidth(4, 115);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C2:C50'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('D2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('D2:D50'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('C52').activate();
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C:C').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('M/d/yyyy H:mm:ss');
  spreadsheet.getRange('D19').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Logs'), true);
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Summary'), true);
  spreadsheet.getRange('B1').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Logs'), true);
  spreadsheet.getRange('A1').activate()
  .setValue('Date');
  spreadsheet.getRange('B1').activate()
  .setValue('Subject');
  spreadsheet.getRange('C1').activate()
  .setValue('To');
  spreadsheet.getRange('A2').activate();
  Logger.log("Created By Abhinav Jain");
};


function Template() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().setValue('First Name');
  spreadsheet.getRange('B1').activate();
  spreadsheet.getCurrentCell().setValue('Last Name');
  spreadsheet.getRange('C1').activate();
  spreadsheet.getCurrentCell().setValue('Recipient');
  spreadsheet.getRange('D1').activate();
  spreadsheet.getCurrentCell().setValue('Email Sent');
  spreadsheet.getRange('A1:D1').activate();
  spreadsheet.getActiveRangeList().setFontWeight(null)
  .setFontWeight('bold')
  .setFontSize(11)
  .setFontColor('white')
  .setBackground('ACCENT1')
  .setBackground('ACCENT4');
};

function Insertnewsheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G6').activate();
  spreadsheet.insertSheet(1);
};

function clearEmailsentcol(colNumber, startRow){
  //colNumber is the numeric value of the colum
  //startRow is the number of the starting row

  var sheet = SpreadsheetApp.getActiveSheet();
  var numRows = sheet.getLastRow() - startRow + 1; // The number of row to clear
  var range = sheet.getRange(startRow, colNumber, numRows).activate();
  range.clear({contentsOnly: true});
}

function emailquota() {
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("one");
  SpreadsheetApp.getActiveSpreadsheet().toast("Remaining Daily email quota: " + emailQuotaRemaining, "Remaining Quota");
}
function toastSend() {
  SpreadsheetApp.getActiveSpreadsheet().toast("Select send email campagin type", "Invalid");
}
//function get() {
//  var countDraft = GmailApp.getDrafts();
//  var len = countDraft.length;
//  for (var i = 0; i < len; i++) {
//    var draft = GmailApp.getDrafts()[i]; // The first draft message in the drafts folder
//    var message = draft.getMessage();
//    Logger.log(message.getSubject());
//  }
//}
function startTimeTrigger() {
// Schedule Daily
  ScriptApp.newTrigger("sendEmails")
           .timeBased()
           .everyDays(1)
           .create();
//    ScriptApp.newTrigger("sendEmails")
//     .timeBased()
//     .everyMinutes(1)
//     .create();
//  Logger.log("One");
  SpreadsheetApp.getActiveSpreadsheet().toast("Your emails are a schedule for every day for a year.","Schedule Email");

};
// Created By Abhinav Jain 
function cancelTimeTrigger(){
  
  var triggers = ScriptApp.getProjectTriggers();
  
  for(var i = 0; i < triggers.length; i++){
    if(triggers[i].getTriggerSource() == ScriptApp.TriggerSource.CLOCK){
      ScriptApp.deleteTrigger(triggers[i]);
    };
  };
  SpreadsheetApp.getActiveSpreadsheet().toast("Your email schedule is deleted.","Schedule Email");
};
function submit(values) {
  var run = "run";
  var test = "test";
//  Logger.log(values);
  var val = values.Send;
//  Logger.log(val);
  if (val == run) {
    sendEmails();
  }
 else if (val == test) {
    testEmail();
  }
  else if(val == null){
    toastSend();
  }
}

function getTrackingGIF() {
  email = PropertiesService.getScriptProperties().getProperty('email');
  subject = PropertiesService.getScriptProperties().getProperty('subject');
  Logger.log(email,subject);
  // Create a url based on the Email Tracker Webhook web app's URL and attaching two URL paramaters 
  // that will pass the Subject and the To line of the email to the web app. Replace [WEBAPP URL] below with the URL of your web app

//  var imgURL = "https://script.google.com/macros/s/AKfycbxpAxLtsIhtBuJqbvhTDWe02N3AQNmApCa0gaz40Qo0k-rBtjbY/exec"
  var imgURL = "https://script.google.com/macros/s/AKfycbwQtPFF4Q8tkNkHDhN11O-IjTenlETHvmO8RyhA8Q/exec"
    // encode the Subject to assure that it will be passed properly as a part of a URL 
    + "?esubject=" + encodeURIComponent(subject.replace(/'/g, ""))
    // encode the To line to assure that it will be passed properly as a part of a URL
    + "&eto=" + encodeURIComponent(email);
  
  //Return an HTML tag for a 1x1 pixel image with the image source as the web app's URL
  return "<img src='" + imgURL + "' width='5' height='5' color='red'/>";
}

function getDraft(subjectLine, sheet=SpreadsheetApp.getActiveSheet()) {
//  var ui = DocumentApp.getUi();
  SpreadsheetApp.getActiveSpreadsheet().toast("Wait a minute", "Waiting");
  if (!subjectLine){
//    subjectLine = Browser.inputBox("Mail Merge", 
//                                      "Enter subject line of the Gmail " +
//                                      "draft message you would like to mail merge with:",
//                                      Browser.Buttons.OK_CANCEL);
    subjectLine = gmaildraft();
    
//    if (subjectLine === "cancel" || subjectLine == ""){
//    // if no subject line finish up
//      return; }
  }
//  PropertiesService.getScriptProperties().setProperty('subject', subjectLine);
//  return;
}
 
function sendEmails(subjectLine, sheet=SpreadsheetApp.getActiveSheet()) {
//
//  var spreadsheet = SpreadsheetApp.getActive();
//  spreadsheet.getRange('D2:D101').activate();
//  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

  SpreadsheetApp.getActiveSpreadsheet().toast("Wait a minute", "Waiting");
  subjectLine = PropertiesService.getScriptProperties().getProperty('subject');
//  subjectLine = "Happy Birthday to";
  SpreadsheetApp.getActiveSpreadsheet().toast(subjectLine);
  if (!subjectLine){
    subjectLine = Browser.inputBox("Mail Merge", 
                                      "Enter subject line of the Gmail " +
                                      "draft message you would like to mail merge with:",
                                      Browser.Buttons.OK_CANCEL);
                                      
    if (subjectLine === "cancel" || subjectLine == ""){ 
    // if no subject line finish up
    return;
    }
  }
  
// get the draft Gmail message to use as a template
  
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  
  // get the data from the passed sheet
  const dataRange = sheet.getDataRange();
  // Fetch displayed values for each row in the Range HT Andrew Roberts 
  // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const data = dataRange.getDisplayValues();

  // assuming row 1 contains our column headings
  const heads = data.shift(); 
  
  // get the index of column named 'Email Status' (Assume header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  
  // convert 2d array into object array
  // @see https://stackoverflow.com/a/22917499/1027723
  // for pretty version see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));
  
  var lastrow = sheet.getLastRow();
  var i = 49.0;
  //Clear Email Sent Column
  var k = PropertiesService.getScriptProperties().getProperty("rowSlot");
  if (k  >= lastrow){
    clearEmailsentcol(4, 2);
  }
  // used to record sent emails
  const out = [];
  
  // Created By Abhinav Jain 
  // loop through all the rows of data
  obj.forEach(function(row, rowIdx){
    // only send emails is email_sent cell is blank and not hidden by filter
    Logger.log(row, rowIdx);
    Logger.log(lastrow);
    if(rowIdx < lastrow) {
      var oldrow = PropertiesService.getScriptProperties().getProperty("rowSlot");;
      if(oldrow + 50 >= lastrow || oldrow < 0){
         oldrow = 0.0;
         PropertiesService.getScriptProperties().setProperty("rowSlot", oldrow);
      }
      oldrow = PropertiesService.getScriptProperties().getProperty("rowSlot");
      Logger.log("oldrow"+ oldrow);
      Logger.log(rowIdx - oldrow);
      if(rowIdx - oldrow <= i && !(rowIdx - oldrow < 0)) {
        Logger.log("d");
        if (row[EMAIL_SENT_COL] == ''){
          try {
            const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
            // @see https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
            // if you need to send emails with unicode/emoji characters change GmailApp for MailApp
            // Uncomment advanced parameters as needed (see docs for limitations)
            PropertiesService.getScriptProperties().setProperty('email', row[RECIPIENT_COL]);
            // Link image to html body
            const body = getTrackingGIF();
            Logger.log(msgObj);
        GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html + body,
          // bcc: 'a.bbc@email.com',
          // cc: 'a.cc@email.com',
          // from: 'an.alias@email.com',
          // name: 'name of the sender',
          // replyTo: 'a.reply@email.com',
          // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
          attachments: emailTemplate.attachments
        });
        // modify cell to record email sent date
            SpreadsheetApp.getActiveSpreadsheet().toast("Sending mail");
            out.push([new Date()]);
          } catch(e) {
        // modify cell to record error
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
   }
  }
  });
  
  var j = 0;
  j = PropertiesService.getScriptProperties().getProperty("rowSlot");
  Logger.log(typeof j);
  j = Number(j);
  Logger.log(typeof j);
  j += 50;
  PropertiesService.getScriptProperties().setProperty("rowSlot", j);
  Logger.log("j" + j);

  // updating the sheet with new data
  sheet.getRange(j - 48, emailSentColIdx+1, out.length).setValues(out);
  
  /**
   * Get a Gmail draft message by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} containing the subject, plain and html message body and attachments
  */
  function getGmailTemplateFromDrafts_(subject_line){
   
    try {
      // get drafts
      const drafts = GmailApp.getDrafts();
      Logger.log(drafts);
      // filter the drafts that match subject line
      const draft = drafts.filter(subjectFilter_(subject_line))[0];
       Logger.log(draft);
      // get the message object
      const msg = draft.getMessage();
       Logger.log(msg);
      // getting attachments so they can be included in the merge
      const attachments = msg.getAttachments();
       Logger.log(attachments);
      return {message: {subject: subject_line, text: msg.getPlainBody(), html:msg.getBody()}, 
              attachments: attachments};
    } catch(e) {
      throw new Error("Oops - can't find Gmail draft");
    }

    /**
     * Filter draft objects with the matching subject linemessage by matching the subject line.
     * @param {string} subject_line to search for draft message
     * @return {object} GmailDraft object
    */
    function subjectFilter_(subject_line){
      return function(element) {
        if (element.getMessage().getSubject() === subject_line) {
          return element;
        }
      }
    }
  }
  
  /**
   * Fill template string with data object
   * @see https://stackoverflow.com/a/378000/1027723
   * @param {string} template string containing {{}} markers which are replaced with data
   * @param {object} data object used to replace {{}} markers
   * @return {object} message replaced with data
  */
  function fillInTemplateFromObject_(template, data) {
    // we have two templates one for plain text and the html body
    // stringifing the object means we can do a global replace
    let template_string = JSON.stringify(template);

    // token replacement
    template_string = template_string.replace(/{{[^{}]+}}/g, key => {
      return data[key.replace(/[{}]+/g, "")] || "";
    });
    return  JSON.parse(template_string);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast("Campagin Complete");
}



function emailValidation()
{ 
  var apikey = 'da33077c93b5444286a91dfb5111628e'
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  count = sheet.getLastRow();
  var row_list = ss.getDataRange().getValues()
  if (count >1001)
  { Browser.msgBox("Maximum Email Address allowed is 1000")
  return;
  }
  else
  {
  payload = create_payload(row_list)
  result = get_result(payload,apikey)
  }
  if (result.getResponseCode() == 200) {
  var params = JSON.parse(result.getContentText());
  var result = params.gamalogic_emailid_vrfy;
  var array_result = [];
  var color = [];
  for (var i =0; i < count ; i++)
  {
   data_result = [];
   color[i] = new Array(8);
   data_result[0] = result[i]["do_you_mean"]
   data_result[1] = result[i]["is_role"]
   data_result[2] = result[i]["is_unknown"]
   data_result[3] = result[i]["is_valid"]
   data_result[4] = result[i]["is_syntax_valid"]
   data_result[5] = result[i]["is_catchall"]
   data_result[6] = result[i]["message"]
   data_result[7] = result[i]["is_disposable"]
   color[i][1] = "red";color[i][2] = "red";color[i][3] = "red";color[i][4]
   = "red";color[i][5] = "red";color[i][6] = "red";color[i][7] = "red";
   if (!data_result[3] && !data_result[2]) {color[i][1] = "red",color[i][2]
    = "red",color[i][3] = "red",color[i][4] = "red",color[i][5] = "red",color[i][6]
    = "red",color[i][7] = "red"}
   else if (data_result[5] && data_result[5]) {color[i][1] = "yellow",color[i][2]
    = "yellow",color[i][3] = "yellow",color[i][4] = "yellow",color[i][5] = "yellow",
    color[i][6] = "yellow",color[i][7] = "yellow" }
   else if (data_result[2]) {color[i][1] = "grey",color[i][2] = "grey",color[i][1] =
    "grey",color[i][3] = "grey",color[i][4] = "grey",color[i][5] = "grey",color[i][6]
    = "grey",color[i][7] = "grey"}
   else if (data_result[3] && !data_result[5]) {color[i][1] = "green",color[i][2]
    = "green",color[i][3] = "green",color[i][4] = "green",color[i][5] = "green",color[i][6]
    = "green",color[i][7] = "green"}
   array_result.push(data_result)
  }
  SpreadsheetApp.getActiveSpreadsheet().toast("Wait a minute");
  sheet.getRange(1, 2, count, 8).setValues(array_result).setBackgroundColors(color);
  sheet.insertRowBefore(1)
  var rows = sheet.getDataRange();
  var lr = sheet.getRange('A1:I1');
  lr.setBackground(null);
  lr = sheet.getRange('A2:I2');
  var head = new Array(8);
  head[0] = 'Email Address'
  head[1] = 'do_you_mean'
  head[2] = 'is_role'
  head[3] = 'is_unknown'
  head[4] = 'is_valid'
  head[5] = 'is_syntax_valid'
  head[6] = 'is_catchall'
  head[7] = 'message'
  head[8] = 'is_disposable'
  sheet.getRange(1, 1, 1, 9).setValues([head]).setFontWeight("bold");
  sheet.insertRowBefore(1)
  sheet.getRange(1,1).setValue(['Total number of email address']).setFontWeight("bold");
  sheet.getRange(1,2).setValue([count]);
  sheet.getRange(1,3).setValue(['Credits Balance']).setFontWeight("bold");
  getBalance(apikey);
  SpreadsheetApp.flush();
 }
 else{Browser.msgBox("Contact support@gamalogic.com")
  return;}
}
function getBalance(apikey) {
var response = UrlFetchApp.fetch("https://gamalogic.com/creditbalance/?apikey="+apikey);
var json = response.getContentText();
var data = JSON.parse(json);
var sheet = SpreadsheetApp.getActiveSheet();
sheet.getRange(1,4).setValue(data['Credit_Balance'])
}
function create_payload(row_list) {
var payload ={"gamalogic_emailid_vrfy": []}
for (var i =0; i < count ; i++)
{
payload["gamalogic_emailid_vrfy"].push({"emailid" : row_list[i][0] });
}
count_address = Object.keys(payload["gamalogic_emailid_vrfy"]).length
payload =JSON.stringify(payload)
return payload
}
function get_result(payload,apikey) {
var url = "https://gamalogic.com/bulkemailvrf/?apikey="+apikey
var options =
{
"method" : "GET",
"payload" : payload,
'contentType': 'application/json'
};
var result = UrlFetchApp.fetch(url, options);
SpreadsheetApp.getActiveSpreadsheet().toast("Almost Done");
return result
}


function saveEmailTemplate(data) {

  var forScope = GmailApp.getInboxUnreadCount(); // needed for auth scope
 
  var raw = "dddddd";

  var draftBody = Utilities.base64Encode(raw, Utilities.Charset.UTF_8).replace(/\//g,'_').replace(/\+/g,'-');

  var params = {
    method      : "post",
    contentType : "application/json",
    headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions:true,
    payload:JSON.stringify({
      "message": {
        "raw": draftBody
      }
    })
  };

  var resp = UrlFetchApp.fetch("https://www.googleapis.com/gmail/v1/users/me/drafts", params);
  Logger.log(resp.getContentText());
}
// Created By Abhinav Jain 