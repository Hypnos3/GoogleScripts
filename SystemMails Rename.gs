/*
    !! Please activate only if you know what you doing !! 
 
 https://gist.github.com/mogsdad/6515581
 https://gist.github.com/mogsdad
 
*/

/*
Used Librarys:
https://sites.google.com/site/scriptsexamples/custom-methods/betterlog

https://sites.google.com/site/scriptsexamples/custom-methods/underscoregs

useful links:
https://sites.google.com/site/scriptsexamples/home

utilitys:
https://gist.github.com/mogsdad/6515581

*/

// startup function
//function doGet() {
//  return HtmlService.createHtmlOutputFromFile('index')
//      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
//}


// The script will check your Gmail mailbox every minute
// with the help of a CLOCK based trigger.
function Setup_Rename_SystemMails()
{

  //Cleanup Trigger:  
  var triggers = ScriptApp.getProjectTriggers();
  
  for(var i in triggers) {
    if (triggers[i].getHandlerFunction() == 'Exec_Gmail_Rename_RequestExpress')
    {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    else if (triggers[i].getHandlerFunction() == 'Exec_Gmail_Rename_SystemMails')
    {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  //Setup Trigger:  
  ScriptApp.newTrigger('Exec_Gmail_Rename_SystemMails')
   .timeBased()
   .everyMinutes(5)
   .create();

  var userMail = Session.getActiveUser().getEmail();
  
  //Setup Propertys:  
  var userProperties = PropertiesService.getUserProperties();
  
  //userProperties.deleteAllProperties();
  userProperties.setProperty('MailDoneLabelName', 'z/processed mail');
  userProperties.setProperty('MailErrorLabelName', 'z/error process mail');
  userProperties.setProperty('LogFileTable', '1J5aEPKiQ1zSQMbANr13kCvaqhx9_NOTzB1Ub8rDhnb0'); //'1J5aEPKiQ1zSQMbANr13kCvaqhx9_NOTzB1Ub8rDhnb0' https://drive.google.com/open?id=1J5aEPKiQ1zSQMbANr13kCvaqhx9_NOTzB1Ub8rDhnb0&authuser=0
  userProperties.setProperty('LogFileTableTab', 'Rename_SystemMails_' + userMail);

  userProperties.setProperty('TestMode', false); //false
  userProperties.setProperty('LogFileLevel', 'INFO'); //The new log level e.g. "OFF","SEVERE","WARNING","INFO","CONFIG","FINE","FINER","FINEST" or "ALL".  
  
  //Test Script
  Exec_Gmail_Rename_SystemMails();
}
 
// This will auto-save the image attachments from Gmail to Google Drive
function Exec_Gmail_Rename_SystemMails()
{
  var userProperties = PropertiesService.getUserProperties();
  //var data = userProperties.getProperties();
  var testMode = (userProperties.getProperty('TestMode')=== 'true');
  var labelNameProcessed = userProperties.getProperty('MailDoneLabelName'); //"z/processed mail";
  var labelNameError = userProperties.getProperty('MailErrorLabelName'); //"z/error process mail";
  
  //https://sites.google.com/site/scriptsexamples/custom-methods/betterlog
  //Best practice for using BetterLog and logging to a spreadsheet:
  //The new log level e.g. "OFF","SEVERE","WARNING","INFO","CONFIG","FINE","FINER","FINEST" or "ALL".
  Logger = BetterLog.setLevel(userProperties.getProperty('LogFileLevel'),userProperties.getProperty('LogFileTableTab')).useSpreadsheet(userProperties.getProperty('LogFileTable'),userProperties.getProperty('LogFileTableTab')); //automatically rolls over at 50,000 rows

  var labelProcessed = GmailApp.getUserLabelByName(labelNameProcessed);   
  if ( ! labelProcessed ) {
      labelProcessed = GmailApp.createLabel(labelNameProcessed);
  }
  labelNameProcessed = labelNameProcessed.replace(/\s|\\|\//g, '-').toLowerCase();
  
  var labelError = GmailApp.getUserLabelByName(labelNameError);   
  if ( ! labelError ) {
      labelError = GmailApp.createLabel(labelNameError);
  }  
  labelNameError  = labelNameError.replace(/\s|\\|\//g, '-').toLowerCase(); 
  
  var searchQuery = "from:(request.application@globalfoundries.com OR globalfoundries@service-now.com OR Request Application OR ServiceNow) is:unread -subject:{redirect} -label:" + labelNameProcessed + " -label:" + labelNameError + " -has:attachment";
   
  var threads = GmailApp.search(searchQuery);
  var cnt = threads.length;
  if (cnt >10)
  {
    threads = GmailApp.search(searchQuery,cnt-10,10);
  }
  Logger.fine("Start!! found " + cnt + " threads, searching in the last " + threads.length);
  Logger.finest("searchQuery='" + searchQuery + "'");
    
  for (var x=threads.length-1; x>=0; x--)
  {  
    var messages = threads[x].getMessages();
    threads[x].addLabel(labelProcessed);    
    for (var y=0; y<messages.length; y++)
    {          
      try
      {		
        processMessageGeneric_(messages[y], testMode, labelProcessed, labelNameProcessed);      
      }catch(err){        
       	Logger.severe('Error occured!! - ' + err + ' - (line=' + err.lineNumber + ')');
       	messages[y].markUnread();
     	threads[x].addLabel(labelError);
        threads[x].removeLabel(labelProcessed);
      }        
    }
  }
  if (threads.length >0)
    Logger.finer("End!!");
}

function processMessageGeneric_(message, testMode, excludeLabel, excludeLabelName)
{
  //var message = messages[y];
  var originalSubject = message.getSubject();
  var bodyPlain = message.getPlainBody();
  var originalSender = message.getFrom();
  var date = message.getDate();
  var messageId = message.getId();
  
  var sender = Session.getActiveUser().getEmail();
  Logger.finer("--------------------------> " + originalSender + " Subject=" + originalSubject);

  var subjects = [{}];
  subjects = originalSubject.split(" ");
  
  if (subjects && subjects.length > 2)
  {
    var orginalSenderID = originalSender.replace(/<\S+>/g, '').trim();
    var subject = formatSubject_(originalSubject.replace(new RegExp(RequestID, 'ig'), ''));
    if (!orginalSenderID)
    {
      orginalSenderID = originalSender;
    }
    orginalSenderID = orginalSenderID.replace(/[()@<>.]/g, '').replace(/globalfoundriescom/gi, '').trim();

    var RequestID = subjects[0].replace(/\s+/g, "");
    var RequestTyp = RequestID.replace(/\d/g, '');
    //Logger.log("RequestID='" + RequestID+ " RequestTyp=" + RequestTyp);

    var replyMessage;
    if (originalSubject.toLowerCase().indexOf("submsr") > -1 || originalSubject.toLowerCase().indexOf("subdomain") > -1 || originalSubject.toLowerCase().indexOf(" sub ") > -1)
    {
      subject = subject.replace(new RegExp(" - ", 'g'), ' ').replace(new RegExp(" to ", 'ig'), ' ').trim();
      subject = subject.replace(new RegExp(" all ", 'ig'), ' ').replace(new RegExp(" no ", 'ig'), ' ').replace(new RegExp(" for ", 'ig'), ' ');
      subject = subject.replace(new RegExp(" in ", 'ig'), ' ').replace(new RegExp(" on ", 'ig'), ' ').replace(new RegExp(" to ", 'ig'), ' ');
      subject = subject.replace("CR ", ' ').replace("/", ' ').replace(",", '');

      var toReplace = ["Customer Review","Not Feasible","Design Complete","Design Ready Review","Pilot Moved Production","Removed from Production",
                       "Pilot Testing","Successful","Production Rollout entities","Ready Release","Waiting Equipment","ready activation wait",
                       "Required","Development","Hold","Research","Resubmitted","Rework","Rollout","Rollback","Paused","Progress", 
                       "eitech","ops", "fsc", "eidev", "InSignOff", "Effective"];
      var index;
      for (index = 0; index < toReplace.length; ++index) {
        subject = subject.replace(new RegExp(toReplace[index], 'i'), '');        
      }
      
      subject = subject.trim();
      Logger.finer("SubRequest identified: " + subject);
      
      if (originalSubject.toLowerCase().indexOf("...") > -1 && subject.length > 30)
      {
        subject = subject.replace("...", ' ').trim();
        var lastIndex = subject.lastIndexOf(" ")
        subject = subject.substring(0, lastIndex);
        Logger.finer("shorten subject " + subject);
      }

      replyMessage = GetMessage_("from:(" + sender + " OR " + originalSender + " OR Request Application OR ServiceNow) subject:(" + subject + ") -subject:((Out of office) OR OOO)", messageId, date, originalSubject);
      //Logger.log("replyMessage is= " + replyMessage.getSubject());
    }
    
    if (!replyMessage)
    {
      replyMessage = GetMessage_("((from:(" + originalSender + ") " + RequestID + ") OR (" + RequestID + " " + RequestTyp + "ˣ  -subject:((Out of office) OR OOO)", messageId, date, originalSubject);
      //"from:(" + originalSender + ") " + RequestID + " -subject:((Out of office) OR OOO)"
      // "from:(noreply@globalfoundries.com) " + RequestExpressID;
      //((from:(Request Application <request.application@globalfoundries.com>) MSR982114) OR (MSR982114 MSRˣ)) -subject:((Out of office) OR OOO) 
    }
    
    var bodyHtml = message.getBody();
    bodyHtml = bodyHtml.replace(' width="90%"',''); //.replace(' align="left"','');


    //Logger.log("orginalSenderID is= " + orginalSenderID);
    var TextContent = parseMessage_(bodyHtml,Session.getActiveUser().getEmail());
    
    var headerline = "<h2>" + originalSubject + "</h2>";
    //-----------------------------------------------------------------------------------------------------
    var HtmlAddtable = "<small><code><table border='0' align='left' width='90%' ><tbody><tr>" + 
            "<tr><td colspan='2'>--------------------------------------------------------------------------------------</td></tr>" +
              "<tr><td>original subject:</td><td>" + originalSubject + "</td></tr>" +
              "<tr><td>original Message ID:</td><td>" + message.getId() + "</td></tr>" +
              "<tr><td>search:</td><td>" + RequestID + " - " + RequestTyp + " - " + RequestID.replace(RequestTyp, '') + " - " + TextContent.txt + "</td></tr>" + 
              "</tbody></table></code></small>";
    
    if (TextContent.TslLnk)
    {
      HtmlAddtable = HtmlAddtable + '<tr><td>Links:</td><td><a href="http://myteamsdrs.gfoundries.com/sites/Fab36_FA/300mmModelling/ToolStartUpCheckListe/Forms/AllItems.aspx">Tool Startup Liste</a> <a href="//file:///G:/Ops-FA/ToolStartUpCheckList/GF.TSL-Launcher/GF.TSL-Launcher.application">Tool Startup Listen Tool</a></td></tr>';      
    }
    if (TextContent.DomainLnk)
    {
      HtmlAddtable = HtmlAddtable + '<tr><td>Links:</td><td><a href="http://myteamsfap/sites/pm/Shared%20Documents/MSR-Help.mht">MSR based on Domain oriented Demand Management</a></td></tr>';
    }    
    HtmlAddtable = HtmlAddtable + "</tbody></table></code></small>"
  
    
                //"<tr><td>auto forwarded from:</td><td>" + originalSender + "</td></tr>" +
                //  "<tr><td>when:</td><td>" + date + "</td></tr>" +
                //    "<tr><td>to:</td><td>" + message.getTo() + "</td></tr>" +
                //      "<tr><td>cc:</td><td>" + message.getCc() + "</td></tr>" +
                //"<tr><td>Filter:</td><td>" + RequestTyp +" " + orginalSenderID.replace(/\s+/g, '') + " </td></tr>" + 
    
    bodyHtml = headerline + "<br>--------------------------------------------------------------------------------------<br>" + 
    bodyHtml + HtmlAddtable;
    
    Logger.finest("body:" + bodyHtml);
      
    /*
    bodyPlain = headerline + "\r\n--------------------------------------------------------------------------------------\r\n" +
      bodyPlain +
        "\r\n--------------------------------------------------------------------------------------\r\n\r\n" +
          "original subject: " + originalSubject + 
            "auto forwarded from: " + originalSender + "\r\n" +
              "when: " + date + "\r\n" +
                "to: " + message.getTo() + "\r\n" +
                  "cc: " + message.getCc() + "\r\n" +
                    "original ID: " + message.getId() + "\r\n" +
                      " Filter Text: " + RequestID + " " + orginalSenderID.replace(/\s+/g, '');
    Logger.finest("body plain:" + bodyPlain);
    */
    
    if (testMode)
    {  
      return;
    }
    
    if (replyMessage)
    {
      var referenceID = replyMessage.getId(); //not working: replyMessage.getThread().getId()
      var labels = replyMessage.getThread().getLabels();
      //labels.push(excludeLabel);
            
            
      	var rawData= message.getRawContent();
      	rawData = rawData.replace(' width="90%"','');
      	rawData = replaceText_(rawData,'<body','>','<body>' +  headerline + "<br>--------------------------------------------------------------------------------------<br>\r\n");
        rawData = replaceText_(rawData, 'Subject:','\n','Subject: ' + replyMessage.getSubject() + '\r\n');
        rawData = replaceText_(rawData, 'References:','\n','References: ' + referenceID + '\r\n', 'Subject: ');
        rawData = replaceText_(rawData, 'In-Reply-To:','\n','In-Reply-To: ' + referenceID + '\r\n', 'Subject: ');
      	rawData = rawData.replace('</body>', HtmlAddtable+'</body>' );

        var resp = createMessage_(rawData, referenceID, labels); 
      
      	var newMessage = GmailApp.getMessageById(resp.id);
      	newMessage.markUnread();
      	newMessage.getThread().addLabel(excludeLabel);
      	Logger.info("created mail in thread; mail id '" + resp.id + "' subject='" + newMessage.getSubject() + "' tothread=" + referenceID);
      	Utilities.sleep(1000);
          
      	message.markRead();
      	message.moveToTrash();           
    }
    else
    {
      	var rawData= message.getRawContent();
      	rawData = rawData.replace(' width="90%"','');
      	rawData = replaceText_(rawData,'<body','>','<body>' +  headerline + "<br>--------------------------------------------------------------------------------------<br>\r\n");
        rawData = replaceText_(rawData, 'Subject:','\n','Subject: [' + RequestID + '] ' + subject + '\r\n');
      	rawData = rawData.replace('</body>', HtmlAddtable+'</body>' );      
      	var resp = createMessage_(rawData );
      	var newMessage = GmailApp.getMessageById(resp.id);
      	newMessage.markUnread();
      	newMessage.getThread().addLabel(excludeLabel);
        Logger.info("created new mail; no other Message found; mail id '" + resp.id + "' subject='" + newMessage.getSubject());
      	Utilities.sleep(1000);
          
      	message.markRead();
      	message.moveToTrash();   
      
      return;      
    }
  }
}

function formatSubject_(subject)
{
  var ok = true;
  var toReplace = ["msr","to","for","sub","domain","open","opened","closed","approved","disapproved","approval","waiting","paused","submitted","started","completed","complete",
                   "incomplete","comment","assigned","reassigned","resolved","resubmitted", "decision", "pending", "cancelled","successful","unsuccessful","handover", "prod", "new status"];
                  
  while (subject.length >= 4 && ok)   
  {
    if (!isNaN(parseInt((subject[0]))) ||
        (subject[0] === ' ') ||
      (subject[0] === '-') || 
      (subject[0] === '/') || 
            (subject[0] === '\\') || 
              (subject[0] === '}') || 
                (subject[0] === ']') ||
                  (subject[0] === ':'))
    {
      subject = subject.substring(1).trim(); 
    }
    else if (subject[0] === '{')    
    {
      subject = subject.substring(1).replace('/}',' ').trim(); 
    }
    else if (subject[0] === '[')
    {
      subject = subject.substring(1).replace('/]',' ').trim(); 
    }
    else
    {
      ok = false;
      for (index = 0; index < toReplace.length; ++index) {
        if (subject.toLowerCase().indexOf(toReplace[index]) == 0)
        {
            subject = subject.substr(toReplace[index].length).trim();     
          	ok = true;
        }
      }
    }
    //Logger.finest("subject change > " + subject);
  }
  return subject;
}

function cleanWord_(text, toremove)
{  
  if (subject.toLowerCase().indexOf(toremove) == 0)
    {
      subject = subject.substr(toremove.length).trim();     
    }
  return text;
}

function GetMessage_(searchQuery, excludeID, date, LogEnhancement)
{  
  var threads = GmailApp.search(searchQuery);
  var cnt = threads.length;
  Utilities.sleep(500);
    
  for (var x=cnt-1; x>=0; x--)
  {  
    var messages = threads[x].getMessages();
    for (var y=0; y<messages.length; y++)
    { 
/*
      Logger.log("message Date:" + messages[y].getDate() + " valueOf=" +messages[y].getDate().valueOf());
      Logger.log("correlation Date:" + date + " valueOf=" + date.valueOf());
      Logger.log(messages[y].getDate().valueOf() < date.valueOf());
*/

      if (messages[y].getId() !== excludeID && messages[y].getDate().valueOf() < date.valueOf())
      {
        Logger.finer("found " + cnt + " mails, used " + y + " message=" + messages[y].getId() + " of query: '" + searchQuery + "' - '" + LogEnhancement + "'");
        return messages[y];
      }
    }
  }
  Logger.finer("no mail found for query: '" + searchQuery + "' - '" + LogEnhancement + "' result= " + cnt);
  return;
}


function parseMessage_(bodyhtml,requestor)
{
  requestor = requestor.toLowerCase();
  var addText = "";
  var addTslLink = false;
  var addDomainLink = false;
  var addLinks = "";
  var addDetails = "";
  var dataArr = bodyhtml.split('<tr>');
  var customer = false;
    
  //Logger.log(dataArr);
  for (var i = 0; i < dataArr.length; i++)
  {
    var data = dataArr[i];
    if (data.indexOf("Details:") > -1 )
    {
      addDetails = (addDetails + " " + data.replace('Details:','').replace('<td>','').replace('</td>','')).trim();
    } 
    else if (data.indexOf("New Comment:") > -1 )
    {
      addDetails = (addDetails + " " + data.match("<pre>(.*)</pre>")).trim();
    } 
    else if (data.indexOf("You are listed in the Customer role.") > -1 )
    {
      //The following event occurred: EITECH: ready for activation/wait for OPS Approval. You are listed in the Customer role.
      customer = true;
    } 
    else if (data.indexOf("Requestor") >0 && data.toLowerCase().indexOf(requestor) > -1 )
    {
      addText = "MeAsRequestor " + addText; 
      customer = false;
    } 
    else if (data.indexOf("Assigned To") >0 && data.toLowerCase().indexOf(requestor) > -1 )
    {
      addText = "AssignedTo " + addText;
      customer = false;
    } 
    else if (data.indexOf("Equipment") > -1)
    {
      addTslLink = true;
    }
    else if (data.indexOf("Type:") > -1 && data.indexOf("Duplicate System") > -1 )
    {
      addText = addText + " DuplicateSystem";
      addTslLink = true;
    }
    else if (data.indexOf("Submitting Organization:") > -1 && data.indexOf("Domains") > -1 )
    {
      addText = addText.replace('MeAsRequestor','SubMSRRequestor') + " SubDomainRequest";
      addDomainLink = true;
    }
    else if (data.indexOf("Working Organization:") > -1 && data.indexOf("Domains") > -1 )
    {
      addText = addText + " DomainRequest";
      addDomainLink = true;
      break;
    }
    //Logger.log(addText);
  }
  
  if (customer)
  {
    addText = addText + " CustomerOfRequest";
  }
  addText = addText.replace(/\s{2,}/g, ' ').trim();
  return { txt: addText, TslLnk: addTslLink, DomainLnk: addDomainLink, Details: addDetails };
}

/*
function TestNewFunction()
{
    var raw = 
        'Subject: testing Draft from Apps Script\n' + 
         'To: cyrus@mydomain.net\n' +
         'Content-Type: multipart/alternative; boundary=1234567890123456789012345678\n' +
         'testing Draft msg\n' + 
         '--1234567890123456789012345678--\n';
  
  var raw = '2015-05-26 14:21:43:614 +0200 021970 INFO raw=Delivered-To: robert.gester@gp.globalfoundries.com\n' + 
    'Received: by 10.70.82.163 with SMTP id j3csp2237607pdy;\n' + 
    '          Tue, 26 May 2015 04:00:19 -0700 (PDT)\n' + 
    'From: Request Application <request.application@globalfoundries.com>\n' + 
    'To: <Hans-Joerg.Buettner@globalfoundries.com>\n' + 
    'Subject: [XXX01234] -  testing creation of mails from Apps Script XXX01234\n' + 
    'Date: Tue, 26 May 2015 06:00:03 -0500\n' + 
    'Message-ID: <4BCED60339A04B459E0B7881F48B6E51@gfoundries.com>\n' + 
    'MIME-Version: 1.0\n' + 
    'Importance: normal\n' +
    'Content-Type: multipart/alternative; boundary=1234567890123456789012345678\n' +
         'testing Draft msg\n' + 
         '--1234567890123456789012345678--\n';
  
  var resp = createMessage_(raw, "14d7c37ad5ebb2e8" );  
  Logger.log(JSON.stringify(resp));
}
*/


function replaceText_(text, startText,endText,newText, addBefore) {
  	var start = text.indexOf(startText);
  	if (start > -1)
    {
      var end = text.indexOf(endText,start+1);      
      if (end > -1)
      {
        end = end + endText.length;
        return text.slice(0,start) + newText + text.slice(end);
      }
      else
      {
        return text.slice(0,start) + newText;
      }
    }
  else if (addBefore)
  {
    var start = text.indexOf(addBefore);
    if (start > -1)
    {
      	return text.slice(0,start) + newText + text.slice(start);
    }
  }
  return text;
}

function createDraft_() {
  //http://stackoverflow.com/questions/17660601/create-draft-mail-using-google-apps-script
  //http://stackoverflow.com/questions/25391740/how-to-use-the-google-apps-script-code-for-creating-a-draft-email-from-985
  
  try{
    var forScope = GmailApp.getInboxUnreadCount(); // needed for auth scope

    var raw = 
        'Subject: testing Draft from Apps Script\n' + 
         //'To: cyrus@mydomain.net\n' +
         'Content-Type: multipart/alternative; boundary=1234567890123456789012345678\n' +
         'testing Draft msg\n' + 
         '--1234567890123456789012345678--\n';

    var draftBody = Utilities.base64Encode(raw);
    //Logger.log(draftBody);
    
    var params = {method:"post",
                  contentType: "application/json",
                  headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
                  muteHttpExceptions:true,
                  payload:JSON.stringify({
                    "message": {
                      "raw": draftBody
                    }
                  })
                 };
    
    var resp = UrlFetchApp.fetch("https://www.googleapis.com/gmail/v1/users/me/drafts", params);
    Logger.log(resp.getContentText());
  /*
   * sample resp: {
   *   "id": "r3322255254535847929",
   *   "message": {
   *     "id": "146d6ec68eb36de8",
   *     "threadId": "146d6ec68eb36de8",
   *     "labelIds": [ "DRAFT" ]
   *   }
   * }
   */
    
  }catch(err){
    Logger.log(err.lineNumber + ' - ' + err);
  }
}

function createMessage_(raw, replyToMessageId, labelIDs)
{
    var forScope = GmailApp.getInboxUnreadCount(); // needed for auth scope

  	var draftBody = Utilities.base64Encode(raw);
    draftBody = draftBody.replace(/\//g,'_').replace(/\+/g,'-'); //http://stackoverflow.com/questions/26663529/invalid-value-for-bytestring-error-when-calling-gmail-send-api-with-base64-encod
    
    var url = "https://www.googleapis.com/gmail/v1/users/me/messages";
  
  var     params = {method:"post",
                    contentType: "application/json",
                    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
                    muteHttpExceptions:true
                   };

  var payload = {};
  
  if (replyToMessageId)
  { 
    payload['threadId'] = replyToMessageId;
  }

  if (labelIDs)
  { 
    payload['labelIds'] = labelIDs;
  }
  
  payload['raw'] = draftBody;

  params['payload'] = JSON.stringify(payload);
    
    Logger.finest('url=' + url);
    Logger.finest('params=' + JSON.stringify(params));
    
    var resp = UrlFetchApp.fetch(url, params);
  /*
   * sample resp: {
   * "id": "14d8f93685b8c197",
   * "threadId": "14d8f93685b8c197",
   * "labelIds": [ "CATEGORY_UPDATES", "Label_176" ]
   * }
   */

    var o  = JSON.parse(resp.getContentText());

    if(o.error)
    {
      Logger.warning('!! Error occured!! ----------------------------------------------------------');
      Logger.info('threadid=' + replyToMessageId);
      Logger.info('labelIDs=' + labelIDs);
      Logger.finest('raw=' + raw);
      Logger.info('url=' + url);
      Logger.info('params=' + JSON.stringify(params));
      Logger.info(resp.getContentText());
      throw o.error.code + ":" +  o.error.message;
    }
    else
    {
      Logger.finer('new message id=' + o.id + " threadId=" + o.threadId);
      return o;
    }
}
