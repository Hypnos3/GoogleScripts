/*
Used Librarys:
https://sites.google.com/site/scriptsexamples/custom-methods/betterlog

https://sites.google.com/site/scriptsexamples/custom-methods/underscoregs

useful links:
https://sites.google.com/site/scriptsexamples/home

utilitys:
https://gist.github.com/mogsdad/6515581

other:
 https://gist.github.com/mogsdad/6515581
 https://gist.github.com/mogsdad

*/
 
var mailExcludeId = "MailChangedByScriptX";
var RenameMailScriptVer_ = "5a";
var RuntimeIsOverMax_ = false;
var RenameMailScriptStartTime_ = new Date(); 

// This will auto-save the image attachments from Gmail to Google Drive
function Exec_Gmail_Rename_SystemMails(runParam)
{
  initLog('RenameMail',arguments.callee.name + RenameMailScriptVer_); //,'TestMode for ' + RenameMailScriptVer_
  //Logger.setLevel("FINEST");

  //define aftzer what time the script should be retriggered
  var retriggertime = 5; //only allowed 1,5,10,15...
  
  var n = new Date().getHours();
  if (n >20 || n < 6)
    retriggertime = 15;  
  
  try
  {
    // to calc elapsed time
    var userProperties = PropertiesService.getUserProperties();
    var testMode = (userProperties.getProperty('TestMode')=== 'true');
    var labelNameProcessed = userProperties.getProperty('renameMailLabelProcessed'); //"z/processed mail";
    var labelNameError = userProperties.getProperty('renameMailLabelError'); //"z/error process mail";
    var enhancedOutput = (runParam === "TEST");
    
    if (!labelNameProcessed || !labelNameError)
    {
      throw 'Run Error, Please save Data first!';
    }
    
    if (checkProperty_('renameMailActive', 'false') === 'true') {
      activateTrigger_(arguments.callee.name, 30);
    }
    
    // Handle max execution times in our outer loop
    // Get start index if we hit max execution time last run
    var LastStart = Date.parse(userProperties.getProperty(arguments.callee.name + "-LastStart")) || -10000;
    
    if (Math.round((RenameMailScriptStartTime_ - LastStart)/1000) < 360) {
      throw new Error('Function is already running!\nPlease wait until complete or at least 6 Minutes!\nScript StartTime=' + RenameMailScriptStartTime_ + '\nLastStart=' + LastStart + '\ndifference='+Math.round((RenameMailScriptStartTime_ - LastStart)/1000));
    }    
    userProperties.setProperty(arguments.callee.name + "-LastStart",RenameMailScriptStartTime_);

    
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
    
    var searchQuery = userProperties.getProperty('renameMailSearchQueryV2'); //"from:(request.application@globalfoundries.com OR globalfoundries@service-now.com OR Request Application OR ServiceNow) is:unread -subject:{redirect} -has:attachment";
    var searchQuery = searchQuery + " -label:" + labelNameProcessed + " -label:" + labelNameError + " -" + mailExcludeId;
    
    var searchQuerySimilar = userProperties.getProperty('renameMailSearchQueryV2Similar');
    //var searchRegex = userProperties.getProperty('renameMailIDRegex');

    var thread = searchThreads_(searchQuery);

    while (thread && !RuntimeIsOverMax_)
    {
      var messages = thread.getMessages();
      
      thread.addLabel(labelError);    
      if (messages) {
        Logger.setCaller(arguments.callee.name + RenameMailScriptVer_).log('processing mail: "' + messages[0].getSubject() + '"\nsearchQuery="' + searchQuery + '"\nretriggertime=' + retriggertime);
        
        for (var y=0; y<messages.length; y++)
        {          
          try
          {		
            processMessageGeneric_(messages[y], testMode, labelProcessed, labelNameProcessed, searchQuerySimilar);      
            if (!testMode) {  
              thread.removeLabel(labelError);
              thread.addLabel(labelProcessed);
            }
          } catch (e) {
            e = (typeof e === 'string') ? new Error(e) : e;
            if (e.message.indexOf('Utilities.sleep') > -1) {
              Logger.setCaller(arguments.callee.name + RenameMailScriptVer_).warning(e);
              //Logger.setCaller(arguments.callee.name + RenameMailScriptVer_).warning('Error occured, which will be ignored!! %s: %s (line %s, file "%s"). Stack: "%s" . While processing "%s".',
              //               e.name||'', e.message||'', e.lineNumber||'', e.fileName||'', e.stack||'', messages[y].getSubject()||'');
              thread.removeLabel(labelError);
              return;
            }
            Logger.setCaller(arguments.callee.name + RenameMailScriptVer_).severe(e);
            messages[y].markUnread();
          }        
          if (Math.round((new Date() - RenameMailScriptStartTime_)/1000) > 280) { //300  - 360 seconds is Google Apps Script max run time
            //We've hit max runtime. 
            RuntimeIsOverMax_ = true;
            break;
          }
        }
      }
      if (!RuntimeIsOverMax_) {
        thread = searchThreads_(searchQuery);
      } else {
        thread = undefined;
      }
      
      /*
      if (Math.round((new Date() - RenameMailScriptStartTime_)/1000) > 250) { //300  - 360 seconds is Google Apps Script max run time
        //We've hit max runtime. 
        RuntimeIsOverMax_ = true;
        thread = undefined;
      } else {
        thread = searchThreads_(searchQuery);
      }*/
    }
  } catch (e) {
    e = (typeof e === 'string') ? new Error(e) : e;
    Logger.setCaller(arguments.callee.name + RenameMailScriptVer_).severe(e);

    if (enhancedOutput) {
      throw e;
    }
  } finally {
    if (RuntimeIsOverMax_) {
      Logger.setCaller(arguments.callee.name + RenameMailScriptVer_).warning('Hit max run time!');
      //retrigger in two Minutes
      retriggertime = 1; //only allowed 1,5,10,15...
    }
    
    userProperties.deleteProperty(arguments.callee.name + "-LastStart");
    if (checkProperty_('renameMailActive', 'false') === 'true') {
      activateTrigger_(arguments.callee.name, retriggertime);
    }
  }
}

function searchThreads_(searchQuery)
{
  
  var threads = GmailApp.search(searchQuery);
  var cnt = threads.length;
  if (cnt < 1 ) {
    return undefined;  
  }
  
  threads = GmailApp.search(searchQuery,cnt-1,1);
  if (threads.length < 1) {
    var e = new Error('thread is null, That should never happen!!');
    Logger.setCaller(arguments.callee.name).severe(e);
    return undefined;
  }
  return threads[0]; //messages[messages.length-1];
}

function processMessageGeneric_(message, testMode, excludeLabel, excludeLabelName, searchQueryEnhancement)
{
  var originalSubject = message.getSubject();
  var bodyHtml = message.getBody().trim();
  var originalSender = message.getFrom();
  var date = message.getDate();
  var messageId = message.getId();
  var messageThreadId = message.getThread().getId();
  var rawData= message.getRawContent();
  var sender = Session.getEffectiveUser().getEmail();
  Logger.setCaller(arguments.callee.name + RenameMailScriptVer_).finer("--------------------------> " + originalSender + " Subject=" + originalSubject);

  if (rawData.indexOf(mailExcludeId) > -1)
  {
    Logger.log("Mail already processed! Subject=" + originalSubject);
    return;
  }
  
  var cleanSubject = originalSubject.replace(/\s\s+/g, ' '); //removing double spaces
 
  var subjIDs = getTextParts_(cleanSubject, /(\b[A-Z]{2,}\d{4,}\b|\b[A-Z]+\d{2,}-\d{4,}\s+(v|V|Version|version|ver|VER|-)?\s*\d+\b|\b[A-Z]+\d{2,}-\d{4,}\b)/g);

  var RequestIDOriginal = subjIDs[0];
  var RequestID = RequestIDOriginal.replace(/\s+/g, ""); //.replace(/\[+(.*?)\]+/g,"$1");
  var RequestTyp = RequestID.replace(/\d/g, '');
  
  cleanSubject = replaceAllIC_(originalSubject, RequestIDOriginal, '');
  cleanSubject = replaceAllIC_(cleanSubject, RequestID, '');  
  cleanSubject = cleanSubject.replace(/\[+(.*?)\]+/g,"$1"); //removing multiple brackets
  cleanSubject = cleanSubject.replace(/(\s\s+|\s*\[+\s*\]+\s*|\s*\(+\s*\)+\s*|\s*\{+\s*\}+\s*)/g, ' '); // removing multiple spaces, empty brackets, etc...  
  cleanSubject = cleanSubject.replace(/\s*-\s*-\s*/g, ' - ').replace(/(^\s*-\s*|\s*\-*\s*\.*\s*$|\s+[a-z]\s*\-*\s*\.*\s*$)/g, '').trim(); // removing multiple slashes
  
  var orginalSenderID = originalSender.replace(/<\S+>/g, '').trim();
  if (!orginalSenderID) {
    orginalSenderID = originalSender;
  }
  
  orginalSenderID = orginalSenderID.replace(/[()@<>.]/g, '').replace(/globalfoundriescom/gi, '').trim();
  //Logger.log("RequestID='" + RequestID+ " RequestTyp=" + RequestTyp);
  
  /**********************************************************************************************************************************************************************************/

  subjIDs = getTextPartsExt_(rawData, subjIDs, new RegExp("(\\b[A-Z]{2,}\\d{4,}\\b|\\b[A-Z]+\\d{2,}-\\d{4,}\\s*([a-zA-Z]+|\\s)\\s*\\d+\\b)", "g")) //SpecIDs werden nur mit Version betrachtet

  if ((orginalSenderID.indexOf("equest") > -1) && bodyHtml) {
    //Is for finding similar Request Express entrys (e.g. Sub Domain Requests)
    //Logger.log("bodyHtml:\n" + bodyHtml);
    //für RequestExpress ermitteln der Subject:
    subjIDs = getTextPartsExt_(bodyHtml, subjIDs, /--\s+(.|\n)+?\<br/g) //"-- (.*?)(\\<br|\\n)"
    //Logger.log("subjIDs: " + JSON.stringify(subjIDs));
  }

  var query = "(from:(" + sender + " OR " + originalSender + ") (" + RequestID;
  
  if (RequestIDOriginal !== RequestID) {
    query += " OR " + RequestIDOriginal;
  }
    
  for (var i = 0, len = subjIDs.length; i < len; i++) {
    if (subjIDs[i] !== RequestID && subjIDs[i] !== RequestIDOriginal ) {      
      if (subjIDs[i].indexOf(' ') >-1) {
        query += " OR (" + subjIDs[i] + ")";
      } else {
        query += " OR " + subjIDs[i];
      }
    }
  }    
  query += ")) " + searchQueryEnhancement;
  
  var replyMessage = GetMessage_(query, messageId, messageThreadId, date, originalSubject);

  if (replyMessage) {
    Logger.info("Found other Message id=" + replyMessage.getId() + "\nsubject=" + replyMessage.getSubject() + "'\nquery='" + query + "'");
  } else {
    Logger.info("No other Message found.\nquery='" + query + "'");
  }

  var TextContent = parseMessage_(bodyHtml, Session.getEffectiveUser().getEmail());
  
  //-----------------------------------------------------------------------------------------------------
  var HtmlAddtable = "<small><code><table border='0' align='left' width='90%' ><tbody><tr>" + 
    "<tr><td colspan='2'>--------------------------------------------------------------------------------------</td></tr>" +
      "<tr><td>original subject:</td><td>" + originalSubject + "</td></tr>" +
        "<tr><td>original Message ID:</td><td>" + messageId + "</td></tr>" +
          '<tr><td>search:</td><td>' + RequestID + ' - ' + RequestTyp + ' - ' + RequestID.replace(RequestTyp, '') + ' - ' + TextContent.txt + ' ' + mailExcludeId + ' </td></tr>';
    
  if (TextContent.TslLnk) {
    HtmlAddtable = HtmlAddtable + '<tr><td>Links:</td><td><a href="http://myteamsdrs.gfoundries.com/sites/Fab36_FA/300mmModelling/ToolStartUpCheckListe/Forms/AllItems.aspx">http://myteamsdrs.gfoundries.com/sites/Fab36_FA/300mmModelling/ToolStartUpCheckListe/Forms/AllItems.aspx Tool Startup Liste</a> - <a href="//file:///G:/Ops-FA/ToolStartUpCheckList/GF.TSL-Launcher/GF.TSL-Launcher.application">//file:///G:/Ops-FA/ToolStartUpCheckList/GF.TSL-Launcher/GF.TSL-Launcher.application Tool Startup Listen Tool</a></td></tr>';      
  }
  
  if (TextContent.DomainLnk) {
    HtmlAddtable = HtmlAddtable + '<tr><td>Links:</td><td><a href="http://myteamsfap/sites/pm/Shared%20Documents/MSR-Help.mht">http://myteamsfap/sites/pm/Shared%20Documents/MSR-Help.mht MSR based on Domain oriented Demand Management</a></td></tr>';
  }
  
  if (replyMessage) {
    HtmlAddtable = HtmlAddtable + '<tr><td>ReplyMessage:</td><td>' + replyMessage.getSubject() + '</td></tr>';
  }

  HtmlAddtable = HtmlAddtable + '</tbody></table>deactivate at <a href="' + ScriptApp.getService().getUrl() + '">' + ScriptApp.getService().getUrl() + '</a></code></small>'
    
  if (testMode)
  {  
    return;
  }    

  rawData = rawData.replace(' width="90%"','');
  
  var labels = message.getThread().getLabels();
  var resp = undefined;

  if (replyMessage) {
    /***********************************************************************************************
    *** Found an other Message where reply to
    ***********************************************************************************************/
    var subject = replyMessage.getSubject();
    var referenceID = replyMessage.getId(); //not working: replyMessage.getThread().getId()
    
    if (!cleanSubject) {
      rawData = replaceText_(rawData,'<body','>','<body><h2>' + RequestIDOriginal + '</h2><br>--------------------------------------------------------------------------------------<br>\r\n');
    } else {
      if (subject.indexOf(RequestID) > -1 || subject.indexOf(RequestIDOriginal) > -1 ) {
        rawData = replaceText_(rawData,'<body','>','<body><h2>' + cleanSubject + '</h2><br>--------------------------------------------------------------------------------------<br>\r\n');
      } else {
        rawData = replaceText_(rawData,'<body','>','<body><h2>' + RequestID + ' - '+ cleanSubject + '</h2><br>--------------------------------------------------------------------------------------<br>\r\n');
      }
    }
    
    /*
    var rmLabels = replyMessage.getThread().getLabels();
    for (var lbl=0; lbl<rmLabels.length; lbl++)
    {
    labels.push(rmLabels[lbl]);
    }*/
        
    rawData = replaceText_(rawData, 'References: ','\n','References: ' + referenceID + '\r\n', 'Subject: ');
    rawData = replaceText_(rawData, 'In-Reply-To: ','\n','In-Reply-To: ' + referenceID + '\r\n', 'Subject: ');
    rawData = replaceText_(rawData, 'Subject: ','\n','Subject: ' + subject + '\r\n');
    if (rawData.indexOf("</body>") > -1) {
      rawData = rawData.replace('</body>', HtmlAddtable+'</body>' );
    } else {
      rawData = rawData + HtmlAddtable+'</body>';
    }
    Logger.setCaller(arguments.callee.name + RenameMailScriptVer_).info("try to created mail in thread threadId=" + replyMessage.getThread().getId() + " threadSubject='" + subject + "'");
    resp = createMessage_(rawData, labels, referenceID); 
    Logger.setCaller(arguments.callee.name + RenameMailScriptVer_).info("created mail in thread; mail id '" + resp.id + "' subject='" + subject + "' tothread=" + referenceID);
  }
  else
  {
    /***********************************************************************************************
    *** Found no other Message, Formatting Subject to be more generic
    ***********************************************************************************************/
    var subject = undefined;
    
    if (cleanSubject) {
      rawData = replaceText_(rawData,'<body','>','<body><h2>' + cleanSubject + '</h2><br>--------------------------------------------------------------------------------------<br>\r\n');
      subject = formatSubject_(cleanSubject); //.replace(/\s\s+/g, ' '); //.replace(/  +/g, ' ');
    }    
    
    subject= ('[' + RequestID + '] ' + subject).trim();
    /*
    if (!subject) {
      subject= '[' + RequestID + '] ' + cleanSubject;
    } else if (subject.indexOf(RequestID) < 0) {
      subject = '[' + RequestID + '] ' + subject;
    }*/

    //rename Message;
    rawData = replaceText_(rawData, 'Subject:','\n','Subject: ' + subject + '\r\n');
    if (rawData.indexOf("</body>") > -1) {
      rawData = rawData.replace('</body>', HtmlAddtable+'</body>' );
    } else {
      rawData = rawData + HtmlAddtable+'</body>';
    }
    //Logger.setCaller(arguments.callee.name + RenameMailScriptVer_).info('try to created mail subject= ' + subject );
    resp = createMessage_(rawData, labels );
    Logger.setCaller(arguments.callee.name + RenameMailScriptVer_).info("created new mail; no other Message found; mail id '" + resp.id + "' subject='" + subject);
  }
  
  if (resp) {
    var newMessage = GmailApp.getMessageById(resp.id);
    newMessage.getThread().addLabel(excludeLabel);      
    message.refresh(); // make sure it's up to date
    
    Utilities.sleep(1000);
    if (message.isInInbox()) {
      newMessage.getThread().moveToInbox();
    } else {
      newMessage.getThread().moveToArchive();
    }
      
    
    if (message.isStarred()) {
      newMessage.star();
    } else {
      newMessage.unstar();
    }
      
    if (message.isUnread()) {
      newMessage.markUnread();
    }
    else {
      newMessage.markRead();
    }

    message.markRead();
    message.moveToTrash();
    Utilities.sleep(1000); //Backend Errors?
  }
}

function formatSubject_(subject)
{ 
  var ok = true;
  var toReplace = ["open","opened","closed","approved","disapproved","approval","waiting","paused","submitted","started","completed","complete",
                   "incomplete","comment","assigned","reassigned","resolved","resubmitted", "decision", "pending", "cancelled","successful","unsuccessful",
                   "release", "required","rework","rollout","rollback","ready","review","design","feasible","development ",
                   "handover", "production" , "prod", "itdc", "let", "new status", "submsr", "sub", "domain", "to ", "no ", "not ", "for ","from ", "re:","fw:","fwd:","aw:","wg:","fyi:","fyi ","ot:", "was:",
                   "ms-a","ms-b","ms-pec","ms-p","ms-q","ms-w","ms-e","msr ", "cr ", "gr ", "setup.fc", "sew status", "is now", "insignoff", "inupdate", "task assignment"];
  var toReplaceLen = toReplace.length;
  for (var index = 0; index < toReplaceLen; ++index) {
    toReplace[index] = toReplace[index].toLowerCase();
  }  
                  
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
      for (var index = 0; index < toReplaceLen; ++index) {
        var subjectt=subject.toLowerCase();
        if (subjectt.indexOf(toReplace[index]) === 0)
        {
            subject = subject.substr(toReplace[index].length).trim();     
          	ok = true;
        }
      }
    }
    //Logger.setCaller(arguments.callee.name).finest("subject change > " + subject);
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

function GetMessage_(searchQuery, excludeID, excludeThreadID, date, LogEnhancement)
{  
  var threads = GmailApp.search(searchQuery);
  var cnt = threads.length;
  Utilities.sleep(1000);

  for (var x=cnt-1; x>=0; x--)
  {  
    var messages = threads[x].getMessages();
    for (var y=0; y<messages.length; y++)
    { 
      if (messages[y].getId() !== excludeID && messages[y].getThread().getId() !== excludeThreadID) {
        var newDate= messages[y].getDate();
        var newDateVal= newDate.valueOf();

        if (newDateVal < date.valueOf())
        {
          //Logger.setCaller(arguments.callee.name).info("found mail id=" + messages[y].getId() + " subject=" + messages[y].getSubject() + "' \nmessagesDate= " + newDate + " = " + newDateVal + " OriginalmessagesDate=" + date + " = " + date.valueOf() + " \nQuery='" + searchQuery + " result= " + cnt + "/" + y + "' LogEnhancement='" + LogEnhancement + "'");
          return messages[y];
        } else if (newDateVal === date.valueOf()) {        
          if (messages[y].getBody().indexOf(mailExcludeId) > -1) {
            //Logger.setCaller(arguments.callee.name).info("found mail id=" + messages[y].getId() + " subject=" + messages[y].getSubject() + "' \nmessagesDate= " + newDate + " = " + newDateVal + " OriginalmessagesDate=" + date + " = " + date.valueOf() + " \nQuery='" + searchQuery + " result= " + cnt + "/" + y + "' LogEnhancement='" + LogEnhancement + "'");
            return messages[y];
          }
          /* else {
            Logger.setCaller(arguments.callee.name).info("Skip message id=" + messages[y].getId() + " subject=" + messages[y].getSubject() + "' \nmessagesDate= " + newDate + " = " + newDateVal + " OriginalmessagesDate=" + date + " = " + date.valueOf() + "' \nLogEnhancement='" + LogEnhancement + "'");
          } */
        }
      }
    }
  }
  //Logger.setCaller(arguments.callee.name).info("no similar mail found for query: '" + searchQuery + "' date= " + date +" result= " + cnt + "' - LogEnhancement='" + LogEnhancement+ "'");
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
    
  //Logger.setCaller(arguments.callee.name).log(dataArr);
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
    //Logger.setCaller(arguments.callee.name).log(addText);
  }
  
  if (customer)
  {
    addText = addText + " CustomerOfRequest";
  }
  addText = addText.replace(/\s{2,}/g, ' ').trim();
  return { txt: addText, TslLnk: addTslLink, DomainLnk: addDomainLink, Details: addDetails };
}

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

function createMessage_(raw, labelIDs, replyToMessageId)
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

  Logger.setCaller(arguments.callee.name + RenameMailScriptVer_);

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
    var respTxt = resp.getContentText();
    var o  = JSON.parse(respTxt);

    if(o.error)
    {
      Logger.warning('!! Error occured!! ----------------------------------------------------------');
      Logger.info('threadid=' + replyToMessageId);
      Logger.info('labelIDs=' + labelIDs);
      //dogInfo('raw=' + raw);
      Logger.info('url=' + url);
      Logger.info('params=' + JSON.stringify(params));
      Logger.info(respTxt);
      throw new Error(o.error.code + ":" +  o.error.message);
    }
    else
    {
      Logger.finer('new message id=' + o.id + " threadId=" + o.threadId);
      return o;
    }
}

/***************************************************************************************************************************
*** Generic Functions:
***************************************************************************************************************************/

//Setup Trigger:  
function activateTrigger_(funcName, newTime) {
  try {
    if (!newTime) {
      newTime = 15; //retriggertime - only allowed 1,5,10,15...
    }    
    
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction()===funcName) {
        call_(function(){ ScriptApp.deleteTrigger(triggers[i]); });
      }
    }
  } catch (ex) {
    ex = (typeof ex === 'string') ? new Error(ex) : ex;
    Logger.setCaller(arguments.callee.name + RenameMailScriptVer_).severe('Error deleteTrigger for "' + funcName + '" with time=' + newTime + '!');
    Logger.severe(ex);
  }
  
  try {
    call_(function(){
      ScriptApp.newTrigger(funcName).timeBased().everyMinutes(newTime).create(); //retriggertime - only allowed 1,5,10,15... 
    });
  } catch (ex) {
    ex = (typeof ex === 'string') ? new Error(ex) : ex;
    Logger.setCaller(arguments.callee.name + RenameMailScriptVer_).severe('Not Possible to activate Trigger for "' + funcName + '" with time=' + newTime + '!! May be deactivated!!!');
    Logger.severe(ex);
  } 
}

/**
 * Replaces in a Text inside a Text with all occurences and ignores Case
 * @param {string} text long text where should be replaced
 * @param {string} find text to search for
 * @param {string} replacement the new text
 * @return {string} returns the text with the replacement
 */
function replaceAllIC_(text, find, replacement)
{ 
  return text.replace(new RegExp(escapeRegExp_(find), 'ig'), replacement);
}

/**
 * Escape a regular Expression
 * @param {string} string regular expression to escape
 * @return {string} returns the escaped regular expression
 */
function escapeRegExp_(string) {
    return string.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
}


function getTextParts_(text, searchRegex)
{
  return getTextPartsExt_(text, undefined, searchRegex);
}

function getTextPartsExt_(text, result, searchRegex)
{
  text = text.replace(/\s\s+/g, ' ').match(searchRegex);

  if (!result)
    result = [];

  if (!text || text.length === 0) {
    return result;  
  }

  var seen = {}; 
  for (var i = 0, len = result.length; i < len; i++) {
      seen[result[i]] = 1;
  }
  
  for (var i = 0, len = text.length, j = result.length; i < len; i++) {
    var item = text[i].replace(/\s*(-+|<br>|<br|<wbr>|<wbr|\/>|\n|\r|\t|\[+|\]+|\(+|\)+|\{+|\}+|&(nbsp|shy|lt|gt|quot|amp|apos|circ|tilde|#x200B);|:|\\+)\s*/g,' ').trim();
    if (item) {
      if (seen[item] !== 1) {
        seen[item] = 1;
        result[j++] = item;
      }
    }
  }
  return result;
}

//copy version 10 lib GASRetry 'MGJu3PS2ZYnANtJ9kyn2vnlLDhaBgl_dE' (changed function name and log line)
/**
* Invokes a function, performing up to 5 retries with exponential backoff.
* Retries with delays of approximately 1, 2, 4, 8 then 16 seconds for a total of 
* about 32 seconds before it gives up and rethrows the last error. 
* See: https://developers.google.com/google-apps/documents-list/#implementing_exponential_backoff 
* <br>Author: peter.herrmann@gmail.com (Peter Herrmann)
<h3>Examples:</h3>
<pre>//Calls an anonymous function that concatenates a greeting with the current Apps user's email
var example1 = call_(function(){return "Hello, " + Session.getActiveUser().getEmail();});
</pre><pre>//Calls an existing function
var example2 = call_(myFunction);
</pre><pre>//Calls an anonymous function that calls an existing function with an argument
var example3 = call_(function(){myFunction("something")});
</pre><pre>//Calls an anonymous function that invokes DocsList.setTrashed on myFile and logs retries with the Logger.log function.
var example4 = call_(function(){myFile.setTrashed(true)}, Logger.log);
</pre>
*
* @param {Function} func The anonymous or named function to call.
* @return {*} The value returned by the called function.
*/
function call_(func) {
  for (var n=0; n<6; n++) {
    try {
      return func();
    } catch(e) {
      e = (typeof e === 'string') ? new Error(e) : e;
      Logger.setCaller("call_ " + n).severe(e);
      
      if (n == 5) {
        throw e;
      } 
      Utilities.sleep((Math.pow(2,n)*1000) + (Math.round(Math.random() * 1000)));
    }    
  }
}
