var Log_Sheet_Name = 'WebServiceLog';
/**
 * Serves HTML of the application for HTTP GET requests.
 * If folderId is provided as a URL parameter, the web app will list
 * the contents of that folder (if permissions allow). Otherwise
 * the web app will list the contents of the root folder.
 *
 * @param {Object} e event parameter that can contain information
 *     about any URL parameters provided.
 */
function doGet(e) {
  try
  {  
    initLog(Log_Sheet_Name,arguments.callee.name, 'doGet - Web Interface loaded');
    if (e) {
      Logger.log("Parameter e=" + JSON.stringify(e));
    }
    var template = HtmlService.createTemplateFromFile('Index');

    /*
    // Retrieve and process any URL parameters, as necessary.
    if (e.parameter.folderId) {
    template.folderId = e.parameter.folderId;
    } else {
    template.folderId = 'root';
    }
    */
    //preInitialize Data:
    SetupAllData();
    Logger.setCaller(arguments.callee.name);
    
    template.userData = PropertiesService.getUserProperties().getProperties();
    template.canClose = false;
    template.testMode = false;
    template.labels = [];
  
    var userlabels = GmailApp.getUserLabels();
    for (var i=0 ; i < userlabels.length; i++) {
      template.labels.push(userlabels[i].getName());
    }
  
    template.folders = driveHelper.listFolders(undefined, 3);
    var rootFolderName = DriveApp.getRootFolder().getName();
    for (var fldr in template.folders) {
      template.folders[fldr] = template.folders[fldr].replace('/' + rootFolderName, '').replace(/^\/|\/$/g, '');
      //Logger.log("actual name:" + template.folders[fldr]);
    }  
  
  // Build and return HTML in IFRAME sandbox mode.
    return template.evaluate()
                   .setTitle('Mail Script Setup')
                   .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    Logger.log("done");
  } catch (e) {
    e = (typeof e === 'string') ? new Error(e) : e;
    Logger.severe(e);
    throw e;
  }    
}

//Function to Initialize Spreadsheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();


  ui.createMenu('Start')
    .addItem("Help", "ShowInfo")
    .addSeparator()
    .addItem("About Info ...", "ShowAbout")
    .addToUi();

  var SsA = SpreadsheetApp.getActiveSpreadsheet();
  SsA.toast("Loading complete... --> Please click the Start menu above to test...", "FS", 15);

  //ShowInfo();
}

/* ------------------------------------------------------------------------------------------------------------ *
 * Menu Functions
 * ------------------------------------------------------------------------------------------------------------ */
function ShowAbout() {
  //var SsA = SpreadsheetApp.getActiveSpreadsheet();
  var SsA = SpreadsheetApp.openById('1J82jl8VMXf6c-8ryYZqNlelY6DKP7pA9vQiZ7wYLxX0');
  var dataSheet = SsA.getSheetId();

  Browser.msgBox("Script created by Robert Gester. \n\r You can send " + MailApp.getRemainingDailyQuota() + " mails today! SpreadSheet ID:'" + SsA.getId() + "' Sheet:'" + dataSheet + "'");
}

function ShowInfo() {
  var html = HtmlService.createHtmlOutputFromFile('Info')
    .setTitle("Google Mail Scripts")
    .setWidth(400)
    .setHeight(260);
  
  var ss = SpreadsheetApp.getActive();
  ss.show(html);
}

/* ------------------------------------------------------------------------------------------------------------ *
 * CallBack Functions
 * ------------------------------------------------------------------------------------------------------------ */

/**
 * Return an array of up to 20 filenames contained in the
 * folder previously specified (or the root folder by default).
 *
 * @param {String} folderId String ID of folder whose contents
 *     are to be retrieved; if this is 'root', the
 *     root folder is used.
 * @return {Object} list of content filenames, along with
 *     the root folder name.
 */
function saveContents(data) {
  try
  {  
    initLog(Log_Sheet_Name,arguments.callee.name,'Data given:\n' + JSON.stringify(data));
    //PropertiesService.getUserProperties().setProperties(properties);
    
    var labels = data["saveMailPDFItems[label]"];
    var folders = data["saveMailPDFItems[folder]"];
    var mailBodys = data["saveMailPDFItems[mailBody]"];
    var mailAttachments = data["saveMailPDFItems[mailAttachment]"];
    var labelProcesseds = data["saveMailPDFItems[labelProcessed]"];
    
    delete data[""];
    delete data["saveMailPDFItems[label]"];
    delete data["saveMailPDFItems[folder]"];
    delete data["saveMailPDFItems[mailBody]"];
    delete data["saveMailPDFItems[mailAttachment]"];
    delete data["saveMailPDFItems[labelProcessed]"];
    
    var saveMailPDFItems = [];
    
    if (labels) {
      //Logger.log(JSON.stringify(labels));
      //Logger.log(JSON.stringify(folders));
      //Logger.log(JSON.stringify(mailBodys));
      //Logger.log(JSON.stringify(mailAttachments));
      //Logger.log(JSON.stringify(labelProcesseds));
      
      if (typeof labels === 'string') {
        //Logger.log('labels is string ' + labels);
        var obj = {};
        obj.label = labels;
        obj.folder = folders;
        obj.mailBody = mailBodys
        obj.mailAttachment = mailAttachments;
        obj.labelProcessed = labelProcesseds;
        saveMailPDFItems.push(obj);
      }
      else 
      {
        //Logger.log('labels is array of length = ' + labels.length);
        for (var i=0 ; i < labels.length; i++) {
          var obj = {};
          obj.label = labels[i];
          obj.folder = folders[i];
          obj.mailBody = mailBodys[i];
          obj.mailAttachment = mailAttachments[i];
          obj.labelProcessed = labelProcesseds[i];
          saveMailPDFItems.push(obj);
        }
      }
    }
    //Logger.log('ok?');
    
    data.saveMailPDFItems = JSON.stringify(saveMailPDFItems);
    Logger.log('preparedData:\n' + JSON.stringify(data));
    PropertiesService.getUserProperties().deleteAllProperties();
    PropertiesService.getUserProperties().setProperties(data);
    
    SetupAllData();
    
    return; // JSON.stringify(PropertiesService.getUserProperties().getProperties());
  } catch (e) {
    e = (typeof e === 'string') ? new Error(e) : e;
    Logger.severe(e);
    throw e;
  }
}

/**
 * Return an array of up to 20 filenames contained in the
 * folder previously specified (or the root folder by default).
 *
 * @param {String} folderId String ID of folder whose contents
 *     are to be retrieved; if this is 'root', the
 *     root folder is used.
 * @return {Object} list of content filenames, along with
 *     the root folder name.
 */
function resetContents() {
  try
  {  
    initLog(Log_Sheet_Name,arguments.callee.name,'resetContents');
    PropertiesService.getUserProperties().deleteAllProperties();
    SetupAllData();
    return '<strong>Please reload/refresh Page (<mark>press F5</mark>) To see changes!</strong><br><br>The Page Reload could not be done automatically due to security restrictions';
    
  } catch (e) {
    e = (typeof e === 'string') ? new Error(e) : e;
    Logger.severe(e);
    throw e;
  }
}


function runScript(name) {
  try
  {
    initLog(Log_Sheet_Name,arguments.callee.name, name);

    if (name === 'saveMailImages') {
      Logger.log('call Exec_Gmail_Save_Image_Attachments');
      return Exec_Gmail_Save_Image_Attachments();
    }
    if (name === 'saveMailPDF') {
      Logger.log('call Exec_Gmail_Save_as_PDF');
      return  Exec_Gmail_Save_as_PDF();
    }
    if (name === 'saveOldMail') {
      Logger.log('call Exec_Gmail_Save_old_as_PDF');
      return  Exec_Gmail_Save_old_as_PDF();
    }
    if (name === 'deleteOldMail') {
      Logger.log('call Exec_Gmail_Delete_Old');
      return  Exec_Gmail_Delete_Old();  
    }
    if (name === 'renameMail') {
      Logger.log('call Exec_Gmail_Rename_SystemMails');
      return  Exec_Gmail_Rename_SystemMails("TEST");
    }
    
    throw "Script '" + name + "' not found!!"
    } catch (e) {
      e = (typeof e === 'string') ? new Error(e) : e;
      Logger.severe(e);      
      throw e;
  }      
}

/* ------------------------------------------------------------------------------------------------------------ */
function initLog(name, caller, firstLine) {
  if (!caller)
    caller = 'undefined';
  
  if (!name)
    name = caller;
  
  try {
    var user = Session.getEffectiveUser().getEmail().split("@")[0]; //get Username
    //(throwErrorByUser_ === '' || Session.getEffectiveUser().getEmail().indexOf(throwErrorByUser_) > -1);
    var userallowed = (user.toLowerCase().indexOf('gester') > -1);
    BetterLog.setThrowError(userallowed).setUser(user);

    if (!userallowed) {
      BetterLog.SHEET_MAX_ROWS = -1;
    }

    Logger = BetterLog.setUser(user)
                      .setCaller(caller)
                      .setLevel(PropertiesService.getScriptProperties().getProperty('BetterLogLevel')) //defaults to 'INFO' level
                      .useSpreadsheet('1yC9jyZjcmMhRqJP1M5vs8FcwQe7P3l-uDFmoE5iUJlc', name); //'1yC9jyZjcmMhRqJP1M5vs8FcwQe7P3l-uDFmoE5iUJlc' //automatically rolls over at 50,000 rows      
  } catch (e) {
    e = (typeof e === 'string') ? new Error(e) : e;
    var errInfo = e.toString(); 
    for (var prop in e)  {  
      errInfo += "\n     " + prop+ ": "+ e[prop]; 
    }     
    Logger.log(errInfo);
    BetterLog.doThrowError(e);
  }
  if (firstLine) {
    Logger.log(firstLine);
   }
}

function randomStr(length)
{
  var s = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  return Array.apply(null, Array(length)).map(function() { return s.charAt(Math.floor(Math.random() * s.length)); }).join('');
}

function testlog()
{ 
  initLog("TestLog",arguments.callee.name, 'starting test');
  for (var i = 0; i < 2000; i++) {
    Logger.log("logging bla " + i + " - " + randomStr(10))
  }
}


// arguments.callee.name
function doLogWebService(msg) {
  try
  {
    initLog(Log_Sheet_Name,msg.caller);

    if (msg.log === 'info') {
      Logger.info(msg.text);
      if (msg.obj) {
        Logger.info(JSON.stringify(obj));
      }
    } else if (msg.log === 'fine') {
      Logger.fine(msg.text);
      if (msg.obj) {
        Logger.fine(JSON.stringify(obj));
      }
    } else if (msg.log === 'finer') {
      Logger.finer(msg.text);
      if (msg.obj) {
        Logger.finer(JSON.stringify(obj));
      }
    } else if (msg.log === 'finest') {
      Logger.finest(msg.text);
      if (msg.obj) {
        Logger.finest(JSON.stringify(obj));
      }
    } else if (msg.log === 'warning') {
      Logger.warning(msg.text);
      if (msg.obj) {
        Logger.warning(JSON.stringify(obj));
      }
    } else if (msg.log === 'severe') {
      Logger.severe(msg.text);
      if (msg.obj) {
        Logger.severe(JSON.stringify(obj));
      }
    } else {
      Logger.log(msg.text);
      if (msg.obj) {
        Logger.log(JSON.stringify(obj));
      }
      throw new Error("Unknown Message type " + msg.log + "!");
    }
  } catch (e) {
    e = (typeof e === 'string') ? new Error(e) : e;
    Logger.severe(e);      
    //throw e;
  }      
}

/* ------------------------------------------------------------------------------------------------------------ */

/*

output:
{

2015-05-28 00:02:39:396 +0200 000272 FINEST rdgester@gmail.com: -> 

"":["z/processed Images","z/skipped Images","x/Save As PDF","on","on","","x/Save Minutes As PDF","on","on","z/Meeting Minutes","x/Save Attachments","on","","z/processed Mail","z/processed mail","z/error process mail"],
"saveOldMailPDFSearch":"(older_than:2y -{This message has been archived} -{Accepted OR Zugesagt OR Abgelehnt OR Tentative OR Weiterleitungsbenachrichtigung OR Nachrichtenrückruf OR Freigabeanfrage}  -filename:ics)",
"renameMailSearchQuerySimilar":"from:(request.application@globalfoundries.com OR globalfoundries@service-now.com OR Request Application OR ServiceNow) -subject:((Out of office) OR OOO)",
"saveMailPDFActive":"false",
"renameMailSearchQuery":"from:(request.application@globalfoundries.com OR globalfoundries@service-now.com OR Request Application OR ServiceNow) is:unread -subject:{redirect} -has:attachment",
"renameMailLabelProcessed":"z/processed mail",
"saveMailImagesLabelSkipped":"z/skipped Images",
"saveMailPDFItems[folder]":["GMail/PDF Storage","GMail/Meeting Minutes","GMail/Attachments"],
"saveMailPDFItems[label]":["x/Save As PDF","x/Save Minutes As PDF","x/Save Attachments"],
"saveOldMailPDFFolder":"GMail/PDF Storage",
"saveMailImagesSearch":"in:all -in:spam -in:trash has:attachment filename:jpg OR filename:png OR filename:gif",
"deleteOldMailSearchQuery":"(older_than:2d label:z/Delete2Days",
"saveOldMailPDFActive":"false",
"canClose":"true",
"testMode":"true",
"saveMailPDFItems[labelProcessed]":["","z/Meeting Minutes",""],
"saveOldMailPDFLabelProcessed":"z/processed Mail",
"saveMailPDFItems[mailAttachment]":["true","true","true"],
"saveMailImagesLabelProcessed":"z/processed Images",
"saveMailImagesDriveFolder":"GMail/Images",
"saveMailImagesActive":"false",
"renameMailActive":"false",
"saveMailPDFItems[mailBody]":["true","true","false"],
"deleteOldMailActive":"false",
"renameMailLabelError":"z/error process mail"
}


*/
