/*
    !! Please activate only if you know what you doing !! 
 
    Auto-Save Gmail message as PDF in Google Drive
    ==============================================
 
    Written by Robert Gester 11/24/2014
    
    To setup script, choose Run -> Setup_Save_as_PDF. 
    For only saving attachments of Mails choose -> Setup_Save_Mail_Attachment. 
    
    If you have an email that you want to archive in Google Drive,
    you can use Google script to save it as PDF in your Google
    Drive account. The following script will save all the messages
    in an email thread as one PDF file in your Google Drive. If it
    comes with attachments, it will create a folder and store
    the messages and attachments within.
    
    Whenver you want to save an email and its attachments
    to Google Drive, simply tag it with the “x/Save As PDF” label.
*/

/*********************************************************************************************************
    Public Functions
*********************************************************************************************************/

function Exec_Gmail_Save_as_PDF(){
  initLog('SaveMailAsPDF',arguments.callee.name);
  try
  {
    
    var ItemsToSave = PropertiesService.getUserProperties().getProperty('saveMailPDFItems');
    if (!ItemsToSave)
    {
      throw 'Run Setup first!!';
    }
  
    // to replace later, when Interface for user Propertys is ready:
    ItemsToSave =  JSON.parse(ItemsToSave);
    
    Logger.finer("Exec_Gmail_Save_as_PDF");  
    for(var i = 0; i < ItemsToSave.length; i++)
    {
      var fldr = driveHelper.getFolder(ItemsToSave[i].folder);
      
      var label = GmailApp.getUserLabelByName(ItemsToSave[i].label);  
      if(label == null)
      {
        label = GmailApp.createLabel(ItemsToSave[i].label);
      }
      
      var labelprocessed = null;
      if(ItemsToSave[i].label_processed !== null && ItemsToSave[i].label_processed !== "")
      {
        labelprocessed = GmailApp.getUserLabelByName(ItemsToSave[i].label_processed);  
        if(labelprocessed == null)
        {
          labelprocessed = GmailApp.createLabel(ItemsToSave[i].label_processed);
        }
      }
      
      Gmail_Save_as_PDF_(fldr, label,ItemsToSave[i].mailBody,ItemsToSave[i].mailAttachment, labelprocessed);
    }
  } catch (e) {
    e = (typeof e === 'string') ? new Error(e) : e;
    Logger.severe(e);
    throw e;
  }  
}

//Saves 2 year old mails to a Google Drive Folder
function Exec_Gmail_Save_old_as_PDF(){
  initLog('SaveOldMailAsPDF',arguments.callee.name);
  try
  {

    var processedOldFolder =  PropertiesService.getUserProperties().getProperty('saveOldMailPDFFolder');
    var folder = driveHelper.getFolder(processedOldFolder);
    
    var labelProcessed_Name =  PropertiesService.getUserProperties().getProperty('saveOldMailPDFLabelProcessed');
    var labelProcessed = GmailApp.getUserLabelByName(labelProcessed_Name);  
    if(labelProcessed == null){
      labelProcessed = GmailApp.createLabel(labelProcessed_Name);
    }
    
    var query = PropertiesService.getUserProperties().getProperty('saveOldMailPDFSearch') + "-label:" + labelProcessed_Name.replace(/[\s]/g, "-");
    
    //(older_than:2y -label:z/processed-Mail -{this message has been archived} -subject:{accepted OR zugesagt OR abgelehnt OR tentative OR weiterleitungsbenachrichtigung OR nachrichtenrückruf OR freigabeanfrage} -filename:ics)
    
    // Scan for old threads
    var threads = GmailApp.search(query, 0, 10); 
    /*
    "(older_than:2y " +
    "-label:" + labelProcessed_Name.replace(/[\s]/g, "-") + " " +
    "-{This message has been archived} " +
    "-{Accepted OR Zugesagt OR Abgelehnt OR Tentative OR Weiterleitungsbenachrichtigung OR Nachrichtenrückruf OR Freigabeanfrage} " +
    "-filename:ics)", 0, 10);  
    */
    
    for (var i = 0; i < threads.length; i++)
    {  
      processMail_(folder, threads[i], true, true);
      threads[i].addLabel(labelProcessed);
      Utilities.sleep(1000);
    }
  } catch (e) {
    e = (typeof e === 'string') ? new Error(e) : e;
    Logger.setCaller(arguments.callee.name).severe(e);
    throw e;
  } 
}

/*********************************************************************************************************
    Private Functions
*********************************************************************************************************/
//Saved mails to a Google Drive Folder
function Gmail_Save_as_PDF_(folder, label, includeBody, includeAttachments, labelProcessed)
{
  var threads = label.getThreads(0, 10); //.getThreads();      
  for (var i = 0; i < threads.length; i++)
  {  
    label.removeFromThread(threads[i]);
    processMail_(folder, threads[i], includeBody, includeAttachments);
    
    if (labelProcessed != null)
    {
      threads[i].addLabel(labelProcessed);
    }
    Utilities.sleep(1000);
  }
}

function processMail_(folder, thread, includeBody, includeAttachments){
    var subject = thread.getFirstMessageSubject();
    if (!subject.trim()) {
      var d = subject.getDate();
      subject =  d.getFullYear() +"-" + d.getMonth() + "-" + d.getDate() + " " + d.getHours() + '-' + d.getMinutes();
    }
    Logger.setCaller(arguments.callee.name).log('process mail "' + subject + '"');
    var subject_folder = driveHelper.getFolder(subject, folder);
    
    //var addLogo = '<img alt="GLOBALFOUNDRIES-Mail" height="24" width="143" src="data:image/gif;base64,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" >';
    mailHelper.saveThread(thread, subject_folder, includeBody, includeAttachments);
    //var labels = thread.getLabels()
    
    if (subject_folder != null) {
    
      var itemCount = driveHelper.getFileCount(subject_folder,1);
      //Logger.setCaller(arguments.callee.name).log(itemCount);
      
      if(itemCount==1)
      {
        driveHelper.moveFolderContent(subject_folder, folder);
        subject_folder.setTrashed(true);
        Utilities.sleep(1000);
      } else if(itemCount==0) {
        subject_folder.setTrashed(true);
        Utilities.sleep(1000);
      }
    }
}
