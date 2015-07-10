/*
    !! Please activate only if you know what you doing !! 
 
    Auto-Save Gmail message as PDF in Google Drive
    ==============================================
 
    Written by Robert Gester 11/24/2014
    
    To setup script, choose Run -> Setup_Delete_Old. 
    
    valid valued for time:
    2y for 2 years

    d for days, m for month and y for years
*/

//var ItemsDeleteOld = [{label: "z/Delete2Days", time: "2d"}];

/*********************************************************************************************************
    Public Functions
*********************************************************************************************************/

function Exec_Gmail_Delete_Old(){
  initLog('DeleteOldMail',arguments.callee.name);
  try
  {
    //  Gmail_Delete_Old_(driveHelper.getFolder("GMail/PDF Storage"), "x/Save As PDF");
    //  Gmail_Delete_Old_(driveHelper.getFolder("GMail/Meeting Minutes"), "z/Save Minutes As PDF");
    // Gmail_Save_Mail_Attachment_(driveHelper.getFolder("GMail/Attachments"), "x/Save Attachments");
    
    Logger.finest("Exec_Gmail_Delete_Old");  
    
    /*
    for(var i = 0; i < ItemsDeleteOld.length; i++)
    {
    var lbl = ItemsDeleteOld[i].label;
    var time = ItemsDeleteOld[i].time;
    
    var label = GmailApp.getUserLabelByName(lbl);  
    if(label == null){
    label = GmailApp.createLabel(lbl);
    }
    */  
    
    
    var searchQuery = PropertiesService.getUserProperties().getProperty('deleteOldMailSearchQuery');
    if (!searchQuery)
    {
      throw 'Run Setup first!!';
    }
    
    // Scan for old threads
    var threads = GmailApp.search(searchQuery);
    // GmailApp.search("(older_than:" + time + " " + "label:" + lbl.replace(/[\s]/g, "-"), 0, 10);  
    
    if (threads)
    {
      for (var i = 0; i < threads.length; i++)
      {  
        Logger.log("delete " + threads[i].getFirstMessageSubject());
        threads[i].moveToTrash();
        Utilities.sleep(1000);
      }
    }
  } catch (e) {
    e = (typeof e === 'string') ? new Error(e) : e;
    Logger.severe(e);
    throw e;
  } 
}
