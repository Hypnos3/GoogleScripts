/*
    !! Please activate only if you know what you doing !! 
 
    Auto-Save your Gmail Image Attachments to Google Drive
    ======================================================
 
    Written by Amit Agarwal on 05/28/2013 
    Rewritten by Robert Gester 11/24/2014
    
    To setup script, choose Run -> Setup_Save_Image_Attachments. 
    
    The default Google Drive folder for saving the image
    attachments is "GMail/Images" and once the message has
    been processed, Gmail applies the label "z/processed Images" to 
    that message.

    original source: http://ctrlq.org/code/19696-emails-awaiting-reply
*/
 
// This will auto-save the image attachments from Gmail to Google Drive
function Exec_Gmail_Save_Image_Attachments()
{
  initLog('SaveImageAttachments',arguments.callee.name);
  try
  {
  
    var folderMain  = driveHelper.getFolder(PropertiesService.getUserProperties().getProperty('saveMailImagesDriveFolder'));
    var label_name = PropertiesService.getUserProperties().getProperty('saveMailImagesLabelProcessed');
    var label_name_skipped = PropertiesService.getUserProperties().getProperty('saveMailImagesLabelSkipped');
    var searchQuery = PropertiesService.getUserProperties().getProperty('saveMailImagesSearch');
    
    if (!label_name)
    {
      throw 'Run Setup first!!';
    }
    
    var label = GmailApp.getUserLabelByName(label_name);   
    if ( ! label ) {
      label = GmailApp.createLabel(label_name);
    }
    
    var labelskipped = GmailApp.getUserLabelByName(label_name_skipped);   
    if ( ! labelskipped ) {
      labelskipped = GmailApp.createLabel(label_name_skipped);
    }
    
    // Scan for threads that have image attachments
    var threads = GmailApp.search(searchQuery +" -in:" + label_name + " -in:" + label_name_skipped, 0, 10);      
    
    //  try {    
    for (var x=0; x<threads.length; x++)
    {  
      var messages = threads[x].getMessages();
      var message = messages[0];
      var subject = message.getSubject();
      var folder = null;
      var skipped = true;
      
      for (var y=0; y<messages.length; y++) {
        
        var attachments = messages[y].getAttachments();
        
        for (var z=0; z<attachments.length; z++) {
          
          var file = attachments[z];
          
          // Only save image attachments that have the MIME type as image.
          if (file.getContentType().match(/image/gi)) {
            if (file.getSize() > 20000)            
            {
              if (folder == null)
              {                
                folder = driveHelper.getFolder(folderMain.getName() + '/' + subject);
              }
              folder.createFile(file);
              skipped=false;
              Utilities.sleep(500);
            }
          }          
        }       
      }
      // Process messages are labelled to skip them in the next iteration.
      if (skipped)
      {
        threads[x].addLabel(labelskipped);
      }
      else
      {
        threads[x].addLabel(label);
      }
      if(folder != null && folder.getFiles().length==0)
      {
        folder.setTrashed(true);
        Utilities.sleep(1000);
      }
    }
  } catch (e) {
    e = (typeof e === 'string') ? new Error(e) : e;
    Logger.severe(e);
    throw e;
  }
}
