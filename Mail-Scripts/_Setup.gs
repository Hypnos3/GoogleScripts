
// The script will check your Gmail mailbox every minute
// with the help of a CLOCK based trigger.
function SetupAllData()
{
  initLog(Log_Sheet_Name,arguments.callee.name);
  
  //Cleanup Trigger:  
  var triggers = ScriptApp.getProjectTriggers();

  //remove all triggers  
  for(var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  //SicherheitsTrigger, der täglich malwieder den Setup aufruft und prüft:
  ScriptApp.newTrigger(arguments.callee.name)
  	.timeBased()
  	.everyHours(12)
  	.create();  
  
  // generic --------------------------------------------
    checkProperty_('BetterLogLevel', 'INFO');
    PropertiesService.getScriptProperties().setProperty('BetterLogLevel', 'INFO');
  
  // renameMail --------------------------------------------
  checkProperty_('renameMailLabelProcessed', 'z/processed mail');
  checkProperty_('renameMailLabelError', 'z/error process mail');

  checkProperty_('renameMailSearchQueryV2', '(request.application@globalfoundries.com OR globalfoundries@service-now.com OR from:setup.FC) -subject:{redirect} -has:attachment');
  //checkProperty_('renameMailSearchQuery', '(request.application@globalfoundries.com OR globalfoundries@service-now.com OR Request Application OR ServiceNow) -subject:{redirect} -has:attachment');  
  //checkProperty_('renameMailIDRegex', '(\\b[A-Z]{2,}\\d{4,}\\b|\\b[A-Z]+\\d{2,}-\\d{4,}\\sv\\d+\\b|\\b[A-Z]+\\d{2,}-\\d{4,}\\b)'); //EICR1001004
  checkProperty_('renameMailSearchQueryV2Similar', '-subject:((Out of office) OR OOO)');
  
  if (checkProperty_('renameMailActive', 'false') === 'true')
  {
    Logger.log("renameMailActive = true")
    checkLabel_('renameMailLabelProcessed');
    checkLabel_('renameMailLabelError');
    
    //Setup Trigger:  
    ScriptApp.newTrigger('Exec_Gmail_Rename_SystemMails')
    .timeBased()
    .everyMinutes(5)
    .create();
  }  
  
  // saveMailImages --------------------------------------------
  checkProperty_('saveMailImagesLabelProcessed', 'z/processed Images');
  checkProperty_('saveMailImagesLabelSkipped', 'z/skipped Images');
  checkProperty_('saveMailImagesDriveFolder', 'GMail/Images');
  checkProperty_('saveMailImagesSearch', 'in:all -in:spam -in:trash has:attachment filename:jpg OR filename:png OR filename:gif');

  if (checkProperty_('saveMailImagesActive','false') === 'true')
  {  
    Logger.log("saveMailImagesActive = true")
    checkLabel_('saveMailImagesLabelProcessed');
    checkLabel_('saveMailImagesLabelSkipped');
    
    ScriptApp.newTrigger('Exec_Gmail_Save_Image_Attachments')
      .timeBased()
      .everyMinutes(10)
      .create();
  }
  
  // saveMailPDF --------------------------------------------
  /*
  var ItemsDefault = [{label: "x/Save As PDF", folder: "GMail/PDF Storage", mailBody: true, mailAttachment: true, labelProcessed: ""},
                      {label: "x/Save Minutes As PDF", folder: "GMail/Meeting Minutes", mailBody: true, mailAttachment: true, labelProcessed: "z/Meeting Minutes"},
                      {label: "x/Save Attachments", folder: "GMail/Attachments", mailBody: false, mailAttachment: true, labelProcessed: ""}];
                      */

  var ItemsDefault = [{label: "Save As PDF", folder: "GMail/PDF Storage", mailBody: true, mailAttachment: true, label_processed: ""}];
    
  var ItemsSavePDF = checkProperty_('saveMailPDFItems', JSON.stringify(ItemsDefault));
  if (ItemsSavePDF) {
    ItemsSavePDF = JSON.parse(ItemsSavePDF)
    if (ItemsSavePDF.length == 0) {
      ItemsSavePDF = null;
    }
  }
  
  if (!ItemsSavePDF) {
    PropertiesService.getUserProperties().setProperty('saveMailPDFActive','false');
    PropertiesService.getUserProperties().setProperty('saveMailPDFItems', JSON.stringify(ItemsDefault));
  }
  
  if (checkProperty_('saveMailPDFActive','false') === 'true')
  {
    Logger.log("saveMailPDFActive = true")
    ScriptApp.newTrigger('Exec_Gmail_Save_as_PDF')
       .timeBased()
       .everyMinutes(5)
       .create();    
  }
  
  // saveOldMail --------------------------------------------
  checkProperty_('saveOldMailPDFLabelProcessed', "z/processed Mail");
  checkProperty_('saveOldMailPDFFolder', "GMail/PDF Storage");
  checkProperty_('saveOldMailPDFSearch', '(older_than:2y -{This message has been archived} -{Accepted OR Zugesagt OR Abgelehnt OR Tentative OR Weiterleitungsbenachrichtigung OR Nachrichtenrückruf OR Freigabeanfrage}  -filename:ics)');

  if (checkProperty_('saveOldMailPDFActive','false')  === 'true')
  {
    Logger.log("saveOldMailPDFActive = true")
    checkLabel_('saveOldMailPDFLabelProcessed');

    // Schedules for the first of every month
    ScriptApp.newTrigger('Exec_Gmail_Save_old_as_PDF')
    .timeBased()
    .atHour(6)
    .onMonthDay(1)
    .create();
  }
  
  // deleteOldMail --------------------------------------------
  checkProperty_('deleteOldMailSearchQuery', 'older_than:2d label:z/Delete2Days');
  
  if (checkProperty_('deleteOldMailActive','false') === 'true')
  {
    Logger.log("deleteOldMailActive = true")
    ScriptApp.newTrigger('Exec_Gmail_Delete_Old')
        .timeBased()
        .everyDays(1)
        .create();  
  }
}

function checkProperty_(propertyName,defaultValue)
{
  var result = PropertiesService.getUserProperties().getProperty(propertyName);
  if (!result)
  {
    PropertiesService.getUserProperties().setProperty(propertyName,defaultValue);
    return defaultValue;
  }
  return result;
}

function checkLabel_(propertyName)
{
  var labelName = PropertiesService.getUserProperties().getProperty(propertyName);
  if (labelName)
  {
    return mailHelper.getLabel(labelName);
  }
}

