/*
    !! Please activate only if you know what you doing !! 
 
    Snooze Mails
    ============
 
    Written by Robert Gester 11/24/2014
    
    First setup your whish in line 16 to 18.
    Then setup script, choose Run -> Setup_SnoozeMails.
    
    When you add a Label "Snooze/Snooze x days" the Mail
    will be moved to your inbox after the period of days.
    
    If you have activated SNOOZE_FOR_WEEKS, additional
    labels "Snooze/due x weeks" are generated. Mails to
    this Label will be moved to your inbox only on mondays.
    
    from http://googleappsdeveloper.blogspot.in/2011/07/gmail-snooze-with-apps-script.html
*/


var MARK_UNREAD = true; //false;
var ADD_UNSNOOZED_LABEL = false;
var SNOOZE_FOR_WEEKS = true;

function getLabelName_(i) {
  return "Snooze/Snooze " + i + " days";
}

function getLabelNameWeek_(i) {
  return "Snooze/due " + i + " weeks";
}

function Setup_SnoozeMails() {
  // Create the labels we’ll need for snoozing
  GmailApp.createLabel("Snooze");
  for (var i = 1; i <= 7; ++i) {
    GmailApp.createLabel(getLabelName_(i));
  }

  if (ADD_UNSNOOZED_LABEL) {
    GmailApp.createLabel("Unsnoozed");
  }
  
  var triggers = ScriptApp.getProjectTriggers();
  
  for(var i in triggers) {
    if (triggers[i].getHandlerFunction() === 'Exec_moveSnoozesForDays'
       || triggers[i].getHandlerFunction() === 'Exec_moveSnoozesForWeeks')
    {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger('Exec_moveSnoozesForDays')
   .timeBased()
   .atHour(5)
   .everyDays(1) // Frequency is required if you are using atHour() or nearMinute()
   .create();
  
  Exec_moveSnoozesForDays();

  if (SNOOZE_FOR_WEEKS)
  {    
    for (var i = 2; i <= 6; ++i) {
      GmailApp.createLabel(getLabelNameWeek_(i));
    }  
    
    ScriptApp.newTrigger('Exec_moveSnoozesForWeeks')
    .timeBased()
    .atHour(5)
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .create();
    
    Exec_moveSnoozesForWeeks();
  }
}

function Exec_moveSnoozesForDays() {
  var oldLabel, newLabel, page;
  for (var i = 1; i <= 7; ++i) {
    newLabel = oldLabel;
    oldLabel = GmailApp.getUserLabelByName(getLabelName_(i));
    page = null;
    // Get threads in "pages" of 100 at a time
    while(!page || page.length == 100) {
      page = oldLabel.getThreads(0, 100);
      if (page.length > 0) {
        if (newLabel) {
          // Move the threads into "today’s" label
          newLabel.addToThreads(page);
        } else {
          // Unless it’s time to unsnooze it
          GmailApp.moveThreadsToInbox(page);
          if (MARK_UNREAD) {
            GmailApp.markThreadsUnread(page);
          }
          if (ADD_UNSNOOZED_LABEL) {
            GmailApp.getUserLabelByName("Unsnoozed")
              .addToThreads(page);
          }          
        }     
        // Move the threads out of "yesterday’s" label
        oldLabel.removeFromThreads(page);
      }  
    }
  }
}

function Exec_moveSnoozesForWeeks() {
  var oldLabel, newLabel, page;
  var dayLabel = GmailApp.getUserLabelByName(getLabelName_(7));
  for (var i = 2; i <= 6; ++i) {
    newLabel = oldLabel;
    oldLabel = GmailApp.getUserLabelByName(getLabelNameWeek_(i));
    page = null;
    // Get threads in "pages" of 100 at a time
    while(!page || page.length == 100) {
      page = oldLabel.getThreads(0, 100);
      if (page.length > 0) {
        if (newLabel) {
          // Move the threads into "today’s" label
          newLabel.addToThreads(page);
        } else {
          // Unless it’s time to snooze daily it
          dayLabel.addToThreads(page);
        }     
        // Move the threads out of old label
        oldLabel.removeFromThreads(page);
      }  
    }
  }
}
