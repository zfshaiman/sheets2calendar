/**
 * Automated version of syncCalendarEvents that runs without user prompts
 * Can be called from time-based triggers
 */
function automatedSyncCalendarEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Find the most recent "Week of" sheet
  const sheet = findMostRecentWeekSheet(ss);
  
  // Setup error logging
  const errorSheet = getOrCreateSheet(ss, 'SyncErrors');
  const logSheet = getOrCreateSheet(ss, 'SyncLog');
  
  if (!sheet) {
    logError(errorSheet, 'Sheet Error', 'No valid "Week of" sheet found in the spreadsheet');
    logSheet.appendRow([new Date(), 'Sync aborted', 'No valid sheet found']);
    return;
  }
  
  // Log sync start with selected sheet
  logSheet.appendRow([new Date(), 'Sync started', 'Selected sheet: ' + sheet.getName()]);
  
  // Load configuration and data
  loadConfiguration();
  const sendInvites = String(config['SEND_INVITES']).toLowerCase() === 'true';
  const proceedWithConflicts = String(config['PROCEED_WITH_CONFLICTS'] || 'true').toLowerCase() === 'true';
  const events = getSheetData(sheet);
  const headers = sheet.getDataRange().getValues()[0];
  const calendarEventIdCol = headers.indexOf("CalendarEventId");

  // --- Conflict Check ---
  const conflicts = [];
  const roomMap = {};
  const staffMap = {};
  
  events.forEach(event => {
    if (!event.topic || event.topic.toUpperCase().includes('LUNCH')) return;
    
    // Create time slot keys
    const dateKey = event.date.toISOString().split('T')[0];
    const start = event.startTime.toLocaleTimeString('en-US', {hour: '2-digit', minute: '2-digit'});
    const end = event.endTime.toLocaleTimeString('en-US', {hour: '2-digit', minute: '2-digit'});
    const timeKey = `${start}-${end}`;

    // Check room conflicts
    if (event.room) {
      const roomKey = `${event.room}@${dateKey}@${timeKey}`;
      if (roomMap[roomKey]) {
        conflicts.push(`ROOM: ${event.room} double-booked at ${dateKey} ${timeKey} (Rows ${roomMap[roomKey]} & ${event.row})`);
      } else {
        roomMap[roomKey] = event.row;
      }
    }

    // Check staff conflicts
    if (event.clinicalStaff) {
      const staffKey = `${event.clinicalStaff}@${dateKey}@${timeKey}`;
      if (staffMap[staffKey]) {
        conflicts.push(`STAFF: ${event.clinicalStaff} double-booked at ${dateKey} ${timeKey} (Rows ${staffMap[staffKey]} & ${event.row})`);
      } else {
        staffMap[staffKey] = event.row;
      }
    }
  });

  if (conflicts.length > 0) {
    logError(errorSheet, 'Conflicts', `${conflicts.length} Conflicts Found: ${conflicts.join('; ')}`);
    
    // Check if we should abort on conflicts
    if (!proceedWithConflicts) {
      logSheet.appendRow([new Date(), 'Sync aborted', 'Due to conflicts']);
      return;
    }
  }
  // --- End Conflict Check ---

  // Process events
  let created = 0, updated = 0, errors = 0, skipped = 0;
  const errorLog = [];

  events.forEach((event, index) => {
    try {
      if (!event.topic || event.topic.toUpperCase().includes('LUNCH')) {
        skipped++;
        return;
      }

      // Build event details
      const start = new Date(event.date);
      start.setHours(event.startTime.getHours(), event.startTime.getMinutes());

      const end = new Date(event.date);
      end.setHours(event.endTime.getHours(), event.endTime.getMinutes());

      const telehealthTag = (event.telehealth === true || event.telehealth === 1 || 
                             String(event.telehealth).toUpperCase() === "TRUE" || 
                             String(event.telehealth) === "1") ? " [TELEHEALTH]" : "";
      
      const title = `${event.program}: ${event.clinicalStaff || ''}${event.peerSupport ? ' & ' + event.peerSupport : ''} - ${event.topic || "Group"}${telehealthTag}`;

      const description = createEventDescription(event);
      const location = event.room || '';    

      const existingEventId = calendarEventIdCol >= 0 ? 
        sheet.getRange(event.row, calendarEventIdCol + 1).getValue() : null;

      // Get calendar ID for the program
      const calendarId = getCalendarIdForProgram(event.program);
      if (!calendarId) throw new Error(`No calendar configured for ${event.program}`);
      
      const calendar = CalendarApp.getCalendarById(calendarId);
      if (!calendar) throw new Error(`Cannot access calendar ${calendarId}`);

      // Handle existing events
      if (existingEventId) {
        let calendarEvent;
        try {
          calendarEvent = calendar.getEventById(existingEventId);
        } catch (e) {
          sheet.getRange(event.row, calendarEventIdCol + 1).clearContent();
          SpreadsheetApp.flush();
        }

        if (calendarEvent) {
          let needsUpdate = false;
          
          // Existing property checks
          if (calendarEvent.getTitle() !== title) {
            calendarEvent.setTitle(title);
            needsUpdate = true;
          }
          if (calendarEvent.getDescription() !== description) {
            calendarEvent.setDescription(description);
            needsUpdate = true;
          }
          if (calendarEvent.getLocation() !== location) {
            calendarEvent.setLocation(location);
            needsUpdate = true;
          }
          if (calendarEvent.getStartTime().getTime() !== start.getTime() ||
              calendarEvent.getEndTime().getTime() !== end.getTime()) {
            calendarEvent.setTime(start, end);
            needsUpdate = true;
          }

          // Guest list handling
          const newGuests = getEventGuests(event).split(',').filter(e => e);
          const currentGuests = calendarEvent.getGuestList().map(g => g.getEmail());
          
          // Use Sets for accurate comparison
          const currentSet = new Set(currentGuests);
          const newSet = new Set(newGuests);
          
          if (!areSetsEqual(currentSet, newSet)) {
            if (calendarEvent.isRecurringEvent()) {
              // Handle recurring event series
              const eventSeries = calendarEvent.getEventSeries();
              eventSeries.setGuests([]);
                if (newGuests.length > 0) {
                  eventSeries.addGuests(newGuests, {sendInvites: sendInvites});
                }
            } else {
              // Handle single event - surgical updates
              currentGuests.forEach(email => {
                if (!newSet.has(email)) {
                  calendarEvent.removeGuest(email);
                }
              });
              newGuests.forEach(email => {
                if (!currentSet.has(email)) {
                  calendarEvent.addGuest(email, {sendInvites: sendInvites});
                }
              });
            }
            needsUpdate = true;
          }

          needsUpdate ? updated++ : skipped++;
          return;
        }
      }

      // New event creation
      const newEvent = calendar.createEvent(title, start, end, {
        description: description,
        location: location,
        guests: getEventGuests(event),
        sendInvites: sendInvites
      });

      if (calendarEventIdCol >= 0) {
        sheet.getRange(event.row, calendarEventIdCol + 1).setValue(newEvent.getId());
      }
      created++;

    } catch (error) {
      errors++;
      const errorMessage = `Row ${event.row}: ${error.message}`;
      errorLog.push(errorMessage);
      logError(errorSheet, 'Event Error', errorMessage);
    }
  });

  // Log results
  const message = `Results: Created: ${created}, Updated: ${updated}, Skipped: ${skipped}, Errors: ${errors}`;
  logSheet.appendRow([new Date(), 'Sync completed', message, events.length]);
  
  // Log detailed errors if any
  if (errorLog.length > 0) {
    logError(errorSheet, 'Batch Errors', errorLog.join('\n'));
  }
}

/**
 * Finds the most recent "Week of" sheet in the spreadsheet
 * @param {Spreadsheet} ss - The spreadsheet to search in
 * @return {Sheet|null} The most recent "Week of" sheet or null if none found
 */
function findMostRecentWeekSheet(ss) {
  const sheets = ss.getSheets();
  const weekSheets = [];
  
  // Filter sheets that start with "Week of"
  for (let i = 0; i < sheets.length; i++) {
    const sheetName = sheets[i].getName();
    if (sheetName.startsWith("Week of ")) {
      try {
        // Extract date from sheet name
        const dateString = sheetName.replace("Week of ", "");
        const sheetDate = new Date(dateString);
        
        if (!isNaN(sheetDate.getTime())) {
          weekSheets.push({
            sheet: sheets[i],
            date: sheetDate
          });
        }
      } catch (e) {
        // Skip sheets with unparseable dates
        continue;
      }
    }
  }
  
  // Sort sheets by date (most recent first)
  weekSheets.sort((a, b) => b.date - a.date);
  
  // Return the most recent week sheet
  return weekSheets.length > 0 ? weekSheets[0].sheet : null;
}

/**
 * Helper function to get or create a sheet
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    
    // Setup headers based on sheet type
    if (sheetName === 'SyncErrors') {
      sheet.appendRow(['Timestamp', 'Error Type', 'Error Details']);
    } else if (sheetName === 'SyncLog') {
      sheet.appendRow(['Timestamp', 'Action', 'Details', 'Event Count']);
    }
    
    // Format headers
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
  }
  return sheet;
}

/**
 * Helper function to log errors
 */
function logError(errorSheet, errorType, errorDetails) {
  errorSheet.appendRow([new Date(), errorType, errorDetails]);
  console.error(`${errorType}: ${errorDetails}`);
}

/**
 * Creates time-driven trigger for automated sync
 * Run this once manually to set up the trigger
 */
/*
function createAutoSyncTrigger() {
  // Delete any existing triggers for this function
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'automatedSyncCalendarEvents') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create a new trigger to run daily at 6 AM
  ScriptApp.newTrigger('automatedSyncCalendarEvents')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();
    
  // Log trigger creation
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = getOrCreateSheet(ss, 'SyncLog');
  logSheet.appendRow([new Date(), 'Trigger created', 'Daily at 6 AM']);
}
*/

/**
 * Helper function to check if two sets are equal
 */
function areSetsEqual(a, b) {
  return a.size === b.size && [...a].every(v => b.has(v));
}
