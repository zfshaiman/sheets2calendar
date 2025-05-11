// CalendarSync.gs v1.2
// Handles all Google Calendar integration
// Contains: syncCalendarEvents(), createEvents(), updateEvents(), etc.
// Dependencies: Google Apps Scripts services, global variables
// Version History 
// 1.2 - cleaned up a title error, moved constants, updated event handling for single events

/**
 * Main function to sync calendar events from the active sheet
 */
function syncCalendarEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // Validate sheet name
  if (!sheet.getName().startsWith('Week of')) {
    SpreadsheetApp.getUi().alert('❌ Please select a "Week of" sheet first.');
    return;
  }

  // Load configuration and data
  loadConfiguration();
  const sendInvites = String(config['SEND_INVITES']).toLowerCase() === 'true';
  const events = getSheetData(sheet);
  const headers = sheet.getDataRange().getValues()[0];
  const calendarEventIdCol = headers.indexOf("CalendarEventId");
  const ui = SpreadsheetApp.getUi();

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
        conflicts.push(`🚨 ROOM: ${event.room} double-booked at ${dateKey} ${timeKey} (Rows ${roomMap[roomKey]} & ${event.row})`);
      } else {
        roomMap[roomKey] = event.row;
      }
    }

    // Check staff conflicts
    if (event.clinicalStaff) {
      const staffKey = `${event.clinicalStaff}@${dateKey}@${timeKey}`;
      if (staffMap[staffKey]) {
        conflicts.push(`🧑⚕️ STAFF: ${event.clinicalStaff} double-booked at ${dateKey} ${timeKey} (Rows ${staffMap[staffKey]} & ${event.row})`);
      } else {
        staffMap[staffKey] = event.row;
      }
    }
  });

  if (conflicts.length > 0) {
    const proceed = ui.alert(
      `${conflicts.length} Conflicts Found!`,
      `CONFLICTS:\n\n${conflicts.join('\n')}\n\nContinue anyway?`,
      ui.ButtonSet.YES_NO
    );
    if (proceed !== ui.Button.YES) return;
  }
  // --- End Conflict Check ---

  // Rest of original function
  let created = 0, updated = 0, errors = 0, skipped = 0;
  const errorLog = [];

  const response = ui.alert(
    'Calendar Sync',
    `This will process ${events.length} events. Continue?`,
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  events.forEach((event, index) => {
    const row = event.row; // Define at top for error handling
    
    // Build event details
    const start = new Date(event.date);
    start.setHours(event.startTime.getHours(), event.startTime.getMinutes());

    const end = new Date(event.date);
    end.setHours(event.endTime.getHours(), event.endTime.getMinutes());

    const telehealthTag = (event.telehealth === true || event.telehealth === 1 || String(event.telehealth).toUpperCase() === "TRUE" || String(event.telehealth) === "1")
      ? " [TELEHEALTH]" : "";
    const title = `${event.program}: ${event.clinicalStaff || ''}${event.peerSupport ? ' & ' + event.peerSupport : ''} - ${event.topic || "Group"}${telehealthTag}`;

    const description = createEventDescription(event);
    const location = event.room || '';    

    try {
      if (!event.topic || event.topic.toUpperCase().includes('LUNCH')) {
        skipped++;
        return;
      }

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

          // IMPROVED GUEST LIST HANDLING
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
              // For single events, remove and add individually
              currentGuests.forEach(email => {
                if (!newSet.has(email)) {
                  calendarEvent.removeGuest(email);
                }
              });
              newGuests.forEach(email => {
                if (!currentSet.has(email)) {
                  calendarEvent.addGuest(email, {sendInvites: sendInvites}); // Only pass the email string
                }
              });
            }
            needsUpdate = true;
          }

          needsUpdate ? updated++ : skipped++;
          return;
        }
      }

      // New event creation (existing code with config)
      const newEvent = calendar.createEvent(title, start, end, {
        description: description,
        location: location,
        guests: getEventGuests(event),
        sendInvites: sendInvites // From config
      });

      if (calendarEventIdCol >= 0) {
        sheet.getRange(event.row, calendarEventIdCol + 1).setValue(newEvent.getId());
      }
      created++;

    } catch (error) {
      errors++;
      errorLog.push(`Row ${event.row}: ${error.message}`);
    }
  });

  // Show results
  const message = `Results:
  ✅ Created: ${created}
  🔄 Updated: ${updated}
  ⏩ Skipped: ${skipped}
  ❌ Errors: ${errors}`;

  ui.alert('Sync Complete', errorLog.length > 0 ? 
    `${message}\n\nError Details:\n${errorLog.join('\n')}` : message, ui.ButtonSet.OK);
}

function areSetsEqual(a, b) {
  return a.size === b.size && [...a].every(v => b.has(v));
}

/**
 * Creates calendar events based on the provided data
 */
function createEvents(events) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const headers = sheet.getDataRange().getValues()[0];
  let calendarEventIdCol = headers.indexOf("CalendarEventId");
  if (calendarEventIdCol === -1) {
    sheet.getRange(1, headers.length + 1).setValue("CalendarEventId");
    calendarEventIdCol = headers.length; // zero-based
  }

  let createdCount = 0, updatedCount = 0, errorCount = 0, skippedCount = 0;
  const errors = [];
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Confirm Sync",
    `This will create new events for rows without a CalendarEventId, and update existing events for rows with an ID. Continue?`,
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  for (const event of events) {
    try {
      const thisId = sheet.getRange(event.row, calendarEventIdCol + 1).getValue();
      const calendarId = getCalendarIdForProgram(event.program);
      if (!calendarId) {
        errors.push(`No calendar ID found for program: ${event.program} (Row ${event.row})`);
        errorCount++; continue;
      }
      const calendar = CalendarApp.getCalendarById(calendarId);
      if (!calendar) {
        errors.push(`Could not access calendar for ID: ${calendarId} (Row ${event.row})`);
        errorCount++; continue;
      }
      const startDateTime = new Date(event.date);
      startDateTime.setHours(event.startTime.getHours(), event.startTime.getMinutes(), 0, 0);
      const endDateTime = new Date(event.date);
      endDateTime.setHours(event.endTime.getHours(), event.endTime.getMinutes(), 0, 0);

      let title = `${event.program}: ${event.clinicalStaff || ''}${event.peerSupport ? ' & ' + event.peerSupport : ''} - ${event.topic || event.program}`;
      if (event.telehealth) title += " [TELEHEALTH]";
      let description = createEventDescription(event);

      // --- Update existing event ---
      if (thisId) {
        let calEvent = null;
        try { calEvent = calendar.getEventById(thisId); } catch(e) {}
        if (calEvent) {
          // Check if any relevant fields have changed
          let needsUpdate = false;
          if (calEvent.getTitle() !== title) { calEvent.setTitle(title); needsUpdate = true; }
          if (calEvent.getDescription() !== description) { calEvent.setDescription(description); needsUpdate = true; }
          if (calEvent.getLocation() !== (event.room || '')) { calEvent.setLocation(event.room || ''); needsUpdate = true; }
          if (
            calEvent.getStartTime().getTime() !== startDateTime.getTime() ||
            calEvent.getEndTime().getTime() !== endDateTime.getTime()
          ) {
            calEvent.setTime(startDateTime, endDateTime);
            needsUpdate = true;
          }
          if (needsUpdate) updatedCount++; else skippedCount++;
          continue;
        } else {
          // Previous event deleted manually; clear ID and create new
          sheet.getRange(event.row, calendarEventIdCol + 1).clear();
        }
      }

      // --- Create new event if not updated above ---
      const calEvent = calendar.createEvent(title, startDateTime, endDateTime, {
        description: description,
        location: event.room || '',
        guests: getEventGuests(event),
        sendInvites: config['SEND_INVITES'] === 'True' || config['SEND_INVITES'] === true
      });
      sheet.getRange(event.row, calendarEventIdCol + 1).setValue(calEvent.getId());
      createdCount++;
    } catch (error) {
      errors.push(`Error in row ${event.row}: ${error}`);
      errorCount++;
    }
  }

  // Show summary
  let message = `Created: ${createdCount}\nUpdated: ${updatedCount}\nSkipped (no change): ${skippedCount}\nErrors: ${errorCount}`;
  if (errors.length > 0) {
    message += '\n\nError details:\n' + errors.join('\n');
    if (message.length > 8000) message = message.substring(0, 7800) + '\n(...truncated)';
  }
  ui.alert("Calendar Event Sync Summary", message, ui.ButtonSet.OK);
}

function previewCalendarSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (!sheet.getName().startsWith('Week of')) {
    ui.alert('❌ Select a "Week of" sheet first');
    return;
  }

  // Load data and config
  loadConfiguration();
  const sendInvites = config['SEND_INVITES'] === 'True';
  const events = getSheetData(sheet);
  const headers = sheet.getDataRange().getValues()[0];
  const calendarEventIdCol = headers.indexOf("CalendarEventId");

  // Data collection for preview
  const changes = {
    create: [],
    update: [],
    conflicts: [],
    errors: []
  };

  // Reuse conflict detection from sync
  const conflicts = detectConflicts(events);
  if (conflicts.length > 0) {
    changes.conflicts = conflicts;
  }

  events.forEach((event, index) => {
    const row = index + 2;
    try {
      if (!event.topic || event.topic.toUpperCase().includes('LUNCH')) return;

      const existingEventId = calendarEventIdCol >= 0 ? 
        sheet.getRange(row, calendarEventIdCol + 1).getValue() : null;

      // Build event details (same as sync)
      const start = new Date(event.date);
      start.setHours(event.startTime.getHours(), event.startTime.getMinutes());
      const end = new Date(event.date);
      end.setHours(event.endTime.getHours(), event.endTime.getMinutes());
      const title = `${event.program}: ${event.clinicalStaff || ''}${event.peerSupport ? ' & ' + event.peerSupport : ''} - ${event.topic || "Group"}`;
      const description = createEventDescription(event);
      const location = event.room || '';

      if (existingEventId) {
        const calendarId = getCalendarIdForProgram(event.program);
        if (!calendarId) throw new Error(`No calendar for ${event.program}`);
        
        const calendar = CalendarApp.getCalendarById(calendarId);
        if (!calendar) throw new Error(`Can't access calendar`);
        
        let calendarEvent;
        try {
          calendarEvent = calendar.getEventById(existingEventId);
        } catch (e) {
          changes.update.push({
            type: '❗ Invalid ID',
            row: row,
            details: `Will clear invalid ID and create new event`
          });
          return;
        }

        if (calendarEvent) {
          const diffs = [];
          
          // Compare properties
          if (calendarEvent.getTitle() !== title) diffs.push('Title');
          if (calendarEvent.getDescription() !== description) diffs.push('Description');
          if (calendarEvent.getLocation() !== location) diffs.push('Location');
          if (calendarEvent.getStartTime().getTime() !== start.getTime() || 
              calendarEvent.getEndTime().getTime() !== end.getTime()) diffs.push('Time');

          // Compare guests
          const newGuests = getEventGuests(event).split(',').filter(e => e);
          const currentGuests = calendarEvent.getGuestList().map(g => g.getEmail());
          const currentSet = new Set(currentGuests);
          const newSet = new Set(newGuests);
          
          if (!areSetsEqual(currentSet, newSet)) {
            diffs.push('Guests');
          }

          if (diffs.length > 0) {
            changes.update.push({
              type: '✏️ Update',
              row: row,
              title: title,
              changes: diffs.join(', ')
            });
          }
        }
      } else {
        changes.create.push({
          type: '🆕 Create',
          row: row,
          title: title,
          time: Utilities.formatDate(start, Session.getScriptTimeZone(), 'MMM d h:mm a')
        });
      }
    } catch (error) {
      changes.errors.push(`Row ${row}: ${error.message}`);
    }
  });

  // Build HTML report
  const html = HtmlService.createHtmlOutput(`
    <style>
      .section { margin: 1rem 0; padding: 1rem; border: 1px solid #ddd; border-radius: 4px; }
      .create { color: #4CAF50; }
      .update { color: #FF9800; }
      .conflict { color: #F44336; }
      .error { color: #9E9E9E; }
    </style>
    
    <div class="section">
      <h3>📝 Would Create (${changes.create.length})</h3>
      ${changes.create.map(e => `<div class="create">${e.type} Row ${e.row}: ${e.title} @ ${e.time}</div>`).join('')}
      ${changes.create.length === 0 ? '<em>No new events</em>' : ''}
    </div>

    <div class="section">
      <h3>✏️ Would Update (${changes.update.length})</h3>
      ${changes.update.map(e => `<div class="update">${e.type} Row ${e.row}: ${e.title}<br>Changes: ${e.changes}</div>`).join('')}
      ${changes.update.length === 0 ? '<em>No updates</em>' : ''}
    </div>

    <div class="section">
      <h3>🚨 Conflicts (${changes.conflicts.length})</h3>
      ${changes.conflicts.map(c => `
        <div class="conflict">
          ${typeof c === 'object' ? `${c.type}: ${c.message}` : c}
        </div>
      `).join('')}
      ${changes.conflicts.length === 0 ? '<em>No conflicts</em>' : ''}
    </div>


    ${changes.errors.length > 0 ? `
    <div class="section">
      <h3>⚠️ Potential Errors (${changes.errors.length})</h3>
      ${changes.errors.map(e => `<div class="error">${e}</div>`).join('')}
    </div>` : ''}
  `)
  .setWidth(800)
  .setHeight(600);

  ui.showModalDialog(html, 'Sync Preview');
}


/**
 * Updates existing calendar events when sheet data changes
 */
function updateEvents(events) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const headers = sheet.getDataRange().getValues()[0];
  const calendarEventIdCol = headers.indexOf("CalendarEventId");
  const ui = SpreadsheetApp.getUi();

  let updatedCount = 0;
  const updates = [];

  for (const event of events) {
    try {
      const rowEventId = sheet.getRange(event.row, calendarEventIdCol + 1).getValue();
      if (!rowEventId) continue; // Skip new events
      
      const calendarId = getCalendarIdForProgram(event.program);
      if (!calendarId) continue;
      
      const calendar = CalendarApp.getCalendarById(calendarId);
      if (!calendar) continue;

      // Get existing event
      const existingEvent = calendar.getEventById(rowEventId);
      if (!existingEvent) {
        sheet.getRange(event.row, calendarEventIdCol + 1).clear(); // Clear invalid ID
        continue;
      }

      // Determine if changes are needed
      const newStart = new Date(event.date);
      newStart.setHours(event.startTime.getHours(), event.startTime.getMinutes());
      const newEnd = new Date(event.date);
      newEnd.setHours(event.endTime.getHours(), event.endTime.getMinutes());
      
      const newTitle = buildEventTitle(event);
      const newLocation = event.room || '';
      const newDescription = createEventDescription(event);

      // Check for changes
      let needsUpdate = false;
      
      if (existingEvent.getTitle() !== newTitle) {
        existingEvent.setTitle(newTitle);
        needsUpdate = true;
      }
      
      if (existingEvent.getDescription() !== newDescription) {
        existingEvent.setDescription(newDescription);
        needsUpdate = true;
      }
      
      if (existingEvent.getLocation() !== newLocation) {
        existingEvent.setLocation(newLocation);
        needsUpdate = true;
      }
      
      if (existingEvent.getStartTime().getTime() !== newStart.getTime() || 
          existingEvent.getEndTime().getTime() !== newEnd.getTime()) {
        existingEvent.setTime(newStart, newEnd);
        needsUpdate = true;
      }

      if (needsUpdate) {
        updatedCount++;
        updates.push(`Row ${event.row}: Updated "${newTitle}"`);
      }
      
    } catch (error) {
      updates.push(`Error in row ${event.row}: ${error}`);
    }
  }
  
  // Show summary
  let message = `Updated ${updatedCount} events`;
  if (updates.length > 0) {
    message += '\n\nChanges:\n' + updates.join('\n');
    if (message.length > 8000) message = message.substring(0, 7800) + '\n(...truncated)';
  }
  ui.alert('Event Update Summary', message, ui.ButtonSet.OK);
}

/**
 * Gets the appropriate calendar ID for a program
 */
function getCalendarIdForProgram(program) {
  if (program.startsWith('MH PC')) {
    return config['PROGRAM_CALENDARS_MH_PC'];
  } else if (program.startsWith('MH IOP')) {
    return config['PROGRAM_CALENDARS_MH_IOP'];
  } else if (program.startsWith('SUD PC/IOP')) {
    return config['PROGRAM_CALENDARS_SUD_PC_IOP'];
  } else if (program.startsWith('SUD PC only')) {
    return config['PROGRAM_CALENDARS_SUD_PC_ONLY'];
  } else if (program.startsWith('MH OP')) {
    return config['PROGRAM_CALENDARS_MH_OP'];
  } else if (program.startsWith('SUD IOP')) {
    return config['PROGRAM_CALENDARS_SUD_IOP'];
  } else if (program.startsWith('SUD OP')) {
    return config['PROGRAM_CALENDARS_SUD_OP'];
  }
  
  return null;
}

/**
 * Gets email addresses for event guests
 */
function getEventGuests(event) {
  const guests = [];
  if (event.clinicalStaff && staffEmails[event.clinicalStaff]) {
    guests.push(staffEmails[event.clinicalStaff]);
  }
  if (event.peerSupport && peerSupportEmails[event.peerSupport]) {
    guests.push(peerSupportEmails[event.peerSupport]);
  }
  return guests.join(','); // Join array into a comma-separated string
}

/**
 * Gets the data from the sheet and formats it for processing
 */
function getSheetData(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const dateCol = headers.indexOf('Date');
  const programCol = headers.indexOf('Program');
  const startTimeCol = headers.indexOf('StartTime');
  const endTimeCol = headers.indexOf('EndTime');
  const roomCol = headers.indexOf('Room');
  const clinicalStaffCol = headers.indexOf('ClinicalStaff');
  const peerSupportCol = headers.indexOf('PeerSupport');
  const topicCol = headers.indexOf('Topic');
  const notesCol = headers.indexOf('Notes');
  const telehealthCol = headers.indexOf('Telehealth');

  const events = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Skip empty rows
    if (!row[dateCol] || !row[programCol]) continue;

    let startTime = row[startTimeCol];
    let endTime = row[endTimeCol];
    // Convert date string to Date object if needed
    let eventDate = row[dateCol];
    if (typeof eventDate === 'string') {
      eventDate = new Date(eventDate); // Add this line
    }

    // Parse time strings into Date objects if needed
    if (typeof startTime === 'string') {
      const [hours, minutes] = startTime.split(':').map(Number);
      startTime = new Date();
      startTime.setHours(hours, minutes, 0, 0);
    }
    if (typeof endTime === 'string') {
      const [hours, minutes] = endTime.split(':').map(Number);
      endTime = new Date();
      endTime.setHours(hours, minutes, 0, 0);
    }

    // Telehealth as boolean
    const telehealth = (
      row[telehealthCol] === true ||
      row[telehealthCol] === 'True' ||
      row[telehealthCol] === 'TRUE' ||
      row[telehealthCol] === 1 ||
      row[telehealthCol] === '1' ||
      row[telehealthCol] === 1.0 ||
      row[telehealthCol] === '1.0'
    );

    events.push({
      date: eventDate,
      program: row[programCol],
      startTime: startTime,
      endTime: endTime,
      room: row[roomCol],
      clinicalStaff: row[clinicalStaffCol],
      peerSupport: row[peerSupportCol],
      topic: row[topicCol],
      notes: row[notesCol],
      telehealth: telehealth,
      row: i + 1
    });
  }
  return events;
}
