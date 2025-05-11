// Conflicts.gs v1.0
// Handles all Google Calendar integration
// Dependencies: getSheetData(), loadTopics(), getAllSheetEvents(), staffAvailability

/**
 * Detects conflicts in room bookings and staff assignments
 */
function detectConflicts(events) {
  const conflicts = [];
  
  // Group events by date for easier processing
  const eventsByDate = {};
  for (const event of events) {
    const dateString = Utilities.formatDate(event.date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!eventsByDate[dateString]) {
      eventsByDate[dateString] = [];
    }
    eventsByDate[dateString].push(event);
  }
  
  // Check conflicts for each date
  for (const dateString in eventsByDate) {
    const dateEvents = eventsByDate[dateString];
    
    // Check room conflicts
    const roomConflicts = findRoomConflicts(dateEvents);
    conflicts.push(...roomConflicts);
    
    // Check staff conflicts
    const staffConflicts = findStaffConflicts(dateEvents);
    conflicts.push(...staffConflicts);
  }
  
  return conflicts;
}

/**
 * Finds room booking conflicts
 */
function findRoomConflicts(events) {
  const conflicts = [];
  
  for (let i = 0; i < events.length; i++) {
    const event1 = events[i];
    
    // Skip events without rooms or with special rooms like "Experiential" or "Office"
    if (!event1.room || event1.room === "Experiential" || event1.room === "Office") continue;
    
    // Skip lunch breaks
    if (event1.topic === 'LUNCH BREAK') continue;
    
    for (let j = i + 1; j < events.length; j++) {
      const event2 = events[j];
      
      // Skip events without rooms or with special rooms
      if (!event2.room || event2.room === "Experiential" || event2.room === "Office") continue;
      
      // Skip lunch breaks
      if (event2.topic === 'LUNCH BREAK') continue;
      
      // Check if same room
      if (event1.room !== event2.room) continue;
      
      // Check for time overlap
      if (doTimesOverlap(event1.startTime, event1.endTime, event2.startTime, event2.endTime)) {
        conflicts.push({
          type: 'Room',
          message: `Room conflict: ${event1.room} is double-booked on ${formatDate(event1.date)} from ${formatTime(event1.startTime)} to ${formatTime(event1.endTime)} with ${event1.program} and ${event2.program}`,
          rows: [event1.row, event2.row]
        });
      }
    }
  }
  
  return conflicts;
}

/**
 * Finds staff assignment conflicts
 */
function findStaffConflicts(events) {
  const conflicts = [];
  
  for (let i = 0; i < events.length; i++) {
    const event1 = events[i];
    
    // Skip events without clinical staff
    if (!event1.clinicalStaff) continue;
    
    for (let j = i + 1; j < events.length; j++) {
      const event2 = events[j];
      
      // Skip events without clinical staff
      if (!event2.clinicalStaff) continue;
      
      // Check if same staff
      if (event1.clinicalStaff !== event2.clinicalStaff) continue;
      
      // Check for time overlap
      if (doTimesOverlap(event1.startTime, event1.endTime, event2.startTime, event2.endTime)) {
        conflicts.push({
          type: 'Staff',
          message: `Staff conflict: ${event1.clinicalStaff} is double-booked on ${formatDate(event1.date)} from ${formatTime(event1.startTime)} to ${formatTime(event1.endTime)} with ${event1.program} and ${event2.program}`,
          rows: [event1.row, event2.row]
        });
      }
    }
  }
  
  return conflicts;
}

/**
 * Checks if two time periods overlap
 */
function doTimesOverlap(start1, end1, start2, end2) {
  // Convert to minutes for easier comparison
  const start1Mins = start1.getHours() * 60 + start1.getMinutes();
  const end1Mins = end1.getHours() * 60 + end1.getMinutes();
  const start2Mins = start2.getHours() * 60 + start2.getMinutes();
  const end2Mins = end2.getHours() * 60 + end2.getMinutes();
  
  return (start1Mins < end2Mins && end1Mins > start2Mins);
}

function getConflicts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = getSheetData(sheet);
  return detectConflicts(data); // Returns [{type, message, rows}]
}

function highlightRows(rows) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getDataRange().setBackground(null); // Clear highlights
  rows.forEach(row => {
    if (row > 1) sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#FFEBEE');
  });
}

function unhighlightRows(rows) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  rows.forEach(row => {
    if (row > 1) sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground(null);
  });
}

/**
 * Displays conflicts to the user
 */
function displayConflicts(conflicts) {
  let message = 'The following conflicts were detected:\n\n';
  
  for (const conflict of conflicts) {
    message += `${conflict.message}\nRows: ${conflict.rows.join(', ')}\n\n`;
  }
  
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Checks if a staff member has a conflict with the specified event time
 */
function hasScheduleConflict(staffName, date, startTime, endTime, allEvents) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    throw new Error(`Invalid date type received: ${typeof date}. Value: ${date}`);
  }
  
  const scriptTimeZone = Session.getScriptTimeZone();
  const dateString = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  for (const event of allEvents) {
    // Validate event.date
    let eventDate;
    if (typeof event.date === 'string') {
      eventDate = Utilities.parseDate(event.date, scriptTimeZone, 'MM/dd/yyyy');
    } else {
      eventDate = new Date(event.date);
    }

    if (isNaN(eventDate.getTime())) {
      console.error('Skipping invalid event date:', event.date);
      continue;
    }

    // Skip events that are not on the same date
    const eventDateString = Utilities.formatDate(eventDate, scriptTimeZone, 'yyyy-MM-dd');
    if (eventDateString !== dateString) continue;    
    if (!(date instanceof Date) || isNaN(date.getTime())) {
      throw new Error('Invalid date type in conflict check');
    }    
    
    // Skip events that don't involve this staff member
    if (event.clinicalStaff !== staffName) continue;
    
    // Check for time overlap
    if (doTimesOverlap(startTime, endTime, event.startTime, event.endTime)) {
      return {
        hasConflict: true,
        conflictEvent: event
      };
    }
  }
  
  return { hasConflict: false };
}

/**
 * Finds coverage recommendations for a list of events
 */
function findCoverageRecommendations(events) {
  const topics = loadTopics();
  const allEvents = getAllSheetEvents();
  const recommendations = [];
  
  for (const event of events) {
    const eventDayOfWeek = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'][event.date.getDay()];
    const eventStart = event.startTime;
    const eventEnd = event.endTime;
    const topicInfo = topics[event.topic] || { 
      type: 'General',
      specialization: '',
      preferredStaff: []
    };
    
    const suitableStaff = [];
    
    // Check each staff member for suitability
    for (const staffName in staffAvailability) {
      // Skip the absent staff member
      if (staffName === event.clinicalStaff) continue;
      
      const staff = staffAvailability[staffName];
      const dayAvail = staff.availability[eventDayOfWeek];
      
      // Initial match information
      let specialtyMatch = false;
      let availabilityMatch = false;
      let hasConflict = false;
      let conflictDetails = null;
      
      // Check if staff is preferred for this topic
      const isPreferred = topicInfo.preferredStaff.includes(staffName);
      
      // Check if staff works on this day
      if (dayAvail) {
        // Convert times to minutes for comparison
        const staffStartMins = timeToMinutes(dayAvail.start);
        const staffEndMins = timeToMinutes(dayAvail.end);
        const eventStartMins = eventStart.getHours() * 60 + eventStart.getMinutes();
        const eventEndMins = eventEnd.getHours() * 60 + eventEnd.getMinutes();
        
        // Check if event is within staff hours
        availabilityMatch = (eventStartMins >= staffStartMins && eventEndMins <= staffEndMins);
      }
      
      // Check if staff is specialized in this topic
      specialtyMatch = isPreferred;
      
      // Check for schedule conflicts
      const conflictCheck = hasScheduleConflict(staffName, event.date, eventStart, eventEnd, allEvents);
      hasConflict = conflictCheck.hasConflict;
      if (hasConflict) {
        conflictDetails = conflictCheck.conflictEvent;
      }
      
      // Calculate match score and description
      let matchDescription = '';
      let matchScore = 0;
      
      if (hasConflict) {
        matchDescription = 'Schedule conflict';
        matchScore = -10;
      } else if (availabilityMatch && specialtyMatch) {
        matchDescription = 'Excellent match';
        matchScore = 10;
      } else if (availabilityMatch) {
        matchDescription = 'Available but not specialized';
        matchScore = 5;
      } else if (specialtyMatch) {
        matchDescription = 'Specialized but outside hours';
        matchScore = 3;
      } else {
        matchDescription = 'Outside regular hours';
        matchScore = 1;
      }
      
      suitableStaff.push({
        name: staffName,
        match: matchDescription,
        matchScore: matchScore,
        specialtyMatch: specialtyMatch,
        availabilityMatch: availabilityMatch,
        hasConflict: hasConflict,
        conflictDetails: conflictDetails
      });
    }
    
    // Sort staff by suitability (highest score first)
    suitableStaff.sort((a, b) => b.matchScore - a.matchScore);
    
    // Take top recommendations, including at least one without conflict if possible
    const topRecommendations = [];
    let nonConflictCount = 0;
    
    for (const staff of suitableStaff) {
      if (!staff.hasConflict) {
        nonConflictCount++;
      }
      
      topRecommendations.push(staff);
      
      if (topRecommendations.length >= 3 && nonConflictCount > 0) {
        break;
      }
    }
    
    recommendations.push({
      event: event,
      recommendations: topRecommendations.slice(0, 3) // Ensure we only take the top 3
    });
  }
  
  return recommendations;
}
