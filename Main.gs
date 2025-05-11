// Treatment Schedule System
// Dependent sheets: Config, PeerSupportEmails, ClinicalStaffEmails
const VERSION = '3.4';
// Version Control:
// 3.4 - fixed staffAvailabilityAssistant
// 3.3 - split code into dedicated script files for various feature set, updated syncCalendarEvents guest handling
// 3.2 - started notification backbone
// 3.1 - cleaned up unused functions
// 3.0 - removed AnalyticsDashboard, split into CounselorWorkload and TopicStats
// 2.9 - updated getDashboardData for topic distribution by level of care
// 2.8 - Clarified syncCalendarEvents logic, removed unused functions
// 2.7 - syncCalendarEvents, updated createEvents(events) to update GCal to reflect changes to existing events by id
// 2.6 - Created WelcomeDialog.html, added reference to script
// 2.5 - Removed ValidationDashboard, replaced with AnalyticsDashboard, added welcome message
// 2.4 - Sync/update events via CalendarEventId
// 2.3 - clearCalendarEvents() and html dialog (work in progress)
// 2.2 - Created ValidationDashboard.html (work in progress)
// 2.1 - Added security scopes and version tracking
// 2.0 - Staff coverage assistant with new Topics sheet integration
// 1.5 - Basic calendar event creation

const CACHE = CacheService.getScriptCache();

// Global variables for configuration
let config = {};
let staffEmails = {};
let peerSupportEmails = {};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Schedule Tools')
    .addItem('🔎 Preview Sync Changes', 'previewCalendarSync')
    .addItem('📅 Sync Calendar Events', 'syncCalendarEvents')
    //.addItem('🧹 Clear Calendar Events (select program)', 'showClearCalendarEventsDialog')
    .addSeparator()
    .addItem('🔃 Staff Availability Assistant', 'showStaffCoverageAssistant')
	  .addItem('👥 Counselor Workload', 'showCounselorWorkload')
    .addItem('📊 Topic Distribution', 'showTopicStats')
    .addItem('🚫 View Active Conflicts', 'showConflictSidebar')
    .addSeparator()
    .addItem('⚙️ Version Info', 'showVersionInfo')
    .addItem('❓ Help / Welcome', 'showWelcomeHelp')
    .addToUi();
  //showWelcomeDialog();
}

function showVersionInfo() {
  const props = PropertiesService.getScriptProperties();
  const lastDeployed = props.getProperty('LAST_DEPLOYED') || 'Not recorded';
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(`Treatment Schedule System\nVersion: ${VERSION}\nLast Deployed: ${lastDeployed}`);
}

function showWelcomeDialog(force) {
  let userKey;
  try {
    const userEmail = Session.getActiveUser().getEmail();
    userKey = "HIDE_WELCOME_" + userEmail;
  } catch (e) {
    // Fallback for non-Google accounts
    userKey = "HIDE_WELCOME_ANONYMOUS";
  }

  const ps = PropertiesService.getUserProperties();
  const hideWelcome = ps.getProperty(userKey);

  if (force || !hideWelcome || hideWelcome === 'false') {
    const html = HtmlService.createHtmlOutputFromFile('WelcomeDialog')
      .setWidth(450)
      .setHeight(340);
    SpreadsheetApp.getUi().showModalDialog(html, "Welcome");
  }
}

// Called from HTML
function setHideWelcome(hide) {
  const userEmail = Session.getActiveUser().getEmail();
  const ps = PropertiesService.getUserProperties();
  const hideKey = "HIDE_WELCOME_" + userEmail;
  ps.setProperty(hideKey, hide ? 'true' : 'false');
}

// Helper for help menu item (always forces dialog)
function showWelcomeHelp() {
  showWelcomeDialog(true);
}

/**
 * Main function that feeds data to the HTML dashboard
 * Called by google.script.run from the HTML
 */
function getCounselorWorkload() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // Validate sheet structure
  if (!sheet.getName().startsWith('Week of')) {
    return { error: 'Please select a "Week of" sheet' };
  }

  // Get sheet data - this function is defined below
  const data = getSheetData(sheet);
  
  // --- Staff Workload ---
  // Count groups per counselor, excluding LUNCH BREAK sessions
  const staffCounts = {};
  data.forEach(event => {
    // Only count valid sessions (has clinical staff and isn't LUNCH BREAK)
    if (event.clinicalStaff && 
        event.topic && 
        !event.topic.toUpperCase().includes('LUNCH')) {
      
      staffCounts[event.clinicalStaff] = (staffCounts[event.clinicalStaff] || 0) + 1;
    }
  });

  // --- Date Range ---
  const dates = data
    .map(e => e.date instanceof Date ? e.date.getTime() : null)
    .filter(Boolean);
  
  const startDate = dates.length ? 
    Utilities.formatDate(new Date(Math.min(...dates)), Session.getScriptTimeZone(), "MMM d") : 'N/A';
  const endDate = dates.length ? 
    Utilities.formatDate(new Date(Math.max(...dates)), Session.getScriptTimeZone(), "MMM d") : 'N/A';

  // --- Detect Conflicts ---
  const conflicts = detectConflicts(data);

  // Return the formatted data
  return {
    staffData: Object.entries(staffCounts).sort((a,b) => b[1] - a[1]),
    summary: {
      totalGroups: data.filter(e => e.topic && !e.topic.toUpperCase().includes('LUNCH')).length,
      uniqueRooms: new Set(data.map(e => e.room).filter(Boolean)).size,
      uniqueStaff: Object.keys(staffCounts).length
    },
    conflicts: conflicts,
    startDate: startDate,
    endDate: endDate
  };
}

// Function to get topic stats data for the dashboard
function getTopicStatsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const topicStatsSheet = ss.getSheetByName('TopicStats');
  
  if (!topicStatsSheet) {
    return [];
  }
  
  const data = topicStatsSheet.getDataRange().getValues();
  
  // Skip header row
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { // Skip empty rows
      result.push({
        program: data[i][0],
        topic: data[i][1],
        week: data[i][2],
        count: data[i][3]
      });
    }
  }
  
  return result;
}

/** --- ERROR HANDLING --- **/
function errorHandler(e) {
  const adminEmail = config['ADMIN_EMAIL'];
  if (adminEmail) {
    MailApp.sendEmail(adminEmail,
      'Script Error Alert',
      `Version: ${VERSION}\nError: ${JSON.stringify(e)}`
    );
  }
}

function setupTriggers() {
  // Remove existing triggers
  ScriptApp.getProjectTriggers().forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });

  // Create new triggers
  ScriptApp.newTrigger('errorHandler')
    .forUser(Session.getEffectiveUser())
    .onFailure()
    .create();

  // Record deployment time
  PropertiesService.getScriptProperties()
    .setProperty('LAST_DEPLOYED', new Date().toISOString());
}

function getConfig() {
  // If you already have a `config` object populated at the top of your code, just return it
  if (typeof config !== "undefined" && Object.keys(config).length > 0) {
    return config;
  }

  // Otherwise, load from the "Config" sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) throw new Error('Config sheet not found');
  const configData = configSheet.getDataRange().getValues();

  let cfg = {};
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0]) {
      cfg[configData[i][0]] = configData[i][1];
    }
  }
  // Optionally populate the global config variable for session-wide use
  config = cfg;
  return cfg;
}

/**
 * Loads configuration from the Config sheet and staff emails
 */
function loadConfiguration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const staffEmailSheet = ss.getSheetByName('EmailsClinicalStaff');
  const peerSupportEmailSheet = ss.getSheetByName('EmailsPeerSupport');
  
  // Load config
  const configData = configSheet.getDataRange().getValues();
  config = {};
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0]) {
      config[configData[i][0]] = configData[i][1];
    }
  }
  
  // Load staff emails
  const staffData = staffEmailSheet.getDataRange().getValues();
  staffEmails = {};
  for (let i = 1; i < staffData.length; i++) {
    if (staffData[i][0]) {
      staffEmails[staffData[i][0]] = staffData[i][1];
    }
  }
  
  // Load peer support emails
  const peerData = peerSupportEmailSheet.getDataRange().getValues();
  peerSupportEmails = {};
  for (let i = 1; i < peerData.length; i++) {
    if (peerData[i][0]) {
      peerSupportEmails[peerData[i][0]] = peerData[i][1];
    }
  }
}

/**
 * Creates a detailed description for the calendar event
 */
function createEventDescription(event) {
  const config = getConfig ? getConfig() : {}; // Defensive for external calls like preview
  let description = '';

  description += `Program: ${event.program}\n`;
  if (event.clinicalStaff) description += `Clinical Staff: ${event.clinicalStaff}\n`;
  if (event.peerSupport) description += `Peer Support: ${event.peerSupport}\n`;
  if (event.topic) description += `Topic: ${event.topic}\n`;
  if (event.notes) description += `Notes: ${event.notes}\n`;

  // Always include room information
  if (event.room) description += `Room: ${event.room}\n`;

  // Add telehealth information if applicable (with program/topic-based logic)
  if (event.telehealth) {
    description += '\nTELEHEALTH SESSION\n';

    let zoomId, zoomPasscode, zoomLink;

    // Prefer event/"per-row" Zoom data if present, else fall back to Config
    if (event.topic && event.topic === 'LGBTQ+ & Allies Group') {
      zoomId = event.zoomMeetingID || config['ZOOM_ID_LGBTQ'];
      zoomPasscode = event.zoomMeetingPasscode || config['ZOOM_PASSCODE_LGBTQ'];
      zoomLink = event.zoomLink || config['ZOOM_LINK_LGBTQ'];
    } else if (event.program && event.program.startsWith('MH IOP')) {
      zoomId = event.zoomMeetingID || config['ZOOM_ID_MHIOP'];
      zoomPasscode = event.zoomMeetingPasscode || config['ZOOM_PASSCODE_MHIOP'];
      zoomLink = event.zoomLink || config['ZOOM_LINK_MHIOP'];
    } else if (
      event.program &&
      (event.program.startsWith('SUD IOP') ||
        event.program.startsWith('SUD PC/IOP') ||
        event.program.startsWith('SUD PC only'))
    ) {
      zoomId = event.zoomMeetingID || config['ZOOM_ID_SUDIOP'];
      zoomPasscode = event.zoomMeetingPasscode || config['ZOOM_PASSCODE_SUDIOP'];
      zoomLink = event.zoomLink || config['ZOOM_LINK_SUDIOP'];
    } else if (event.program && event.program.startsWith('SUD OP')) {
      zoomId = event.zoomMeetingID || config['ZOOM_ID_SUDOP'];
      zoomPasscode = event.zoomMeetingPasscode || config['ZOOM_PASSCODE_SUDOP'];
      zoomLink = event.zoomLink || config['ZOOM_LINK_SUDOP'];
    } else {
      // Generic fallback for programs not matching above
      zoomId = event.zoomMeetingID || event.zoomMeetingId || '';
      zoomPasscode = event.zoomMeetingPasscode || '';
      zoomLink = event.zoomLink || '';
    }

    if (zoomId) description += `Zoom Meeting ID: ${zoomId}\n`;
    if (zoomPasscode) description += `Zoom Passcode: ${zoomPasscode}\n`;
    if (zoomLink) description += `Zoom Link: ${zoomLink}\n`;
  }

  description += `\nCreated via Treatment Schedule System v${config.VERSION || VERSION}`;
  return description;
}

/**
 * Formats a date as MM/DD/YYYY
 */
function formatDate(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) {
    return 'Invalid Date';
  }

  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
  
}

/**
 * Formats a time as HH:MM AM/PM
 */
function formatTime(time) {
  return Utilities.formatDate(time, Session.getScriptTimeZone(), 'h:mm a');
}

/**
 * Improved date parser with multiple format support and validation
 */
function parseDate(dateString) {
  const timeZone = Session.getScriptTimeZone();
  
  // Try official Utilities.parseDate first with expected format
  try {
    const parsedDate = Utilities.parseDate(dateString, timeZone, 'MM/dd/yyyy');
    if (parsedDate && !isNaN(parsedDate.getTime())) return parsedDate;
  } catch(e) { /* Continue to fallback methods */ }

  // Fallback 1: Handle dates with 1-digit months/days (e.g. 1/5/2023)
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(dateString)) {
    const [month, day, year] = dateString.split('/').map(Number);
    const date = new Date(year, month - 1, day);
    if (isValidDate(date, month, day, year)) return date;
  }

  // Fallback 2: Handle dates with different delimiters (e.g. 12-31-2023)
  if (/^\d{1,2}[-\/]\d{1,2}[-\/]\d{4}$/.test(dateString)) {
    const cleaned = dateString.replace(/[-\/]/g, '/');
    const date = new Date(cleaned);
    if (!isNaN(date.getTime())) return date;
  }

  // Final fallback: Native Date parsing with validation
  const date = new Date(dateString);
  if (!isNaN(date.getTime())) return date;

  return null;
}

/**
 * Validates date components against JS Date object resolution
 */
function isValidDate(date, originalMonth, originalDay, originalYear) {
  return date.getMonth() + 1 === originalMonth &&
         date.getDate() === originalDay &&
         date.getFullYear() === originalYear;
}


/**
 * Helper to parse time values
 */
function parseTime(timeValue) {
  if (!timeValue) return null;
  
  // If already a date object, return it
  if (timeValue instanceof Date) return timeValue;
  
  // Parse time string into Date object if needed
  if (typeof timeValue === 'string') {
    const [hours, minutes] = timeValue.split(':').map(Number);
    const time = new Date();
    time.setHours(hours, minutes, 0, 0);
    return time;
  }
  
  return null;
}

/**
 * Loads topic information from the Topics sheet
 */
function loadTopics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const topicsSheet = ss.getSheetByName('Topics');
  if (!topicsSheet) return {};
  
  const topicsData = topicsSheet.getDataRange().getValues();
  const headers = topicsData[0];
  
  const topicCol = headers.indexOf('Topic');
  const typeCol = headers.indexOf('Type');
  const specialCol = headers.indexOf('Specialization');
  const staffCol = headers.indexOf('Preferred Staff');
  const notesCol = headers.indexOf('Notes');
  
  if (topicCol === -1) return {};
  
  const topics = {};
  
  for (let i = 1; i < topicsData.length; i++) {
    const row = topicsData[i];
    const topic = row[topicCol];
    
    if (!topic) continue;
    
    topics[topic] = {
      type: row[typeCol] || 'General',
      specialization: specialCol !== -1 ? (row[specialCol] || '') : '',
      preferredStaff: staffCol !== -1 ? (row[staffCol] ? row[staffCol].split(',').map(s => s.trim()) : []) : [],
      notes: notesCol !== -1 ? (row[notesCol] || '') : ''
    };
  }
  
  return topics;
}

/**
 * Loads staff availability data
 */
let staffAvailability = {};

function loadStaffAvailability() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const availSheet = ss.getSheetByName('StaffAvailability');
  const availData = availSheet.getDataRange().getValues();
  
  const headers = availData[0];
  const staffNameCol = headers.indexOf('Staff');
  const maxGroupsCol = headers.indexOf('MaxDailyGroups');
  const roleCol = headers.indexOf('Role');
  
  // Get day columns
  const dayColumns = {
    'MON': { start: headers.indexOf('MON_START'), end: headers.indexOf('MON_END') },
    'TUE': { start: headers.indexOf('TUE_START'), end: headers.indexOf('TUE_END') },
    'WED': { start: headers.indexOf('WED_START'), end: headers.indexOf('WED_END') },
    'THU': { start: headers.indexOf('THU_START'), end: headers.indexOf('THU_END') },
    'FRI': { start: headers.indexOf('FRI_START'), end: headers.indexOf('FRI_END') },
    'SAT': { start: headers.indexOf('SAT_START'), end: headers.indexOf('SAT_END') }
  };
  
  staffAvailability = {};
  
  for (let i = 1; i < availData.length; i++) {
    const row = availData[i];
    const staffName = row[staffNameCol];
    
    if (!staffName || row[roleCol] !== 'Clinical') continue;
    
    const maxGroups = row[maxGroupsCol] || 3; // Default to 3 if not specified
    
    const availability = {};
    for (const day in dayColumns) {
      const startCol = dayColumns[day].start;
      const endCol = dayColumns[day].end;
      
      if (row[startCol] && row[endCol]) {
        // Convert to Date objects regardless of source type
        const startTime = new Date(row[startCol]);
        const endTime = new Date(row[endCol]);

        // If time-only (no date), set to 1970-01-01 base
        if (startTime.getFullYear() === 1970) {
          startTime.setFullYear(1970, 0, 1);
        }
        if (endTime.getFullYear() === 1970) {
          endTime.setFullYear(1970, 0, 1);
        }

        availability[day] = {
          start: startTime,
          end: endTime
        };
      }
    }
    
    staffAvailability[staffName] = {
      maxGroups: maxGroups,
      availability: availability
    };
  }
}

/**
 * Gets all events for a specific staff member on a given date
 */
function getStaffEvents(sheet, staffName, date) {
  // Ensure 'date' is a Date object
  if (!(date instanceof Date)) date = new Date(date);
  
  /*
  if (!(targetDate instanceof Date)) {
    throw new Error('Target date must be Date object');
  }
  */

  const data = getSheetData(sheet);
  const dateString = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  return data.filter(event => {
    // Ensure event.date is a Date
    const eventDate = new Date(event.date);
    const eventDateString = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return eventDateString === dateString && event.clinicalStaff === staffName;
  });
}

/**
 * Gets all events from all Week sheets
 */
function getAllSheetEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let allEvents = [];
  
  for (const sheet of sheets) {
    if (sheet.getName().startsWith('Week of')) {
      const data = getSheetData(sheet);
      allEvents = allEvents.concat(data);
    }
  }
  
  return allEvents;
}

/**
 * Converts a time object to minutes since midnight
 */
function timeToMinutes(time) {
  if (typeof time === 'string') {
    const [hours, minutes] = time.split(':').map(Number);
    return hours * 60 + minutes;
  }
  
  return time.getHours() * 60 + time.getMinutes();
}


function compressData(data) {
  const jsonString = JSON.stringify(data);
  const compressedBlob = Utilities.gzip(Utilities.newBlob(jsonString));
  return Utilities.base64EncodeWebSafe(compressedBlob.getBytes());
}

function decompressData(compressedData) {
  try {
    const decodedBytes = Utilities.base64DecodeWebSafe(compressedData);
    const gzipBlob = Utilities.newBlob(decodedBytes, 'application/gzip');
    const ungzippedBlob = Utilities.ungzip(gzipBlob);
    return JSON.parse(ungzippedBlob.getDataAsString());
  } catch (e) {
    console.error('Decompression error:', e);
    return null;
  }
}

