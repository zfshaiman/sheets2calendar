// UI v1.0
// Dependencies: loadConfiguration(), loadStaffAvailability(), getStaffEvents(), findCoverageRecommendations(), displayCoverageRecommendations(), parseDate(), formatDate()

function showConflictSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ConflictSidebar')
    .setTitle('Dashboard')
    .setWidth(800);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showCounselorWorkload() {
  const html = HtmlService.createHtmlOutputFromFile('CounselorWorkload')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Counselor Workload');
}

function showTopicStats() {
  const html = HtmlService.createHtmlOutputFromFile('TopicStats')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Topic Distribution');
}

/**
 * Shows a sidebar with staff coverage recommendations
 */
function showStaffCoverageAssistant() {
  const ui = SpreadsheetApp.getUi();
  
  // Prompt for absent staff member
  const staffResponse = ui.prompt(
    'Staff Coverage Assistant',
    'Enter the name of the absent staff member:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (staffResponse.getSelectedButton() !== ui.Button.OK) return;
  const absentStaff = staffResponse.getResponseText().trim();
  
  // Prompt for absence date
  const dateResponse = ui.prompt(
    'Staff Coverage Assistant',
    'Enter the date of absence (MM/DD/YYYY):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (dateResponse.getSelectedButton() !== ui.Button.OK) return;
  const absenceDate = parseDate(dateResponse.getResponseText().trim());
  
  if (!absenceDate) {
    ui.alert('Invalid date format. Please use MM/DD/YYYY format.');
    return;
  }
  
  // Load configuration and data
  loadConfiguration();
  loadStaffAvailability();
  
  if (!staffAvailability[absentStaff]) {
    ui.alert(`${absentStaff} not found in staff availability data.\nPlease check the StaffAvailability sheet.`);
    return;
  }

  // Get the active sheet
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!sheet.getName().startsWith('Week of')) {
    ui.alert('Please select a "Week of" sheet before running this function.');
    return;
  }
  
  // Get events for the absent staff member
  const events = getStaffEvents(sheet, absentStaff, absenceDate);
  if (events.length === 0) {
    ui.alert(`No events found for ${absentStaff} on ${formatDate(absenceDate)}.`);
    return;
  }
  
  // Find coverage recommendations
  const recommendations = findCoverageRecommendations(events);
  
  // Display recommendations
  displayCoverageRecommendations(absentStaff, absenceDate, recommendations);
}

/**
 * Displays coverage recommendations in a user-friendly format
 */
function displayCoverageRecommendations(absentStaff, date, recommendations) {
  let htmlOutput = HtmlService.createHtmlOutput()
    .setTitle('Staff Coverage Recommendations')
    .setWidth(600)
    .setHeight(500);
  
  let html = `
    <style>
      body { font-family: Arial, sans-serif; margin: 10px; }
      h1 { font-size: 18px; color: #3c78d8; }
      h2 { font-size: 16px; margin-top: 20px; border-bottom: 1px solid #ddd; padding-bottom: 5px; }
      .event { margin-bottom: 20px; padding: 10px; border: 1px solid #ddd; border-radius: 4px; }
      .event-header { font-weight: bold; margin-bottom: 5px; }
      .event-details { margin-bottom: 10px; font-size: 14px; }
      .recommendations { margin-left: 15px; }
      .recommendation { margin-bottom: 10px; padding: 8px; border-left: 3px solid #ddd; }
      .excellent { border-left-color: #6aa84f; }
      .good { border-left-color: #f1c232; }
      .fair { border-left-color: #e69138; }
      .poor { border-left-color: #cc0000; }
      .conflict { border-left-color: #cc0000; background-color: #ffeeee; }
      .conflict-details { font-size: 12px; color: #cc0000; margin-top: 3px; }
      .specialty-match { color: #6aa84f; font-size: 12px; }
      .availability-match { color: #3c78d8; font-size: 12px; }
    </style>
    <h1>Coverage Recommendations for ${absentStaff} on ${formatDate(date)}</h1>
  `;
  
  if (recommendations.length === 0) {
    html += '<p>No events found for this staff member on the selected date.</p>';
  } else {
    recommendations.forEach((rec, index) => {
      const event = rec.event;
      html += `
        <div class="event">
          <div class="event-header">${formatTime(event.startTime)} - ${formatTime(event.endTime)}: ${event.program}</div>
          <div class="event-details">
            Topic: ${event.topic || 'N/A'}<br>
            Room: ${event.room || 'N/A'}<br>
            ${event.notes ? 'Notes: ' + event.notes + '<br>' : ''}
          </div>
          <div class="recommendations">
            <strong>Recommended Coverage:</strong><br>
      `;
      
      if (rec.recommendations.length === 0) {
        html += '<p>No suitable staff found for coverage.</p>';
      } else {
        rec.recommendations.forEach(staff => {
          let cssClass = '';
          if (staff.hasConflict) {
            cssClass = 'conflict';
          } else if (staff.match === 'Excellent match') {
            cssClass = 'excellent';
          } else if (staff.match === 'Available but not specialized') {
            cssClass = 'good';
          } else if (staff.match === 'Specialized but outside hours') {
            cssClass = 'fair';
          } else {
            cssClass = 'poor';
          }
          
          html += `<div class="recommendation ${cssClass}">
            <strong>${staff.name}</strong> - ${staff.match}`;
          
          if (staff.specialtyMatch) {
            html += `<br><span class="specialty-match">✓ Preferred for topic: ${event.topic}</span>`;
          }
          
          if (staff.availabilityMatch) {
            html += `<br><span class="availability-match">✓ Within regular hours on ${['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'][event.date.getDay()]}</span>`;
          }
          
          if (staff.hasConflict && staff.conflictDetails) {
            const conflict = staff.conflictDetails;
            html += `<div class="conflict-details">⚠️ Conflict: ${formatTime(conflict.startTime)} - ${formatTime(conflict.endTime)} ${conflict.program} (${conflict.topic || 'No topic'})</div>`;
          }
          
          html += `</div>`;
        });
      }
      
      html += `
          </div>
        </div>
      `;
    });
  }
  
  htmlOutput.setContent(html);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}
