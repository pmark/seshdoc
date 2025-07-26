/**
 * MenuHandler.gs - Main menu and dialog orchestration (Fixed)
 * Handles the custom menu creation and dialog flow coordination
 */

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Therapy Tools")
    .addItem("📋 Document Sessions", "showSessionAppointments")
    .addItem("✏️ Document Session (Manual)", "showClientSelector")
    .addSeparator()
    .addItem("🎯 Manage Client Goals", "showGoalManagement")
    .addSeparator()
    .addItem("🔧 Setup Form Configuration", "showFormConfiguration")
    .addItem("🔑 Authorize All Permissions", "authorizeAllPermissions")
    .addItem("🔍 Test Calendar Access", "testCalendarAccess")
    .addSeparator()
    .addItem("🔄 Sync Form Responses", "syncFormResponses")
    .addToUi();
}

/**
 * Shows session appointments with date selection (primary workflow)
 */
function showSessionAppointments() {
  const html = HtmlService.createHtmlOutputFromFile("SessionAppointments")
    .setWidth(700)
    .setHeight(600)
    .setTitle("Session Appointments");

  SpreadsheetApp.getUi().showModalDialog(html, "Document Sessions");
}

/**
 * Shows the client selector dialog (manual workflow)
 */
function showClientSelector() {
  const html = HtmlService.createHtmlOutputFromFile("ClientSelector")
    .setWidth(400)
    .setHeight(500)
    .setTitle("Select Client");

  SpreadsheetApp.getUi().showModalDialog(
    html,
    "Document Session - Manual Selection"
  );
}

/**
 * Shows the goal selector dialog (second step)
 * @param {Object} sessionData - Session data with client and optional appointment
 */
function showGoalSelector(sessionData) {
  const template = HtmlService.createTemplateFromFile("GoalSelector");
  template.sessionData = JSON.stringify(sessionData);

  const html = template
    .evaluate()
    .setWidth(450)
    .setHeight(400)
    .setTitle("Select Goal");

  SpreadsheetApp.getUi().showModalDialog(
    html,
    "Document Session - Select Goal"
  );
}

/**
 * Shows goal management interface
 */
function showGoalManagement() {
  // First show client selector for goal management
  const html = HtmlService.createHtmlOutputFromFile(
    "GoalManagementClientSelector"
  )
    .setWidth(400)
    .setHeight(450)
    .setTitle("Select Client for Goal Management");

  SpreadsheetApp.getUi().showModalDialog(html, "Manage Client Goals");
}
function showGoalManagementInterface(selectedClient) {
  const template = HtmlService.createTemplateFromFile(
    "GoalManagementInterface"
  );
  template.client = JSON.stringify(selectedClient);

  const html = template
    .evaluate()
    .setWidth(600)
    .setHeight(500)
    .setTitle("Manage Goals - " + selectedClient.name);

  SpreadsheetApp.getUi().showModalDialog(html, "Goal Management");
}

/**
 * Shows form configuration dialog
 */
function showFormConfiguration() {
  const html = HtmlService.createHtmlOutputFromFile("FormConfiguration")
    .setWidth(500)
    .setHeight(400)
    .setTitle("Form Configuration");

  SpreadsheetApp.getUi().showModalDialog(html, "Setup Google Form Integration");
}

/**
 * Test calendar access for troubleshooting
 */
function testCalendarAccess() {
  try {
    const results = testAllCalendarMethods();

    let message = "Calendar Access Test Results:\n\n";
    message += `✓ Default Calendar: ${
      results.defaultCalendar ? "Working" : "Failed"
    }\n`;
    message += `✓ Primary Calendar: ${
      results.primaryCalendar ? "Working" : "Failed"
    }\n`;
    message += `✓ All Calendars: ${
      results.allCalendars ? "Working" : "Failed"
    }\n`;
    message += `✓ Calendar API: ${
      results.calendarAPI ? "Working" : "Not Enabled"
    }\n`;

    if (results.errors.length > 0) {
      message += "\nErrors:\n" + results.errors.join("\n");
    }

    SpreadsheetApp.getUi().alert(
      "Calendar Test Results",
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      "Test Failed",
      "Calendar test failed: " + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Generates pre-populated form URL and shows completion dialog
 * @param {Object} sessionData - Object containing client and goal info
 */
function generateSessionForm(sessionData) {
  try {
    // Use the correct function name instead of class method
    const formUrl = createEnhancedPrePopulatedUrl(sessionData);

    const html = HtmlService.createHtmlOutput(
      `
      <div style="padding: 20px; font-family: Arial, sans-serif;">
        <h3>Session Form Ready</h3>
        <p><strong>Client:</strong> ${sessionData.client.name}</p>
        <p><strong>Goal:</strong> ${sessionData.goal}</p>
        <p><strong>Date:</strong> ${new Date().toLocaleDateString()}</p>
        <br>
        <a href="${formUrl}" target="_blank" style="
          display: inline-block;
          padding: 10px 20px;
          background-color: #4285f4;
          color: white;
          text-decoration: none;
          border-radius: 4px;
          font-weight: bold;
        ">Open Session Documentation Form</a>
        <br><br>
        <button onclick="google.script.host.close()" style="
          padding: 8px 16px;
          background-color: #f1f3f4;
          border: 1px solid #dadce0;
          border-radius: 4px;
          cursor: pointer;
        ">Close</button>
      </div>
    `
    )
      .setWidth(400)
      .setHeight(250)
      .setTitle("Session Form Generated");

    SpreadsheetApp.getUi().showModalDialog(html, "Ready to Document");

    // Record the session if we have appointment data
    if (sessionData.appointmentId) {
      recordSessionStart(sessionData);
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      "Error",
      "Failed to generate form: " + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Record session start for tracking
 * @param {Object} sessionData - Session data
 */
function recordSessionStart(sessionData) {
  try {
    // Record session with tracking service if available
    if (typeof recordSession !== "undefined") {
      recordSession(sessionData);
    }
  } catch (error) {
    console.error("Error recording session:", error);
    // Don't block the main flow if tracking fails
  }
}

/**
 * Sync form responses manually
 */
function syncFormResponses() {
  try {
    const config = getFormConfig();

    if (!config.FORM_ID || config.FORM_ID === "YOUR_GOOGLE_FORM_ID_HERE") {
      SpreadsheetApp.getUi().alert(
        "Configuration Required",
        "Please configure your Google Form ID first.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Show progress dialog
    const html = HtmlService.createHtmlOutput(
      `
      <div style="padding: 20px; font-family: Arial, sans-serif; text-align: center;">
        <h3>Syncing Form Responses</h3>
        <p>Please wait while we sync recent form responses with client data...</p>
        <div style="margin: 20px 0;">
          <div style="width: 100%; height: 4px; background-color: #f1f3f4; border-radius: 2px;">
            <div style="width: 0%; height: 100%; background-color: #4285f4; border-radius: 2px; animation: progress 3s ease-in-out infinite;"></div>
          </div>
        </div>
        <p style="font-size: 12px; color: #5f6368;">This may take a few moments...</p>
        <style>
          @keyframes progress {
            0% { width: 0%; }
            50% { width: 70%; }
            100% { width: 100%; }
          }
        </style>
      </div>
    `
    )
      .setWidth(400)
      .setHeight(200)
      .setTitle("Syncing Data");

    SpreadsheetApp.getUi().showModalDialog(html, "Sync in Progress");

    // Perform the sync
    let syncCount = 0;
    if (typeof syncExistingResponses !== "undefined") {
      syncCount = syncExistingResponses(config.FORM_ID, 7); // Last 7 days
    }

    // Show completion message
    SpreadsheetApp.getUi().alert(
      "Sync Complete",
      `Successfully synced ${syncCount} form responses.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      "Sync Error",
      "Failed to sync form responses: " + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
