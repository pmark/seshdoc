/**
 * AuthorizationSetup.gs - Handle Calendar Authorization and Permissions
 * This file helps diagnose and fix authorization issues with Google Calendar access
 */

/**
 * Test and authorize all required permissions - RUN THIS FIRST
 * This function should be run manually to trigger OAuth authorization for all scopes
 */
function authorizeAllPermissions() {
  try {
    console.log('Testing all required permissions...');
    
    // Test spreadsheet access
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetName = spreadsheet.getName();
    console.log('✓ Spreadsheet access authorized. Name:', spreadsheetName);
    
    // Test UI access (this will trigger container.ui scope)
    const ui = SpreadsheetApp.getUi();
    console.log('✓ UI access authorized');
    
    // Test calendar access
    const calendar = CalendarApp.getDefaultCalendar();
    const calendarName = calendar.getName();
    console.log('✓ Calendar access authorized. Calendar name:', calendarName);
    
    // Test user email access
    const userEmail = Session.getActiveUser().getEmail();
    console.log('✓ User email access authorized:', userEmail);
    
    // Test getting events for today
    const today = new Date();
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    
    const events = calendar.getEvents(today, tomorrow);
    console.log('✓ Successfully retrieved', events.length, 'events for today');
    
    // Test getting all calendars
    const allCalendars = CalendarApp.getAllCalendars();
    console.log('✓ Found', allCalendars.length, 'accessible calendars');
    
    // Show success dialog (this tests UI permissions)
    ui.alert(
      'Authorization Successful', 
      `All permissions are working properly!\n\n✓ Spreadsheet: ${spreadsheetName}\n✓ Calendar: ${calendarName} (${events.length} events today)\n✓ UI Access: Working\n✓ User Email: ${userEmail}\n✓ Total Calendars: ${allCalendars.length}`, 
      ui.ButtonSet.OK
    );
    
    return {
      success: true,
      spreadsheetName: spreadsheetName,
      calendarName: calendarName,
      eventCount: events.length,
      calendarCount: allCalendars.length,
      userEmail: userEmail
    };
    
  } catch (error) {
    console.error('❌ Authorization failed:', error);
    
    // Try to show error without modal dialog if UI permission is the issue
    try {
      SpreadsheetApp.getUi().alert(
        'Authorization Required', 
        `Permissions need to be granted.\n\nError: ${error.message}\n\nPlease run this function again and grant all requested permissions.`, 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiError) {
      console.error('Cannot show UI dialog:', uiError);
      // Log the error for manual review
      console.error('AUTHORIZATION NEEDED: Please grant all permissions when prompted');
    }
    
    throw error;
  }
}

/**
 * Check current authorization status
 * @return {Object} Authorization status information
 */
function checkAuthorizationStatus() {
  try {
    // Test basic calendar access
    const calendar = CalendarApp.getDefaultCalendar();
    const calendarName = calendar.getName();
    
    return {
      authorized: true,
      calendarName: calendarName,
      error: null
    };
    
  } catch (error) {
    return {
      authorized: false,
      calendarName: null,
      error: error.message
    };
  }
}

/**
 * Setup OAuth scopes explicitly (add to manifest)
 * This function doesn't run but documents required scopes
 */
function setupOAuthScopes() {
  /*
  Add this to your appsscript.json manifest file:
  
  {
    "timeZone": "America/New_York",
    "dependencies": {
      "enabledAdvancedServices": []
    },
    "exceptionLogging": "STACKDRIVER",
    "runtimeVersion": "V8",
    "oauthScopes": [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/calendar.readonly",
      "https://www.googleapis.com/auth/script.scriptapp"
    ]
  }
  */
}

/**
 * Alternative calendar access method using Advanced Google Services
 * Enable this if CalendarApp continues to fail
 */
function enableAdvancedCalendarService() {
  /*
  TO ENABLE ADVANCED CALENDAR API:
  
  1. In Apps Script editor, click "Services" in left sidebar
  2. Click "Add a service"
  3. Select "Google Calendar API"
  4. Click "Add"
  
  Then you can use Calendar API directly:
  
  function getEventsViaAPI() {
    try {
      const today = new Date().toISOString();
      const tomorrow = new Date(Date.now() + 24*60*60*1000).toISOString();
      
      const events = Calendar.Events.list('primary', {
        timeMin: today,
        timeMax: tomorrow,
        singleEvents: true,
        orderBy: 'startTime'
      });
      
      return events.items || [];
    } catch (error) {
      console.error('Calendar API error:', error);
      return [];
    }
  }
  */
}

/**
 * Test different calendar access methods
 * @return {Object} Test results
 */
function testAllCalendarMethods() {
  const results = {
    defaultCalendar: false,
    primaryCalendar: false,
    allCalendars: false,
    calendarAPI: false,
    errors: []
  };
  
  // Test 1: Default Calendar
  try {
    const defaultCal = CalendarApp.getDefaultCalendar();
    defaultCal.getName();
    results.defaultCalendar = true;
    console.log('✓ Default calendar access works');
  } catch (error) {
    results.errors.push('Default calendar: ' + error.message);
    console.log('❌ Default calendar failed:', error.message);
  }
  
  // Test 2: Primary Calendar
  try {
    const primaryCal = CalendarApp.getCalendarById('primary');
    primaryCal.getName();
    results.primaryCalendar = true;
    console.log('✓ Primary calendar access works');
  } catch (error) {
    results.errors.push('Primary calendar: ' + error.message);
    console.log('❌ Primary calendar failed:', error.message);
  }
  
  // Test 3: All Calendars
  try {
    const allCals = CalendarApp.getAllCalendars();
    if (allCals.length > 0) {
      results.allCalendars = true;
      console.log('✓ All calendars access works, found:', allCals.length);
    }
  } catch (error) {
    results.errors.push('All calendars: ' + error.message);
    console.log('❌ All calendars failed:', error.message);
  }
  
  // Test 4: Calendar API (if enabled)
  try {
    if (typeof Calendar !== 'undefined') {
      Calendar.Events.list('primary', { maxResults: 1 });
      results.calendarAPI = true;
      console.log('✓ Calendar API access works');
    }
  } catch (error) {
    results.errors.push('Calendar API: ' + error.message);
    console.log('❌ Calendar API failed:', error.message);
  }
  
  return results;
}

/**
 * Safe calendar access wrapper - use this instead of direct CalendarApp calls
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {Array} Array of events or empty array if authorization fails
 */
function safeGetCalendarEvents(startDate, endDate) {
  try {
    // Check authorization first
    const authStatus = checkAuthorizationStatus();
    if (!authStatus.authorized) {
      console.warn('Calendar not authorized:', authStatus.error);
      return [];
    }
    
    // Try to get events
    const calendar = CalendarApp.getDefaultCalendar();
    const events = calendar.getEvents(startDate, endDate);
    
    console.log(`Successfully retrieved ${events.length} events`);
    return events;
    
  } catch (error) {
    console.error('Safe calendar access failed:', error);
    
    // Show user-friendly error
    if (error.message.includes('authorization') || error.message.includes('permission')) {
      SpreadsheetApp.getUi().alert(
        'Calendar Authorization Needed',
        'This script needs permission to access your Google Calendar.\n\nPlease run "Therapy Tools > Authorize Calendar Access" from the menu first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    return [];
  }
}

/**
 * Initialize and test all permissions
 * Run this function manually to set up all required permissions
 */
function initializeAllPermissions() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert('Permission Setup', 'This will test and request all necessary permissions. Click OK to continue.', ui.ButtonSet.OK);
    
    // Test spreadsheet access
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    console.log('✓ Spreadsheet access works');
    
    // Test calendar access
    const authResult = authorizeCalendarAccess();
    
    // Test email access (for therapist name)
    const userEmail = Session.getActiveUser().getEmail();
    console.log('✓ User email access works:', userEmail);
    
    ui.alert(
      'Setup Complete',
      `All permissions have been authorized successfully!\n\n✓ Spreadsheet access\n✓ Calendar access (${authResult.eventCount} events found)\n✓ User email access\n\nYou can now use all Therapy Tools features.`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Permission setup failed:', error);
    ui.alert(
      'Setup Failed',
      `Permission setup encountered an error:\n\n${error.message}\n\nPlease try running this function again.`,
      ui.ButtonSet.OK
    );
  }
}