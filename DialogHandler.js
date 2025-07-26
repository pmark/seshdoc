/**
 * DialogHandlers.gs - Server-side dialog event handlers
 * Handles communication between HTML dialogs and server-side functions
 */

/**
 * Calendar Integration Dialog Handlers
 */

/**
 * Load today's appointments with client matching and documentation status
 * @return {Array<Object>} Array of appointment objects with enhanced data
 */
function getTodaysAppointmentsForDialog() {
  try {
    // Get today's appointments from calendar
    const appointments = getTodaysAppointments();
    
    // Add client matching and documentation status
    return enhanceAppointmentsWithStatus(appointments);
    
  } catch (error) {
    console.error('Error loading today\'s appointments:', error);
    throw new Error('Unable to load calendar appointments: ' + error.message);
  }
}

/**
 * Get appointments for a specific date
 * @param {Date} targetDate - Target date for appointments
 * @return {Array<Object>} Array of appointment objects with enhanced data
 */
function getAppointmentsForDateDialog(targetDate) {
  try {
    const appointments = getAppointmentsForDate(targetDate);
    return enhanceAppointmentsWithStatus(appointments);
    
  } catch (error) {
    console.error('Error loading appointments for date:', error);
    throw new Error('Unable to load appointments for selected date: ' + error.message);
  }
}

/**
 * Get appointments for a date range
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {Array<Object>} Array of appointment objects with enhanced data
 */
function getAppointmentsForDateRangeDialog(startDate, endDate) {
  try {
    const appointments = getAppointmentsForDateRange(startDate, endDate);
    return enhanceAppointmentsWithStatus(appointments);
    
  } catch (error) {
    console.error('Error loading appointments for date range:', error);
    throw new Error('Unable to load appointments for selected date range: ' + error.message);
  }
}

/**
 * Enhance appointments with client matching and documentation status
 * @param {Array<Object>} appointments - Raw appointment objects
 * @return {Array<Object>} Enhanced appointment objects
 */
function enhanceAppointmentsWithStatus(appointments) {
  try {
    // Add client matching
    const appointmentsWithMatches = appointments.map(appointment => {
      try {
        const matchedClient = matchAppointmentToClient(appointment, 'Clients');
        
        return {
          ...appointment,
          matchedClient: matchedClient,
          hasClientMatch: !!matchedClient,
          matchConfidence: calculateMatchConfidence(appointment, matchedClient)
        };
      } catch (matchError) {
        console.error('Error matching appointment to client:', matchError);
        return {
          ...appointment,
          matchedClient: null,
          hasClientMatch: false,
          matchConfidence: 'none'
        };
      }
    });
    
    // Add documentation status if tracking service is available
    if (typeof getAppointmentStatuses !== 'undefined') {
      return getAppointmentStatuses(appointmentsWithMatches);
    } else {
      // Add default status
      return appointmentsWithMatches.map(appointment => ({
        ...appointment,
        documentationStatus: 'not_started',
        sessionRecord: null,
        isDocumented: false
      }));
    }
    
  } catch (error) {
    console.error('Error enhancing appointments:', error);
    throw error;
  }
}

/**
 * Show goal selector for a specific appointment
 * @param {Object} appointment - Appointment object with matched client
 */
function showGoalSelectorForAppointment(appointment) {
  try {
    if (!appointment.hasClientMatch) {
      throw new Error('No client matched for this appointment');
    }
    
    // Mark session as in progress if tracking is available
    if (typeof markSessionInProgress !== 'undefined') {
      markSessionInProgress(appointment.id, appointment.matchedClient, appointment);
    }
    
    // Create session data from appointment
    const sessionData = {
      appointment: appointment,
      client: appointment.matchedClient,
      fromCalendar: true
    };
    
    // Show goal selector with appointment context
    showGoalSelector(sessionData);
    
  } catch (error) {
    console.error('Error showing goal selector for appointment:', error);
    throw new Error('Failed to show goal selector: ' + error.message);
  }
}

/**
 * Show client selector for manual appointment matching
 * @param {Object} appointment - Appointment object without client match
 */
function showClientSelectorForAppointment(appointment) {
  try {
    const template = HtmlService.createTemplateFromFile('ClientSelector');
    template.appointment = JSON.stringify(appointment);
    template.isManualMatch = true;
    
    const html = template.evaluate()
      .setWidth(400)
      .setHeight(500)
      .setTitle('Match Client to Appointment');
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Link Client to Appointment');
    
  } catch (error) {
    console.error('Error showing client selector for appointment:', error);
    throw new Error('Failed to show client selector: ' + error.message);
  }
}

/**
 * Get documentation URL for an appointment
 * @param {string} appointmentId - Appointment ID
 * @return {string} Documentation URL or null
 */
function getDocumentationForAppointment(appointmentId) {
  try {
    if (typeof getSessionRecord !== 'undefined') {
      const sessionRecord = getSessionRecord(appointmentId);
      return sessionRecord ? sessionRecord.formUrl : null;
    }
    return null;
  } catch (error) {
    console.error('Error getting documentation for appointment:', error);
    return null;
  }
}

/**
 * Continue session documentation for in-progress appointment
 * @param {string} appointmentId - Appointment ID
 * @return {string} Documentation URL or null
 */
function continueSessionDocumentation(appointmentId) {
  try {
    // Same as getting documentation for now
    return getDocumentationForAppointment(appointmentId);
  } catch (error) {
    console.error('Error continuing session documentation:', error);
    return null;
  }
}/**
 * DialogHandlers.gs - Server-side dialog event handlers
 * Handles communication between HTML dialogs and server-side functions
 */

/**
 * Client Selector Dialog Handlers
 */

/**
 * Load all clients for the client selector dialog
 * @return {Array<Object>} Array of client objects
 */
function loadClientsForDialog() {
  try {
    return ClientService.getAllClientsForSelection();
  } catch (error) {
    console.error('Error loading clients for dialog:', error);
    throw new Error('Unable to load client data: ' + error.message);
  }
}

/**
 * Search clients based on search term
 * @param {string} searchTerm - Search term to filter clients
 * @return {Array<Object>} Filtered array of client objects
 */
function searchClientsForDialog(searchTerm) {
  try {
    return ClientService.searchClients(searchTerm);
  } catch (error) {
    console.error('Error searching clients:', error);
    throw new Error('Search failed: ' + error.message);
  }
}

/**
 * Handle client selection and proceed to goal selector
 * @param {Object} selectedClient - The selected client object
 */
function handleClientSelection(selectedClient) {
  try {
    // Validate the selected client
    if (!ClientService.validateClient(selectedClient)) {
      throw new Error('Invalid client selected');
    }
    
    // Close current dialog and show goal selector
    showGoalSelector(selectedClient);
    
  } catch (error) {
    console.error('Error handling client selection:', error);
    throw new Error('Failed to proceed to goal selection: ' + error.message);
  }
}

/**
 * Goal Selector Dialog Handlers
 */

/**
 * Load goals for the selected client
 * @param {string} clientId - ID of the selected client
 * @return {Array<string>} Array of goal strings
 */
function loadGoalsForDialog(clientId) {
  try {
    return ClientService.getClientGoals(clientId);
  } catch (error) {
    console.error('Error loading goals for dialog:', error);
    throw new Error('Unable to load goals: ' + error.message);
  }
}

/**
 * Handle goal selection and generate form
 * @param {Object} sessionData - Complete session data with client and goal
 */
function handleGoalSelection(sessionData) {
  try {
    // Validate session data
    if (!sessionData || !sessionData.client || !sessionData.goal) {
      throw new Error('Incomplete session data');
    }
    
    if (!ClientService.validateClient(sessionData.client)) {
      throw new Error('Invalid client data');
    }
    
    if (!sessionData.goal || sessionData.goal.trim() === '') {
      throw new Error('No goal selected');
    }
    
    // Generate the pre-populated form
    generateSessionForm(sessionData);
    
  } catch (error) {
    console.error('Error handling goal selection:', error);
    throw new Error('Failed to generate session form: ' + error.message);
  }
}

/**
 * Handle return to client selector from goal selector
 */
function handleBackToClientSelector() {
  try {
    showClientSelector();
  } catch (error) {
    console.error('Error returning to client selector:', error);
    throw new Error('Failed to return to client selection: ' + error.message);
  }
}

/**
 * Utility Functions for Dialogs
 */

/**
 * Get configuration status for debugging
 * @return {Object} Configuration validation results
 */
function getConfigurationStatus() {
  try {
    return FormUrlGenerator.validateConfiguration();
  } catch (error) {
    return {
      isValid: false,
      errors: ['Configuration validation failed: ' + error.message]
    };
  }
}

/**
 * Test form URL generation
 * @return {string} Test form URL
 */
function testFormGeneration() {
  try {
    return FormUrlGenerator.createTestUrl();
  } catch (error) {
    throw new Error('Form generation test failed: ' + error.message);
  }
}