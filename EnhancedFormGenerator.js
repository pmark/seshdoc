/**
 * EnhancedFormGenerator.gs - Advanced form pre-population with calendar and historical data
 * 
 * IMPORTANT: This file must be loaded AFTER FormUrlGenerator.gs
 */

/**
 * Enhanced form configuration with more comprehensive field mappings
 */
function getEnhancedFormConfig() {
  const baseConfig = getFormConfig();
  
  return {
    ...baseConfig,
    
    // Extended field mappings for comprehensive pre-population
    EXTENDED_FIELDS: {
      // Basic session info
      SESSION_DATE: 'entry.111111111',
      SESSION_START_TIME: 'entry.222222222',
      SESSION_END_TIME: 'entry.333333333',
      SESSION_DURATION: 'entry.444444444',
      SESSION_TYPE: 'entry.555555555', // 'Individual', 'Group', 'Family', etc.
      
      // Client information
      CLIENT_EMAIL: 'entry.666666666',
      CLIENT_PHONE: 'entry.777777777',
      INSURANCE_PROVIDER: 'entry.888888888',
      
      // Session details
      SESSION_LOCATION: 'entry.999999999', // 'In-Person', 'Telehealth', 'Phone'
      APPOINTMENT_TYPE: 'entry.101010101', // 'Initial', 'Follow-up', 'Crisis', etc.
      
      // Previous session reference
      LAST_SESSION_DATE: 'entry.121212121',
      PREVIOUS_GOALS_STATUS: 'entry.131313131',
      
      // Clinical notes
      MOOD_ASSESSMENT: 'entry.141414141',
      RISK_ASSESSMENT: 'entry.151515151',
      MEDICATION_CHANGES: 'entry.161616161',
      
      // Administrative
      THERAPIST_NAME: 'entry.171717171',
      SUPERVISOR_REVIEW: 'entry.181818181',
      BILLING_CODE: 'entry.191919191'
    }
  };
}

/**
 * Create comprehensively pre-populated form URL using calendar and historical data
 * @param {Object} sessionData - Enhanced session data
 * @return {string} Pre-populated form URL
 */
function createEnhancedPrePopulatedUrl(sessionData) {
  const config = getEnhancedFormConfig();
  
  // Validate input
  if (!sessionData || !sessionData.client || !sessionData.goal) {
    throw new Error('Invalid session data provided');
  }
  
  // Get historical prefill data
  const prefillData = getPrefillDataForClient(sessionData.client.id);
  
  // Build the base URL
  const baseUrl = `https://docs.google.com/forms/d/${config.FORM_ID}/formResponse`;
  const params = [];
  
  // Basic client information (from original system)
  addBasicClientFields(params, sessionData, config);
  
  // Calendar-derived information
  addCalendarFields(params, sessionData, config);
  
  // Historical data prefilling
  addHistoricalFields(params, sessionData, prefillData, config);
  
  // Administrative fields
  addAdministrativeFields(params, sessionData, config);
  
  return `${baseUrl}?${params.join('&')}`;
}

/**
 * Add basic client fields to form parameters
 * @param {Array} params - URL parameters array
 * @param {Object} sessionData - Session data
 * @param {Object} config - Form configuration
 */
function addBasicClientFields(params, sessionData, config) {
  // Core client data
  if (config.FIELD_MAPPINGS.CLIENT_NAME) {
    params.push(`${config.FIELD_MAPPINGS.CLIENT_NAME}=${encodeURIComponent(sessionData.client.name)}`);
  }
  
  if (config.FIELD_MAPPINGS.CLIENT_ID) {
    params.push(`${config.FIELD_MAPPINGS.CLIENT_ID}=${encodeURIComponent(sessionData.client.id)}`);
  }
  
  if (config.FIELD_MAPPINGS.SELECTED_GOAL) {
    params.push(`${config.FIELD_MAPPINGS.SELECTED_GOAL}=${encodeURIComponent(sessionData.goal)}`);
  }
  
  // Extended client information
  if (config.EXTENDED_FIELDS.CLIENT_EMAIL && sessionData.client.email) {
    params.push(`${config.EXTENDED_FIELDS.CLIENT_EMAIL}=${encodeURIComponent(sessionData.client.email)}`);
  }
  
  if (config.EXTENDED_FIELDS.CLIENT_PHONE && sessionData.client.phone) {
    params.push(`${config.EXTENDED_FIELDS.CLIENT_PHONE}=${encodeURIComponent(sessionData.client.phone)}`);
  }
}

/**
 * Add calendar-derived fields to form parameters
 * @param {Array} params - URL parameters array
 * @param {Object} sessionData - Session data
 * @param {Object} config - Form configuration
 */
function addCalendarFields(params, sessionData, config) {
  if (!sessionData.appointment) return;
  
  const appointment = sessionData.appointment;
  
  // Session timing from calendar
  if (config.FIELD_MAPPINGS.SESSION_DATE) {
    params.push(`${config.FIELD_MAPPINGS.SESSION_DATE}=${encodeURIComponent(appointment.startTime.toLocaleDateString())}`);
  }
  
  if (config.EXTENDED_FIELDS.SESSION_START_TIME) {
    const startTime = appointment.startTime.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', hour12: true });
    params.push(`${config.EXTENDED_FIELDS.SESSION_START_TIME}=${encodeURIComponent(startTime)}`);
  }
  
  if (config.EXTENDED_FIELDS.SESSION_END_TIME) {
    const endTime = appointment.endTime.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', hour12: true });
    params.push(`${config.EXTENDED_FIELDS.SESSION_END_TIME}=${encodeURIComponent(endTime)}`);
  }
  
  if (config.EXTENDED_FIELDS.SESSION_DURATION) {
    params.push(`${config.EXTENDED_FIELDS.SESSION_DURATION}=${encodeURIComponent(appointment.duration + ' minutes')}`);
  }
  
  // Location information
  if (config.EXTENDED_FIELDS.SESSION_LOCATION && appointment.location) {
    const sessionLocation = determineSessionLocation(appointment.location);
    params.push(`${config.EXTENDED_FIELDS.SESSION_LOCATION}=${encodeURIComponent(sessionLocation)}`);
  }
  
  // Appointment type inference
  if (config.EXTENDED_FIELDS.APPOINTMENT_TYPE) {
    const appointmentType = inferAppointmentType(appointment, sessionData);
    params.push(`${config.EXTENDED_FIELDS.APPOINTMENT_TYPE}=${encodeURIComponent(appointmentType)}`);
  }
}

/**
 * Add historical data fields to form parameters
 * @param {Array} params - URL parameters array
 * @param {Object} sessionData - Session data
 * @param {Object} prefillData - Historical prefill data
 * @param {Object} config - Form configuration
 */
function addHistoricalFields(params, sessionData, prefillData, config) {
  // Previous session information
  if (config.EXTENDED_FIELDS.LAST_SESSION_DATE && prefillData.lastSessionDate) {
    params.push(`${config.EXTENDED_FIELDS.LAST_SESSION_DATE}=${encodeURIComponent(prefillData.lastSessionDate)}`);
  }
  
  // Insurance information
  if (config.EXTENDED_FIELDS.INSURANCE_PROVIDER && prefillData.insuranceProvider) {
    params.push(`${config.EXTENDED_FIELDS.INSURANCE_PROVIDER}=${encodeURIComponent(prefillData.insuranceProvider)}`);
  }
  
  // Session type preference
  if (config.EXTENDED_FIELDS.SESSION_TYPE && prefillData.sessionType) {
    params.push(`${config.EXTENDED_FIELDS.SESSION_TYPE}=${encodeURIComponent(prefillData.sessionType)}`);
  }
  
  // Clinical context
  if (config.EXTENDED_FIELDS.PREVIOUS_GOALS_STATUS && prefillData.previousGoalsAchieved) {
    params.push(`${config.EXTENDED_FIELDS.PREVIOUS_GOALS_STATUS}=${encodeURIComponent(prefillData.previousGoalsAchieved)}`);
  }
  
  // Medication notes
  if (config.EXTENDED_FIELDS.MEDICATION_CHANGES && prefillData.medicationNotes) {
    params.push(`${config.EXTENDED_FIELDS.MEDICATION_CHANGES}=${encodeURIComponent(prefillData.medicationNotes)}`);
  }
}

/**
 * Add administrative fields to form parameters
 * @param {Array} params - URL parameters array
 * @param {Object} sessionData - Session data
 * @param {Object} config - Form configuration
 */
function addAdministrativeFields(params, sessionData, config) {
  // Therapist information
  if (config.EXTENDED_FIELDS.THERAPIST_NAME) {
    const currentUser = Session.getActiveUser().getEmail();
    const therapistName = getTherapistNameFromEmail(currentUser);
    params.push(`${config.EXTENDED_FIELDS.THERAPIST_NAME}=${encodeURIComponent(therapistName)}`);
  }
  
  // Billing code (can be inferred from appointment type and duration)
  if (config.EXTENDED_FIELDS.BILLING_CODE && sessionData.appointment) {
    const billingCode = determineBillingCode(sessionData.appointment);
    params.push(`${config.EXTENDED_FIELDS.BILLING_CODE}=${encodeURIComponent(billingCode)}`);
  }
  
  // Auto-populate form completion date/time
  params.push(`entry.form_completed_at=${encodeURIComponent(new Date().toISOString())}`);
}

/**
 * Determine session location type from calendar location
 * @param {string} location - Calendar event location
 * @return {string} Session location type
 */
function determineSessionLocation(location) {
  if (!location) return 'In-Person';
  
  const locationLower = location.toLowerCase();
  
  if (locationLower.includes('zoom') || locationLower.includes('teams') || 
      locationLower.includes('telehealth') || locationLower.includes('video')) {
    return 'Telehealth';
  }
  
  if (locationLower.includes('phone') || locationLower.includes('call')) {
    return 'Phone';
  }
  
  return 'In-Person';
}

/**
 * Infer appointment type from calendar data and client history
 * @param {Object} appointment - Calendar appointment
 * @param {Object} sessionData - Session data
 * @return {string} Appointment type
 */
function inferAppointmentType(appointment, sessionData) {
  const title = appointment.title.toLowerCase();
  
  // Check for keywords in appointment title
  if (title.includes('initial') || title.includes('intake') || title.includes('first')) {
    return 'Initial';
  }
  
  if (title.includes('crisis') || title.includes('emergency') || title.includes('urgent')) {
    return 'Crisis';
  }
  
  if (title.includes('family') || title.includes('couple')) {
    return 'Family/Couple';
  }
  
  if (title.includes('group')) {
    return 'Group';
  }
  
  // Check session history to determine if this is a new client
  const prefillData = getPrefillDataForClient(sessionData.client.id);
  if (!prefillData.lastSessionDate) {
    return 'Initial';
  }
  
  return 'Follow-up';
}

/**
 * Get therapist name from email or configuration
 * @param {string} email - Therapist email
 * @return {string} Therapist name
 */
function getTherapistNameFromEmail(email) {
  // This could be enhanced to look up therapist names from a configuration sheet
  // For now, extract name from email
  const namePart = email.split('@')[0];
  return namePart.replace(/[._]/g, ' ').replace(/\b\w/g, function(l) { return l.toUpperCase(); });
}

/**
 * Determine billing code based on appointment details
 * @param {Object} appointment - Calendar appointment
 * @return {string} Billing code
 */
function determineBillingCode(appointment) {
  const duration = appointment.duration;
  
  // Standard CPT codes for psychotherapy
  if (duration <= 30) {
    return '90834'; // 30-minute session
  } else if (duration <= 45) {
    return '90837'; // 45-minute session
  } else if (duration <= 60) {
    return '90834'; // Individual therapy
  } else {
    return '90837'; // Extended session
  }
}

/**
 * Create session data with calendar integration
 * @param {Object} appointment - Calendar appointment
 * @param {Object} client - Client object
 * @param {string} goal - Selected goal
 * @return {Object} Enhanced session data
 */
function createSessionDataFromAppointment(appointment, client, goal) {
  return {
    client: client,
    goal: goal,
    appointment: appointment,
    sessionDate: appointment.startTime.toLocaleDateString(),
    sessionTime: appointment.timeDisplay,
    appointmentId: appointment.id,
    extractedFromCalendar: true
  };
}

/**
 * Get prefill data for a client (wrapper function)
 * @param {string} clientId - Client ID
 * @return {Object} Prefill data object
 */
function getPrefillDataForClient(clientId) {
  try {
    // This calls the SessionTrackingService function if it exists
    if (typeof getPrefillData !== 'undefined') {
      return getPrefillData(clientId);
    }
    return {};
  } catch (error) {
    console.error('Error getting prefill data:', error);
    return {};
  }
}

/**
 * Update prefill data after session completion
 * @param {Object} sessionData - Completed session data
 * @param {Object} formData - Additional form data collected
 */
function updatePrefillDataFromSession(sessionData, formData) {
  formData = formData || {};
  
  const updatedPrefillData = {
    lastSessionDate: sessionData.sessionDate,
    commonDuration: sessionData.appointment ? (sessionData.appointment.duration + ' minutes') : '',
    preferredTime: sessionData.appointment ? (sessionData.appointment.startTime.getHours() + ':00') : '',
    sessionType: formData.sessionType || '',
    locationPreference: formData.sessionLocation || '',
    insuranceProvider: formData.insuranceProvider || '',
    medicationNotes: formData.medicationChanges || ''
  };
  
  // Add any additional form data
  Object.assign(updatedPrefillData, formData);
  
  try {
    // This calls the SessionTrackingService function if it exists
    if (typeof updatePrefillData !== 'undefined') {
      updatePrefillData(sessionData.client.id, updatedPrefillData);
    }
  } catch (error) {
    console.error('Error updating prefill data:', error);
  }
}