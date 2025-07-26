/**
 * CalendarService.gs - Google Calendar integration with robust error handling
 * Handles calendar event retrieval and client matching
 */

/**
 * Configuration for calendar integration
 */
function getCalendarConfig() {
  return {
    // Default calendar (primary calendar)
    CALENDAR_ID: 'primary',
    
    // Keywords to identify therapy sessions in event titles
    THERAPY_KEYWORDS: [
      'therapy', 'session', 'counseling', 'appointment',
      'meeting', 'consultation', 'treatment'
    ],
    
    // How many days to look ahead for appointments
    DAYS_AHEAD: 7,
    
    // Business hours (24-hour format)
    START_HOUR: 8,
    END_HOUR: 18,
    
    // Retry configuration
    MAX_RETRIES: 3,
    RETRY_DELAY_MS: 1000
  };
}

/**
 * Get today's therapy appointments with authorization check
 * @return {Array<Object>} Array of appointment objects
 */
function getTodaysAppointments() {
  const today = new Date();
  const startOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const endOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
  
  return safeGetCalendarEvents(startOfDay, endOfDay);
}

/**
 * Get appointments for a specific date with authorization check
 * @param {Date} targetDate - Target date
 * @return {Array<Object>} Array of appointment objects
 */
function getAppointmentsForDate(targetDate) {
  const startOfDay = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate());
  const endOfDay = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 23, 59, 59);
  
  return safeGetCalendarEvents(startOfDay, endOfDay);
}

/**
 * Get appointments for a date range with authorization check
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {Array<Object>} Array of appointment objects
 */
function getAppointmentsForDateRange(startDate, endDate) {
  return safeGetCalendarEvents(startDate, endDate);
}

/**
 * Safe calendar events retrieval with authorization handling
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {Array<Object>} Array of formatted appointment objects
 */
function safeGetCalendarEvents(startDate, endDate) {
  try {
    // Check authorization first
    const authStatus = checkAuthorizationStatus();
    if (!authStatus.authorized) {
      console.warn('Calendar not authorized:', authStatus.error);
      
      // Return empty array with helpful message for UI
      throw new Error('Calendar access requires authorization. Please run "Authorize Calendar Access" from the Therapy Tools menu.');
    }
    
    // Try to get events using the safest method
    const calendar = CalendarApp.getDefaultCalendar();
    const events = calendar.getEvents(startDate, endDate);
    
    if (!events) {
      console.log('No events returned from calendar');
      return [];
    }
    
    console.log(`Found ${events.length} total events in date range`);
    
    // Filter and format events safely
    const therapyEvents = [];
    
    for (let i = 0; i < events.length; i++) {
      try {
        const event = events[i];
        
        if (isTherapyAppointment(event)) {
          const formattedEvent = formatAppointmentData(event);
          if (formattedEvent) {
            therapyEvents.push(formattedEvent);
          }
        }
      } catch (eventError) {
        console.error(`Error processing event ${i}:`, eventError);
        // Continue with other events
      }
    }
    
    // Sort by start time
    therapyEvents.sort((a, b) => new Date(a.startTime) - new Date(b.startTime));
    
    console.log(`Filtered to ${therapyEvents.length} therapy appointments`);
    return therapyEvents;
    
  } catch (error) {
    console.error('Safe calendar access failed:', error);
    
    // For authorization errors, provide helpful guidance
    if (error.message.includes('authorization') || error.message.includes('permission') || 
        error.message.includes('transport') || error.message.includes('wardeninit')) {
      
      throw new Error('Calendar authorization required. Please go to Therapy Tools menu > "Authorize Calendar Access" and grant permissions.');
    }
    
    // For other errors, provide the original error
    throw error;
  }
}

/**
 * Get appointments for a specific date range with retry logic
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {Array<Object>} Array of appointment objects
 */
function getAppointmentsInRangeWithRetry(startDate, endDate) {
  const config = getCalendarConfig();
  let lastError = null;
  
  for (let attempt = 1; attempt <= config.MAX_RETRIES; attempt++) {
    try {
      console.log(`Calendar API attempt ${attempt}/${config.MAX_RETRIES}`);
      return getAppointmentsInRange(startDate, endDate);
      
    } catch (error) {
      lastError = error;
      console.error(`Calendar API attempt ${attempt} failed:`, error);
      
      // Don't retry on certain types of errors
      if (isNonRetryableError(error)) {
        break;
      }
      
      // Wait before retrying (except on last attempt)
      if (attempt < config.MAX_RETRIES) {
        console.log(`Waiting ${config.RETRY_DELAY_MS}ms before retry...`);
        Utilities.sleep(config.RETRY_DELAY_MS * attempt); // Exponential backoff
      }
    }
  }
  
  // If all retries failed, try fallback method
  console.log('All calendar API attempts failed, trying fallback...');
  try {
    return getAppointmentsFallback(startDate, endDate);
  } catch (fallbackError) {
    console.error('Fallback method also failed:', fallbackError);
    throw new Error('Unable to access calendar data. Please check your calendar permissions and try again. Original error: ' + (lastError ? lastError.message : 'Unknown error'));
  }
}

/**
 * Get appointments for a specific date range (main implementation)
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {Array<Object>} Array of appointment objects
 */
function getAppointmentsInRange(startDate, endDate) {
  const config = getCalendarConfig();
  
  try {
    // Try to get the calendar
    let calendar;
    try {
      calendar = CalendarApp.getDefaultCalendar();
    } catch (error) {
      // If default calendar fails, try primary
      calendar = CalendarApp.getCalendarById(config.CALENDAR_ID);
    }
    
    if (!calendar) {
      throw new Error('Unable to access calendar');
    }
    
    // Get events with proper error handling
    const events = calendar.getEvents(startDate, endDate);
    
    if (!events) {
      console.log('No events returned from calendar');
      return [];
    }
    
    console.log(`Found ${events.length} total events in date range`);
    
    // Filter and format events
    const therapyEvents = events
      .filter(event => {
        try {
          return isTherapyAppointment(event);
        } catch (filterError) {
          console.error('Error filtering event:', filterError);
          return false; // Skip problematic events
        }
      })
      .map(event => {
        try {
          return formatAppointmentData(event);
        } catch (formatError) {
          console.error('Error formatting event:', formatError);
          return null;
        }
      })
      .filter(appointment => appointment !== null) // Remove failed formatting attempts
      .sort((a, b) => new Date(a.startTime) - new Date(b.startTime));
    
    console.log(`Filtered to ${therapyEvents.length} therapy appointments`);
    return therapyEvents;
    
  } catch (error) {
    console.error('Error in getAppointmentsInRange:', error);
    throw error;
  }
}

/**
 * Fallback method for getting appointments when primary method fails
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {Array<Object>} Array of appointment objects
 */
function getAppointmentsFallback(startDate, endDate) {
  console.log('Using fallback calendar method');
  
  try {
    // Try a simpler approach with calendar list
    const calendars = CalendarApp.getAllCalendars();
    console.log(`Found ${calendars.length} accessible calendars`);
    
    let allEvents = [];
    
    // Check the first few calendars (usually primary is first)
    const calendarsToCheck = calendars.slice(0, 3);
    
    for (let i = 0; i < calendarsToCheck.length; i++) {
      try {
        const calendar = calendarsToCheck[i];
        console.log(`Checking calendar: ${calendar.getName()}`);
        
        const events = calendar.getEvents(startDate, endDate);
        
        if (events && events.length > 0) {
          const therapyEvents = events
            .filter(event => isTherapyAppointment(event))
            .map(event => formatAppointmentData(event));
          
          allEvents = allEvents.concat(therapyEvents);
        }
        
      } catch (calendarError) {
        console.error(`Error accessing calendar ${i}:`, calendarError);
        continue; // Try next calendar
      }
    }
    
    // Remove duplicates and sort
    const uniqueEvents = removeDuplicateEvents(allEvents);
    return uniqueEvents.sort((a, b) => new Date(a.startTime) - new Date(b.startTime));
    
  } catch (error) {
    console.error('Fallback method failed:', error);
    throw error;
  }
}

/**
 * Remove duplicate events based on start time and title
 * @param {Array} events - Array of events
 * @return {Array} Deduplicated events
 */
function removeDuplicateEvents(events) {
  const seen = new Set();
  return events.filter(event => {
    const key = `${event.startTime.getTime()}-${event.title}`;
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

/**
 * Check if an error should not be retried
 * @param {Error} error - The error to check
 * @return {boolean} True if error should not be retried
 */
function isNonRetryableError(error) {
  const errorMessage = error.message ? error.message.toLowerCase() : '';
  
  // Don't retry permission errors
  if (errorMessage.includes('permission') || errorMessage.includes('access denied')) {
    return true;
  }
  
  // Don't retry invalid calendar errors
  if (errorMessage.includes('calendar not found') || errorMessage.includes('invalid calendar')) {
    return true;
  }
  
  return false;
}

/**
 * Check if an event is likely a therapy appointment with error handling
 * @param {CalendarEvent} event - Calendar event
 * @return {boolean} True if likely a therapy appointment
 */
function isTherapyAppointment(event) {
  try {
    const config = getCalendarConfig();
    
    // Basic null checks
    if (!event) return false;
    
    // Get title safely
    let title = '';
    try {
      title = event.getTitle() || '';
    } catch (error) {
      console.error('Error getting event title:', error);
      return false;
    }
    
    // Get description safely
    let description = '';
    try {
      description = event.getDescription() || '';
    } catch (error) {
      // Description errors are non-fatal
      description = '';
    }
    
    const titleLower = title.toLowerCase();
    const descriptionLower = description.toLowerCase();
    
    // Check for therapy keywords in title or description
    const hasKeyword = config.THERAPY_KEYWORDS.some(keyword => 
      titleLower.includes(keyword) || descriptionLower.includes(keyword)
    );
    
    // Get start time safely
    let startTime;
    try {
      startTime = event.getStartTime();
    } catch (error) {
      console.error('Error getting event start time:', error);
      return false;
    }
    
    if (!startTime) return false;
    
    // Check if it's during business hours
    const startHour = startTime.getHours();
    const isDuringBusinessHours = startHour >= config.START_HOUR && startHour < config.END_HOUR;
    
    // Check if it's not an all-day event
    let isNotAllDay = true;
    try {
      isNotAllDay = !event.isAllDayEvent();
    } catch (error) {
      // Assume not all-day if we can't determine
      isNotAllDay = true;
    }
    
    return hasKeyword && isDuringBusinessHours && isNotAllDay;
    
  } catch (error) {
    console.error('Error in isTherapyAppointment:', error);
    return false;
  }
}

/**
 * Format calendar event data for the application with error handling
 * @param {CalendarEvent} event - Calendar event
 * @return {Object} Formatted appointment object
 */
function formatAppointmentData(event) {
  try {
    // Get basic event data with error handling
    const startTime = event.getStartTime();
    const endTime = event.getEndTime();
    const title = event.getTitle() || 'Untitled Event';
    
    if (!startTime || !endTime) {
      throw new Error('Invalid event times');
    }
    
    // Get optional data safely
    let description = '';
    let location = '';
    let eventId = '';
    let guestList = [];
    let isRecurring = false;
    
    try { description = event.getDescription() || ''; } catch (e) { /* ignore */ }
    try { location = event.getLocation() || ''; } catch (e) { /* ignore */ }
    try { eventId = event.getId() || generateEventId(startTime, title); } catch (e) { 
      eventId = generateEventId(startTime, title);
    }
    try { 
      guestList = event.getGuestList().map(guest => guest.getEmail()); 
    } catch (e) { 
      guestList = [];
    }
    try { isRecurring = event.isRecurringEvent(); } catch (e) { isRecurring = false; }
    
    // Calculate duration safely
    const duration = Math.round((endTime - startTime) / (1000 * 60)); // minutes
    
    // Extract client name from event title
    const extractedClient = extractClientFromTitle(title);
    
    return {
      id: eventId,
      title: title,
      description: description,
      startTime: startTime,
      endTime: endTime,
      duration: duration,
      timeDisplay: formatTimeRange(startTime, endTime),
      extractedClientName: extractedClient.name,
      extractedClientId: extractedClient.id,
      location: location,
      attendees: guestList,
      isRecurring: isRecurring
    };
    
  } catch (error) {
    console.error('Error formatting appointment data:', error);
    throw error;
  }
}

/**
 * Generate a simple event ID when the real ID is not available
 * @param {Date} startTime - Event start time
 * @param {string} title - Event title
 * @return {string} Generated ID
 */
function generateEventId(startTime, title) {
  const timestamp = startTime.getTime();
  const titleHash = title.split('').reduce((a, b) => {
    a = ((a << 5) - a) + b.charCodeAt(0);
    return a & a;
  }, 0);
  return `generated_${timestamp}_${Math.abs(titleHash)}`;
}

/**
 * Extract client information from event title
 * @param {string} title - Event title
 * @return {Object} Object with extracted name and potential ID
 */
function extractClientFromTitle(title) {
  // Common patterns for client names in appointment titles:
  // "Therapy Session - John Doe"
  // "John Doe - Therapy"
  // "Session: John Doe"
  // "John Doe (C001)"
  
  let name = '';
  let id = '';
  
  try {
    // Extract ID in parentheses
    const idMatch = title.match(/\(([A-Z0-9]+)\)/);
    if (idMatch) {
      id = idMatch[1];
    }
    
    // Remove common session keywords and extract name
    let cleanTitle = title
      .replace(/\s*-\s*(therapy|session|counseling|appointment|meeting|consultation|treatment)\s*/gi, '')
      .replace(/\s*(therapy|session|counseling|appointment|meeting|consultation|treatment)\s*-?\s*/gi, '')
      .replace(/\([^)]*\)/g, '') // Remove parentheses content
      .trim();
    
    name = cleanTitle || title;
    
  } catch (error) {
    console.error('Error extracting client from title:', error);
    name = title; // Fallback to full title
  }
  
  return { name, id };
}

/**
 * Format time range for display
 * @param {Date} startTime - Start time
 * @param {Date} endTime - End time
 * @return {string} Formatted time range
 */
function formatTimeRange(startTime, endTime) {
  try {
    const options = { 
      hour: 'numeric', 
      minute: '2-digit',
      hour12: true 
    };
    
    const start = startTime.toLocaleTimeString('en-US', options);
    const end = endTime.toLocaleTimeString('en-US', options);
    
    return `${start} - ${end}`;
    
  } catch (error) {
    console.error('Error formatting time range:', error);
    return 'Time not available';
  }
}

/**
 * Get appointments with matched client data
 * @param {Date} startDate - Start date (optional, defaults to today)
 * @param {Date} endDate - End date (optional, defaults to today)
 * @param {string} sheetName - Name of the clients sheet
 * @return {Array<Object>} Array of appointments with client matches
 */
function getAppointmentsWithClientMatches(startDate, endDate, sheetName) {
  sheetName = sheetName || 'Clients';
  
  if (!startDate || !endDate) {
    const today = new Date();
    startDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    endDate = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
  }
  
  try {
    const appointments = getAppointmentsInRangeWithRetry(startDate, endDate);
    
    return appointments.map(appointment => {
      try {
        const matchedClient = matchAppointmentToClient(appointment, sheetName);
        
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
    
  } catch (error) {
    console.error('Error getting appointments with client matches:', error);
    throw error;
  }
}

/**
 * Match calendar appointment with client in spreadsheet
 * @param {Object} appointment - Appointment object
 * @param {string} sheetName - Name of the clients sheet
 * @return {Object|null} Matched client object or null
 */
function matchAppointmentToClient(appointment, sheetName) {
  try {
    // Get all clients safely
    let allClients = [];
    try {
      if (typeof getAllClients !== 'undefined') {
        allClients = getAllClients(sheetName);
      } else {
        console.warn('getAllClients function not available');
        return null;
      }
    } catch (error) {
      console.error('Error getting all clients:', error);
      return null;
    }
    
    if (!allClients || allClients.length === 0) {
      console.log('No clients found in spreadsheet');
      return null;
    }
    
    // First try to match by extracted ID
    if (appointment.extractedClientId) {
      const byId = allClients.find(client => 
        client.id && client.id.toLowerCase() === appointment.extractedClientId.toLowerCase()
      );
      if (byId) return byId;
    }
    
    // Then try to match by name (fuzzy matching)
    if (appointment.extractedClientName) {
      const extractedName = appointment.extractedClientName.toLowerCase().trim();
      
      // Exact name match
      const exactMatch = allClients.find(client => 
        client.name && client.name.toLowerCase() === extractedName
      );
      if (exactMatch) return exactMatch;
      
      // Partial name match (client name contains extracted name or vice versa)
      const partialMatch = allClients.find(client => {
        if (!client.name) return false;
        const clientName = client.name.toLowerCase();
        return clientName.includes(extractedName) || extractedName.includes(clientName);
      });
      if (partialMatch) return partialMatch;
    }
    
    return null;
    
  } catch (error) {
    console.error('Error matching appointment to client:', error);
    return null;
  }
}

/**
 * Calculate confidence score for client matching
 * @param {Object} appointment - Appointment object
 * @param {Object|null} client - Matched client or null
 * @return {string} Confidence level: 'high', 'medium', 'low', 'none'
 */
function calculateMatchConfidence(appointment, client) {
  if (!client) return 'none';
  
  try {
    // High confidence: ID match or exact name match
    if (appointment.extractedClientId && client.id &&
        client.id.toLowerCase() === appointment.extractedClientId.toLowerCase()) {
      return 'high';
    }
    
    if (appointment.extractedClientName && client.name &&
        client.name.toLowerCase() === appointment.extractedClientName.toLowerCase()) {
      return 'high';
    }
    
    // Medium confidence: partial name match
    if (appointment.extractedClientName && client.name) {
      const extractedName = appointment.extractedClientName.toLowerCase();
      const clientName = client.name.toLowerCase();
      
      if (clientName.includes(extractedName) || extractedName.includes(clientName)) {
        return 'medium';
      }
    }
    
    return 'low';
    
  } catch (error) {
    console.error('Error calculating match confidence:', error);
    return 'low';
  }
}