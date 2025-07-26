/**
 * DialogHandlers.gs - Server-side dialog event handlers (FIXED)
 * Handles communication between HTML dialogs and server-side functions
 * Fixed to prevent serialization errors with Date objects and complex data
 */

/**
 * Calendar Integration Dialog Handlers
 */

/**
 * Load today's appointments with client matching and documentation status
 * @return {Array<Object>} Array of appointment objects with enhanced data (serialization-safe)
 */
function getTodaysAppointmentsForDialog() {
  try {
    // Get today's appointments from calendar
    const appointments = getTodaysAppointments();

    // Add client matching and documentation status
    const enhancedAppointments = enhanceAppointmentsWithStatus(appointments);

    // Make appointments serialization-safe
    return makeAppointmentsSerializationSafe(enhancedAppointments);
  } catch (error) {
    console.error("Error loading today's appointments:", error);
    throw new Error("Unable to load calendar appointments: " + error.message);
  }
}

/**
 * Get appointments for a specific date
 * @param {string} targetDateString - Target date as ISO string
 * @return {Array<Object>} Array of appointment objects with enhanced data
 */
function getAppointmentsForDateDialog(targetDateString) {
  try {
    // Convert string back to Date object
    const targetDate = new Date(targetDateString);

    if (isNaN(targetDate.getTime())) {
      throw new Error("Invalid date provided");
    }

    const appointments = getAppointmentsForDate(targetDate);
    const enhancedAppointments = enhanceAppointmentsWithStatus(appointments);

    return makeAppointmentsSerializationSafe(enhancedAppointments);
  } catch (error) {
    console.error("Error loading appointments for date:", error);
    throw new Error(
      "Unable to load appointments for selected date: " + error.message
    );
  }
}

/**
 * Get appointments for a date range
 * @param {string} startDateString - Start date as ISO string
 * @param {string} endDateString - End date as ISO string
 * @return {Array<Object>} Array of appointment objects with enhanced data
 */
function getAppointmentsForDateRangeDialog(startDateString, endDateString) {
  try {
    // Convert strings back to Date objects
    const startDate = new Date(startDateString);
    const endDate = new Date(endDateString);

    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      throw new Error("Invalid date range provided");
    }

    const appointments = getAppointmentsForDateRange(startDate, endDate);
    const enhancedAppointments = enhanceAppointmentsWithStatus(appointments);

    return makeAppointmentsSerializationSafe(enhancedAppointments);
  } catch (error) {
    console.error("Error loading appointments for date range:", error);
    throw new Error(
      "Unable to load appointments for selected date range: " + error.message
    );
  }
}

/**
 * Make appointments serialization-safe by converting Date objects to strings
 * @param {Array<Object>} appointments - Raw appointment objects
 * @return {Array<Object>} Serialization-safe appointment objects
 */
function makeAppointmentsSerializationSafe(appointments) {
  try {
    if (!Array.isArray(appointments)) {
      return [];
    }

    return appointments.map((appointment) => {
      const safeAppointment = {};

      // Copy all properties, converting Date objects to ISO strings
      for (const [key, value] of Object.entries(appointment)) {
        if (value instanceof Date) {
          safeAppointment[key] = value.toISOString();
        } else if (
          value &&
          typeof value === "object" &&
          value.constructor === Object
        ) {
          // Handle nested objects
          safeAppointment[key] = makeObjectSerializationSafe(value);
        } else if (Array.isArray(value)) {
          // Handle arrays
          safeAppointment[key] = value.map((item) =>
            item instanceof Date ? item.toISOString() : item
          );
        } else {
          // Primitive values and null/undefined
          safeAppointment[key] = value;
        }
      }

      // Add formatted date strings for display
      if (appointment.startTime instanceof Date) {
        safeAppointment.startTimeFormatted =
          appointment.startTime.toLocaleString();
        safeAppointment.dateFormatted =
          appointment.startTime.toLocaleDateString();
        safeAppointment.timeFormatted =
          appointment.startTime.toLocaleTimeString();
      }

      if (appointment.endTime instanceof Date) {
        safeAppointment.endTimeFormatted = appointment.endTime.toLocaleString();
      }

      return safeAppointment;
    });
  } catch (error) {
    console.error("Error making appointments serialization-safe:", error);
    return [];
  }
}

/**
 * Make a nested object serialization-safe
 * @param {Object} obj - Object to make safe
 * @return {Object} Serialization-safe object
 */
function makeObjectSerializationSafe(obj) {
  try {
    if (!obj || typeof obj !== "object") {
      return obj;
    }

    const safeObj = {};

    for (const [key, value] of Object.entries(obj)) {
      if (value instanceof Date) {
        safeObj[key] = value.toISOString();
      } else if (
        value &&
        typeof value === "object" &&
        value.constructor === Object
      ) {
        safeObj[key] = makeObjectSerializationSafe(value);
      } else if (Array.isArray(value)) {
        safeObj[key] = value.map((item) =>
          item instanceof Date ? item.toISOString() : item
        );
      } else {
        safeObj[key] = value;
      }
    }

    return safeObj;
  } catch (error) {
    console.error("Error making object serialization-safe:", error);
    return {};
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
    const appointmentsWithMatches = appointments.map((appointment) => {
      try {
        const matchedClient = matchAppointmentToClient(appointment, "Clients");

        return {
          ...appointment,
          matchedClient: matchedClient,
          hasClientMatch: !!matchedClient,
          matchConfidence: calculateMatchConfidence(appointment, matchedClient),
        };
      } catch (matchError) {
        console.error("Error matching appointment to client:", matchError);
        return {
          ...appointment,
          matchedClient: null,
          hasClientMatch: false,
          matchConfidence: "none",
        };
      }
    });

    // Add documentation status if tracking service is available
    if (typeof getAppointmentStatuses !== "undefined") {
      return getAppointmentStatuses(appointmentsWithMatches);
    } else {
      // Add default status
      return appointmentsWithMatches.map((appointment) => ({
        ...appointment,
        documentationStatus: "not_started",
        sessionRecord: null,
        isDocumented: false,
      }));
    }
  } catch (error) {
    console.error("Error enhancing appointments:", error);
    throw error;
  }
}

/**
 * Show goal selector for a specific appointment
 * @param {Object} appointmentData - Serialized appointment object
 */
function showGoalSelectorForAppointment(appointmentData) {
  try {
    // Deserialize the appointment data
    const appointment = deserializeAppointmentData(appointmentData);

    if (!appointment.hasClientMatch) {
      throw new Error("No client matched for this appointment");
    }

    // Mark session as in progress if tracking is available
    if (typeof markSessionInProgress !== "undefined") {
      markSessionInProgress(
        appointment.id,
        appointment.matchedClient,
        appointment
      );
    }

    // Create session data from appointment
    const sessionData = {
      appointment: appointment,
      client: appointment.matchedClient,
      fromCalendar: true,
    };

    // Show goal selector with appointment context
    showGoalSelector(sessionData);
  } catch (error) {
    console.error("Error showing goal selector for appointment:", error);
    throw new Error("Failed to show goal selector: " + error.message);
  }
}

/**
 * Show client selector for manual appointment matching
 * @param {Object} appointmentData - Serialized appointment object
 */
function showClientSelectorForAppointment(appointmentData) {
  try {
    // Deserialize the appointment data
    const appointment = deserializeAppointmentData(appointmentData);

    const template = HtmlService.createTemplateFromFile("ClientSelector");
    template.appointment = JSON.stringify(appointmentData); // Keep as serialized for HTML
    template.isManualMatch = true;

    const html = template
      .evaluate()
      .setWidth(400)
      .setHeight(500)
      .setTitle("Match Client to Appointment");

    SpreadsheetApp.getUi().showModalDialog(html, "Link Client to Appointment");
  } catch (error) {
    console.error("Error showing client selector for appointment:", error);
    throw new Error("Failed to show client selector: " + error.message);
  }
}

/**
 * Deserialize appointment data (convert ISO strings back to Date objects)
 * @param {Object} appointmentData - Serialized appointment data
 * @return {Object} Deserialized appointment data
 */
function deserializeAppointmentData(appointmentData) {
  try {
    if (!appointmentData || typeof appointmentData !== "object") {
      throw new Error("Invalid appointment data");
    }

    const appointment = { ...appointmentData };

    // Convert ISO strings back to Date objects
    if (appointment.startTime && typeof appointment.startTime === "string") {
      appointment.startTime = new Date(appointment.startTime);
    }

    if (appointment.endTime && typeof appointment.endTime === "string") {
      appointment.endTime = new Date(appointment.endTime);
    }

    // Handle any other date fields that might exist
    const dateFields = ["createdDate", "lastUpdated", "sessionDate"];
    dateFields.forEach((field) => {
      if (appointment[field] && typeof appointment[field] === "string") {
        try {
          appointment[field] = new Date(appointment[field]);
        } catch (dateError) {
          console.warn(`Failed to parse date field ${field}:`, dateError);
        }
      }
    });

    return appointment;
  } catch (error) {
    console.error("Error deserializing appointment data:", error);
    throw error;
  }
}

/**
 * Get documentation URL for an appointment
 * @param {string} appointmentId - Appointment ID
 * @return {string} Documentation URL or null
 */
function getDocumentationForAppointment(appointmentId) {
  try {
    if (typeof getSessionRecord !== "undefined") {
      const sessionRecord = getSessionRecord(appointmentId);
      return sessionRecord ? sessionRecord.formUrl : null;
    }
    return null;
  } catch (error) {
    console.error("Error getting documentation for appointment:", error);
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
    console.error("Error continuing session documentation:", error);
    return null;
  }
}

/**
 * Client Selector Dialog Handlers
 */

/**
 * Load all clients for the client selector dialog
 * @return {Array<Object>} Array of client objects (serialization-safe)
 */
function loadClientsForDialog() {
  try {
    const clients = ClientService.getAllClientsForSelection();

    // Make sure client data is serialization-safe
    return clients.map((client) => {
      const safeClient = {};

      for (const [key, value] of Object.entries(client)) {
        if (value instanceof Date) {
          safeClient[key] = value.toISOString();
        } else {
          safeClient[key] = value;
        }
      }

      return safeClient;
    });
  } catch (error) {
    console.error("Error loading clients for dialog:", error);
    throw new Error("Unable to load client data: " + error.message);
  }
}

/**
 * Search clients based on search term
 * @param {string} searchTerm - Search term to filter clients
 * @return {Array<Object>} Filtered array of client objects (serialization-safe)
 */
function searchClientsForDialog(searchTerm) {
  try {
    const clients = ClientService.searchClients(searchTerm);

    // Make sure client data is serialization-safe
    return clients.map((client) => {
      const safeClient = {};

      for (const [key, value] of Object.entries(client)) {
        if (value instanceof Date) {
          safeClient[key] = value.toISOString();
        } else {
          safeClient[key] = value;
        }
      }

      return safeClient;
    });
  } catch (error) {
    console.error("Error searching clients:", error);
    throw new Error("Search failed: " + error.message);
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
      throw new Error("Invalid client selected");
    }

    // Close current dialog and show goal selector
    showGoalSelector(selectedClient);
  } catch (error) {
    console.error("Error handling client selection:", error);
    throw new Error("Failed to proceed to goal selection: " + error.message);
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
    console.error("Error loading goals for dialog:", error);
    throw new Error("Unable to load goals: " + error.message);
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
      throw new Error("Incomplete session data");
    }

    if (!ClientService.validateClient(sessionData.client)) {
      throw new Error("Invalid client data");
    }

    if (!sessionData.goal || sessionData.goal.trim() === "") {
      throw new Error("No goal selected");
    }

    // Generate the pre-populated form
    generateSessionForm(sessionData);
  } catch (error) {
    console.error("Error handling goal selection:", error);
    throw new Error("Failed to generate session form: " + error.message);
  }
}

/**
 * Handle return to client selector from goal selector
 */
function handleBackToClientSelector() {
  try {
    showClientSelector();
  } catch (error) {
    console.error("Error returning to client selector:", error);
    throw new Error("Failed to return to client selection: " + error.message);
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
      errors: ["Configuration validation failed: " + error.message],
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
    throw new Error("Form generation test failed: " + error.message);
  }
}

/**
 * Load all clients for goal management dialog
 * @return {Array<Object>} Array of client objects with goal information
 */
function getAllClientsForGoalManagement() {
  try {
    const clients = ClientService.getAllClients();

    // Enhance each client with goal information
    return clients.map((client) => {
      const config = ConfigurationService.getSpreadsheetConfig();
      const goalsColumn = config.CLIENT_GOALS_COLUMN || "Goals";

      return {
        id: client[config.CLIENT_ID_COLUMN] || client.ID || "",
        name: client[config.CLIENT_NAME_COLUMN] || client.Name || "",
        email: client[config.CLIENT_EMAIL_COLUMN] || client.Email || "",
        phone: client[config.CLIENT_PHONE_COLUMN] || client.Phone || "",
        Goals: client[goalsColumn] || "",
        // Include any other fields you want to access in goal management
        _rawData: client,
      };
    });
  } catch (error) {
    console.error("Error loading clients for goal management:", error);
    throw new Error("Unable to load client data: " + error.message);
  }
}
