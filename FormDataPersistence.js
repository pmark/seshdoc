/**
 * FormDataPersistence.gs
 * Form response processing with configurable field updates
 * Handles automatic data updates from completed forms
 */

const FormDataPersistence = {
  
  /**
   * Process form response and update client data
   * @param {Object} formResponse - Form response object
   * @returns {boolean} Success status
   */
  processFormResponse(formResponse) {
    try {
      if (!formResponse || typeof formResponse !== 'object') {
        throw new Error('Invalid form response data');
      }
      
      const persistentFields = ConfigurationService.getPersistentFields();
      const fieldMappings = ConfigurationService.getFieldMappings();
      
      // Extract client ID from form response
      const clientId = this.extractClientId(formResponse, fieldMappings);
      if (!clientId) {
        Logger.log('Warning: Could not extract client ID from form response');
        return false;
      }
      
      // Extract session ID if available
      const sessionId = this.extractSessionId(formResponse, fieldMappings);
      
      // Process each persistent field
      const updateData = {};
      const appendData = {};
      
      for (const [fieldKey, config] of Object.entries(persistentFields)) {
        const value = this.extractFieldValue(formResponse, fieldKey, fieldMappings);
        
        if (value !== null && value !== undefined && value !== '') {
          if (config.updateType === 'replace') {
            updateData[config.column] = value;
          } else if (config.updateType === 'append') {
            appendData[config.column] = value;
          }
        }
      }
      
      // Update client record
      let updateSuccess = true;
      
      if (Object.keys(updateData).length > 0) {
        updateSuccess = ClientService.updateClient(clientId, updateData);
      }
      
      // Handle append operations
      if (updateSuccess && Object.keys(appendData).length > 0) {
        updateSuccess = this.appendToClientData(clientId, appendData);
      }
      
      // Update session status
      if (updateSuccess && sessionId) {
        SessionTrackingService.updateSessionStatus(sessionId, 'completed', {
          Form_Response_ID: formResponse.id || '',
          Notes: this.extractSessionNotes(formResponse, fieldMappings)
        });
      }
      
      // Log the processing
      this.logFormProcessing(formResponse, clientId, sessionId, updateSuccess);
      
      return updateSuccess;
    } catch (error) {
      Logger.log('Error processing form response: ' + error.toString());
      return false;
    }
  },
  
  /**
   * Extract client ID from form response
   * @param {Object} formResponse - Form response
   * @param {Object} fieldMappings - Field mappings configuration
   * @returns {string|null} Client ID
   */
  extractClientId(formResponse, fieldMappings) {
    try {
      // Try multiple possible field mappings for client ID
      const possibleFields = ['Client_ID', 'CLIENT_ID', 'Client_Name', 'CLIENT_NAME'];
      
      for (const field of possibleFields) {
        const entryId = fieldMappings[field];
        if (entryId && formResponse[entryId]) {
          const value = formResponse[entryId];
          
          // If this is a name field, try to find the client by name
          if (field.includes('Name')) {
            const client = ClientService.getClientByName(value);
            if (client) {
              const config = ConfigurationService.getSpreadsheetConfig();
              return client[config.CLIENT_ID_COLUMN];
            }
          } else {
            // Assume this is the client ID
            return value;
          }
        }
      }
      
      return null;
    } catch (error) {
      Logger.log('Error extracting client ID: ' + error.toString());
      return null;
    }
  },
  
  /**
   * Extract session ID from form response
   * @param {Object} formResponse - Form response
   * @param {Object} fieldMappings - Field mappings configuration
   * @returns {string|null} Session ID
   */
  extractSessionId(formResponse, fieldMappings) {
    try {
      const sessionIdField = fieldMappings['Session_ID'] || fieldMappings['SESSION_ID'];
      
      if (sessionIdField && formResponse[sessionIdField]) {
        return formResponse[sessionIdField];
      }
      
      return null;
    } catch (error) {
      Logger.log('Error extracting session ID: ' + error.toString());
      return null;
    }
  },
  
  /**
   * Extract field value from form response
   * @param {Object} formResponse - Form response
   * @param {string} fieldKey - Field key
   * @param {Object} fieldMappings - Field mappings configuration
   * @returns {any} Field value
   */
  extractFieldValue(formResponse, fieldKey, fieldMappings) {
    try {
      const entryId = fieldMappings[fieldKey];
      
      if (!entryId || !formResponse[entryId]) {
        return null;
      }
      
      return formResponse[entryId];
    } catch (error) {
      Logger.log(`Error extracting field value for ${fieldKey}: ${error.toString()}`);
      return null;
    }
  },
  
  /**
   * Extract session notes from form response
   * @param {Object} formResponse - Form response
   * @param {Object} fieldMappings - Field mappings configuration
   * @returns {string} Session notes
   */
  extractSessionNotes(formResponse, fieldMappings) {
    try {
      const notesFields = ['Session_Notes', 'SESSION_NOTES', 'Notes', 'NOTES'];
      
      for (const field of notesFields) {
        const entryId = fieldMappings[field];
        if (entryId && formResponse[entryId]) {
          return formResponse[entryId];
        }
      }
      
      return '';
    } catch (error) {
      Logger.log('Error extracting session notes: ' + error.toString());
      return '';
    }
  },
  
  /**
   * Append data to client fields (for fields that accumulate data)
   * @param {string} clientId - Client ID
   * @param {Object} appendData - Data to append
   * @returns {boolean} Success status
   */
  appendToClientData(clientId, appendData) {
    try {
      if (!clientId || !appendData || Object.keys(appendData).length === 0) {
        return true;
      }
      
      const client = ClientService.getClientById(clientId);
      if (!client) {
        throw new Error('Client not found');
      }
      
      const updateData = {};
      const timestamp = new Date().toLocaleDateString();
      
      for (const [column, newValue] of Object.entries(appendData)) {
        const currentValue = client[column] || '';
        
        // Format new entry with timestamp
        const newEntry = `[${timestamp}] ${newValue}`;
        
        // Append using pipe-delimited format
        const updatedValue = PipeDelimitedHelpers.addItem(currentValue, newEntry, true);
        
        updateData[column] = updatedValue;
      }
      
      return ClientService.updateClient(clientId, updateData);
    } catch (error) {
      Logger.log('Error appending to client data: ' + error.toString());
      return false;
    }
  },
  
  /**
   * Set up form submission trigger
   * @param {string} formId - Google Form ID
   * @returns {boolean} Success status
   */
  setupFormSubmissionTrigger(formId) {
    try {
      if (!formId) {
        throw new Error('Form ID is required');
      }
      
      const form = FormApp.openById(formId);
      
      // Delete existing triggers for this form
      this.removeFormSubmissionTriggers(formId);
      
      // Create new trigger
      const trigger = ScriptApp.newTrigger('onFormSubmit')
        .timeBased()
        .everyMinutes(1) // Check for new responses every minute
        .create();
      
      // Store trigger info in script properties
      const triggerInfo = {
        triggerId: trigger.getUniqueId(),
        formId: formId,
        created: new Date().toISOString()
      };
      
      PropertiesService.getScriptProperties().setProperty(
        `FORM_TRIGGER_${formId}`,
        JSON.stringify(triggerInfo)
      );
      
      Logger.log(`Set up form submission trigger for form: ${formId}`);
      
      return true;
    } catch (error) {
      Logger.log('Error setting up form submission trigger: ' + error.toString());
      return false;
    }
  },
  
  /**
   * Remove form submission triggers
   * @param {string} formId - Google Form ID (optional - if not provided, removes all)
   * @returns {boolean} Success status
   */
  removeFormSubmissionTriggers(formId = null) {
    try {
      const triggers = ScriptApp.getProjectTriggers();
      let removedCount = 0;
      
      triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'onFormSubmit') {
          if (!formId) {
            // Remove all form submission triggers
            ScriptApp.deleteTrigger(trigger);
            removedCount++;
          } else {
            // Check if this trigger is for the specific form
            const triggerInfo = PropertiesService.getScriptProperties()
              .getProperty(`FORM_TRIGGER_${formId}`);
            
            if (triggerInfo) {
              const info = JSON.parse(triggerInfo);
              if (info.triggerId === trigger.getUniqueId()) {
                ScriptApp.deleteTrigger(trigger);
                PropertiesService.getScriptProperties()
                  .deleteProperty(`FORM_TRIGGER_${formId}`);
                removedCount++;
              }
            }
          }
        }
      });
      
      Logger.log(`Removed ${removedCount} form submission triggers`);
      
      return true;
    } catch (error) {
      Logger.log('Error removing form submission triggers: ' + error.toString());
      return false;
    }
  },
  
  /**
   * Handle form submission trigger event
   * @param {Object} e - Trigger event object (not used in time-based triggers)
   */
  onFormSubmit(e) {
    try {
      // Get all configured forms and check for new responses
      const formTriggers = this.getActiveFormTriggers();
      
      formTriggers.forEach(triggerInfo => {
        this.processNewFormResponses(triggerInfo.formId);
      });
    } catch (error) {
      Logger.log('Error in form submission handler: ' + error.toString());
    }
  },
  
  /**
   * Get active form triggers
   * @returns {Array} Array of active trigger info
   */
  getActiveFormTriggers() {
    try {
      const properties = PropertiesService.getScriptProperties().getProperties();
      const triggers = [];
      
      for (const [key, value] of Object.entries(properties)) {
        if (key.startsWith('FORM_TRIGGER_')) {
          try {
            const triggerInfo = JSON.parse(value);
            triggers.push(triggerInfo);
          } catch (parseError) {
            Logger.log(`Error parsing trigger info for ${key}: ${parseError.toString()}`);
          }
        }
      }
      
      return triggers;
    } catch (error) {
      Logger.log('Error getting active form triggers: ' + error.toString());
      return [];
    }
  },
  
  /**
   * Process new form responses for a specific form
   * @param {string} formId - Google Form ID
   * @returns {number} Number of responses processed
   */
  processNewFormResponses(formId) {
    try {
      if (!formId) {
        return 0;
      }
      
      const form = FormApp.openById(formId);
      const responses = form.getResponses();
      
      // Get timestamp of last processed response
      const lastProcessedKey = `LAST_PROCESSED_${formId}`;
      const lastProcessedTimestamp = PropertiesService.getScriptProperties()
        .getProperty(lastProcessedKey);
      
      const lastProcessed = lastProcessedTimestamp ? 
        new Date(lastProcessedTimestamp) : 
        new Date(0);
      
      let processedCount = 0;
      let latestTimestamp = lastProcessed;
      
      responses.forEach(response => {
        const responseTimestamp = response.getTimestamp();
        
        if (responseTimestamp > lastProcessed) {
          const formResponseData = this.convertFormResponseToObject(response);
          
          if (this.processFormResponse(formResponseData)) {
            processedCount++;
          }
          
          if (responseTimestamp > latestTimestamp) {
            latestTimestamp = responseTimestamp;
          }
        }
      });
      
      // Update last processed timestamp
      if (latestTimestamp > lastProcessed) {
        PropertiesService.getScriptProperties().setProperty(
          lastProcessedKey,
          latestTimestamp.toISOString()
        );
      }
      
      if (processedCount > 0) {
        Logger.log(`Processed ${processedCount} new form responses for form: ${formId}`);
      }
      
      return processedCount;
    } catch (error) {
      Logger.log(`Error processing new form responses for ${formId}: ${error.toString()}`);
      return 0;
    }
  },
  
  /**
   * Convert Google Forms response to object
   * @param {GoogleAppsScript.Forms.FormResponse} response - Form response
   * @returns {Object} Response data object
   */
  convertFormResponseToObject(response) {
    try {
      const responseData = {
        id: response.getId(),
        timestamp: response.getTimestamp(),
        editResponseUrl: response.getEditResponseUrl()
      };
      
      const itemResponses = response.getItemResponses();
      
      itemResponses.forEach(itemResponse => {
        const item = itemResponse.getItem();
        const entryId = `entry.${item.getId()}`;
        const responseValue = itemResponse.getResponse();
        
        responseData[entryId] = responseValue;
      });
      
      return responseData;
    } catch (error) {
      Logger.log('Error converting form response to object: ' + error.toString());
      return {};
    }
  },
  
  /**
   * Log form processing activity
   * @param {Object} formResponse - Form response
   * @param {string} clientId - Client ID
   * @param {string} sessionId - Session ID
   * @param {boolean} success - Processing success status
   */
  logFormProcessing(formResponse, clientId, sessionId, success) {
    try {
      const logEntry = {
        timestamp: new Date().toISOString(),
        formResponseId: formResponse.id || 'unknown',
        clientId: clientId || 'unknown',
        sessionId: sessionId || 'none',
        success: success,
        processedBy: Session.getActiveUser().getEmail()
      };
      
      // Log to console
      Logger.log(`Form processing: ${JSON.stringify(logEntry)}`);
      
      // Could also log to a dedicated sheet for audit trail
      this.logToProcessingSheet(logEntry);
    } catch (error) {
      Logger.log('Error logging form processing: ' + error.toString());
    }
  },
  
  /**
   * Log to processing sheet for audit trail
   * @param {Object} logEntry - Log entry object
   */
  logToProcessingSheet(logEntry) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let logSheet = ss.getSheetByName('Form_Processing_Log');
      
      if (!logSheet) {
        logSheet = ss.insertSheet('Form_Processing_Log');
        
        // Set up headers
        const headers = [
          'Timestamp',
          'Form_Response_ID',
          'Client_ID',
          'Session_ID',
          'Success',
          'Processed_By',
          'Error_Message'
        ];
        
        logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        logSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        logSheet.setFrozenRows(1);
      }
      
      const rowData = [
        logEntry.timestamp,
        logEntry.formResponseId,
        logEntry.clientId,
        logEntry.sessionId,
        logEntry.success,
        logEntry.processedBy,
        logEntry.errorMessage || ''
      ];
      
      logSheet.appendRow(rowData);
    } catch (error) {
      Logger.log('Error logging to processing sheet: ' + error.toString());
    }
  },
  
  /**
   * Get form processing statistics
   * @param {Date} startDate - Optional start date for filtering
   * @param {Date} endDate - Optional end date for filtering
   * @returns {Object} Processing statistics
   */
  getProcessingStatistics(startDate = null, endDate = null) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = ss.getSheetByName('Form_Processing_Log');
      
      if (!logSheet) {
        return { totalProcessed: 0, successRate: 0 };
      }
      
      const data = logSheet.getDataRange().getValues();
      
      if (data.length <= 1) {
        return { totalProcessed: 0, successRate: 0 };
      }
      
      const headers = data[0];
      const timestampColumn = headers.indexOf('Timestamp');
      const successColumn = headers.indexOf('Success');
      
      let totalProcessed = 0;
      let successfulProcessed = 0;
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const timestamp = new Date(row[timestampColumn]);
        
        // Apply date filters if provided
        if (startDate && timestamp < startDate) continue;
        if (endDate && timestamp > endDate) continue;
        
        totalProcessed++;
        
        if (row[successColumn] === true || row[successColumn] === 'true') {
          successfulProcessed++;
        }
      }
      
      const successRate = totalProcessed > 0 ? 
        Math.round((successfulProcessed / totalProcessed) * 100) : 
        0;
      
      return {
        totalProcessed: totalProcessed,
        successfulProcessed: successfulProcessed,
        failedProcessed: totalProcessed - successfulProcessed,
        successRate: successRate
      };
    } catch (error) {
      Logger.log('Error getting processing statistics: ' + error.toString());
      return {
        totalProcessed: 0,
        successRate: 0,
        error: error.message
      };
    }
  }
};