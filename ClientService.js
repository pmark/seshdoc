/**
 * ClientService.gs
 * Client data operations including search, validation, and CRUD operations
 * Handles client data stored in the spreadsheet
 */

const ClientService = {
  
  /**
   * Get all clients from the spreadsheet
   * @returns {Array} Array of client objects
   */
  getAllClients() {
    try {
      const config = ConfigurationService.getSpreadsheetConfig();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(config.CLIENTS_SHEET);
      
      if (!sheet) {
        throw new Error(`Client sheet "${config.CLIENTS_SHEET}" not found`);
      }
      
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) {
        return []; // No data or only headers
      }
      
      const headers = data[0];
      const clients = [];
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const client = {};
        
        headers.forEach((header, index) => {
          client[header] = row[index] || '';
        });
        
        // Add row number for reference
        client._rowNumber = i + 1;
        
        clients.push(client);
      }
      
      return clients;
    } catch (error) {
      Logger.log('Error getting all clients: ' + error.toString());
      return [];
    }
  },
  
  /**
   * Search clients by name or ID
   * @param {string} searchTerm - Search term
   * @param {boolean} exactMatch - Whether to use exact matching
   * @returns {Array} Array of matching client objects
   */
  searchClients(searchTerm, exactMatch = false) {
    try {
      if (!searchTerm || typeof searchTerm !== 'string') {
        return [];
      }
      
      const allClients = this.getAllClients();
      const config = ConfigurationService.getSpreadsheetConfig();
      const searchLower = searchTerm.toLowerCase().trim();
      
      return allClients.filter(client => {
        const id = String(client[config.CLIENT_ID_COLUMN] || '').toLowerCase();
        const name = String(client[config.CLIENT_NAME_COLUMN] || '').toLowerCase();
        
        if (exactMatch) {
          return id === searchLower || name === searchLower;
        } else {
          return id.includes(searchLower) || name.includes(searchLower);
        }
      });
    } catch (error) {
      Logger.log('Error searching clients: ' + error.toString());
      return [];
    }
  },
  
  /**
   * Get client by ID
   * @param {string} clientId - Client ID
   * @returns {Object|null} Client object or null if not found
   */
  getClientById(clientId) {
    try {
      if (!clientId) {
        return null;
      }
      
      const config = ConfigurationService.getSpreadsheetConfig();
      const clients = this.searchClients(clientId, true);
      
      return clients.find(client => 
        String(client[config.CLIENT_ID_COLUMN]) === String(clientId)
      ) || null;
    } catch (error) {
      Logger.log('Error getting client by ID: ' + error.toString());
      return null;
    }
  },
  
  /**
   * Get client by name
   * @param {string} clientName - Client name
   * @returns {Object|null} Client object or null if not found
   */
  getClientByName(clientName) {
    try {
      if (!clientName) {
        return null;
      }
      
      const config = ConfigurationService.getSpreadsheetConfig();
      const clients = this.searchClients(clientName, true);
      
      return clients.find(client => 
        String(client[config.CLIENT_NAME_COLUMN]).toLowerCase() === clientName.toLowerCase()
      ) || null;
    } catch (error) {
      Logger.log('Error getting client by name: ' + error.toString());
      return null;
    }
  },
  
  /**
   * Update client data
   * @param {string} clientId - Client ID
   * @param {Object} updateData - Data to update
   * @returns {boolean} Success status
   */
  updateClient(clientId, updateData) {
    try {
      if (!clientId || !updateData || typeof updateData !== 'object') {
        return false;
      }
      
      const config = ConfigurationService.getSpreadsheetConfig();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(config.CLIENTS_SHEET);
      
      if (!sheet) {
        throw new Error(`Client sheet "${config.CLIENTS_SHEET}" not found`);
      }
      
      const client = this.getClientById(clientId);
      if (!client) {
        throw new Error(`Client with ID "${clientId}" not found`);
      }
      
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const rowNumber = client._rowNumber;
      
      // Update each field
      for (const [field, value] of Object.entries(updateData)) {
        const columnIndex = headers.indexOf(field);
        if (columnIndex !== -1) {
          sheet.getRange(rowNumber, columnIndex + 1).setValue(value);
        }
      }
      
      // Log the update
      Logger.log(`Updated client ${clientId} with data: ${JSON.stringify(updateData)}`);
      
      return true;
    } catch (error) {
      Logger.log('Error updating client: ' + error.toString());
      return false;
    }
  },
  
  /**
   * Add new client
   * @param {Object} clientData - Client data
   * @returns {string|null} New client ID or null if failed
   */
  addClient(clientData) {
    try {
      if (!clientData || typeof clientData !== 'object') {
        return null;
      }
      
      const config = ConfigurationService.getSpreadsheetConfig();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(config.CLIENTS_SHEET);
      
      if (!sheet) {
        throw new Error(`Client sheet "${config.CLIENTS_SHEET}" not found`);
      }
      
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // Generate new client ID if not provided
      if (!clientData[config.CLIENT_ID_COLUMN]) {
        clientData[config.CLIENT_ID_COLUMN] = this.generateClientId();
      }
      
      // Prepare row data
      const newRow = headers.map(header => clientData[header] || '');
      
      // Add to sheet
      sheet.appendRow(newRow);
      
      Logger.log(`Added new client: ${clientData[config.CLIENT_ID_COLUMN]}`);
      
      return clientData[config.CLIENT_ID_COLUMN];
    } catch (error) {
      Logger.log('Error adding client: ' + error.toString());
      return null;
    }
  },
  
  /**
   * Generate unique client ID
   * @returns {string} New client ID
   */
  generateClientId() {
    try {
      const config = ConfigurationService.getSpreadsheetConfig();
      const allClients = this.getAllClients();
      const existingIds = allClients.map(client => 
        String(client[config.CLIENT_ID_COLUMN])
      );
      
      let newId;
      let attempts = 0;
      const maxAttempts = 100;
      
      do {
        // Generate ID in format: C-YYYYMMDD-XXX
        const date = new Date();
        const dateStr = date.getFullYear().toString() + 
                       (date.getMonth() + 1).toString().padStart(2, '0') + 
                       date.getDate().toString().padStart(2, '0');
        const randomNum = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
        
        newId = `C-${dateStr}-${randomNum}`;
        attempts++;
      } while (existingIds.includes(newId) && attempts < maxAttempts);
      
      if (attempts >= maxAttempts) {
        // Fallback to timestamp-based ID
        newId = `C-${Date.now()}`;
      }
      
      return newId;
    } catch (error) {
      Logger.log('Error generating client ID: ' + error.toString());
      return `C-${Date.now()}`;
    }
  },
  
  /**
   * Validate client data
   * @param {Object} clientData - Client data to validate
   * @returns {Object} Validation result
   */
  validateClientData(clientData) {
    try {
      const errors = [];
      const warnings = [];
      
      if (!clientData || typeof clientData !== 'object') {
        errors.push('Client data must be an object');
        return { valid: false, errors, warnings };
      }
      
      const config = ConfigurationService.getSpreadsheetConfig();
      
      // Required fields validation
      const requiredFields = [config.CLIENT_NAME_COLUMN];
      
      for (const field of requiredFields) {
        if (!clientData[field] || String(clientData[field]).trim() === '') {
          errors.push(`${field} is required`);
        }
      }
      
      // Email validation
      if (clientData[config.CLIENT_EMAIL_COLUMN]) {
        const email = String(clientData[config.CLIENT_EMAIL_COLUMN]);
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(email)) {
          errors.push('Invalid email format');
        }
      }
      
      // Phone validation
      if (clientData[config.CLIENT_PHONE_COLUMN]) {
        const phone = String(clientData[config.CLIENT_PHONE_COLUMN]);
        const phoneRegex = /^[\d\s\-\(\)\+]+$/;
        if (!phoneRegex.test(phone)) {
          warnings.push('Phone number format may be invalid');
        }
      }
      
      // Goals validation
      if (clientData[config.CLIENT_GOALS_COLUMN]) {
        const goalsValidation = PipeDelimitedHelpers.validate(
          clientData[config.CLIENT_GOALS_COLUMN]
        );
        if (!goalsValidation.valid) {
          warnings.push('Goals format issues: ' + goalsValidation.issues.join(', '));
        }
      }
      
      return {
        valid: errors.length === 0,
        errors,
        warnings
      };
    } catch (error) {
      Logger.log('Error validating client data: ' + error.toString());
      return {
        valid: false,
        errors: ['Validation error: ' + error.message],
        warnings: []
      };
    }
  },
  
  /**
   * Get client statistics
   * @returns {Object} Statistics object
   */
  getClientStatistics() {
    try {
      const allClients = this.getAllClients();
      const config = ConfigurationService.getSpreadsheetConfig();
      
      const stats = {
        totalClients: allClients.length,
        clientsWithEmail: 0,
        clientsWithPhone: 0,
        clientsWithGoals: 0,
        totalGoals: 0,
        avgGoalsPerClient: 0,
        insuranceProviders: {},
        recentlyAdded: 0
      };
      
      const oneWeekAgo = new Date();
      oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
      
      allClients.forEach(client => {
        // Count clients with contact info
        if (client[config.CLIENT_EMAIL_COLUMN]) {
          stats.clientsWithEmail++;
        }
        
        if (client[config.CLIENT_PHONE_COLUMN]) {
          stats.clientsWithPhone++;
        }
        
        // Count goals
        const goals = PipeDelimitedHelpers.parse(client[config.CLIENT_GOALS_COLUMN]);
        if (goals.length > 0) {
          stats.clientsWithGoals++;
          stats.totalGoals += goals.length;
        }
        
        // Count insurance providers
        const insurance = client[config.INSURANCE_COLUMN];
        if (insurance) {
          stats.insuranceProviders[insurance] = (stats.insuranceProviders[insurance] || 0) + 1;
        }
        
        // Check if recently added (if there's a date field)
        // This would need to be customized based on your actual date field
      });
      
      if (stats.clientsWithGoals > 0) {
        stats.avgGoalsPerClient = Math.round((stats.totalGoals / stats.clientsWithGoals) * 100) / 100;
      }
      
      return stats;
    } catch (error) {
      Logger.log('Error getting client statistics: ' + error.toString());
      return {
        totalClients: 0,
        error: error.message
      };
    }
  },
  
  /**
   * Export clients to CSV format
   * @param {Array} clientIds - Optional array of specific client IDs to export
   * @returns {string} CSV formatted string
   */
  exportClientsToCSV(clientIds = null) {
    try {
      let clients = this.getAllClients();
      
      if (clientIds && Array.isArray(clientIds)) {
        const config = ConfigurationService.getSpreadsheetConfig();
        clients = clients.filter(client => 
          clientIds.includes(client[config.CLIENT_ID_COLUMN])
        );
      }
      
      if (clients.length === 0) {
        return '';
      }
      
      // Get headers (excluding internal fields)
      const headers = Object.keys(clients[0]).filter(key => !key.startsWith('_'));
      
      // Create CSV content
      let csv = headers.map(header => `"${header}"`).join(',') + '\n';
      
      clients.forEach(client => {
        const row = headers.map(header => {
          const value = String(client[header] || '');
          // Escape quotes and wrap in quotes
          return `"${value.replace(/"/g, '""')}"`;
        }).join(',');
        
        csv += row + '\n';
      });
      
      return csv;
    } catch (error) {
      Logger.log('Error exporting clients to CSV: ' + error.toString());
      return '';
    }
  },
  
  /**
   * Match calendar appointment to client
   * @param {string} appointmentTitle - Calendar appointment title
   * @returns {Object|null} Matched client or null
   */
  matchCalendarAppointment(appointmentTitle) {
    try {
      if (!appointmentTitle) {
        return null;
      }
      
      const allClients = this.getAllClients();
      const config = ConfigurationService.getSpreadsheetConfig();
      const titleLower = appointmentTitle.toLowerCase();
      
      // Try exact ID match first
      for (const client of allClients) {
        const clientId = String(client[config.CLIENT_ID_COLUMN]).toLowerCase();
        if (titleLower.includes(clientId)) {
          return client;
        }
      }
      
      // Try name matching (look for name parts in title)
      for (const client of allClients) {
        const clientName = String(client[config.CLIENT_NAME_COLUMN]).toLowerCase();
        const nameParts = clientName.split(' ').filter(part => part.length > 2);
        
        const matchingParts = nameParts.filter(part => titleLower.includes(part));
        
        // If most name parts match, consider it a match
        if (matchingParts.length >= Math.ceil(nameParts.length / 2) && matchingParts.length > 0) {
          return client;
        }
      }
      
      return null;
    } catch (error) {
      Logger.log('Error matching calendar appointment: ' + error.toString());
      return null;
    }
  }
};