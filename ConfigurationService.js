/**
 * ConfigurationService.gs
 * Centralized configuration management system
 * Handles field mappings, persistence settings, and admin interface
 */

const ConfigurationService = {
  
  // Default configuration constants
  DEFAULT_FIELD_MAPPINGS: {
    'Client_Name': 'entry.123456789',
    'Client_Email': 'entry.987654321',
    'Client_Phone': 'entry.456789123',
    'Insurance_Provider': 'entry.789123456',
    'Session_Date': 'entry.321654987',
    'Goals_Selected': 'entry.654987321',
    'Session_Notes': 'entry.987321654',
    'Homework_Assigned': 'entry.147258369',
    'Next_Session_Date': 'entry.369258147'
  },
  
  DEFAULT_PERSISTENT_FIELDS: {
    'CLIENT_EMAIL': { column: 'Email', updateType: 'replace' },
    'CLIENT_PHONE': { column: 'Phone', updateType: 'replace' },
    'INSURANCE_PROVIDER': { column: 'Insurance_Provider', updateType: 'replace' },
    'MEDICAL_HISTORY': { column: 'Medical_History', updateType: 'append' },
    'EMERGENCY_CONTACT': { column: 'Emergency_Contact', updateType: 'replace' },
    'THERAPY_GOALS': { column: 'Goals', updateType: 'append' },
    'SESSION_NOTES': { column: 'Session_History', updateType: 'append' }
  },
  
  DEFAULT_SPREADSHEET_CONFIG: {
    'CLIENTS_SHEET': 'Clients',
    'SESSIONS_SHEET': 'Sessions',
    'CONFIG_SHEET': 'Config',
    'CLIENT_ID_COLUMN': 'ID',
    'CLIENT_NAME_COLUMN': 'Name',
    'CLIENT_GOALS_COLUMN': 'Goals',
    'CLIENT_EMAIL_COLUMN': 'Email',
    'CLIENT_PHONE_COLUMN': 'Phone',
    'INSURANCE_COLUMN': 'Insurance_Provider'
  },
  
  /**
   * Initialize configuration system
   */
  initializeConfiguration() {
    try {
      // Check if configuration exists
      const existingConfig = this.getConfiguration();
      
      if (!existingConfig || Object.keys(existingConfig).length === 0) {
        // Initialize with defaults
        this.setConfiguration({
          fieldMappings: this.DEFAULT_FIELD_MAPPINGS,
          persistentFields: this.DEFAULT_PERSISTENT_FIELDS,
          spreadsheetConfig: this.DEFAULT_SPREADSHEET_CONFIG,
          lastUpdated: new Date().toISOString(),
          version: '1.0.0'
        });
        
        Logger.log('Configuration initialized with defaults');
      }
      
      return true;
    } catch (error) {
      Logger.log('Error initializing configuration: ' + error.toString());
      return false;
    }
  },
  
  /**
   * Get full configuration object
   */
  getConfiguration() {
    try {
      const configString = PropertiesService.getScriptProperties().getProperty('THERAPY_CONFIG');
      if (!configString) {
        return null;
      }
      
      return JSON.parse(configString);
    } catch (error) {
      Logger.log('Error getting configuration: ' + error.toString());
      return null;
    }
  },
  
  /**
   * Set complete configuration
   */
  setConfiguration(config) {
    try {
      config.lastUpdated = new Date().toISOString();
      
      PropertiesService.getScriptProperties().setProperty(
        'THERAPY_CONFIG', 
        JSON.stringify(config)
      );
      
      // Also backup to spreadsheet if Config sheet exists
      this.backupConfigToSheet(config);
      
      return true;
    } catch (error) {
      Logger.log('Error setting configuration: ' + error.toString());
      return false;
    }
  },
  
  /**
   * Get field mappings
   */
  getFieldMappings() {
    try {
      const config = this.getConfiguration();
      return config ? config.fieldMappings : this.DEFAULT_FIELD_MAPPINGS;
    } catch (error) {
      Logger.log('Error getting field mappings: ' + error.toString());
      return this.DEFAULT_FIELD_MAPPINGS;
    }
  },
  
  /**
   * Update field mappings
   */
  updateFieldMappings(newMappings) {
    try {
      const config = this.getConfiguration() || {};
      config.fieldMappings = { ...config.fieldMappings, ...newMappings };
      
      return this.setConfiguration(config);
    } catch (error) {
      Logger.log('Error updating field mappings: ' + error.toString());
      return false;
    }
  },
  
  /**
   * Get persistent fields configuration
   */
  getPersistentFields() {
    try {
      const config = this.getConfiguration();
      return config ? config.persistentFields : this.DEFAULT_PERSISTENT_FIELDS;
    } catch (error) {
      Logger.log('Error getting persistent fields: ' + error.toString());
      return this.DEFAULT_PERSISTENT_FIELDS;
    }
  },
  
  /**
   * Update persistent fields configuration
   */
  updatePersistentFields(newFields) {
    try {
      const config = this.getConfiguration() || {};
      config.persistentFields = { ...config.persistentFields, ...newFields };
      
      return this.setConfiguration(config);
    } catch (error) {
      Logger.log('Error updating persistent fields: ' + error.toString());
      return false;
    }
  },
  
  /**
   * Get spreadsheet configuration
   */
  getSpreadsheetConfig() {
    try {
      const config = this.getConfiguration();
      return config ? config.spreadsheetConfig : this.DEFAULT_SPREADSHEET_CONFIG;
    } catch (error) {
      Logger.log('Error getting spreadsheet config: ' + error.toString());
      return this.DEFAULT_SPREADSHEET_CONFIG;
    }
  },
  
  /**
   * Update spreadsheet configuration
   */
  updateSpreadsheetConfig(newConfig) {
    try {
      const config = this.getConfiguration() || {};
      config.spreadsheetConfig = { ...config.spreadsheetConfig, ...newConfig };
      
      return this.setConfiguration(config);
    } catch (error) {
      Logger.log('Error updating spreadsheet config: ' + error.toString());
      return false;
    }
  },
  
  /**
   * Validate form entry IDs
   */
  validateFormEntryIds(formId, entryIds) {
    try {
      const form = FormApp.openById(formId);
      const items = form.getItems();
      const validEntryIds = [];
      
      items.forEach(item => {
        const itemId = item.getId();
        validEntryIds.push('entry.' + itemId);
      });
      
      const results = {};
      for (const [field, entryId] of Object.entries(entryIds)) {
        results[field] = {
          valid: validEntryIds.includes(entryId),
          entryId: entryId
        };
      }
      
      return results;
    } catch (error) {
      Logger.log('Error validating form entry IDs: ' + error.toString());
      return null;
    }
  },
  
  /**
   * Export configuration to JSON
   */
  exportConfiguration() {
    try {
      const config = this.getConfiguration();
      if (!config) {
        return null;
      }
      
      return {
        ...config,
        exportDate: new Date().toISOString(),
        exportedBy: Session.getActiveUser().getEmail()
      };
    } catch (error) {
      Logger.log('Error exporting configuration: ' + error.toString());
      return null;
    }
  },
  
  /**
   * Import configuration from JSON
   */
  importConfiguration(configData) {
    try {
      // Validate configuration structure
      if (!configData || typeof configData !== 'object') {
        throw new Error('Invalid configuration data');
      }
      
      // Ensure required fields exist
      const requiredFields = ['fieldMappings', 'persistentFields', 'spreadsheetConfig'];
      for (const field of requiredFields) {
        if (!configData[field]) {
          throw new Error(`Missing required field: ${field}`);
        }
      }
      
      // Set imported configuration
      configData.importDate = new Date().toISOString();
      configData.importedBy = Session.getActiveUser().getEmail();
      
      return this.setConfiguration(configData);
    } catch (error) {
      Logger.log('Error importing configuration: ' + error.toString());
      return false;
    }
  },
  
  /**
   * Reset configuration to defaults
   */
  resetConfiguration() {
    try {
      const defaultConfig = {
        fieldMappings: this.DEFAULT_FIELD_MAPPINGS,
        persistentFields: this.DEFAULT_PERSISTENT_FIELDS,
        spreadsheetConfig: this.DEFAULT_SPREADSHEET_CONFIG,
        lastUpdated: new Date().toISOString(),
        version: '1.0.0',
        resetDate: new Date().toISOString(),
        resetBy: Session.getActiveUser().getEmail()
      };
      
      return this.setConfiguration(defaultConfig);
    } catch (error) {
      Logger.log('Error resetting configuration: ' + error.toString());
      return false;
    }
  },
  
  /**
   * Backup configuration to spreadsheet
   */
  backupConfigToSheet(config) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let configSheet = ss.getSheetByName('Config');
      
      if (!configSheet) {
        configSheet = ss.insertSheet('Config');
      }
      
      // Clear existing data
      configSheet.clear();
      
      // Set headers
      configSheet.getRange(1, 1, 1, 3).setValues([['Key', 'Value', 'Type']]);
      
      // Add configuration data
      const rows = [];
      
      // Field mappings
      for (const [key, value] of Object.entries(config.fieldMappings || {})) {
        rows.push([`fieldMappings.${key}`, value, 'string']);
      }
      
      // Persistent fields
      for (const [key, value] of Object.entries(config.persistentFields || {})) {
        rows.push([`persistentFields.${key}`, JSON.stringify(value), 'json']);
      }
      
      // Spreadsheet config
      for (const [key, value] of Object.entries(config.spreadsheetConfig || {})) {
        rows.push([`spreadsheetConfig.${key}`, value, 'string']);
      }
      
      // Metadata
      rows.push(['lastUpdated', config.lastUpdated, 'date']);
      rows.push(['version', config.version, 'string']);
      
      if (rows.length > 0) {
        configSheet.getRange(2, 1, rows.length, 3).setValues(rows);
      }
      
      return true;
    } catch (error) {
      Logger.log('Error backing up config to sheet: ' + error.toString());
      return false;
    }
  },
  
  /**
   * Get configuration summary for display
   */
  getConfigurationSummary() {
    try {
      const config = this.getConfiguration();
      if (!config) {
        return {
          status: 'Not configured',
          fieldMappings: 0,
          persistentFields: 0,
          lastUpdated: 'Never'
        };
      }
      
      return {
        status: 'Configured',
        fieldMappings: Object.keys(config.fieldMappings || {}).length,
        persistentFields: Object.keys(config.persistentFields || {}).length,
        lastUpdated: config.lastUpdated || 'Unknown',
        version: config.version || '1.0.0'
      };
    } catch (error) {
      Logger.log('Error getting configuration summary: ' + error.toString());
      return {
        status: 'Error',
        error: error.message
      };
    }
  }
};