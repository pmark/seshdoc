/**
 * SessionTrackingService.gs
 * Session documentation status tracking
 * Handles session records and form submission tracking
 */

const SessionTrackingService = {
  
  /**
   * Get or create sessions sheet
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} Sessions sheet
   */
  getSessionsSheet() {
    try {
      const config = ConfigurationService.getSpreadsheetConfig();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName(config.SESSIONS_SHEET);
      
      if (!sheet) {
        sheet = ss.insertSheet(config.SESSIONS_SHEET);
        
        // Set up headers
        const headers = [
          'Session_ID',
          'Date',
          'Client_ID',
          'Client_Name',
          'Goals_Selected',
          'Form_ID',
          'Form_Response_ID',
          'Status',
          'Created_Date',
          'Last_Updated',
          'Notes'
        ];
        
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        sheet.setFrozenRows(1);
      }
      
      return sheet;
    } catch (error) {
      Logger.log('Error getting sessions sheet: ' + error.toString());
      throw new Error('Unable to access sessions sheet');
    }
  },
  
  /**
   * Get session status for a specific date and client
   * @param {Date} sessionDate - Session date
   * @param {string} clientId - Client ID
   * @returns {string} Status: 'not_started', 'in_progress', 'completed'
   */
  getSessionStatus(sessionDate, clientId) {
    try {
      if (!sessionDate || !clientId) {
        return 'not_started';
      }
      
      const sheet = this.getSessionsSheet();
      const data = sheet.getDataRange().getValues();
      
      if (data.length <= 1) {
        return 'not_started';
      }
      
      const headers = data[0];
      const dateColumn = headers.indexOf('Date');
      const clientIdColumn = headers.indexOf('Client_ID');
      const statusColumn = headers.indexOf('Status');
      
      if (dateColumn === -1 || clientIdColumn === -1 || statusColumn === -1) {
        return 'not_started';
      }
      
      const sessionDateStr = this.formatDateForComparison(sessionDate);
      
      // Find matching session
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowDateStr = this.formatDateForComparison(row[dateColumn]);
        const rowClientId = String(row[clientIdColumn]);
        
        if (rowDateStr === sessionDateStr && rowClientId === String(clientId)) {
          return row[statusColumn] || 'not_started';
        }
      }
      
      return 'not_started';
    } catch (error) {
      Logger.log('Error getting session status: ' + error.toString());
      return 'not_started';
    }
  },
  
  /**
   * Create or update session record
   * @param {Object} sessionData - Session data
   * @returns {string|null} Session ID or null if failed
   */
  createOrUpdateSession(sessionData) {
    try {
      if (!sessionData || typeof sessionData !== 'object') {
        throw new Error('Session data is required');
      }
      
      const {
        sessionDate,
        clientId,
        clientName,
        goalsSelected = '',
        formId = '',
        formResponseId = '',
        status = 'in_progress',
        notes = ''
      } = sessionData;
      
      if (!sessionDate || !clientId) {
        throw new Error('Session date and client ID are required');
      }
      
      const sheet = this.getSessionsSheet();
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      
      const sessionDateStr = this.formatDateForComparison(sessionDate);
      let existingRowIndex = -1;
      
      // Check if session already exists
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowDateStr = this.formatDateForComparison(row[headers.indexOf('Date')]);
        const rowClientId = String(row[headers.indexOf('Client_ID')]);
        
        if (rowDateStr === sessionDateStr && rowClientId === String(clientId)) {
          existingRowIndex = i;
          break;
        }
      }
      
      const sessionId = existingRowIndex > -1 ? 
        data[existingRowIndex][headers.indexOf('Session_ID')] : 
        this.generateSessionId();
      
      const now = new Date();
      const rowData = [
        sessionId,
        sessionDate,
        clientId,
        clientName,
        goalsSelected,
        formId,
        formResponseId,
        status,
        existingRowIndex > -1 ? data[existingRowIndex][headers.indexOf('Created_Date')] : now,
        now,
        notes
      ];
      
      if (existingRowIndex > -1) {
        // Update existing session
        sheet.getRange(existingRowIndex + 1, 1, 1, rowData.length).setValues([rowData]);
        Logger.log(`Updated session: ${sessionId}`);
      } else {
        // Create new session
        sheet.appendRow(rowData);
        Logger.log(`Created session: ${sessionId}`);
      }
      
      return sessionId;
    } catch (error) {
      Logger.log('Error creating/updating session: ' + error.toString());
      return null;
    }
  },
  
  /**
   * Update session status
   * @param {string} sessionId - Session ID
   * @param {string} status - New status
   * @param {Object} additionalData - Additional data to update
   * @returns {boolean} Success status
   */
  updateSessionStatus(sessionId, status, additionalData = {}) {
    try {
      if (!sessionId || !status) {
        return false;
      }
      
      const sheet = this.getSessionsSheet();
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      
      const sessionIdColumn = headers.indexOf('Session_ID');
      const statusColumn = headers.indexOf('Status');
      const lastUpdatedColumn = headers.indexOf('Last_Updated');
      
      if (sessionIdColumn === -1 || statusColumn === -1) {
        return false;
      }
      
      // Find session row
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][sessionIdColumn]) === String(sessionId)) {
          // Update status
          sheet.getRange(i + 1, statusColumn + 1).setValue(status);
          
          // Update last updated time
          if (lastUpdatedColumn > -1) {
            sheet.getRange(i + 1, lastUpdatedColumn + 1).setValue(new Date());
          }
          
          // Update additional data
          for (const [field, value] of Object.entries(additionalData)) {
            const columnIndex = headers.indexOf(field);
            if (columnIndex > -1) {
              sheet.getRange(i + 1, columnIndex + 1).setValue(value);
            }
          }
          
          Logger.log(`Updated session status: ${sessionId} -> ${status}`);
          return true;
        }
      }
      
      return false;
    } catch (error) {
      Logger.log('Error updating session status: ' + error.toString());
      return false;
    }
  },
  
  /**
   * Get all sessions for a client
   * @param {string} clientId - Client ID
   * @param {number} limit - Maximum number of sessions to return
   * @returns {Array} Array of session objects
   */
  getClientSessions(clientId, limit = 50) {
    try {
      if (!clientId) {
        return [];
      }
      
      const sheet = this.getSessionsSheet();
      const data = sheet.getDataRange().getValues();
      
      if (data.length <= 1) {
        return [];
      }
      
      const headers = data[0];
      const sessions = [];
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const session = {};
        
        headers.forEach((header, index) => {
          session[header] = row[index];
        });
        
        if (String(session.Client_ID) === String(clientId)) {
          sessions.push(session);
        }
      }
      
      // Sort by date (most recent first)
      sessions.sort((a, b) => {
        const dateA = new Date(a.Date);
        const dateB = new Date(b.Date);
        return dateB.getTime() - dateA.getTime();
      });
      
      return sessions.slice(0, limit);
    } catch (error) {
      Logger.log('Error getting client sessions: ' + error.toString());
      return [];
    }
  },
  
  /**
   * Get sessions by date range
   * @param {Date} startDate - Start date
   * @param {Date} endDate - End date
   * @returns {Array} Array of session objects
   */
  getSessionsByDateRange(startDate, endDate) {
    try {
      if (!startDate || !endDate) {
        return [];
      }
      
      const sheet = this.getSessionsSheet();
      const data = sheet.getDataRange().getValues();
      
      if (data.length <= 1) {
        return [];
      }
      
      const headers = data[0];
      const sessions = [];
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const session = {};
        
        headers.forEach((header, index) => {
          session[header] = row[index];
        });
        
        const sessionDate = new Date(session.Date);
        if (sessionDate >= startDate && sessionDate <= endDate) {
          sessions.push(session);
        }
      }
      
      return sessions;
    } catch (error) {
      Logger.log('Error getting sessions by date range: ' + error.toString());
      return [];
    }
  },
  
  /**
   * Get session statistics
   * @param {Date} startDate - Optional start date for filtering
   * @param {Date} endDate - Optional end date for filtering
   * @returns {Object} Statistics object
   */
  getSessionStatistics(startDate = null, endDate = null) {
    try {
      let sessions = [];
      
      if (startDate && endDate) {
        sessions = this.getSessionsByDateRange(startDate, endDate);
      } else {
        const sheet = this.getSessionsSheet();
        const data = sheet.getDataRange().getValues();
        
        if (data.length > 1) {
          const headers = data[0];
          for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const session = {};
            headers.forEach((header, index) => {
              session[header] = row[index];
            });
            sessions.push(session);
          }
        }
      }
      
      const stats = {
        totalSessions: sessions.length,
        completedSessions: 0,
        inProgressSessions: 0,
        notStartedSessions: 0,
        uniqueClients: new Set(),
        avgSessionsPerClient: 0,
        mostActiveClient: null,
        recentActivity: []
      };
      
      const clientSessionCounts = {};
      
      sessions.forEach(session => {
        // Count by status
        switch (session.Status) {
          case 'completed':
            stats.completedSessions++;
            break;
          case 'in_progress':
            stats.inProgressSessions++;
            break;
          default:
            stats.notStartedSessions++;
        }
        
        // Track unique clients
        if (session.Client_ID) {
          stats.uniqueClients.add(session.Client_ID);
          clientSessionCounts[session.Client_ID] = (clientSessionCounts[session.Client_ID] || 0) + 1;
        }
        
        // Track recent activity (last 7 days)
        const sessionDate = new Date(session.Date);
        const sevenDaysAgo = new Date();
        sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
        
        if (sessionDate >= sevenDaysAgo) {
          stats.recentActivity.push({
            date: sessionDate,
            clientId: session.Client_ID,
            clientName: session.Client_Name,
            status: session.Status
          });
        }
      });
      
      // Calculate averages
      if (stats.uniqueClients.size > 0) {
        stats.avgSessionsPerClient = Math.round((stats.totalSessions / stats.uniqueClients.size) * 100) / 100;
      }
      
      // Find most active client
      let maxSessions = 0;
      for (const [clientId, count] of Object.entries(clientSessionCounts)) {
        if (count > maxSessions) {
          maxSessions = count;
          stats.mostActiveClient = { clientId, sessionCount: count };
        }
      }
      
      stats.uniqueClients = stats.uniqueClients.size;
      
      return stats;
    } catch (error) {
      Logger.log('Error getting session statistics: ' + error.toString());
      return {
        totalSessions: 0,
        error: error.message
      };
    }
  },
  
  /**
   * Generate unique session ID
   * @returns {string} New session ID
   */
  generateSessionId() {
    try {
      const date = new Date();
      const dateStr = date.getFullYear().toString() + 
                    (date.getMonth() + 1).toString().padStart(2, '0') + 
                    date.getDate().toString().padStart(2, '0');
      
      const timeStr = date.getHours().toString().padStart(2, '0') + 
                     date.getMinutes().toString().padStart(2, '0') + 
                     date.getSeconds().toString().padStart(2, '0');
      
      const randomNum = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
      
      return `S-${dateStr}-${timeStr}-${randomNum}`;
    } catch (error) {
      Logger.log('Error generating session ID: ' + error.toString());
      return `S-${Date.now()}`;
    }
  },
  
  /**
   * Format date for comparison
   * @param {Date|string} date - Date to format
   * @returns {string} Formatted date string
   */
  formatDateForComparison(date) {
    try {
      if (!date) {
        return '';
      }
      
      const dateObj = date instanceof Date ? date : new Date(date);
      
      if (isNaN(dateObj.getTime())) {
        return '';
      }
      
      return dateObj.getFullYear() + '-' + 
             (dateObj.getMonth() + 1).toString().padStart(2, '0') + '-' + 
             dateObj.getDate().toString().padStart(2, '0');
    } catch (error) {
      Logger.log('Error formatting date for comparison: ' + error.toString());
      return '';
    }
  },
  
  /**
   * Delete session record
   * @param {string} sessionId - Session ID to delete
   * @returns {boolean} Success status
   */
  deleteSession(sessionId) {
    try {
      if (!sessionId) {
        return false;
      }
      
      const sheet = this.getSessionsSheet();
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      
      const sessionIdColumn = headers.indexOf('Session_ID');
      
      if (sessionIdColumn === -1) {
        return false;
      }
      
      // Find and delete session row
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][sessionIdColumn]) === String(sessionId)) {
          sheet.deleteRow(i + 1);
          Logger.log(`Deleted session: ${sessionId}`);
          return true;
        }
      }
      
      return false;
    } catch (error) {
      Logger.log('Error deleting session: ' + error.toString());
      return false;
    }
  },
  
  /**
   * Export sessions to CSV
   * @param {Date} startDate - Optional start date filter
   * @param {Date} endDate - Optional end date filter
   * @returns {string} CSV formatted string
   */
  exportSessionsToCSV(startDate = null, endDate = null) {
    try {
      let sessions = [];
      
      if (startDate && endDate) {
        sessions = this.getSessionsByDateRange(startDate, endDate);
      } else {
        const sheet = this.getSessionsSheet();
        const data = sheet.getDataRange().getValues();
        
        if (data.length > 1) {
          const headers = data[0];
          for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const session = {};
            headers.forEach((header, index) => {
              session[header] = row[index];
            });
            sessions.push(session);
          }
        }
      }
      
      if (sessions.length === 0) {
        return '';
      }
      
      const headers = Object.keys(sessions[0]);
      let csv = headers.map(header => `"${header}"`).join(',') + '\n';
      
      sessions.forEach(session => {
        const row = headers.map(header => {
          const value = String(session[header] || '');
          return `"${value.replace(/"/g, '""')}"`;
        }).join(',');
        
        csv += row + '\n';
      });
      
      return csv;
    } catch (error) {
      Logger.log('Error exporting sessions to CSV: ' + error.toString());
      return '';
    }
  }
};