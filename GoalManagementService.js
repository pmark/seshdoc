/**
 * GoalManagementService.gs - Client goal management system
 * Provides safe and convenient ways to add, edit, and manage client goals
 */

class GoalManagementService {
  
  /**
   * Add a new goal for a client
   * @param {string} clientId - Client ID
   * @param {string} goalText - Goal text to add
   * @param {string} sheetName - Name of the clients sheet
   * @return {Array<string>} Updated goals array
   */
  static addGoal(clientId, goalText, sheetName = 'Clients') {
    try {
      if (!goalText || goalText.trim() === '') {
        throw new Error('Goal text cannot be empty');
      }
      
      const currentGoals = getGoals(clientId, sheetName);
      const trimmedGoal = goalText.trim();
      
      // Check for duplicates
      if (currentGoals.some(goal => goal.toLowerCase() === trimmedGoal.toLowerCase())) {
        throw new Error('This goal already exists for the client');
      }
      
      const updatedGoals = [...currentGoals, trimmedGoal];
      setGoals(clientId, updatedGoals, sheetName);
      
      // Log the change
      this.logGoalChange(clientId, 'ADDED', trimmedGoal);
      
      return updatedGoals;
      
    } catch (error) {
      console.error('Error adding goal:', error);
      throw new Error('Failed to add goal: ' + error.message);
    }
  }
  
  /**
   * Remove a goal for a client
   * @param {string} clientId - Client ID
   * @param {string} goalText - Goal text to remove
   * @param {string} sheetName - Name of the clients sheet
   * @return {Array<string>} Updated goals array
   */
  static removeGoal(clientId, goalText, sheetName = 'Clients') {
    try {
      const currentGoals = getGoals(clientId, sheetName);
      const updatedGoals = currentGoals.filter(goal => goal !== goalText);
      
      if (updatedGoals.length === currentGoals.length) {
        throw new Error('Goal not found');
      }
      
      setGoals(clientId, updatedGoals, sheetName);
      
      // Log the change
      this.logGoalChange(clientId, 'REMOVED', goalText);
      
      return updatedGoals;
      
    } catch (error) {
      console.error('Error removing goal:', error);
      throw new Error('Failed to remove goal: ' + error.message);
    }
  }
  
  /**
   * Update an existing goal for a client
   * @param {string} clientId - Client ID
   * @param {string} oldGoalText - Original goal text
   * @param {string} newGoalText - New goal text
   * @param {string} sheetName - Name of the clients sheet
   * @return {Array<string>} Updated goals array
   */
  static updateGoal(clientId, oldGoalText, newGoalText, sheetName = 'Clients') {
    try {
      if (!newGoalText || newGoalText.trim() === '') {
        throw new Error('New goal text cannot be empty');
      }
      
      const currentGoals = getGoals(clientId, sheetName);
      const trimmedNewGoal = newGoalText.trim();
      
      // Find the goal to update
      const goalIndex = currentGoals.findIndex(goal => goal === oldGoalText);
      if (goalIndex === -1) {
        throw new Error('Original goal not found');
      }
      
      // Check for duplicates (excluding the goal being updated)
      const otherGoals = currentGoals.filter((_, index) => index !== goalIndex);
      if (otherGoals.some(goal => goal.toLowerCase() === trimmedNewGoal.toLowerCase())) {
        throw new Error('A goal with this text already exists');
      }
      
      const updatedGoals = [...currentGoals];
      updatedGoals[goalIndex] = trimmedNewGoal;
      
      setGoals(clientId, updatedGoals, sheetName);
      
      // Log the change
      this.logGoalChange(clientId, 'UPDATED', `"${oldGoalText}" → "${trimmedNewGoal}"`);
      
      return updatedGoals;
      
    } catch (error) {
      console.error('Error updating goal:', error);
      throw new Error('Failed to update goal: ' + error.message);
    }
  }
  
  /**
   * Reorder goals for a client
   * @param {string} clientId - Client ID
   * @param {Array<string>} reorderedGoals - Goals in new order
   * @param {string} sheetName - Name of the clients sheet
   * @return {Array<string>} Updated goals array
   */
  static reorderGoals(clientId, reorderedGoals, sheetName = 'Clients') {
    try {
      const currentGoals = getGoals(clientId, sheetName);
      
      // Validate that the reordered array contains the same goals
      if (reorderedGoals.length !== currentGoals.length) {
        throw new Error('Reordered goals must contain the same number of goals');
      }
      
      const currentSet = new Set(currentGoals);
      const reorderedSet = new Set(reorderedGoals);
      
      if (currentSet.size !== reorderedSet.size || 
          ![...currentSet].every(goal => reorderedSet.has(goal))) {
        throw new Error('Reordered goals must contain exactly the same goals');
      }
      
      setGoals(clientId, reorderedGoals, sheetName);
      
      // Log the change
      this.logGoalChange(clientId, 'REORDERED', 'Goals reordered');
      
      return reorderedGoals;
      
    } catch (error) {
      console.error('Error reordering goals:', error);
      throw new Error('Failed to reorder goals: ' + error.message);
    }
  }
  
  /**
   * Mark a goal as completed
   * @param {string} clientId - Client ID
   * @param {string} goalText - Goal text to mark as completed
   * @param {string} sheetName - Name of the clients sheet
   * @return {Object} Result with updated goals and archived goal
   */
  static completeGoal(clientId, goalText, sheetName = 'Clients') {
    try {
      // Remove from active goals
      const updatedGoals = this.removeGoal(clientId, goalText, sheetName);
      
      // Add to completed goals archive
      this.archiveCompletedGoal(clientId, goalText);
      
      // Log the completion
      this.logGoalChange(clientId, 'COMPLETED', goalText);
      
      return {
        activeGoals: updatedGoals,
        completedGoal: goalText,
        completedDate: new Date()
      };
      
    } catch (error) {
      console.error('Error completing goal:', error);
      throw new Error('Failed to complete goal: ' + error.message);
    }
  }
  
  /**
   * Archive a completed goal
   * @param {string} clientId - Client ID
   * @param {string} goalText - Completed goal text
   */
  static archiveCompletedGoal(clientId, goalText) {
    try {
      this.initializeGoalArchiveSheet();
      
      const archiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Completed_Goals');
      const archiveRecord = [
        clientId,
        goalText,
        new Date(),
        Session.getActiveUser().getEmail()
      ];
      
      const nextRow = archiveSheet.getLastRow() + 1;
      archiveSheet.getRange(nextRow, 1, 1, archiveRecord.length).setValues([archiveRecord]);
      
    } catch (error) {
      console.error('Error archiving completed goal:', error);
    }
  }
  
  /**
   * Get completed goals for a client
   * @param {string} clientId - Client ID
   * @return {Array<Object>} Array of completed goal objects
   */
  static getCompletedGoals(clientId) {
    try {
      this.initializeGoalArchiveSheet();
      
      const archiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Completed_Goals');
      const data = archiveSheet.getDataRange().getValues();
      
      const completedGoals = [];
      
      // Skip header row and find client's completed goals
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === clientId) {
          completedGoals.push({
            goalText: data[i][1],
            completedDate: data[i][2],
            completedBy: data[i][3]
          });
        }
      }
      
      return completedGoals.sort((a, b) => b.completedDate - a.completedDate);
      
    } catch (error) {
      console.error('Error getting completed goals:', error);
      return [];
    }
  }
  
  /**
   * Get goal statistics for a client
   * @param {string} clientId - Client ID
   * @param {string} sheetName - Name of the clients sheet
   * @return {Object} Goal statistics
   */
  static getGoalStatistics(clientId, sheetName = 'Clients') {
    try {
      const activeGoals = getGoals(clientId, sheetName);
      const completedGoals = this.getCompletedGoals(clientId);
      const goalHistory = this.getGoalHistory(clientId);
      
      return {
        activeGoalsCount: activeGoals.length,
        completedGoalsCount: completedGoals.length,
        totalGoalsEver: activeGoals.length + completedGoals.length,
        lastGoalAdded: goalHistory.length > 0 ? goalHistory[0].date : null,
        lastGoalCompleted: completedGoals.length > 0 ? completedGoals[0].completedDate : null,
        goalCompletionRate: calculateCompletionRate(activeGoals.length, completedGoals.length)
      };
      
    } catch (error) {
      console.error('Error getting goal statistics:', error);
      return {};
    }
  }
  
  /**
   * Log goal changes for audit trail
   * @param {string} clientId - Client ID
   * @param {string} action - Action performed (ADDED, REMOVED, UPDATED, etc.)
   * @param {string} details - Details of the change
   */
  static logGoalChange(clientId, action, details) {
    try {
      this.initializeGoalHistorySheet();
      
      const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Goal_History');
      const logEntry = [
        new Date(),
        clientId,
        action,
        details,
        Session.getActiveUser().getEmail()
      ];
      
      const nextRow = historySheet.getLastRow() + 1;
      historySheet.getRange(nextRow, 1, 1, logEntry.length).setValues([logEntry]);
      
    } catch (error) {
      console.error('Error logging goal change:', error);
    }
  }
  
  /**
   * Get goal change history for a client
   * @param {string} clientId - Client ID
   * @param {number} limit - Maximum number of entries to return
   * @return {Array<Object>} Goal change history
   */
  static getGoalHistory(clientId, limit = 50) {
    try {
      this.initializeGoalHistorySheet();
      
      const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Goal_History');
      const data = historySheet.getDataRange().getValues();
      
      const clientHistory = [];
      
      // Skip header row and find client's history
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === clientId) {
          clientHistory.push({
            date: data[i][0],
            action: data[i][2],
            details: data[i][3],
            user: data[i][4]
          });
        }
      }
      
      // Sort by date (most recent first) and limit results
      return clientHistory
        .sort((a, b) => b.date - a.date)
        .slice(0, limit);
        
    } catch (error) {
      console.error('Error getting goal history:', error);
      return [];
    }
  }
  
  /**
   * Initialize goal archive sheet
   */
  static initializeGoalArchiveSheet() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    if (!spreadsheet.getSheetByName('Completed_Goals')) {
      const archiveSheet = spreadsheet.insertSheet('Completed_Goals');
      
      const headers = ['Client ID', 'Goal Text', 'Completed Date', 'Completed By'];
      archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      archiveSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      archiveSheet.setFrozenRows(1);
    }
  }
  
  /**
   * Initialize goal history sheet
   */
  static initializeGoalHistorySheet() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    if (!spreadsheet.getSheetByName('Goal_History')) {
      const historySheet = spreadsheet.insertSheet('Goal_History');
      
      const headers = ['Date', 'Client ID', 'Action', 'Details', 'User'];
      historySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      historySheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      historySheet.setFrozenRows(1);
    }
  }
  
  /**
   * Bulk update goals with validation
   * @param {string} clientId - Client ID
   * @param {Array<string>} newGoals - New goals array
   * @param {string} sheetName - Name of the clients sheet
   * @return {Array<string>} Updated goals array
   */
  static bulkUpdateGoals(clientId, newGoals, sheetName = 'Clients') {
    try {
      // Validate goals
      const validatedGoals = newGoals
        .filter(goal => goal && goal.trim() !== '')
        .map(goal => goal.trim())
        .filter((goal, index, array) => array.indexOf(goal) === index); // Remove duplicates
      
      setGoals(clientId, validatedGoals, sheetName);
      
      // Log the change
      this.logGoalChange(clientId, 'BULK_UPDATE', `Updated to ${validatedGoals.length} goals`);
      
      return validatedGoals;
      
    } catch (error) {
      console.error('Error bulk updating goals:', error);
      throw new Error('Failed to bulk update goals: ' + error.message);
    }
  }
}

/**
 * Helper function to calculate goal completion rate
 * @param {number} activeCount - Number of active goals
 * @param {number} completedCount - Number of completed goals
 * @return {number} Completion rate percentage
 */
function calculateCompletionRate(activeCount, completedCount) {
  const total = activeCount + completedCount;
  return total > 0 ? Math.round((completedCount / total) * 100) : 0;
}