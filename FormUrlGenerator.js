/**
 * FormUrlGenerator.gs - Google Form URL generation
 * Base class for creating pre-populated Google Form URLs
 * 
 * IMPORTANT: This must be defined before EnhancedFormGenerator.gs
 */

/**
 * Form URL Generator - Base functionality
 */
function getFormConfig() {
  return {
    // Replace with your actual Google Form ID
    FORM_ID: '1dJSpL3_yLcy1mBh9wf-of1BVakpPO6aQnvHm-7VqeJo',

    // Replace with your actual form field entry IDs
    // To find these: open your form, go to preview, inspect element on each field
    FIELD_MAPPINGS: {
      CLIENT_NAME: 'entry.123456789',    // Replace with actual entry ID
      CLIENT_ID: 'entry.987654321',      // Replace with actual entry ID  
      SELECTED_GOAL: 'entry.456789123',  // Replace with actual entry ID
      SESSION_DATE: 'entry.789123456'    // Replace with actual entry ID
    }
  };
}

/**
 * Creates a pre-populated Google Form URL
 * @param {Object} sessionData - Session data object
 * @param {Object} sessionData.client - Client object with id and name
 * @param {string} sessionData.goal - Selected goal
 * @return {string} Pre-populated form URL
 */
function createPrePopulatedUrl(sessionData) {
  const config = getFormConfig();
  
  // Validate input
  if (!sessionData || !sessionData.client || !sessionData.goal) {
    throw new Error('Invalid session data provided');
  }
  
  if (!validateClient(sessionData.client)) {
    throw new Error('Invalid client data provided');
  }
  
  // Build the base URL
  const baseUrl = `https://docs.google.com/forms/d/${config.FORM_ID}/formResponse`;
  
  // Create URL parameters
  const params = [];
  
  // Add pre-filled form data
  if (config.FIELD_MAPPINGS.CLIENT_NAME) {
    params.push(`${config.FIELD_MAPPINGS.CLIENT_NAME}=${encodeURIComponent(sessionData.client.name)}`);
  }
  
  if (config.FIELD_MAPPINGS.CLIENT_ID) {
    params.push(`${config.FIELD_MAPPINGS.CLIENT_ID}=${encodeURIComponent(sessionData.client.id)}`);
  }
  
  if (config.FIELD_MAPPINGS.SELECTED_GOAL) {
    params.push(`${config.FIELD_MAPPINGS.SELECTED_GOAL}=${encodeURIComponent(sessionData.goal)}`);
  }
  
  if (config.FIELD_MAPPINGS.SESSION_DATE) {
    params.push(`${config.FIELD_MAPPINGS.SESSION_DATE}=${encodeURIComponent(new Date().toLocaleDateString())}`);
  }
  
  // Return the complete URL
  return `${baseUrl}?${params.join('&')}`;
}

/**
 * Validates form configuration
 * @return {Object} Validation result with isValid boolean and errors array
 */
function validateFormConfiguration() {
  const config = getFormConfig();
  const errors = [];
  
  if (!config.FORM_ID || config.FORM_ID === 'YOUR_GOOGLE_FORM_ID_HERE') {
    errors.push('Form ID not configured');
  }
  
  if (!config.FIELD_MAPPINGS.CLIENT_NAME || config.FIELD_MAPPINGS.CLIENT_NAME.startsWith('entry.123')) {
    errors.push('Client Name field mapping not configured');
  }
  
  if (!config.FIELD_MAPPINGS.SELECTED_GOAL || config.FIELD_MAPPINGS.SELECTED_GOAL.startsWith('entry.456')) {
    errors.push('Selected Goal field mapping not configured');
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

/**
 * Helper method to test form URL generation
 * @return {string} Test URL for debugging
 */
function createTestFormUrl() {
  const testData = {
    client: {
      id: 'TEST001',
      name: 'Test Client'
    },
    goal: 'Test Goal'
  };
  
  return createPrePopulatedUrl(testData);
}

/**
 * Validate client data
 * @param {Object} client - Client object to validate
 * @return {boolean} True if valid
 */
function validateClient(client) {
  return client && 
         typeof client.id === 'string' && 
         typeof client.name === 'string' &&
         client.id.trim() !== '' && 
         client.name.trim() !== '';
}