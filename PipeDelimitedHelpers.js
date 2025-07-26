/**
 * PipeDelimitedHelpers.gs
 * Generic utilities for handling pipe-delimited data
 * Provides safe handling of data containing commas, quotes, and special characters
 */

const PipeDelimitedHelpers = {
  
  /**
   * Parse pipe-delimited string into array
   * @param {string} data - Pipe-delimited string
   * @param {boolean} trim - Whether to trim whitespace
   * @returns {Array} Array of parsed items
   */
  parse(data, trim = true) {
    try {
      if (!data || typeof data !== 'string') {
        return [];
      }
      
      const items = data.split('|');
      return trim ? items.map(item => item.trim()).filter(item => item.length > 0) : items;
    } catch (error) {
      Logger.log('Error parsing pipe-delimited data: ' + error.toString());
      return [];
    }
  },
  
  /**
   * Convert array to pipe-delimited string
   * @param {Array} items - Array of items
   * @param {boolean} trim - Whether to trim whitespace
   * @returns {string} Pipe-delimited string
   */
  stringify(items, trim = true) {
    try {
      if (!Array.isArray(items)) {
        return '';
      }
      
      const processedItems = trim ? 
        items.map(item => String(item).trim()).filter(item => item.length > 0) : 
        items.map(item => String(item));
      
      return processedItems.join('|');
    } catch (error) {
      Logger.log('Error stringifying to pipe-delimited: ' + error.toString());
      return '';
    }
  },
  
  /**
   * Add item to pipe-delimited string
   * @param {string} data - Existing pipe-delimited string
   * @param {string} newItem - Item to add
   * @param {boolean} allowDuplicates - Whether to allow duplicate items
   * @returns {string} Updated pipe-delimited string
   */
  addItem(data, newItem, allowDuplicates = false) {
    try {
      if (!newItem || typeof newItem !== 'string') {
        return data || '';
      }
      
      const items = this.parse(data);
      const trimmedNewItem = newItem.trim();
      
      if (!allowDuplicates && items.includes(trimmedNewItem)) {
        return data;
      }
      
      items.push(trimmedNewItem);
      return this.stringify(items);
    } catch (error) {
      Logger.log('Error adding item to pipe-delimited data: ' + error.toString());
      return data || '';
    }
  },
  
  /**
   * Remove item from pipe-delimited string
   * @param {string} data - Existing pipe-delimited string
   * @param {string} itemToRemove - Item to remove
   * @returns {string} Updated pipe-delimited string
   */
  removeItem(data, itemToRemove) {
    try {
      if (!data || !itemToRemove) {
        return data || '';
      }
      
      const items = this.parse(data);
      const trimmedItem = itemToRemove.trim();
      
      const filteredItems = items.filter(item => item !== trimmedItem);
      return this.stringify(filteredItems);
    } catch (error) {
      Logger.log('Error removing item from pipe-delimited data: ' + error.toString());
      return data || '';
    }
  },
  
  /**
   * Update item in pipe-delimited string
   * @param {string} data - Existing pipe-delimited string
   * @param {string} oldItem - Item to replace
   * @param {string} newItem - New item
   * @returns {string} Updated pipe-delimited string
   */
  updateItem(data, oldItem, newItem) {
    try {
      if (!data || !oldItem || !newItem) {
        return data || '';
      }
      
      const items = this.parse(data);
      const trimmedOldItem = oldItem.trim();
      const trimmedNewItem = newItem.trim();
      
      const updatedItems = items.map(item => 
        item === trimmedOldItem ? trimmedNewItem : item
      );
      
      return this.stringify(updatedItems);
    } catch (error) {
      Logger.log('Error updating item in pipe-delimited data: ' + error.toString());
      return data || '';
    }
  },
  
  /**
   * Reorder items in pipe-delimited string
   * @param {string} data - Existing pipe-delimited string
   * @param {Array} newOrder - Array of items in new order
   * @returns {string} Reordered pipe-delimited string
   */
  reorderItems(data, newOrder) {
    try {
      if (!data || !Array.isArray(newOrder)) {
        return data || '';
      }
      
      const currentItems = this.parse(data);
      const validNewOrder = newOrder.filter(item => currentItems.includes(item));
      
      // Add any items not in newOrder to the end
      const missingItems = currentItems.filter(item => !newOrder.includes(item));
      const finalOrder = [...validNewOrder, ...missingItems];
      
      return this.stringify(finalOrder);
    } catch (error) {
      Logger.log('Error reordering pipe-delimited data: ' + error.toString());
      return data || '';
    }
  },
  
  /**
   * Check if item exists in pipe-delimited string
   * @param {string} data - Pipe-delimited string
   * @param {string} item - Item to check
   * @returns {boolean} True if item exists
   */
  contains(data, item) {
    try {
      if (!data || !item) {
        return false;
      }
      
      const items = this.parse(data);
      return items.includes(item.trim());
    } catch (error) {
      Logger.log('Error checking if pipe-delimited data contains item: ' + error.toString());
      return false;
    }
  },
  
  /**
   * Get count of items in pipe-delimited string
   * @param {string} data - Pipe-delimited string
   * @returns {number} Count of items
   */
  getCount(data) {
    try {
      if (!data) {
        return 0;
      }
      
      const items = this.parse(data);
      return items.length;
    } catch (error) {
      Logger.log('Error getting count of pipe-delimited data: ' + error.toString());
      return 0;
    }
  },
  
  /**
   * Merge multiple pipe-delimited strings
   * @param {Array} dataArray - Array of pipe-delimited strings
   * @param {boolean} removeDuplicates - Whether to remove duplicates
   * @returns {string} Merged pipe-delimited string
   */
  merge(dataArray, removeDuplicates = true) {
    try {
      if (!Array.isArray(dataArray)) {
        return '';
      }
      
      let allItems = [];
      
      dataArray.forEach(data => {
        if (data && typeof data === 'string') {
          const items = this.parse(data);
          allItems = allItems.concat(items);
        }
      });
      
      if (removeDuplicates) {
        allItems = [...new Set(allItems)];
      }
      
      return this.stringify(allItems);
    } catch (error) {
      Logger.log('Error merging pipe-delimited data: ' + error.toString());
      return '';
    }
  },
  
  /**
   * Filter items in pipe-delimited string
   * @param {string} data - Pipe-delimited string
   * @param {Function} filterFunction - Function to filter items
   * @returns {string} Filtered pipe-delimited string
   */
  filter(data, filterFunction) {
    try {
      if (!data || typeof filterFunction !== 'function') {
        return data || '';
      }
      
      const items = this.parse(data);
      const filteredItems = items.filter(filterFunction);
      
      return this.stringify(filteredItems);
    } catch (error) {
      Logger.log('Error filtering pipe-delimited data: ' + error.toString());
      return data || '';
    }
  },
  
  /**
   * Map items in pipe-delimited string
   * @param {string} data - Pipe-delimited string
   * @param {Function} mapFunction - Function to map items
   * @returns {string} Mapped pipe-delimited string
   */
  map(data, mapFunction) {
    try {
      if (!data || typeof mapFunction !== 'function') {
        return data || '';
      }
      
      const items = this.parse(data);
      const mappedItems = items.map(mapFunction);
      
      return this.stringify(mappedItems);
    } catch (error) {
      Logger.log('Error mapping pipe-delimited data: ' + error.toString());
      return data || '';
    }
  },
  
  /**
   * Validate pipe-delimited string format
   * @param {string} data - Pipe-delimited string
   * @returns {Object} Validation result
   */
  validate(data) {
    try {
      if (!data) {
        return { valid: true, itemCount: 0, issues: [] };
      }
      
      if (typeof data !== 'string') {
        return { valid: false, itemCount: 0, issues: ['Data must be a string'] };
      }
      
      const items = this.parse(data);
      const issues = [];
      
      // Check for empty items
      const emptyItems = items.filter(item => item.trim() === '');
      if (emptyItems.length > 0) {
        issues.push(`Found ${emptyItems.length} empty items`);
      }
      
      // Check for duplicate items
      const uniqueItems = [...new Set(items)];
      if (uniqueItems.length !== items.length) {
        issues.push(`Found ${items.length - uniqueItems.length} duplicate items`);
      }
      
      // Check for items with special characters that might cause issues
      const problematicItems = items.filter(item => 
        item.includes('\n') || item.includes('\r') || item.includes('\t')
      );
      if (problematicItems.length > 0) {
        issues.push(`Found ${problematicItems.length} items with line breaks or tabs`);
      }
      
      return {
        valid: issues.length === 0,
        itemCount: items.length,
        uniqueItemCount: uniqueItems.length,
        issues: issues
      };
    } catch (error) {
      Logger.log('Error validating pipe-delimited data: ' + error.toString());
      return { valid: false, itemCount: 0, issues: ['Validation error: ' + error.message] };
    }
  },
  
  /**
   * Clean pipe-delimited string
   * @param {string} data - Pipe-delimited string
   * @param {Object} options - Cleaning options
   * @returns {string} Cleaned pipe-delimited string
   */
  clean(data, options = {}) {
    try {
      if (!data) {
        return '';
      }
      
      const defaultOptions = {
        removeDuplicates: true,
        removeEmpty: true,
        trimItems: true,
        removeLineBreaks: true
      };
      
      const opts = { ...defaultOptions, ...options };
      
      let items = this.parse(data, opts.trimItems);
      
      // Remove empty items
      if (opts.removeEmpty) {
        items = items.filter(item => item.trim().length > 0);
      }
      
      // Remove line breaks and tabs
      if (opts.removeLineBreaks) {
        items = items.map(item => 
          item.replace(/[\n\r\t]/g, ' ').replace(/\s+/g, ' ').trim()
        );
      }
      
      // Remove duplicates
      if (opts.removeDuplicates) {
        items = [...new Set(items)];
      }
      
      return this.stringify(items);
    } catch (error) {
      Logger.log('Error cleaning pipe-delimited data: ' + error.toString());
      return data || '';
    }
  }
};