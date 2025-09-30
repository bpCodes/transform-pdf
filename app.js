
/* app.js
 * Main entry point for Excel processing application
 * 
 * This application processes Excel files in the browser:
 * - Detects range by text: from START_TEXT to END_TEXT
 *   searching from FIND_COL_START column (1-based).
 * - Within the range, checks 3 consecutive columns starting at VALUE_COL_START (1-based).
 *   If it finds a value in any of them, moves it to the first of those 3 and empties the other 2.
 *   If there's no value in the 3, removes the row.
 * - Downloads file with _clean.xlsx suffix
 * 
 * Requires: SheetJS (XLSX) global, and elements #file, #go, #status in the DOM.
 */

// Application initialization
console.log('Excel Processor initialized');

// The application is now modularized:
// - utils.js: Helper functions for parsing, validation, and utilities
// - processor.js: Main Excel processing logic and configuration
// - ui.js: UI event handlers and status management
// - app.js: Entry point (this file)

// All modules are loaded via script tags in the HTML
  
  