/**
 * OnEdit Router - Thin routing layer for onEdit triggers
 * @fileoverview Routes onEdit events to appropriate handlers with guard clauses,
 * sheet allowlist validation, debouncing via LockService, and centralized logging
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

/**
 * Sheet allowlist configuration
 * @const {Array<string|RegExp>}
 */
const ALLOWED_SHEETS = [
  'PreferredNames',
  'Player\'s Bonus Points',
  'Players_Prize-Wall-Points',
  'Player\'s Prize-Wall-Points',
  // Event sheet pattern: MM-DD-YYYY or MM-DD-X-YYYY (e.g., 01-15-2024, 12-25-A-2024)
  /^\d{2}-\d{2}(-|[A-Za-z]-)\d{4}$/,
  // Mission source sheets
  'Attendance_Missions',
  'Flag_Missions',
  'Dice Roll Points',
  'Dice_Points',
  // Employee log
  'Employee_Log'
];

/**
 * Lock timeout in milliseconds (5 seconds)
 * @const {number}
 */
const LOCK_TIMEOUT_MS = 5000;

// ============================================================================
// MAIN ROUTER
// ============================================================================

/**
 * Main onEdit trigger router
 * Routes edit events to appropriate handlers with guard clauses and debouncing
 * Called by the onEdit function in Code.js
 * @param {Object} e - Edit event object
 */
function onEditRouter(e) {
  const startedAt = new Date();
  let logContext = {
    sheet: null,
    range: null,
    user: currentUser(),
    action: 'UNKNOWN',
    startedAt: startedAt,
    endedAt: null,
    status: 'UNKNOWN',
    error: null
  };

  try {
    // ========================================================================
    // GUARD CLAUSE: Event object validation
    // ========================================================================
    if (!e) {
      logEvent({
        ...logContext,
        action: 'ONEDIT_IGNORED',
        endedAt: new Date(),
        status: 'IGNORED',
        error: 'Missing event object'
      });
      return;
    }

    if (!e.range) {
      logEvent({
        ...logContext,
        action: 'ONEDIT_IGNORED',
        endedAt: new Date(),
        status: 'IGNORED',
        error: 'Missing range in event object'
      });
      return;
    }

    // ========================================================================
    // GUARD CLAUSE: Multi-cell paste check
    // ========================================================================
    // Ignore edits that modify more than one cell (paste operations)
    // This prevents performance issues and unintended batch operations
    const numRows = e.range.getNumRows();
    const numCols = e.range.getNumColumns();
    const isMultiCell = (numRows > 1 || numCols > 1);

    if (isMultiCell) {
      logEvent({
        ...logContext,
        sheet: e.range.getSheet().getName(),
        range: e.range.getA1Notation(),
        action: 'ONEDIT_IGNORED',
        endedAt: new Date(),
        status: 'IGNORED',
        error: `Multi-cell edit not supported (${numRows}x${numCols})`
      });
      return;
    }

    // ========================================================================
    // GUARD CLAUSE: Sheet allowlist check
    // ========================================================================
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();

    if (!isSheetAllowed(sheetName)) {
      logEvent({
        ...logContext,
        sheet: sheetName,
        range: e.range.getA1Notation(),
        action: 'ONEDIT_IGNORED',
        endedAt: new Date(),
        status: 'IGNORED',
        error: 'Sheet not in allowlist'
      });
      return;
    }

    // Update log context with sheet and range info
    logContext.sheet = sheetName;
    logContext.range = e.range.getA1Notation();

    // ========================================================================
    // DEBOUNCE: LockService to prevent concurrent runs
    // ========================================================================
    const lock = LockService.getScriptLock();
    const lockKey = `onEdit_${sheetName}_${e.range.getA1Notation()}`;
    
    try {
      if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
        logEvent({
          ...logContext,
          action: 'ONEDIT_LOCKED',
          endedAt: new Date(),
          status: 'SKIPPED',
          error: 'Another onEdit is already running for this location'
        });
        return;
      }

      // ======================================================================
      // ROUTING: Dispatch to appropriate handlers
      // ======================================================================
      routeEditEvent(e, logContext);

    } finally {
      // Always release the lock
      try {
        lock.releaseLock();
      } catch (unlockErr) {
        console.error('Failed to release lock:', unlockErr);
      }
    }

  } catch (err) {
    // Top-level error handler
    logContext.endedAt = new Date();
    logContext.status = 'ERROR';
    logContext.error = err.message || String(err);
    logEvent(logContext);
    console.error('onEdit router error:', err);
  }
}

// ============================================================================
// ROUTING LOGIC
// ============================================================================

/**
 * Routes edit event to the appropriate handler based on sheet name
 * @param {Object} e - Edit event object
 * @param {Object} logContext - Logging context
 */
function routeEditEvent(e, logContext) {
  const sheetName = e.range.getSheet().getName();
  
  // Mission source sheets -> BP Total sync
  const missionSheets = ['Attendance_Missions', 'Flag_Missions', 'Dice Roll Points', 'Dice_Points'];
  if (missionSheets.includes(sheetName)) {
    logContext.action = 'MISSION_SHEET_EDIT';
    handleMissionSheetEdit(e, logContext);
    return;
  }

  // Dice Points checkbox handler
  if (sheetName === 'Dice_Points') {
    logContext.action = 'DICE_CHECKBOX_EDIT';
    handleDicePointsEdit(e, logContext);
    return;
  }

  // Employee Log edit handler
  if (sheetName === 'Employee_Log') {
    logContext.action = 'EMPLOYEE_LOG_EDIT';
    handleEmployeeLogEdit(e, logContext);
    return;
  }

  // PreferredNames edit handler
  if (sheetName === 'PreferredNames') {
    logContext.action = 'PREFERRED_NAMES_EDIT';
    handlePreferredNamesEdit(e, logContext);
    return;
  }

  // Player's Bonus Points variants
  const bonusPointsSheets = ['Player\'s Bonus Points', 'Players_Prize-Wall-Points', 'Player\'s Prize-Wall-Points'];
  if (bonusPointsSheets.includes(sheetName)) {
    logContext.action = 'BONUS_POINTS_EDIT';
    handleBonusPointsEdit(e, logContext);
    return;
  }

  // Event sheets (MM-DD-YYYY or MM-DD-X-YYYY pattern)
  if (isEventSheet(sheetName)) {
    logContext.action = 'EVENT_SHEET_EDIT';
    handleEventSheetEdit(e, logContext);
    return;
  }

  // If we get here, log as unhandled (should not happen due to allowlist)
  logContext.action = 'ONEDIT_UNHANDLED';
  logContext.status = 'SKIPPED';
  logContext.endedAt = new Date();
  logEvent(logContext);
}

// ============================================================================
// HANDLER FUNCTIONS
// ============================================================================

/**
 * Handles edits to mission source sheets (Attendance_Missions, Flag_Missions, Dice_Points)
 * Triggers BP_Total sync when data rows are edited
 * @param {Object} e - Edit event object
 * @param {Object} logContext - Logging context
 */
function handleMissionSheetEdit(e, logContext) {
  try {
    const editRow = e.range.getRow();
    
    // Only sync if the edit was to data rows (not header)
    if (editRow > 1 && typeof updateBPTotalFromSources === 'function') {
      updateBPTotalFromSources();
      logContext.status = 'SUCCESS';
      logContext.endedAt = new Date();
      logEvent(logContext);
    } else {
      logContext.status = 'SKIPPED';
      logContext.error = 'Header row edit or no sync function available';
      logContext.endedAt = new Date();
      logEvent(logContext);
    }
  } catch (err) {
    logContext.status = 'ERROR';
    logContext.error = err.message || String(err);
    logContext.endedAt = new Date();
    logEvent(logContext);
    console.error('Mission sheet edit handler error:', err);
  }
}

/**
 * Handles Dice Points checkbox edits
 * @param {Object} e - Edit event object
 * @param {Object} logContext - Logging context
 */
function handleDicePointsEdit(e, logContext) {
  try {
    if (typeof onDicePointCheckboxEdit === 'function') {
      onDicePointCheckboxEdit(e);
      logContext.status = 'SUCCESS';
    } else {
      logContext.status = 'SKIPPED';
      logContext.error = 'onDicePointCheckboxEdit function not available';
    }
    logContext.endedAt = new Date();
    logEvent(logContext);
  } catch (err) {
    logContext.status = 'ERROR';
    logContext.error = err.message || String(err);
    logContext.endedAt = new Date();
    logEvent(logContext);
    console.error('Dice points edit handler error:', err);
  }
}

/**
 * Handles Employee Log edits
 * @param {Object} e - Edit event object
 * @param {Object} logContext - Logging context
 */
function handleEmployeeLogEdit(e, logContext) {
  try {
    if (typeof handleEmployeeLogEdit_ === 'function') {
      handleEmployeeLogEdit_(e);
      logContext.status = 'SUCCESS';
    } else {
      logContext.status = 'SKIPPED';
      logContext.error = 'handleEmployeeLogEdit_ function not available';
    }
    logContext.endedAt = new Date();
    logEvent(logContext);
  } catch (err) {
    logContext.status = 'ERROR';
    logContext.error = err.message || String(err);
    logContext.endedAt = new Date();
    logEvent(logContext);
    console.error('Employee log edit handler error:', err);
  }
}

/**
 * Handles PreferredNames sheet edits
 * @param {Object} e - Edit event object
 * @param {Object} logContext - Logging context
 */
function handlePreferredNamesEdit(e, logContext) {
  try {
    // Currently no specific handler for PreferredNames edits
    // This is a placeholder for future functionality
    logContext.status = 'SKIPPED';
    logContext.error = 'No handler implemented for PreferredNames';
    logContext.endedAt = new Date();
    logEvent(logContext);
  } catch (err) {
    logContext.status = 'ERROR';
    logContext.error = err.message || String(err);
    logContext.endedAt = new Date();
    logEvent(logContext);
    console.error('PreferredNames edit handler error:', err);
  }
}

/**
 * Handles Player's Bonus Points / Prize-Wall-Points edits
 * @param {Object} e - Edit event object
 * @param {Object} logContext - Logging context
 */
function handleBonusPointsEdit(e, logContext) {
  try {
    // Currently no specific handler for bonus points edits
    // This is a placeholder for future functionality
    logContext.status = 'SKIPPED';
    logContext.error = 'No handler implemented for Bonus Points sheets';
    logContext.endedAt = new Date();
    logEvent(logContext);
  } catch (err) {
    logContext.status = 'ERROR';
    logContext.error = err.message || String(err);
    logContext.endedAt = new Date();
    logEvent(logContext);
    console.error('Bonus points edit handler error:', err);
  }
}

/**
 * Handles event sheet edits (MM-DD-YYYY or MM-DD-X-YYYY)
 * @param {Object} e - Edit event object
 * @param {Object} logContext - Logging context
 */
function handleEventSheetEdit(e, logContext) {
  try {
    // Currently no specific handler for event sheet edits
    // This is a placeholder for future functionality
    logContext.status = 'SKIPPED';
    logContext.error = 'No handler implemented for event sheets';
    logContext.endedAt = new Date();
    logEvent(logContext);
  } catch (err) {
    logContext.status = 'ERROR';
    logContext.error = err.message || String(err);
    logContext.endedAt = new Date();
    logEvent(logContext);
    console.error('Event sheet edit handler error:', err);
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Checks if a sheet name is in the allowlist
 * @param {string} sheetName - Name of the sheet
 * @return {boolean} True if sheet is allowed
 */
function isSheetAllowed(sheetName) {
  for (const allowed of ALLOWED_SHEETS) {
    if (typeof allowed === 'string') {
      if (sheetName === allowed) {
        return true;
      }
    } else if (allowed instanceof RegExp) {
      if (allowed.test(sheetName)) {
        return true;
      }
    }
  }
  return false;
}

/**
 * Checks if a sheet name matches the event sheet pattern
 * @param {string} sheetName - Name of the sheet
 * @return {boolean} True if sheet matches event pattern
 */
function isEventSheet(sheetName) {
  const eventPattern = /^\d{2}-\d{2}(-|[A-Za-z]-)\d{4}$/;
  return eventPattern.test(sheetName);
}

// ============================================================================
// LOGGING
// ============================================================================

/**
 * Central logging function for onEdit events
 * Logs to Integrity_Log sheet or falls back to Logger
 * @param {Object} context - Logging context
 * @param {string} context.sheet - Sheet name
 * @param {string} context.range - Range in A1 notation
 * @param {string} context.user - User email
 * @param {string} context.action - Action type
 * @param {Date} context.startedAt - Start timestamp
 * @param {Date} context.endedAt - End timestamp
 * @param {string} context.status - Status (SUCCESS, ERROR, SKIPPED, IGNORED)
 * @param {string} context.error - Error message (if any)
 */
function logEvent(context) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('Integrity_Log');
    
    // Create Integrity_Log sheet if it doesn't exist
    if (!logSheet) {
      logSheet = ss.insertSheet('Integrity_Log');
      logSheet.appendRow([
        'Timestamp',
        'StoreID',
        'Event_ID',
        'Action',
        'Operator',
        'PreferredName',
        'Seed',
        'Checksum_Before',
        'Checksum_After',
        'RL_Band',
        'DF_Tags',
        'Details',
        'Status'
      ]);
      logSheet.setFrozenRows(1);
      logSheet.getRange('A1:M1')
        .setFontWeight('bold')
        .setBackground('#2196f3')
        .setFontColor('#ffffff');
    }

    // Format timestamp
    const timestamp = context.startedAt ? 
      Utilities.formatDate(context.startedAt, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'") :
      dateISO();

    // Calculate duration if both timestamps are available
    let details = '';
    if (context.startedAt && context.endedAt) {
      const durationMs = context.endedAt.getTime() - context.startedAt.getTime();
      details += `Duration: ${durationMs}ms`;
    }
    
    if (context.sheet) {
      details += (details ? ' | ' : '') + `Sheet: ${context.sheet}`;
    }
    
    if (context.range) {
      details += (details ? ' | ' : '') + `Range: ${context.range}`;
    }
    
    if (context.error) {
      details += (details ? ' | ' : '') + `Error: ${context.error}`;
    }

    // Append row to Integrity_Log
    const row = [
      timestamp,
      'MAIN', // StoreID
      '', // Event_ID (not applicable for onEdit)
      context.action || 'ONEDIT',
      context.user || currentUser(),
      '', // PreferredName (not applicable for onEdit)
      '', // Seed
      '', // Checksum_Before
      '', // Checksum_After
      '', // RL_Band
      '', // DF_Tags
      details,
      context.status || 'UNKNOWN'
    ];
    
    logSheet.appendRow(row);
    
  } catch (err) {
    // Fallback to Logger if Integrity_Log write fails
    console.error('Failed to write to Integrity_Log:', err);
    Logger.log(`[onEdit] ${context.action} - ${context.status} - ${context.sheet || 'N/A'}:${context.range || 'N/A'} - ${context.error || ''}`);
  }
}
