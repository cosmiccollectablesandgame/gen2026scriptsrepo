/**
 * BP Total Pipeline Service v8.0.0
 * @fileoverview CANONICAL pipeline for synchronizing BP_Total from source sheets:
 *   - Attendance_Missions (Attendance Mission Points)
 *   - Flag_Missions (Flag Mission Points)
 *   - Dice Roll Points (Dice Roll Points)
 *
 * AUTHORITATIVE BP_Total SCHEMA (v8.0.0):
 *   PreferredName | Current_BP | Historical_BP | Redeemed_Total | Prestige_BP |
 *   Attendance Mission Points | Flag Mission Points | Dice Roll Points | 
 *   Manual_Adjustment_Points | LastUpdated
 *
 * This is the ONLY file that should write to the BP_Total state columns.
 * 
 * DEPENDENCIES:
 *   - bpHeaderResolver.js (BP_HEADERS, resolveHeaderIndex)
 */

// ============================================================================
// MAIN PIPELINE: updateBPTotalFromSources
// ============================================================================

/**
 * Synchronizes BP_Total from the three mission source sheets.
 * This is the CANONICAL entry point for BP sync operations.
 *
 * Pipeline:
 *   1. Reads mission totals from Attendance_Missions, Flag_Missions, Dice Roll Points
 *   2. Writes them into BP_Total mission columns
 *   3. Computes totalFromSources = att + flag + dice
 *   4. Applies BP_Global_Cap from Prize_Throttle to clamp Current_BP
 *   5. Updates BP_Historical as monotonic lifetime max
 *   6. Updates LastUpdated timestamp
 *
 * @return {number} Count of players updated
 */
function updateBPTotalFromSources() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ensure BP_Total has correct schema
  ensureBPTotalSchemaEnhanced_();

  const bpSheet = ss.getSheetByName('BP_Total');
  if (!bpSheet) {
    console.error('BP_Total sheet not found after schema ensure');
    return 0;
  }

  // Get throttle cap
  const globalCap = getBPGlobalCap_();

  // Read source data
  const attendanceMap = getAttendanceMissionPoints_();
  const flagMap = getFlagMissionPoints_();
  const diceMap = getDiceRollPoints_();

  // Build set of all player names across all sources
  const allNames = new Set([
    ...attendanceMap.keys(),
    ...flagMap.keys(),
    ...diceMap.keys()
  ]);

  if (allNames.size === 0) {
    console.log('No players found in source sheets');
    return 0;
  }

  // Read BP_Total data
  const bpData = bpSheet.getDataRange().getValues();
  const headers = bpData[0];

  // Use header resolver for column mapping
  const colMap = {};
  try {
    colMap[BP_HEADERS.PREFERRED_NAME] = resolveHeaderIndex(headers, BP_HEADERS.PREFERRED_NAME);
    colMap[BP_HEADERS.CURRENT_BP] = resolveHeaderIndex(headers, BP_HEADERS.CURRENT_BP);
    colMap[BP_HEADERS.ATTENDANCE_POINTS] = resolveHeaderIndex(headers, BP_HEADERS.ATTENDANCE_POINTS);
    colMap[BP_HEADERS.FLAG_POINTS] = resolveHeaderIndex(headers, BP_HEADERS.FLAG_POINTS);
    colMap[BP_HEADERS.DICE_POINTS] = resolveHeaderIndex(headers, BP_HEADERS.DICE_POINTS);
    colMap[BP_HEADERS.LAST_UPDATED] = resolveHeaderIndex(headers, BP_HEADERS.LAST_UPDATED);
    colMap[BP_HEADERS.HISTORICAL_BP] = resolveHeaderIndex(headers, BP_HEADERS.HISTORICAL_BP);
  } catch (e) {
    console.error('BP_Total schema validation failed:', e.message);
    return 0;
  }

  // Build map of existing BP_Total rows by PreferredName
  const existingRows = new Map();
  for (let i = 1; i < bpData.length; i++) {
    const name = String(bpData[i][colMap[BP_HEADERS.PREFERRED_NAME]] || '').trim();
    if (name) {
      existingRows.set(name, i); // Row index in data array (0-based)
    }
  }

  const now = new Date();
  let updatedCount = 0;
  const newRows = [];

  // Process all players
  for (const playerName of allNames) {
    const attPoints = attendanceMap.get(playerName) || 0;
    const flagPts = flagMap.get(playerName) || 0;
    const dicePts = diceMap.get(playerName) || 0;

    const totalFromSources = attPoints + flagPts + dicePts;
    const newCurrentBP = Math.min(totalFromSources, globalCap);

    if (existingRows.has(playerName)) {
      // Update existing row
      const rowIdx = existingRows.get(playerName);
      const rowNum = rowIdx + 1; // 1-based for sheet

      // Get current values
      const oldAttendance = coerceNumber(bpData[rowIdx][colMap[BP_HEADERS.ATTENDANCE_POINTS]], 0);
      const oldFlag = coerceNumber(bpData[rowIdx][colMap[BP_HEADERS.FLAG_POINTS]], 0);
      const oldDice = coerceNumber(bpData[rowIdx][colMap[BP_HEADERS.DICE_POINTS]], 0);
      const currentHistorical = coerceNumber(bpData[rowIdx][colMap[BP_HEADERS.HISTORICAL_BP]], 0);

      // Check if anything changed
      const hasChanges = (
        oldAttendance !== attPoints ||
        oldFlag !== flagPts ||
        oldDice !== dicePts
      );

      if (hasChanges) {
        // Update BP_Historical as monotonic max
        const newHistorical = Math.max(currentHistorical, totalFromSources);

        // Write updates
        bpSheet.getRange(rowNum, colMap[BP_HEADERS.ATTENDANCE_POINTS] + 1).setValue(attPoints);
        bpSheet.getRange(rowNum, colMap[BP_HEADERS.FLAG_POINTS] + 1).setValue(flagPts);
        bpSheet.getRange(rowNum, colMap[BP_HEADERS.DICE_POINTS] + 1).setValue(dicePts);
        bpSheet.getRange(rowNum, colMap[BP_HEADERS.CURRENT_BP] + 1).setValue(newCurrentBP);
        bpSheet.getRange(rowNum, colMap[BP_HEADERS.HISTORICAL_BP] + 1).setValue(newHistorical);
        bpSheet.getRange(rowNum, colMap[BP_HEADERS.LAST_UPDATED] + 1).setValue(now);

        updatedCount++;
      }
    } else {
      // New player - prepare row for batch append
      const newRow = new Array(headers.length).fill('');
      newRow[colMap[BP_HEADERS.PREFERRED_NAME]] = playerName;
      newRow[colMap[BP_HEADERS.CURRENT_BP]] = newCurrentBP;
      newRow[colMap[BP_HEADERS.ATTENDANCE_POINTS]] = attPoints;
      newRow[colMap[BP_HEADERS.FLAG_POINTS]] = flagPts;
      newRow[colMap[BP_HEADERS.DICE_POINTS]] = dicePts;
      newRow[colMap[BP_HEADERS.LAST_UPDATED]] = now;
      newRow[colMap[BP_HEADERS.HISTORICAL_BP]] = totalFromSources;

      newRows.push(newRow);
      updatedCount++;
    }
  }

  // Batch append new rows
  if (newRows.length > 0) {
    const startRow = bpSheet.getLastRow() + 1;
    bpSheet.getRange(startRow, 1, newRows.length, headers.length).setValues(newRows);
  }

  // Log the sync
  if (typeof logIntegrityAction === 'function') {
    logIntegrityAction('BP_TOTAL_SYNC', {
      details: `Synced ${updatedCount} player(s) from sources. Cap: ${globalCap}`,
      status: 'SUCCESS'
    });
  }

  return updatedCount;
}

// ============================================================================
// SOURCE SHEET READERS
// ============================================================================

/**
 * Reads Attendance_Missions and returns Map of PreferredName -> Attendance Mission Points
 * @return {Map<string, number>}
 * @private
 */
function getAttendanceMissionPoints_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Attendance_Missions');

  if (!sheet || sheet.getLastRow() <= 1) {
    return new Map();
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Use header resolver
  let nameCol, pointsCol;
  try {
    nameCol = resolveHeaderIndex(headers, BP_HEADERS.PREFERRED_NAME);
    pointsCol = resolveHeaderIndex(headers, BP_HEADERS.ATTENDANCE_POINTS);
  } catch (e) {
    console.warn('Attendance_Missions schema issue:', e.message);
    return new Map();
  }

  const result = new Map();
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][nameCol] || '').trim();
    if (name) {
      const points = coerceNumber(data[i][pointsCol], 0);
      result.set(name, points);
    }
  }

  return result;
}

/**
 * Reads Flag_Missions and returns Map of PreferredName -> Flag Mission Points
 * @return {Map<string, number>}
 * @private
 */
function getFlagMissionPoints_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Flag_Missions');

  if (!sheet || sheet.getLastRow() <= 1) {
    return new Map();
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Use header resolver
  let nameCol, pointsCol;
  try {
    nameCol = resolveHeaderIndex(headers, BP_HEADERS.PREFERRED_NAME);
    pointsCol = resolveHeaderIndex(headers, BP_HEADERS.FLAG_POINTS);
  } catch (e) {
    console.warn('Flag_Missions schema issue:', e.message);
    return new Map();
  }

  const result = new Map();
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][nameCol] || '').trim();
    if (name) {
      const points = coerceNumber(data[i][pointsCol], 0);
      result.set(name, points);
    }
  }

  return result;
}

/**
 * Reads Dice Roll Points (or legacy Dice_Points) and returns Map of PreferredName -> Dice Roll Points
 * @return {Map<string, number>}
 * @private
 */
function getDiceRollPoints_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Try primary sheet name first, then legacy fallback
  let sheet = ss.getSheetByName('Dice Roll Points');
  if (!sheet) {
    sheet = ss.getSheetByName('Dice_Points');
  }

  if (!sheet || sheet.getLastRow() <= 1) {
    return new Map();
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Use header resolver
  let nameCol, pointsCol;
  try {
    nameCol = resolveHeaderIndex(headers, BP_HEADERS.PREFERRED_NAME);
    pointsCol = resolveHeaderIndex(headers, BP_HEADERS.DICE_POINTS);
  } catch (e) {
    console.warn('Dice Roll Points schema issue:', e.message);
    return new Map();
  }

  const result = new Map();
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][nameCol] || '').trim();
    if (name) {
      const points = coerceNumber(data[i][pointsCol], 0);
      result.set(name, points);
    }
  }

  return result;
}

// ============================================================================
// SCHEMA HELPERS
// ============================================================================

/**
 * Ensures BP_Total has the authoritative schema with all required headers.
 * Creates the sheet if missing, or adds missing columns to existing sheet.
 * Does NOT remove or rename existing columns.
 * @private
 */
function ensureBPTotalSchemaEnhanced_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('BP_Total');

  const requiredHeaders = getBPTotalRequiredHeaders();

  if (!sheet) {
    // Create new sheet with full schema
    sheet = ss.insertSheet('BP_Total');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    formatHeaderRow_(sheet, requiredHeaders.length);
    return;
  }

  if (sheet.getLastRow() === 0) {
    // Empty sheet - add headers
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    formatHeaderRow_(sheet, requiredHeaders.length);
    return;
  }

  // Check existing headers and add missing ones
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = requiredHeaders.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });

    console.log('Added missing BP_Total columns:', missing.join(', '));
  }
}

/**
 * Formats the header row with styling
 * @param {Sheet} sheet - The sheet
 * @param {number} numCols - Number of columns
 * @private
 */
function formatHeaderRow_(sheet, numCols) {
  sheet.getRange(1, 1, 1, numCols)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
}

/**
 * Gets BP_Global_Cap from Prize_Throttle (default: 100)
 * @return {number}
 * @private
 */
function getBPGlobalCap_() {
  try {
    if (typeof getThrottleParam === 'function') {
      const cap = getThrottleParam('BP_Global_Cap', 100);
      return coerceNumber(cap, 100);
    }

    if (typeof getThrottleKV === 'function') {
      const throttle = getThrottleKV();
      return coerceNumber(throttle.BP_Global_Cap, 100);
    }
  } catch (e) {
    console.warn('Could not get BP_Global_Cap from throttle:', e.message);
  }

  return 100; // Default cap
}

// ============================================================================
// LEGACY FUNCTION STUBS (DEPRECATED)
// These are kept for backward compatibility but redirect to the canonical pipeline
// ============================================================================

/**
 * @deprecated Use updateBPTotalFromSources() instead
 * Legacy function - redirects to canonical pipeline
 */
function syncBPTotals(silent) {
  console.warn('syncBPTotals is deprecated. Use updateBPTotalFromSources instead.');
  const count = updateBPTotalFromSources();
  return {
    synced: count,
    errors: []
  };
}

/**
 * @deprecated Use ensureBPTotalSchemaEnhanced_() instead
 * Legacy consolidated schema function - now a no-op
 */
function ensureBPTotalConsolidatedSchema() {
  console.warn('ensureBPTotalConsolidatedSchema is deprecated. Using canonical schema.');
  ensureBPTotalSchemaEnhanced_();
}

/**
 * @deprecated No longer needed
 * Legacy migration function - now a no-op
 */
function migrateBPTotalSchema_() {
  console.warn('migrateBPTotalSchema_ is deprecated. Schema is now managed by ensureBPTotalSchemaEnhanced_.');
  ensureBPTotalSchemaEnhanced_();
}

// ============================================================================
// VALIDATION HELPERS
// ============================================================================

/**
 * Validates mission points integrity across source sheets and BP_Total
 * @return {Object} {pass: boolean, issues: Array}
 */
function validateMissionPointsIntegrity() {
  const issues = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check source sheets exist
  const sourceSheets = [
    { name: 'Attendance_Missions', pointsCol: 'Attendance Mission Points' },
    { name: 'Flag_Missions', pointsCol: 'Flag Mission Points' },
    { name: 'Dice Roll Points', pointsCol: 'Dice Roll Points' }
  ];

  sourceSheets.forEach(spec => {
    const sheet = ss.getSheetByName(spec.name);
    if (!sheet) {
      issues.push({
        sheet: spec.name,
        row: 0,
        issue: 'Sheet not found'
      });
      return;
    }

    if (sheet.getLastRow() <= 1) {
      // Empty sheet is OK
      return;
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    if (!headers.includes('PreferredName')) {
      issues.push({
        sheet: spec.name,
        row: 1,
        issue: 'Missing PreferredName column'
      });
    }

    // Check for points column (allowing synonyms)
    const hasPointsCol = headers.some(h =>
      h === spec.pointsCol ||
      h.toLowerCase().includes('points')
    );

    if (!hasPointsCol) {
      issues.push({
        sheet: spec.name,
        row: 1,
        issue: `Missing ${spec.pointsCol} column`
      });
    }
  });

  // Check BP_Total schema
  const bpSheet = ss.getSheetByName('BP_Total');
  if (!bpSheet) {
    issues.push({
      sheet: 'BP_Total',
      row: 0,
      issue: 'Sheet not found'
    });
  } else if (bpSheet.getLastRow() > 0) {
    const headers = bpSheet.getRange(1, 1, 1, bpSheet.getLastColumn()).getValues()[0];
    
    // Use header resolver validation
    const validation = validateBPTotalHeaders(headers);
    
    validation.missing.forEach(missing => {
      issues.push({
        sheet: 'BP_Total',
        row: 1,
        issue: `Missing column: ${missing}`
      });
    });
  }

  return {
    pass: issues.length === 0,
    issues: issues
  };
}