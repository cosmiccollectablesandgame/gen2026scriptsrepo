/**
 * Mission Points Service - Points Pipeline Consolidation
 * @fileoverview Manages the three mission point sources (Flag_Missions, Attendance_Missions, Dice_Points)
 *               and consolidates them into BP_Total
 */

// ============================================================================
// CONSTANTS - Mission Point Categories
// ============================================================================

/**
 * Flag Mission categories (discretionary awards)
 */
const FLAG_MISSION_CATEGORIES = [
  'Cosmic_Selfie',
  'Review_Writer',
  'Social_Media_Star',
  'App_Explorer',
  'Cosmic_Merchant',
  'Precon_Pioneer',
  'Gravitational_Pull',
  'Rogue_Planet',
  'Quantum_Collector'
];

/**
 * Default points per flag mission category
 */
const FLAG_MISSION_DEFAULTS = {
  'Cosmic_Selfie': 5,
  'Review_Writer': 10,
  'Social_Media_Star': 5,
  'App_Explorer': 5,
  'Cosmic_Merchant': 15,
  'Precon_Pioneer': 10,
  'Gravitational_Pull': 10,
  'Rogue_Planet': 20,
  'Quantum_Collector': 25
};

/**
 * Attendance milestone thresholds and point awards
 */
const ATTENDANCE_MILESTONES = {
  5: 5,     // 5 events = 5 bonus points
  10: 10,   // 10 events = 10 bonus points
  25: 15,   // 25 events = 15 bonus points
  50: 25,   // 50 events = 25 bonus points
  100: 50   // 100 events = 50 bonus points
};

// ============================================================================
// PREFERRED NAMES - Canonical Player Registry
// ============================================================================

// NOTE: ensurePreferredNamesSchema() moved to PlayerProvisioning.js (canonical)
// NOTE: getAllPreferredNames() moved to PlayerProvisioning.js (canonical)

/**
 * Gets canonical player names as a Set for quick lookup
 * @return {Set<string>} Set of preferred_name_id values
 */
function getPreferredNameSet() {
  const players = getAllPreferredNames();
  return new Set(players.map(p => p.preferred_name_id));
}

/**
 * Adds a new player to PreferredNames (and provisions all source sheets)
 * @param {string} preferredNameId - Unique player identifier
 * @param {string} displayName - Display name (optional, defaults to preferredNameId)
 * @return {Object} Result {success, message}
 */
function registerPlayer(preferredNameId, displayName = null) {
  if (!preferredNameId || typeof preferredNameId !== 'string') {
    throwError('Invalid player ID', 'INVALID_INPUT', 'preferred_name_id is required');
  }

  preferredNameId = preferredNameId.trim();
  displayName = displayName ? displayName.trim() : preferredNameId;

  ensurePreferredNamesSchema();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PreferredNames');

  // Check if player already exists
  const existing = getPreferredNameSet();
  if (existing.has(preferredNameId)) {
    return { success: false, message: 'Player already exists' };
  }

  // Add to PreferredNames
  sheet.appendRow([
    preferredNameId,
    displayName,
    dateISO(),
    dateISO(),
    'ACTIVE'
  ]);

  // Provision all source sheets
  provisionPlayerInSourceSheets_(preferredNameId);

  logIntegrityAction('PLAYER_REGISTER', {
    preferredName: preferredNameId,
    details: `Registered new player: ${displayName}`,
    dfTags: ['DF-PROVISION'],
    status: 'SUCCESS'
  });

  return { success: true, message: `Player ${preferredNameId} registered successfully` };
}

/**
 * Provisions a player in all source sheets (Flag_Missions, Attendance_Missions, Dice_Points, BP_Total)
 * @param {string} preferredNameId - Player ID
 * @private
 */
function provisionPlayerInSourceSheets_(preferredNameId) {
  // Provision in Flag_Missions
  ensureFlagMissionsSchema();
  const flagSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Flag_Missions');
  if (flagSheet && !playerExistsInSheet_(flagSheet, preferredNameId, 'preferred_name_id')) {
    const newRow = [preferredNameId, 0]; // preferred_name_id, Flag_Points (will be computed)
    FLAG_MISSION_CATEGORIES.forEach(() => newRow.push(0));
    newRow.push(dateISO()); // LastUpdated
    flagSheet.appendRow(newRow);
  }

  // Provision in Attendance_Missions
  ensureAttendanceMissionsSchema();
  const attSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance_Missions');
  if (attSheet && !playerExistsInSheet_(attSheet, preferredNameId, 'preferred_name_id')) {
    attSheet.appendRow([
      preferredNameId,  // preferred_name_id
      0,                // Events_Attended
      0,                // Current_Streak
      0,                // Best_Streak
      0,                // Attendance_Points
      '',               // Last_Event_Date
      '',               // Milestones_Achieved (comma-separated)
      dateISO()         // LastUpdated
    ]);
  }

  // Provision in Dice_Points
  ensureDicePointsSchema();
  const diceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dice_Points');
  if (diceSheet && !playerExistsInSheet_(diceSheet, preferredNameId, 'preferred_name_id')) {
    diceSheet.appendRow([
      preferredNameId,  // preferred_name_id
      0,                // Points
      false,            // Add_Point_Checkbox
      dateISO()         // LastUpdated
    ]);
  }

  // Provision in BP_Total (uses consolidated schema)
  ensureBPTotalConsolidatedSchema();
  const bpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BP_Total');
  if (bpSheet && !playerExistsInSheet_(bpSheet, preferredNameId, 'preferred_name_id')) {
    bpSheet.appendRow([
      preferredNameId,  // preferred_name_id
      0,                // Flag_Mission_Points
      0,                // Attendance_Mission_Points
      0,                // Dice_Points
      0,                // Raw_Total
      0,                // Capped_BP
      0,                // Overflow_to_Prestige
      dateISO()         // LastUpdated
    ]);
  }
}

/**
 * Checks if a player exists in a sheet
 * @param {Sheet} sheet - The sheet to check
 * @param {string} preferredNameId - Player ID
 * @param {string} colName - Column name to check
 * @return {boolean} True if player exists
 * @private
 */
function playerExistsInSheet_(sheet, preferredNameId, colName) {
  if (!sheet || sheet.getLastRow() <= 1) return false;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf(colName);

  if (nameCol === -1) return false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredNameId) {
      return true;
    }
  }

  return false;
}

// ============================================================================
// FLAG MISSIONS
// ============================================================================

/**
 * Ensures Flag_Missions sheet exists with proper schema
 */
function ensureFlagMissionsSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Flag_Missions');

  const requiredHeaders = [
    'preferred_name_id',
    'Flag_Points'  // Computed sum of all mission columns
  ].concat(FLAG_MISSION_CATEGORIES).concat(['LastUpdated']);

  if (!sheet) {
    sheet = ss.insertSheet('Flag_Missions');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);

    // Set up ARRAYFORMULA for Flag_Points column (B)
    // Formula: sum of columns C through K for each row
    sheet.getRange('B2').setFormula(
      '=ARRAYFORMULA(IF(A2:A="","",SUMIF(ROW(A2:A),"="&ROW(A2:A),C2:C)+SUMIF(ROW(A2:A),"="&ROW(A2:A),D2:D)+SUMIF(ROW(A2:A),"="&ROW(A2:A),E2:E)+SUMIF(ROW(A2:A),"="&ROW(A2:A),F2:F)+SUMIF(ROW(A2:A),"="&ROW(A2:A),G2:G)+SUMIF(ROW(A2:A),"="&ROW(A2:A),H2:H)+SUMIF(ROW(A2:A),"="&ROW(A2:A),I2:I)+SUMIF(ROW(A2:A),"="&ROW(A2:A),J2:J)+SUMIF(ROW(A2:A),"="&ROW(A2:A),K2:K)))'
    );

    // Style header
    sheet.getRange(1, 1, 1, requiredHeaders.length)
      .setFontWeight('bold')
      .setBackground('#9c27b0')
      .setFontColor('#ffffff');
    return;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = requiredHeaders.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });
  }
}

// NOTE: awardFlagMission() moved to flagMissionService.js (canonical)

/**
 * Gets a player's flag mission points
 * @param {string} preferredNameId - Player ID
 * @return {Object|null} Player's flag missions data
 */
function getPlayerFlagMissions(preferredNameId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Flag_Missions');

  if (!sheet || sheet.getLastRow() <= 1) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('preferred_name_id');

  if (nameCol === -1) return null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredNameId) {
      const result = {};
      headers.forEach((header, idx) => {
        result[header] = data[i][idx];
      });
      return result;
    }
  }

  return null;
}

// ============================================================================
// ATTENDANCE MISSIONS
// ============================================================================

/**
 * Ensures Attendance_Missions sheet exists with proper schema
 */
function ensureAttendanceMissionsSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Attendance_Missions');

  const requiredHeaders = [
    'preferred_name_id',
    'Events_Attended',
    'Current_Streak',
    'Best_Streak',
    'Attendance_Points',
    'Last_Event_Date',
    'Milestones_Achieved',
    'LastUpdated'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('Attendance_Missions');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:H1')
      .setFontWeight('bold')
      .setBackground('#ff9800')
      .setFontColor('#ffffff');
    return;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = requiredHeaders.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });
  }
}

/**
 * Records attendance for a player and awards milestone points
 * @param {string} preferredNameId - Player ID
 * @param {string} eventDate - Event date (ISO format or Date object)
 * @return {Object} Result {eventsAttended, milestonesAwarded, pointsAwarded}
 */
function recordAttendance(preferredNameId, eventDate = null) {
  ensureAttendanceMissionsSchema();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Attendance_Missions');

  if (!sheet) {
    throwError('Attendance_Missions sheet not found', 'SHEET_MISSING');
  }

  eventDate = eventDate || dateISO();
  if (eventDate instanceof Date) {
    eventDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('preferred_name_id');
  const eventsCol = headers.indexOf('Events_Attended');
  const streakCol = headers.indexOf('Current_Streak');
  const bestStreakCol = headers.indexOf('Best_Streak');
  const pointsCol = headers.indexOf('Attendance_Points');
  const lastEventCol = headers.indexOf('Last_Event_Date');
  const milestonesCol = headers.indexOf('Milestones_Achieved');
  const updatedCol = headers.indexOf('LastUpdated');

  // Find player row
  let playerRow = -1;
  let currentData = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredNameId) {
      playerRow = i;
      currentData = {
        events: coerceNumber(data[i][eventsCol], 0),
        streak: coerceNumber(data[i][streakCol], 0),
        bestStreak: coerceNumber(data[i][bestStreakCol], 0),
        points: coerceNumber(data[i][pointsCol], 0),
        lastEvent: data[i][lastEventCol],
        milestones: data[i][milestonesCol] ? String(data[i][milestonesCol]).split(',').map(m => parseInt(m, 10)).filter(m => !isNaN(m)) : []
      };
      break;
    }
  }

  if (playerRow === -1) {
    // Auto-provision
    const canonicalNames = getPreferredNameSet();
    if (!canonicalNames.has(preferredNameId)) {
      throwError('Player not in PreferredNames registry', 'PLAYER_NOT_FOUND');
    }
    provisionPlayerInSourceSheets_(preferredNameId);
    return recordAttendance(preferredNameId, eventDate);
  }

  // Increment events attended
  const newEvents = currentData.events + 1;

  // Update streak (simplified: just increment for now)
  const newStreak = currentData.streak + 1;
  const newBestStreak = Math.max(currentData.bestStreak, newStreak);

  // Check for new milestones
  const milestonesAwarded = [];
  let pointsAwarded = 0;

  Object.keys(ATTENDANCE_MILESTONES).forEach(threshold => {
    const thresholdNum = parseInt(threshold, 10);
    if (newEvents >= thresholdNum && !currentData.milestones.includes(thresholdNum)) {
      milestonesAwarded.push(thresholdNum);
      pointsAwarded += ATTENDANCE_MILESTONES[threshold];
    }
  });

  const newPoints = currentData.points + pointsAwarded;
  const newMilestones = currentData.milestones.concat(milestonesAwarded).sort((a, b) => a - b);

  // Update sheet
  sheet.getRange(playerRow + 1, eventsCol + 1).setValue(newEvents);
  sheet.getRange(playerRow + 1, streakCol + 1).setValue(newStreak);
  sheet.getRange(playerRow + 1, bestStreakCol + 1).setValue(newBestStreak);
  sheet.getRange(playerRow + 1, pointsCol + 1).setValue(newPoints);
  sheet.getRange(playerRow + 1, lastEventCol + 1).setValue(eventDate);
  sheet.getRange(playerRow + 1, milestonesCol + 1).setValue(newMilestones.join(','));
  sheet.getRange(playerRow + 1, updatedCol + 1).setValue(dateISO());

  if (milestonesAwarded.length > 0) {
    logIntegrityAction('ATTENDANCE_MILESTONE', {
      preferredName: preferredNameId,
      details: `Events: ${newEvents}. Milestones: ${milestonesAwarded.join(', ')}. Points: +${pointsAwarded}`,
      dfTags: ['DF-MILESTONE'],
      status: 'SUCCESS'
    });
  }

  return {
    eventsAttended: newEvents,
    milestonesAwarded,
    pointsAwarded,
    totalPoints: newPoints
  };
}

/**
 * Gets a player's attendance data
 * @param {string} preferredNameId - Player ID
 * @return {Object|null} Player's attendance data
 */
function getPlayerAttendance(preferredNameId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Attendance_Missions');

  if (!sheet || sheet.getLastRow() <= 1) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('preferred_name_id');

  if (nameCol === -1) return null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredNameId) {
      const result = {};
      headers.forEach((header, idx) => {
        result[header] = data[i][idx];
      });
      return result;
    }
  }

  return null;
}

// ============================================================================
// DICE POINTS
// ============================================================================

/**
 * Ensures Dice_Points sheet exists with proper schema
 */
function ensureDicePointsSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Dice_Points');

  const requiredHeaders = [
    'preferred_name_id',
    'Points',
    'Add_Point_Checkbox',
    'LastUpdated'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('Dice_Points');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:D1')
      .setFontWeight('bold')
      .setBackground('#4caf50')
      .setFontColor('#ffffff');

    // Set checkbox data validation for column C
    // (Applied per-row when players are added)
    return;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = requiredHeaders.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });
  }
}

// NOTE: awardDicePoints() moved to dicePointsBackEndService.js (canonical)

/**
 * Gets a player's dice points
 * @param {string} preferredNameId - Player ID
 * @return {number} Current dice points (0 if not found)
 */
function getPlayerDicePoints(preferredNameId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Dice_Points');

  if (!sheet || sheet.getLastRow() <= 1) return 0;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('preferred_name_id');
  const pointsCol = headers.indexOf('Points');

  if (nameCol === -1 || pointsCol === -1) return 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredNameId) {
      return coerceNumber(data[i][pointsCol], 0);
    }
  }

  return 0;
}

/**
 * Handles checkbox click for dice point increment (for onEdit trigger)
 * @param {Object} e - Edit event object
 */
function onDicePointCheckboxEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Dice_Points') return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const checkboxCol = headers.indexOf('Add_Point_Checkbox') + 1;

  if (e.range.getColumn() !== checkboxCol) return;
  if (e.value !== 'TRUE' && e.value !== true) return;

  const row = e.range.getRow();
  if (row <= 1) return; // Header row

  const nameCol = headers.indexOf('preferred_name_id') + 1;
  const pointsCol = headers.indexOf('Points') + 1;
  const updatedCol = headers.indexOf('LastUpdated') + 1;

  const preferredNameId = sheet.getRange(row, nameCol).getValue();
  const currentPoints = coerceNumber(sheet.getRange(row, pointsCol).getValue(), 0);

  // Increment points
  sheet.getRange(row, pointsCol).setValue(currentPoints + 1);
  if (updatedCol > 0) {
    sheet.getRange(row, updatedCol).setValue(dateISO());
  }

  // Reset checkbox
  e.range.setValue(false);

  logIntegrityAction('DICE_POINTS_CHECKBOX', {
    preferredName: preferredNameId,
    details: `Quick-add: ${currentPoints} → ${currentPoints + 1}`,
    status: 'SUCCESS'
  });
}

// ============================================================================
// BP_TOTAL CONSOLIDATION
// NOTE: ensureBPTotalConsolidatedSchema() moved to bpTotalPipeline.js (use ensureBPTotalSchemaEnhanced_)
// NOTE: migrateBPTotalSchema_() moved to bpTotalPipeline.js
// ============================================================================

// NOTE: syncBPTotals() moved to bpTotalPipeline.js (use updateBPTotalFromSources)

/**
 * Builds a lookup map from a source sheet
 * @param {Sheet} sheet - Source sheet
 * @param {string} keyCol - Key column name
 * @param {string} valueCol - Value column name
 * @return {Map} Map of key -> value
 * @private
 */
function buildSourceMap_(sheet, keyCol, valueCol) {
  const map = new Map();

  if (!sheet || sheet.getLastRow() <= 1) return map;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyIdx = headers.indexOf(keyCol);
  const valueIdx = headers.indexOf(valueCol);

  if (keyIdx === -1 || valueIdx === -1) return map;

  for (let i = 1; i < data.length; i++) {
    const key = data[i][keyIdx];
    const value = data[i][valueIdx];
    if (key) {
      map.set(key, coerceNumber(value, 0));
    }
  }

  return map;
}

/**
 * Routes overflow points to Prestige_Overflow sheet
 * @param {string} preferredNameId - Player ID
 * @param {number} overflow - Overflow amount
 * @private
 */
function routeOverflowToPrestige_(preferredNameId, overflow) {
  // This uses the existing addPrestigeOverflow_ from bpService.gs
  // But we need to track it separately for the sync
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Prestige_Overflow');

  if (!sheet) {
    sheet = ss.insertSheet('Prestige_Overflow');
    sheet.appendRow(['PreferredName', 'Total_Overflow', 'Last_Updated', 'Prestige_Tier']);
    sheet.setFrozenRows(1);
  }

  const data = sheet.getDataRange().getValues();
  let playerRow = -1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === preferredNameId) {
      playerRow = i;
      break;
    }
  }

  if (playerRow === -1) {
    sheet.appendRow([preferredNameId, overflow, dateISO(), computePrestigeTier_(overflow)]);
  } else {
    // Update existing - set to current overflow (not cumulative from sync)
    sheet.getRange(playerRow + 1, 2).setValue(overflow);
    sheet.getRange(playerRow + 1, 3).setValue(dateISO());
    sheet.getRange(playerRow + 1, 4).setValue(computePrestigeTier_(overflow));
  }
}

/**
 * Computes prestige tier based on overflow amount
 * @param {number} overflow - Total overflow
 * @return {string} Tier name
 * @private
 */
function computePrestigeTier_(overflow) {
  if (overflow >= 500) return 'Diamond';
  if (overflow >= 250) return 'Platinum';
  if (overflow >= 100) return 'Gold';
  if (overflow >= 50) return 'Silver';
  return 'Bronze';
}

// ============================================================================
// VALIDATION & INTEGRITY
// ============================================================================

/**
 * Validates all mission point data for integrity
 * @return {Object} Validation result {pass, issues}
 */
function validateMissionPointsIntegrity() {
  const issues = [];

  // Get canonical names
  const canonicalNames = getPreferredNameSet();

  // Check Flag_Missions
  const flagSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Flag_Missions');
  if (flagSheet && flagSheet.getLastRow() > 1) {
    const data = flagSheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = headers.indexOf('preferred_name_id');

    for (let i = 1; i < data.length; i++) {
      const playerId = data[i][nameCol];
      if (playerId && !canonicalNames.has(playerId)) {
        issues.push({
          sheet: 'Flag_Missions',
          row: i + 1,
          player: playerId,
          issue: 'Not in PreferredNames registry'
        });
      }

      // Check for negative values
      FLAG_MISSION_CATEGORIES.forEach(cat => {
        const catCol = headers.indexOf(cat);
        if (catCol !== -1 && coerceNumber(data[i][catCol], 0) < 0) {
          issues.push({
            sheet: 'Flag_Missions',
            row: i + 1,
            player: playerId,
            issue: `Negative value in ${cat}`
          });
        }
      });
    }
  }

  // Check Attendance_Missions
  const attSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance_Missions');
  if (attSheet && attSheet.getLastRow() > 1) {
    const data = attSheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = headers.indexOf('preferred_name_id');
    const pointsCol = headers.indexOf('Attendance_Points');

    for (let i = 1; i < data.length; i++) {
      const playerId = data[i][nameCol];
      if (playerId && !canonicalNames.has(playerId)) {
        issues.push({
          sheet: 'Attendance_Missions',
          row: i + 1,
          player: playerId,
          issue: 'Not in PreferredNames registry'
        });
      }

      if (pointsCol !== -1 && coerceNumber(data[i][pointsCol], 0) < 0) {
        issues.push({
          sheet: 'Attendance_Missions',
          row: i + 1,
          player: playerId,
          issue: 'Negative attendance points'
        });
      }
    }
  }

  // Check Dice_Points
  const diceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dice_Points');
  if (diceSheet && diceSheet.getLastRow() > 1) {
    const data = diceSheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = headers.indexOf('preferred_name_id');
    const pointsCol = headers.indexOf('Points');

    for (let i = 1; i < data.length; i++) {
      const playerId = data[i][nameCol];
      if (playerId && !canonicalNames.has(playerId)) {
        issues.push({
          sheet: 'Dice_Points',
          row: i + 1,
          player: playerId,
          issue: 'Not in PreferredNames registry'
        });
      }

      if (pointsCol !== -1 && coerceNumber(data[i][pointsCol], 0) < 0) {
        issues.push({
          sheet: 'Dice_Points',
          row: i + 1,
          player: playerId,
          issue: 'Negative dice points'
        });
      }
    }
  }

  return {
    pass: issues.length === 0,
    issues
  };
}

// NOTE: provisionAllPlayers() moved to PlayerProvisioning.js (canonical)

// NOTE: getCanonicalNames() - multiple definitions exist across files
// TODO: Consolidate to utils.js as canonical location

// ============================================================================
// PLAYER LOOKUP - Unified Player Data
// ============================================================================

/**
 * Gets comprehensive player data from all sources
 * @param {string} preferredNameId - Player ID
 * @return {Object} Complete player profile
 */
function getPlayerProfile(preferredNameId) {
  const profile = {
    id: preferredNameId,
    registered: false,
    flag_missions: null,
    attendance: null,
    dice_points: 0,
    bp_total: null,
    keys: null
  };

  // Check PreferredNames
  const canonicalNames = getPreferredNameSet();
  profile.registered = canonicalNames.has(preferredNameId);

  // Get Flag Missions
  profile.flag_missions = getPlayerFlagMissions(preferredNameId);

  // Get Attendance
  profile.attendance = getPlayerAttendance(preferredNameId);

  // Get Dice Points
  profile.dice_points = getPlayerDicePoints(preferredNameId);

  // Get BP Total
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bpSheet = ss.getSheetByName('BP_Total');
  if (bpSheet && bpSheet.getLastRow() > 1) {
    const data = bpSheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = headers.indexOf('preferred_name_id');

    if (nameCol !== -1) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][nameCol] === preferredNameId) {
          profile.bp_total = {};
          headers.forEach((h, idx) => {
            profile.bp_total[h] = data[i][idx];
          });
          break;
        }
      }
    }
  }

  // Get Keys
  profile.keys = getPlayerKeys(preferredNameId);

  return profile;
}