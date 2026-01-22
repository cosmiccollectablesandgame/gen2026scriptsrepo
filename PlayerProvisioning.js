/**
 * ════════════════════════════════════════════════════════════════════════════
 * PLAYER PROVISIONING SYSTEM
 * ════════════════════════════════════════════════════════════════════════════
 *
 * @fileoverview Ensures all players in PreferredNames exist in all tracking sheets
 * 
 * This system:
 *   - Maintains PreferredNames as the canonical player registry
 *   - Provisions new players to all tracking sheets automatically
 *   - Discovers unprovisioned players from event tabs
 *   - Provides batch provisioning for maintenance
 *
 * Compatible with: Engine v8.0.0+, attendanceMissionScanner.gs
 * Version: 1.0.0
 * ════════════════════════════════════════════════════════════════════════════
 */

/**
 * Sheet configurations for provisioning
 * Each entry defines: sheetName, keyColumn, and default values for new rows
 */
const PROVISION_TARGETS = {
  BP_Total: {
    sheetName: 'BP_Total',
    keyColumn: 'preferred_name_id',
    defaults: {
      'BP_Historical': 0,
      'BP_Redeemed': 0,
      'BP_Current': 0,
      'Flag_Points': 0,
      'Attendance_Points': 0,
      'Dice_Points': 0,
      'LastUpdated': function() { return new Date().toISOString(); }
    }
  },
  Attendance_Missions: {
    sheetName: 'Attendance_Missions',
    keyColumn: 'PreferredName',
    defaults: {
      'First Contact': 0,
      'Stellar Explorer': 0,
      'Deck Diver': 0,
      'Lunar Loyalty': 0,
      'Meteor Shower': 0,
      'Sealed Voyager': 0,
      'Draft Navigator': 0,
      'Stellar Scholar': 0,
      'Casual Commander Events': 0,
      'Transitional Commander Events': 0,
      'cEDH Events': 0,
      'Limited Events': 0,
      'Academy Events': 0,
      'Outreach Events': 0,
      'Free Play Events': 0,
      'Interstellar Strategist': 0,
      'Black Hole Survivor': 0,
      'Points': 0
    }
  },
  Flag_Missions: {
    sheetName: 'Flag_Missions',
    keyColumn: 'preferred_name_id',
    defaults: {
      'Cosmic_Selfie': false,
      'Review_Writer': false,
      'Social_Media_Star': false,
      'App_Explorer': false,
      'Cosmic_Merchant': false,
      'Flag_Points': 0,
      'LastUpdated': function() { return new Date().toISOString(); }
    }
  },
  Dice_Points: {
    sheetName: 'Dice_Points',
    keyColumn: 'preferred_name_id',
    defaults: {
      'Points': 0,
      'LastUpdated': function() { return new Date().toISOString(); }
    }
  },
  Key_Tracker: {
    sheetName: 'Key_Tracker',
    keyColumn: 'PreferredName',
    defaults: {
      'White': 0,
      'Blue': 0,
      'Black': 0,
      'Red': 0,
      'Green': 0,
      'Total': 0,
      'Rainbow': false
    }
  }
};

/**
 * Normalizes player name for consistent comparison
 * @param {string} name - Raw player name
 * @return {string} Normalized name
 */
function normalizePlayerName(name) {
  return String(name || '').trim().replace(/\s+/g, ' ');
}

/**
 * Case-insensitive player name comparison
 * @param {string} a - First name
 * @param {string} b - Second name
 * @return {boolean} True if names match
 */
function playerNamesMatch(a, b) {
  return normalizePlayerName(a).toLowerCase() === normalizePlayerName(b).toLowerCase();
}

/**
 * Gets all canonical player names from PreferredNames sheet
 * @return {string[]} Array of preferred_name_id values
 */
function getAllPreferredNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetCI_('PreferredNames');
  
  if (!sheet) {
    return [];
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  
  // Get header map to find preferred_name_id column
  const headerMap = getHeaderMap_(sheet);
  
  // Try to find the column with synonyms
  const possibleHeaders = ['preferrednameid', 'preferredname', 'name', 'player'];
  let colIndex = -1;
  
  for (const header of possibleHeaders) {
    if (headerMap[header] !== undefined) {
      colIndex = headerMap[header];
      break;
    }
  }
  
  // Fallback to column A if not found
  if (colIndex === -1) {
    colIndex = 0;
  }
  
  // Read all data starting from row 2
  const data = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  
  // Return array of non-empty, normalized names
  return data
    .map(row => normalizePlayerName(row[0]))
    .filter(name => name.length > 0);
}

/**
 * Checks if a player exists in PreferredNames
 * @param {string} name - Player name to check
 * @return {boolean}
 */
function playerExists(name) {
  const allNames = getAllPreferredNames();
  const normalizedName = normalizePlayerName(name);
  
  for (const existingName of allNames) {
    if (playerNamesMatch(normalizedName, existingName)) {
      return true;
    }
  }
  
  return false;
}

/**
 * Gets sheet by name (case-insensitive)
 * @param {string} name - Sheet name
 * @return {Sheet|null}
 */
function getSheetCI_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const lowerName = name.toLowerCase();
  for (const sheet of sheets) {
    if (sheet.getName().toLowerCase() === lowerName) {
      return sheet;
    }
  }
  return null;
}

/**
 * Gets header map for a sheet (column name -> index)
 * @param {Sheet} sheet
 * @return {Object} Map of lowercase header names to column indices
 */
function getHeaderMap_(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return {};
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i]).toLowerCase().replace(/[_\s-]/g, '');
    map[h] = i;
    map[String(headers[i])] = i; // Also store original
  }
  return map;
}

/**
 * Provisions a single player to a specific sheet if not already present
 * @param {string} sheetName - Target sheet name
 * @param {string} playerName - Player's preferred_name_id
 * @param {string} keyColumn - Column name containing player identifier
 * @param {Object} defaults - Default column values (can include functions)
 * @return {Object} { created: boolean, existed: boolean, error: string|null }
 */
function provisionToSheet_(sheetName, playerName, keyColumn, defaults) {
  try {
    // Get sheet (case-insensitive)
    const sheet = getSheetCI_(sheetName);
    
    if (!sheet) {
      return { created: false, existed: false, error: 'Sheet not found' };
    }
    
    // Get header map
    const headerMap = getHeaderMap_(sheet);
    
    // Find key column index
    const keyColNormalized = keyColumn.toLowerCase().replace(/[_\s-]/g, '');
    let keyColIndex = headerMap[keyColNormalized];
    
    // If not found by normalized name, try original
    if (keyColIndex === undefined) {
      keyColIndex = headerMap[keyColumn];
    }
    
    if (keyColIndex === undefined) {
      return { created: false, existed: false, error: 'Key column not found' };
    }
    
    // Check if player already exists (case-insensitive search)
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const existingData = sheet.getRange(2, keyColIndex + 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < existingData.length; i++) {
        const existingName = String(existingData[i][0] || '').trim();
        if (playerNamesMatch(existingName, playerName)) {
          return { created: false, existed: true, error: null };
        }
      }
    }
    
    // Build new row array with player name in key column and defaults in other columns
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = new Array(headers.length).fill('');
    
    // Set player name in key column
    newRow[keyColIndex] = playerName;
    
    // Set defaults for other columns
    for (const colName in defaults) {
      const colNormalized = colName.toLowerCase().replace(/[_\s-]/g, '');
      let colIndex = headerMap[colNormalized];
      
      // Try original name if normalized not found
      if (colIndex === undefined) {
        colIndex = headerMap[colName];
      }
      
      if (colIndex !== undefined) {
        let value = defaults[colName];
        // For defaults that are functions, call them to get value
        if (typeof value === 'function') {
          value = value();
        }
        newRow[colIndex] = value;
      }
    }
    
    // Append row
    sheet.appendRow(newRow);
    
    return { created: true, existed: false, error: null };
  } catch (e) {
    return { created: false, existed: false, error: e.message };
  }
}

/**
 * Provisions a player to ALL tracking sheets
 * @param {string} playerName - Player's preferred_name_id
 * @return {Object} { createdBySheet: {}, existedBySheet: {}, skippedSheets: [] }
 */
function provisionPlayerProfile(playerName) {
  const createdBySheet = {};
  const existedBySheet = {};
  const skippedSheets = [];
  
  // Loop through PROVISION_TARGETS
  for (const targetKey in PROVISION_TARGETS) {
    const target = PROVISION_TARGETS[targetKey];
    const result = provisionToSheet_(
      target.sheetName,
      playerName,
      target.keyColumn,
      target.defaults
    );
    
    if (result.error) {
      skippedSheets.push(target.sheetName);
    } else {
      createdBySheet[target.sheetName] = result.created;
      existedBySheet[target.sheetName] = result.existed;
    }
  }
  
  return {
    createdBySheet: createdBySheet,
    existedBySheet: existedBySheet,
    skippedSheets: skippedSheets
  };
}

/**
 * Adds a new player to PreferredNames and provisions to all tracking sheets
 * @param {string} name - Player's preferred_name_id
 * @return {Object} { success: boolean, alreadyExisted: boolean, profileResult: Object }
 */
function addNewPlayer(name) {
  try {
    // Normalize name
    const normalizedName = normalizePlayerName(name);
    
    if (!normalizedName) {
      return { success: false, alreadyExisted: false, profileResult: null, error: 'Invalid name' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getSheetCI_('PreferredNames');
    
    if (!sheet) {
      return { success: false, alreadyExisted: false, profileResult: null, error: 'PreferredNames sheet not found' };
    }
    
    // Check if already exists in PreferredNames
    let alreadyExisted = false;
    let canonicalName = normalizedName;
    
    const existingNames = getAllPreferredNames();
    for (const existing of existingNames) {
      if (playerNamesMatch(existing, normalizedName)) {
        alreadyExisted = true;
        canonicalName = existing; // Use existing casing
        break;
      }
    }
    
    // If not exists, add to PreferredNames
    if (!alreadyExisted) {
      const headerMap = getHeaderMap_(sheet);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      const newRow = new Array(headers.length).fill('');
      
      // Set preferred_name_id
      const nameColIndex = headerMap['preferrednameid'] !== undefined ? headerMap['preferrednameid'] : 0;
      newRow[nameColIndex] = normalizedName;
      
      // Set display_name if column exists
      const displayNameIndex = headerMap['displayname'];
      if (displayNameIndex !== undefined) {
        newRow[displayNameIndex] = normalizedName;
      }
      
      // Set created_at if column exists
      const createdAtIndex = headerMap['createdat'];
      if (createdAtIndex !== undefined) {
        newRow[createdAtIndex] = new Date().toISOString();
      }
      
      // Set last_active if column exists
      const lastActiveIndex = headerMap['lastactive'];
      if (lastActiveIndex !== undefined) {
        newRow[lastActiveIndex] = new Date().toISOString();
      }
      
      // Set status if column exists
      const statusIndex = headerMap['status'];
      if (statusIndex !== undefined) {
        newRow[statusIndex] = 'ACTIVE';
      }
      
      sheet.appendRow(newRow);
    }
    
    // Call provisionPlayerProfile()
    const profileResult = provisionPlayerProfile(canonicalName);
    
    // Log to Integrity_Log if available
    logProvisioningAction_('ADD_NEW_PLAYER', {
      playerName: canonicalName,
      alreadyExisted: alreadyExisted,
      profileResult: profileResult
    });
    
    return {
      success: true,
      alreadyExisted: alreadyExisted,
      profileResult: profileResult
    };
  } catch (e) {
    return {
      success: false,
      alreadyExisted: false,
      profileResult: null,
      error: e.message
    };
  }
}

/**
 * Batch provisions multiple players
 * @param {string[]} names - Array of player names
 * @return {Object} { totalCreated: number, totalExisted: number, errors: string[] }
 */
function provisionMultiplePlayers(names) {
  let totalCreated = 0;
  let totalExisted = 0;
  const errors = [];
  
  for (const name of names) {
    try {
      const result = addNewPlayer(name);
      
      if (result.success && result.profileResult) {
        for (const sheet in result.profileResult.createdBySheet) {
          if (result.profileResult.createdBySheet[sheet]) {
            totalCreated++;
          }
        }
        for (const sheet in result.profileResult.existedBySheet) {
          if (result.profileResult.existedBySheet[sheet]) {
            totalExisted++;
          }
        }
      } else if (result.error) {
        errors.push(name + ': ' + result.error);
      }
    } catch (e) {
      errors.push(name + ': ' + e.message);
    }
  }
  
  return {
    totalCreated: totalCreated,
    totalExisted: totalExisted,
    errors: errors
  };
}

/**
 * Gets list of players in event tabs but NOT in PreferredNames
 * Uses same event detection pattern as attendanceMissionScanner.gs
 * @return {string[]} Array of unprovisioned player names
 */
function getUnprovisionedPlayers() {
  // Get all preferred names (lowercase set for fast lookup)
  const preferredNames = getAllPreferredNames();
  const preferredNamesSet = new Set(preferredNames.map(n => n.toLowerCase()));
  
  // Get all event sheets (pattern: /^\d{2}-\d{2}[A-Za-z]+-\d{4}$/)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const EVENT_TAB_REGEX = /^\d{2}-\d{2}[A-Za-z]+-\d{4}$/;
  const eventSheets = allSheets.filter(sheet => EVENT_TAB_REGEX.test(sheet.getName()));
  
  const unprovisionedSet = new Set();
  
  // For each event sheet, read Column B (preferred_name_id)
  for (const sheet of eventSheets) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    
    // Column B = preferred_name_id
    const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    
    for (let i = 0; i < data.length; i++) {
      const name = normalizePlayerName(data[i][0]);
      if (name && !preferredNamesSet.has(name.toLowerCase())) {
        unprovisionedSet.add(name);
      }
    }
  }
  
  // Deduplicate and return
  return Array.from(unprovisionedSet);
}

/**
 * Scans event tabs for unprovisioned players and provisions them
 * @return {Object} { newPlayersFound: number, provisioned: number, errors: string[] }
 */
function discoverAndProvisionNewPlayers() {
  try {
    // Call getUnprovisionedPlayers()
    const unprovisionedPlayers = getUnprovisionedPlayers();
    
    // If none found, return early with zeros
    if (unprovisionedPlayers.length === 0) {
      return {
        newPlayersFound: 0,
        provisioned: 0,
        errors: []
      };
    }
    
    // Call provisionMultiplePlayers()
    const result = provisionMultiplePlayers(unprovisionedPlayers);
    
    // Log results to Integrity_Log if available
    logProvisioningAction_('DISCOVER_AND_PROVISION', {
      newPlayersFound: unprovisionedPlayers.length,
      provisioned: unprovisionedPlayers.length,
      totalCreated: result.totalCreated,
      totalExisted: result.totalExisted,
      errors: result.errors
    });
    
    return {
      newPlayersFound: unprovisionedPlayers.length,
      provisioned: unprovisionedPlayers.length,
      errors: result.errors
    };
  } catch (e) {
    return {
      newPlayersFound: 0,
      provisioned: 0,
      errors: [e.message]
    };
  }
}

/**
 * Menu handler for "Provision All Players"
 * Ensures all PreferredNames players exist in all tracking sheets
 */
function runFullProvisioning() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  ss.toast('Provisioning all players...', '⏳ Please wait', -1);
  
  try {
    const allPlayers = getAllPreferredNames();
    
    if (allPlayers.length === 0) {
      ui.alert('No Players', 'No players found in PreferredNames sheet.', ui.ButtonSet.OK);
      return;
    }
    
    let totalCreated = 0;
    let totalExisted = 0;
    
    for (const player of allPlayers) {
      const result = provisionPlayerProfile(player);
      for (const sheet in result.createdBySheet) {
        if (result.createdBySheet[sheet]) totalCreated++;
      }
      for (const sheet in result.existedBySheet) {
        if (result.existedBySheet[sheet]) totalExisted++;
      }
    }
    
    ss.toast('Complete!', '✅ Provisioning', 3);
    ui.alert('✅ Provisioning Complete',
      'Players checked: ' + allPlayers.length + '\n' +
      'New entries created: ' + totalCreated + '\n' +
      'Already existed: ' + totalExisted,
      ui.ButtonSet.OK);
      
  } catch (e) {
    ss.toast('Error: ' + e.message, '❌ Failed', 5);
    Logger.log('Provisioning error: ' + e.message + '\n' + e.stack);
  }
}

/**
 * Logs provisioning action to Integrity_Log if available
 * @param {string} action - Action name
 * @param {Object} details - Details object
 */
function logProvisioningAction_(action, details) {
  if (typeof logIntegrityAction === 'function') {
    logIntegrityAction(action, details);
  } else {
    Logger.log('[' + action + '] ' + JSON.stringify(details));
  }
}
