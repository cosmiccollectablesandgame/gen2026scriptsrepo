/**
 * ══════════════════════════════════════════════════════════════════════════
 * PHASE 5 SERVICE - Transactional Truth Consumers
 * ══════════════════════════════════════════════════════════════════════════
 * 
 * @fileoverview Extends PreferredNames identity hygiene into transactional systems
 * 
 * Core Functions:
 *   - normalizePlayerName(): Unified normalization for all identity operations
 *   - scanStoreCreditLedgerForUnknownNames(): Queue unknown ledger names
 *   - scanPreordersForUnknownNames(): Queue unknown preorder names
 *   - createNewPlayerFromUndiscovered(): Retail onboarding workflow
 *   - refreshResolveUnknownsDashboard(): Staff dashboard view
 *   - safeRunPlayerLookupBuild(): Enforced PlayerLookup build
 * 
 * Phase 5 Design Principle:
 *   PreferredNames is the registrar.
 *   Transactions must name real people.
 *   PlayerLookup is the unified read model.
 * 
 * Version: 1.0.0
 * Compatible with: Phases 1-4 Identity Hygiene System
 * ══════════════════════════════════════════════════════════════════════════
 */

// ============================================================================
// CONFIGURATION - Truth Consumer Enforcement Modes
// ============================================================================

/**
 * Configuration for how different truth consumers handle unknown names
 * @constant
 */
const TRUTH_CONSUMERS = {
  storeCreditLedger: {
    allowUnknownNamesOnWrite: true,        // Retail-first: allow new customers
    mustQueueUnknowns: true,               // But queue them for onboarding
    blocksPlayerLookupRollupsWhenDirty: false,  // Don't block rollups
    sheetNames: ['Store_Credit_Ledger'],
    scanWindowRows: 1000                   // Recent rows to scan
  },
  preorders: {
    allowUnknownNamesOnWrite: false,       // Preorders must use canonical names
    mustQueueUnknowns: true,
    blocksPlayerLookupRollupsWhenDirty: true,   // Block rollups if dirty
    sheetNames: ['Preorders', 'PreOrders', 'Preorder_Requests', 'Orders_Preorder', 'Preorders_Sold'],
    scanWindowRows: 1000
  }
};

/**
 * Legacy sheets to exclude from provisioning and audits
 * @constant
 */
const RETIRED_SHEETS = [
  'Players_Prize-Wall-Points',
  'Player\'s Prize-Wall-Points'
];

// ============================================================================
// CORE NORMALIZATION FUNCTION
// ============================================================================

/**
 * Normalizes a player name to canonical form for identity matching
 * 
 * Rules:
 *   - Trim whitespace
 *   - Collapse internal whitespace to single spaces
 *   - Remove leading/trailing punctuation (safe subset: . , - _ ')
 *   - Preserve internal punctuation (e.g., O'Connor, Mary-Jane)
 *   - Case-insensitive comparison (lowercase for comparison only)
 * 
 * @param {string} name - Raw player name
 * @return {string} Normalized name (preserves original case for display)
 */
function normalizePlayerName(name) {
  if (!name || typeof name !== 'string') {
    return '';
  }
  
  // Step 1: Trim
  let normalized = name.trim();
  
  // Step 2: Collapse internal whitespace
  normalized = normalized.replace(/\s+/g, ' ');
  
  // Step 3: Remove leading/trailing punctuation (safe subset)
  // But preserve internal punctuation
  normalized = normalized.replace(/^[\.\,\-\_\']+|[\.\,\-\_\']+$/g, '');
  
  return normalized;
}

/**
 * Case-insensitive comparison of two normalized names
 * @param {string} name1 - First name
 * @param {string} name2 - Second name
 * @return {boolean} True if names match (case-insensitive)
 */
function namesMatch(name1, name2) {
  const n1 = normalizePlayerName(name1).toLowerCase();
  const n2 = normalizePlayerName(name2).toLowerCase();
  return n1 === n2 && n1.length > 0;
}

// ============================================================================
// CANONICAL NAMES LOADER
// ============================================================================

/**
 * Loads canonical player names from PreferredNames sheet
 * @return {Object} { canonicalSet: Set<string>, canonicalList: Array<string>, canonicalMap: Map<string, string> }
 * @private
 */
function loadCanonicalNames_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const canonicalSet = new Set();
  const canonicalList = [];
  const canonicalMap = new Map(); // lowercase -> original case
  
  // Try PreferredNames sheet
  const prefSheet = ss.getSheetByName('PreferredNames');
  if (prefSheet && prefSheet.getLastRow() > 1) {
    const data = prefSheet.getRange(2, 1, prefSheet.getLastRow() - 1, 1).getValues();
    data.forEach(row => {
      const rawName = String(row[0] || '').trim();
      if (rawName) {
        const normalized = normalizePlayerName(rawName);
        const normalizedLower = normalized.toLowerCase();
        
        if (!canonicalSet.has(normalizedLower)) {
          canonicalSet.add(normalizedLower);
          canonicalList.push(normalized);
          canonicalMap.set(normalizedLower, normalized);
        }
      }
    });
  }
  
  return {
    canonicalSet: canonicalSet,
    canonicalList: canonicalList.sort(),
    canonicalMap: canonicalMap
  };
}

/**
 * Checks if a name exists in PreferredNames (canonical registry)
 * @param {string} name - Name to check
 * @return {boolean} True if name is canonical
 */
function isCanonicalName(name) {
  const normalized = normalizePlayerName(name).toLowerCase();
  if (!normalized) return false;
  
  const { canonicalSet } = loadCanonicalNames_();
  return canonicalSet.has(normalized);
}

/**
 * Gets the canonical spelling for a name (returns exact case from PreferredNames)
 * @param {string} name - Name to look up
 * @return {string|null} Canonical name or null if not found
 */
function getCanonicalName(name) {
  const normalized = normalizePlayerName(name).toLowerCase();
  if (!normalized) return null;
  
  const { canonicalMap } = loadCanonicalNames_();
  return canonicalMap.get(normalized) || null;
}

// ============================================================================
// UNDISCOVERED NAMES QUEUE MANAGEMENT
// ============================================================================

/**
 * Queues an unknown name to UndiscoveredNames sheet
 * 
 * @param {string} rawName - The unknown name to queue
 * @param {string} sourceType - Source type (EVENT_SCAN, STORE_CREDIT_LEDGER, PREORDERS)
 * @param {string} sourceSheet - Source sheet name
 * @param {string|number} [sourceRef] - Optional source reference (row number, ID, etc.)
 * @return {boolean} True if queued (or already exists), false on error
 */
function queueUnknownName(rawName, sourceType, sourceSheet, sourceRef) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const normalized = normalizePlayerName(rawName);
    
    if (!normalized) {
      return false; // Empty name, skip
    }
    
    // Check if already canonical
    if (isCanonicalName(normalized)) {
      return true; // Already canonical, nothing to queue
    }
    
    // Get or create UndiscoveredNames sheet
    let undiscoveredSheet = ss.getSheetByName('UndiscoveredNames');
    if (!undiscoveredSheet) {
      undiscoveredSheet = ss.insertSheet('UndiscoveredNames');
      initUndiscoveredNamesHeaders_(undiscoveredSheet);
    }
    
    // Check if already in undiscovered queue
    const data = undiscoveredSheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0];
      const nameColIdx = headers.indexOf('NormalizedName') !== -1 ? headers.indexOf('NormalizedName') : headers.indexOf('Potential_Name');
      const statusColIdx = headers.indexOf('Status');
      
      for (let i = 1; i < data.length; i++) {
        const existingName = normalizePlayerName(data[i][nameColIdx]).toLowerCase();
        const status = String(data[i][statusColIdx] || 'OPEN').toUpperCase();
        
        if (existingName === normalized.toLowerCase()) {
          if (status === 'RESOLVED') {
            return true; // Already resolved
          }
          // Already queued and OPEN, update counts
          updateUnknownNameSeenCount_(undiscoveredSheet, i + 1, sourceSheet);
          return true;
        }
      }
    }
    
    // Add new unknown name
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
    
    undiscoveredSheet.appendRow([
      normalized,           // NormalizedName
      rawName,              // RawName
      sourceType,           // SourceType
      sourceSheet,          // SourceSheets
      sourceRef || '',      // SourceRefs
      1,                    // SeenCount
      timestamp,            // FirstSeen
      timestamp,            // LastSeen
      'OPEN',               // Status
      '',                   // ResolvedAs
      '',                   // ResolutionType
      '',                   // ResolvedBy
      ''                    // ResolvedAt
    ]);
    
    return true;
    
  } catch (e) {
    console.error('queueUnknownName error:', e);
    return false;
  }
}

/**
 * Initializes UndiscoveredNames sheet with Phase 5 headers
 * @param {Sheet} sheet - The UndiscoveredNames sheet
 * @private
 */
function initUndiscoveredNamesHeaders_(sheet) {
  const headers = [
    'NormalizedName',
    'RawName',
    'SourceType',
    'SourceSheets',
    'SourceRefs',
    'SeenCount',
    'FirstSeen',
    'LastSeen',
    'Status',
    'ResolvedAs',
    'ResolutionType',
    'ResolvedBy',
    'ResolvedAt'
  ];
  
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#f3f3f3');
  sheet.setFrozenRows(1);
}

/**
 * Updates the seen count for an existing unknown name
 * @param {Sheet} sheet - UndiscoveredNames sheet
 * @param {number} rowIndex - Row index (1-based)
 * @param {string} sourceSheet - Source sheet to add to refs
 * @private
 */
function updateUnknownNameSeenCount_(sheet, rowIndex, sourceSheet) {
  try {
    const data = sheet.getRange(rowIndex, 1, 1, 13).getValues()[0];
    const headers = sheet.getRange(1, 1, 1, 13).getValues()[0];
    
    const seenCountIdx = headers.indexOf('SeenCount');
    const lastSeenIdx = headers.indexOf('LastSeen');
    const sourceSheetsIdx = headers.indexOf('SourceSheets');
    
    // Update seen count
    const newCount = (parseInt(data[seenCountIdx]) || 0) + 1;
    sheet.getRange(rowIndex, seenCountIdx + 1).setValue(newCount);
    
    // Update last seen timestamp
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
    sheet.getRange(rowIndex, lastSeenIdx + 1).setValue(timestamp);
    
    // Add source sheet if not already listed
    const existingSources = String(data[sourceSheetsIdx] || '');
    if (!existingSources.includes(sourceSheet)) {
      const newSources = existingSources ? `${existingSources}, ${sourceSheet}` : sourceSheet;
      sheet.getRange(rowIndex, sourceSheetsIdx + 1).setValue(newSources);
    }
    
  } catch (e) {
    console.error('updateUnknownNameSeenCount_ error:', e);
  }
}

// ============================================================================
// STORE CREDIT LEDGER SCANNING
// ============================================================================

/**
 * Scans Store_Credit_Ledger for unknown player names
 * Queues unknowns to UndiscoveredNames
 * 
 * @param {number} [scanWindowRows] - Number of recent rows to scan (default: 1000)
 * @return {Object} { scannedRows: number, unknownNames: Array<string>, queuedCount: number }
 */
function SCAN_STORE_CREDIT_LEDGER_FOR_UNKNOWN_NAMES(scanWindowRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = TRUTH_CONSUMERS.storeCreditLedger;
  const windowSize = scanWindowRows || config.scanWindowRows;
  
  const result = {
    scannedRows: 0,
    unknownNames: [],
    queuedCount: 0
  };
  
  // Find Store_Credit_Ledger sheet
  let ledgerSheet = null;
  for (const sheetName of config.sheetNames) {
    ledgerSheet = ss.getSheetByName(sheetName);
    if (ledgerSheet) break;
  }
  
  if (!ledgerSheet) {
    result.message = 'Store_Credit_Ledger sheet not found';
    return result;
  }
  
  const lastRow = ledgerSheet.getLastRow();
  if (lastRow <= 1) {
    result.message = 'Store_Credit_Ledger is empty';
    return result;
  }
  
  // Get headers
  const headers = ledgerSheet.getRange(1, 1, 1, ledgerSheet.getLastColumn()).getValues()[0];
  const nameColIdx = findHeaderIndex_(headers, ['PreferredName', 'preferred_name_id', 'Player', 'Customer', 'Name']);
  
  if (nameColIdx === -1) {
    result.message = 'Could not find player name column in Store_Credit_Ledger';
    return result;
  }
  
  // Determine scan range (last N rows)
  const startRow = Math.max(2, lastRow - windowSize + 1);
  const numRows = lastRow - startRow + 1;
  
  const data = ledgerSheet.getRange(startRow, nameColIdx + 1, numRows, 1).getValues();
  const seenNames = new Set();
  
  data.forEach((row, idx) => {
    const rawName = String(row[0] || '').trim();
    if (!rawName) return;
    
    const normalized = normalizePlayerName(rawName);
    const normalizedLower = normalized.toLowerCase();
    
    // Skip if already processed this name in this scan
    if (seenNames.has(normalizedLower)) return;
    seenNames.add(normalizedLower);
    
    result.scannedRows++;
    
    // Check if canonical
    if (!isCanonicalName(normalized)) {
      result.unknownNames.push(rawName);
      
      // Queue to UndiscoveredNames
      const queued = queueUnknownName(
        rawName,
        'STORE_CREDIT_LEDGER',
        ledgerSheet.getName(),
        `Row ${startRow + idx}`
      );
      
      if (queued) {
        result.queuedCount++;
      }
    }
  });
  
  // Mark ledger status in document properties
  const isDirty = result.unknownNames.length > 0;
  setSheetHygieneStatus_('Store_Credit_Ledger', !isDirty);
  
  result.message = isDirty
    ? `Found ${result.unknownNames.length} unknown names (queued ${result.queuedCount})`
    : 'All names are canonical';
  
  return result;
}

// ============================================================================
// PREORDERS SCANNING
// ============================================================================

/**
 * Scans Preorders sheet for unknown player names
 * Queues unknowns to UndiscoveredNames
 * 
 * @param {number} [scanWindowRows] - Number of recent rows to scan (default: 1000)
 * @return {Object} { scannedRows: number, unknownNames: Array<string>, queuedCount: number }
 */
function SCAN_PREORDERS_FOR_UNKNOWN_NAMES(scanWindowRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = TRUTH_CONSUMERS.preorders;
  const windowSize = scanWindowRows || config.scanWindowRows;
  
  const result = {
    scannedRows: 0,
    unknownNames: [],
    queuedCount: 0
  };
  
  // Find Preorders sheet (try multiple names)
  let preordersSheet = null;
  for (const sheetName of config.sheetNames) {
    preordersSheet = ss.getSheetByName(sheetName);
    if (preordersSheet) break;
  }
  
  if (!preordersSheet) {
    result.message = 'Preorders sheet not found';
    return result;
  }
  
  const lastRow = preordersSheet.getLastRow();
  if (lastRow <= 1) {
    result.message = 'Preorders sheet is empty';
    return result;
  }
  
  // Get headers
  const headers = preordersSheet.getRange(1, 1, 1, preordersSheet.getLastColumn()).getValues()[0];
  const nameColIdx = findHeaderIndex_(headers, ['PreferredName', 'Customer', 'Player', 'Name']);
  
  if (nameColIdx === -1) {
    result.message = 'Could not find customer name column in Preorders sheet';
    return result;
  }
  
  // Determine scan range
  const startRow = Math.max(2, lastRow - windowSize + 1);
  const numRows = lastRow - startRow + 1;
  
  const data = preordersSheet.getRange(startRow, nameColIdx + 1, numRows, 1).getValues();
  const seenNames = new Set();
  
  data.forEach((row, idx) => {
    const rawName = String(row[0] || '').trim();
    if (!rawName) return;
    
    const normalized = normalizePlayerName(rawName);
    const normalizedLower = normalized.toLowerCase();
    
    if (seenNames.has(normalizedLower)) return;
    seenNames.add(normalizedLower);
    
    result.scannedRows++;
    
    if (!isCanonicalName(normalized)) {
      result.unknownNames.push(rawName);
      
      const queued = queueUnknownName(
        rawName,
        'PREORDERS',
        preordersSheet.getName(),
        `Row ${startRow + idx}`
      );
      
      if (queued) {
        result.queuedCount++;
      }
    }
  });
  
  // Mark preorders status
  const isDirty = result.unknownNames.length > 0;
  setSheetHygieneStatus_('Preorders', !isDirty);
  
  result.message = isDirty
    ? `Found ${result.unknownNames.length} unknown names (queued ${result.queuedCount})`
    : 'All names are canonical';
  
  return result;
}

// ============================================================================
// HYGIENE STATUS TRACKING
// ============================================================================

/**
 * Sets hygiene status for a sheet (clean/dirty)
 * Uses document properties for persistence
 * 
 * @param {string} sheetName - Sheet name
 * @param {boolean} isClean - True if clean, false if dirty
 * @private
 */
function setSheetHygieneStatus_(sheetName, isClean) {
  try {
    const props = PropertiesService.getDocumentProperties();
    const key = `HYGIENE_STATUS_${sheetName}`;
    const timestamp = new Date().toISOString();
    
    const status = {
      clean: isClean,
      lastChecked: timestamp
    };
    
    props.setProperty(key, JSON.stringify(status));
  } catch (e) {
    console.error('setSheetHygieneStatus_ error:', e);
  }
}

/**
 * Gets hygiene status for a sheet
 * @param {string} sheetName - Sheet name
 * @return {Object} { clean: boolean, lastChecked: string }
 */
function getSheetHygieneStatus(sheetName) {
  try {
    const props = PropertiesService.getDocumentProperties();
    const key = `HYGIENE_STATUS_${sheetName}`;
    const statusJson = props.getProperty(key);
    
    if (statusJson) {
      return JSON.parse(statusJson);
    }
  } catch (e) {
    console.error('getSheetHygieneStatus error:', e);
  }
  
  return { clean: false, lastChecked: null };
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Finds column index from array of header synonyms
 * @param {Array<string>} headers - Header row
 * @param {Array<string>} synonyms - Possible column names
 * @return {number} Column index (0-based) or -1 if not found
 * @private
 */
function findHeaderIndex_(headers, synonyms) {
  const normalizedHeaders = headers.map(h => String(h || '').trim().toLowerCase());
  
  for (const synonym of synonyms) {
    const idx = normalizedHeaders.indexOf(synonym.toLowerCase());
    if (idx !== -1) return idx;
  }
  
  return -1;
}

// ============================================================================
// CREATE NEW PLAYER WORKFLOW (Retail Onboarding)
// ============================================================================

/**
 * Creates a new player from an undiscovered name entry
 * This is the retail-side onboarding workflow
 * 
 * @param {string} normalizedName - The normalized name from UndiscoveredNames
 * @param {string} [canonicalName] - Optional: override canonical spelling
 * @return {Object} { success: boolean, message: string, canonicalName: string }
 */
function CREATE_NEW_PLAYER_FROM_UNDISCOVERED(normalizedName, canonicalName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Step 1: Validate input
    if (!normalizedName || typeof normalizedName !== 'string') {
      return {
        success: false,
        message: 'Normalized name is required'
      };
    }
    
    const normalized = normalizePlayerName(normalizedName);
    if (!normalized) {
      return {
        success: false,
        message: 'Invalid name provided'
      };
    }
    
    // Step 2: Use provided canonical name or default to normalized
    const finalCanonicalName = canonicalName ? normalizePlayerName(canonicalName) : normalized;
    
    // Step 3: Check if already exists in PreferredNames
    if (isCanonicalName(finalCanonicalName)) {
      return {
        success: false,
        message: `Player "${finalCanonicalName}" already exists in PreferredNames`,
        canonicalName: finalCanonicalName
      };
    }
    
    // Step 4: Add to PreferredNames
    const prefSheet = ss.getSheetByName('PreferredNames');
    if (!prefSheet) {
      return {
        success: false,
        message: 'PreferredNames sheet not found'
      };
    }
    
    prefSheet.appendRow([finalCanonicalName]);
    
    // Step 5: Mark UndiscoveredNames entry as RESOLVED
    const undiscoveredSheet = ss.getSheetByName('UndiscoveredNames');
    if (undiscoveredSheet) {
      markUndiscoveredAsResolved_(
        undiscoveredSheet,
        normalized,
        finalCanonicalName,
        'ADD_CANONICAL'
      );
    }
    
    // Step 6: Provision player to tracking sheets (if provisioner exists)
    if (typeof PROVISION_PLAYER_PROFILE === 'function') {
      try {
        PROVISION_PLAYER_PROFILE(finalCanonicalName);
      } catch (provisionError) {
        console.warn('Provisioning failed:', provisionError);
        // Continue anyway - provisioning is optional
      }
    }
    
    // Step 7: Log to Integrity_Log
    if (typeof logIntegrityAction === 'function') {
      try {
        logIntegrityAction('NEW_PLAYER_FROM_UNDISCOVERED', {
          preferredName: finalCanonicalName,
          details: `Created new player "${finalCanonicalName}" from undiscovered name "${normalized}"`,
          status: 'SUCCESS'
        });
      } catch (logError) {
        console.warn('Integrity logging failed:', logError);
      }
    }
    
    return {
      success: true,
      message: `Successfully created player "${finalCanonicalName}"`,
      canonicalName: finalCanonicalName
    };
    
  } catch (e) {
    console.error('CREATE_NEW_PLAYER_FROM_UNDISCOVERED error:', e);
    return {
      success: false,
      message: 'Error creating player: ' + e.message
    };
  }
}

/**
 * Marks an UndiscoveredNames entry as RESOLVED
 * @param {Sheet} sheet - UndiscoveredNames sheet
 * @param {string} normalizedName - The normalized name to resolve
 * @param {string} resolvedAs - The canonical name it resolved to
 * @param {string} resolutionType - Type: ADD_CANONICAL, MAP_EXISTING, IGNORE
 * @private
 */
function markUndiscoveredAsResolved_(sheet, normalizedName, resolvedAs, resolutionType) {
  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    
    const headers = data[0];
    const nameColIdx = headers.indexOf('NormalizedName') !== -1 ? headers.indexOf('NormalizedName') : headers.indexOf('Potential_Name');
    const statusColIdx = headers.indexOf('Status');
    const resolvedAsIdx = headers.indexOf('ResolvedAs');
    const resolutionTypeIdx = headers.indexOf('ResolutionType');
    const resolvedByIdx = headers.indexOf('ResolvedBy');
    const resolvedAtIdx = headers.indexOf('ResolvedAt');
    
    const normalizedLower = normalizePlayerName(normalizedName).toLowerCase();
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
    const user = Session.getActiveUser().getEmail() || 'system';
    
    for (let i = 1; i < data.length; i++) {
      const existingName = normalizePlayerName(data[i][nameColIdx]).toLowerCase();
      
      if (existingName === normalizedLower) {
        // Update this row
        if (statusColIdx !== -1) {
          sheet.getRange(i + 1, statusColIdx + 1).setValue('RESOLVED');
        }
        if (resolvedAsIdx !== -1) {
          sheet.getRange(i + 1, resolvedAsIdx + 1).setValue(resolvedAs);
        }
        if (resolutionTypeIdx !== -1) {
          sheet.getRange(i + 1, resolutionTypeIdx + 1).setValue(resolutionType);
        }
        if (resolvedByIdx !== -1) {
          sheet.getRange(i + 1, resolvedByIdx + 1).setValue(user);
        }
        if (resolvedAtIdx !== -1) {
          sheet.getRange(i + 1, resolvedAtIdx + 1).setValue(timestamp);
        }
        
        break;
      }
    }
  } catch (e) {
    console.error('markUndiscoveredAsResolved_ error:', e);
  }
}

// ============================================================================
// RESOLVE UNKNOWNS DASHBOARD
// ============================================================================

/**
 * Refreshes the Resolve_Unknowns_Dashboard sheet
 * Creates a staff-friendly view of OPEN unknown names by source type
 * 
 * @return {Object} { success: boolean, message: string, sections: Object }
 */
function REFRESH_RESOLVE_UNKNOWN_DASHBOARD() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get UndiscoveredNames data
    const undiscoveredSheet = ss.getSheetByName('UndiscoveredNames');
    if (!undiscoveredSheet) {
      return {
        success: false,
        message: 'UndiscoveredNames sheet not found'
      };
    }
    
    const data = undiscoveredSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return {
        success: false,
        message: 'UndiscoveredNames sheet is empty'
      };
    }
    
    const headers = data[0];
    const nameColIdx = headers.indexOf('NormalizedName') !== -1 ? headers.indexOf('NormalizedName') : 0;
    const rawNameColIdx = headers.indexOf('RawName') !== -1 ? headers.indexOf('RawName') : 1;
    const sourceTypeColIdx = headers.indexOf('SourceType') !== -1 ? headers.indexOf('SourceType') : 2;
    const sourceSheetsColIdx = headers.indexOf('SourceSheets') !== -1 ? headers.indexOf('SourceSheets') : 3;
    const seenCountColIdx = headers.indexOf('SeenCount') !== -1 ? headers.indexOf('SeenCount') : 5;
    const firstSeenColIdx = headers.indexOf('FirstSeen') !== -1 ? headers.indexOf('FirstSeen') : 6;
    const lastSeenColIdx = headers.indexOf('LastSeen') !== -1 ? headers.indexOf('LastSeen') : 7;
    const statusColIdx = headers.indexOf('Status') !== -1 ? headers.indexOf('Status') : 8;
    
    // Filter for OPEN items only
    const openItems = [];
    for (let i = 1; i < data.length; i++) {
      const status = String(data[i][statusColIdx] || 'OPEN').toUpperCase();
      if (status === 'OPEN') {
        openItems.push({
          normalizedName: data[i][nameColIdx],
          rawName: data[i][rawNameColIdx],
          sourceType: data[i][sourceTypeColIdx] || 'UNKNOWN',
          sourceSheets: data[i][sourceSheetsColIdx],
          seenCount: data[i][seenCountColIdx],
          firstSeen: data[i][firstSeenColIdx],
          lastSeen: data[i][lastSeenColIdx]
        });
      }
    }
    
    // Categorize by source type
    const sections = {
      storeCreditLedger: [],
      preorders: [],
      events: [],
      other: []
    };
    
    openItems.forEach(item => {
      const sourceType = String(item.sourceType).toUpperCase();
      
      if (sourceType.includes('STORE_CREDIT') || sourceType.includes('LEDGER')) {
        sections.storeCreditLedger.push(item);
      } else if (sourceType.includes('PREORDER')) {
        sections.preorders.push(item);
      } else if (sourceType.includes('EVENT') || sourceType.includes('SCAN')) {
        sections.events.push(item);
      } else {
        sections.other.push(item);
      }
    });
    
    // Create or update dashboard sheet
    let dashboardSheet = ss.getSheetByName('Resolve_Unknowns_Dashboard');
    if (!dashboardSheet) {
      dashboardSheet = ss.insertSheet('Resolve_Unknowns_Dashboard');
    }
    
    dashboardSheet.clear();
    
    // Build dashboard content
    let currentRow = 1;
    const dashboardHeaders = [
      'NormalizedName',
      'RawName',
      'SourceType',
      'SourceSheets',
      'SeenCount',
      'FirstSeen',
      'LastSeen',
      'SuggestedNextAction'
    ];
    
    // Section 1: Retail-first (Store Credit)
    if (sections.storeCreditLedger.length > 0) {
      dashboardSheet.getRange(currentRow, 1, 1, 1).setValue('═══ RETAIL-FIRST (Store Credit) ═══');
      dashboardSheet.getRange(currentRow, 1, 1, 8).setFontWeight('bold').setBackground('#fce5cd');
      currentRow++;
      
      dashboardSheet.getRange(currentRow, 1, 1, dashboardHeaders.length).setValues([dashboardHeaders]);
      dashboardSheet.getRange(currentRow, 1, 1, dashboardHeaders.length).setFontWeight('bold').setBackground('#f3f3f3');
      currentRow++;
      
      sections.storeCreditLedger.forEach(item => {
        dashboardSheet.appendRow([
          item.normalizedName,
          item.rawName,
          item.sourceType,
          item.sourceSheets,
          item.seenCount,
          item.firstSeen,
          item.lastSeen,
          'Create New Player'
        ]);
      });
      
      currentRow = dashboardSheet.getLastRow() + 2;
    }
    
    // Section 2: Event-first (Events/Scans)
    if (sections.events.length > 0) {
      dashboardSheet.getRange(currentRow, 1, 1, 1).setValue('═══ EVENT-FIRST (Events / Scans) ═══');
      dashboardSheet.getRange(currentRow, 1, 1, 8).setFontWeight('bold').setBackground('#d9ead3');
      currentRow++;
      
      dashboardSheet.getRange(currentRow, 1, 1, dashboardHeaders.length).setValues([dashboardHeaders]);
      dashboardSheet.getRange(currentRow, 1, 1, dashboardHeaders.length).setFontWeight('bold').setBackground('#f3f3f3');
      currentRow++;
      
      sections.events.forEach(item => {
        dashboardSheet.appendRow([
          item.normalizedName,
          item.rawName,
          item.sourceType,
          item.sourceSheets,
          item.seenCount,
          item.firstSeen,
          item.lastSeen,
          'Spellcheck + Canonicalize'
        ]);
      });
      
      currentRow = dashboardSheet.getLastRow() + 2;
    }
    
    // Section 3: Preorders
    if (sections.preorders.length > 0) {
      dashboardSheet.getRange(currentRow, 1, 1, 1).setValue('═══ PREORDERS ═══');
      dashboardSheet.getRange(currentRow, 1, 1, 8).setFontWeight('bold').setBackground('#cfe2f3');
      currentRow++;
      
      dashboardSheet.getRange(currentRow, 1, 1, dashboardHeaders.length).setValues([dashboardHeaders]);
      dashboardSheet.getRange(currentRow, 1, 1, dashboardHeaders.length).setFontWeight('bold').setBackground('#f3f3f3');
      currentRow++;
      
      sections.preorders.forEach(item => {
        dashboardSheet.appendRow([
          item.normalizedName,
          item.rawName,
          item.sourceType,
          item.sourceSheets,
          item.seenCount,
          item.firstSeen,
          item.lastSeen,
          'Must Resolve Before Processing'
        ]);
      });
      
      currentRow = dashboardSheet.getLastRow() + 2;
    }
    
    // Section 4: Other
    if (sections.other.length > 0) {
      dashboardSheet.getRange(currentRow, 1, 1, 1).setValue('═══ OTHER ═══');
      dashboardSheet.getRange(currentRow, 1, 1, 8).setFontWeight('bold').setBackground('#efefef');
      currentRow++;
      
      dashboardSheet.getRange(currentRow, 1, 1, dashboardHeaders.length).setValues([dashboardHeaders]);
      dashboardSheet.getRange(currentRow, 1, 1, dashboardHeaders.length).setFontWeight('bold').setBackground('#f3f3f3');
      currentRow++;
      
      sections.other.forEach(item => {
        dashboardSheet.appendRow([
          item.normalizedName,
          item.rawName,
          item.sourceType,
          item.sourceSheets,
          item.seenCount,
          item.firstSeen,
          item.lastSeen,
          'Review Manually'
        ]);
      });
    }
    
    // Auto-resize columns
    dashboardSheet.autoResizeColumns(1, dashboardHeaders.length);
    
    return {
      success: true,
      message: `Dashboard refreshed: ${openItems.length} open items`,
      sections: {
        storeCreditLedger: sections.storeCreditLedger.length,
        events: sections.events.length,
        preorders: sections.preorders.length,
        other: sections.other.length
      }
    };
    
  } catch (e) {
    console.error('REFRESH_RESOLVE_UNKNOWN_DASHBOARD error:', e);
    return {
      success: false,
      message: 'Error refreshing dashboard: ' + e.message
    };
  }
}

// ============================================================================
// PROVISIONAL LEDGER NAMES (for retail onboarding visibility)
// ============================================================================

/**
 * Generates Provisional_Ledger_Names sheet
 * Shows non-canonical Store Credit customers awaiting onboarding
 * 
 * @return {Object} { success: boolean, message: string, provisionalCount: number }
 */
function generateProvisionalLedgerNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get Store_Credit_Ledger
    const ledgerSheet = ss.getSheetByName('Store_Credit_Ledger');
    if (!ledgerSheet || ledgerSheet.getLastRow() <= 1) {
      return {
        success: false,
        message: 'Store_Credit_Ledger not found or empty'
      };
    }
    
    // Get all ledger data
    const ledgerData = ledgerSheet.getDataRange().getValues();
    const ledgerHeaders = ledgerData[0];
    const nameColIdx = findHeaderIndex_(ledgerHeaders, ['PreferredName', 'preferred_name_id', 'Player', 'Customer']);
    const amountColIdx = findHeaderIndex_(ledgerHeaders, ['Amount', 'RunningBalance']);
    const timestampColIdx = findHeaderIndex_(ledgerHeaders, ['Timestamp', 'Date']);
    
    if (nameColIdx === -1) {
      return {
        success: false,
        message: 'Could not find name column in Store_Credit_Ledger'
      };
    }
    
    // Aggregate by name
    const nameAggregates = new Map();
    
    for (let i = 1; i < ledgerData.length; i++) {
      const rawName = String(ledgerData[i][nameColIdx] || '').trim();
      if (!rawName) continue;
      
      const normalized = normalizePlayerName(rawName);
      const normalizedLower = normalized.toLowerCase();
      
      // Skip if canonical
      if (isCanonicalName(normalized)) continue;
      
      const amount = amountColIdx !== -1 ? parseFloat(ledgerData[i][amountColIdx]) || 0 : 0;
      const timestamp = timestampColIdx !== -1 ? ledgerData[i][timestampColIdx] : '';
      
      if (!nameAggregates.has(normalizedLower)) {
        nameAggregates.set(normalizedLower, {
          rawName: rawName,
          normalizedName: normalized,
          balance: 0,
          firstSeen: timestamp,
          lastSeen: timestamp,
          transactionCount: 0
        });
      }
      
      const agg = nameAggregates.get(normalizedLower);
      agg.balance += amount;
      agg.transactionCount++;
      if (timestamp) {
        agg.lastSeen = timestamp;
      }
    }
    
    // Check UndiscoveredNames for status
    const undiscoveredSheet = ss.getSheetByName('UndiscoveredNames');
    const undiscoveredStatus = new Map();
    
    if (undiscoveredSheet && undiscoveredSheet.getLastRow() > 1) {
      const undiscData = undiscoveredSheet.getDataRange().getValues();
      const undiscHeaders = undiscData[0];
      const undiscNameIdx = headers.indexOf('NormalizedName') !== -1 ? undiscHeaders.indexOf('NormalizedName') : 0;
      const undiscStatusIdx = undiscHeaders.indexOf('Status') !== -1 ? undiscHeaders.indexOf('Status') : 8;
      
      for (let i = 1; i < undiscData.length; i++) {
        const name = normalizePlayerName(undiscData[i][undiscNameIdx]).toLowerCase();
        const status = String(undiscData[i][undiscStatusIdx] || 'OPEN');
        undiscoveredStatus.set(name, status);
      }
    }
    
    // Create Provisional_Ledger_Names sheet
    let provisionalSheet = ss.getSheetByName('Provisional_Ledger_Names');
    if (!provisionalSheet) {
      provisionalSheet = ss.insertSheet('Provisional_Ledger_Names');
    }
    
    provisionalSheet.clear();
    
    const headers = [
      'RawName',
      'NormalizedName',
      'LedgerBalance',
      'TransactionCount',
      'FirstSeen',
      'LastSeen',
      'Status',
      'ActionHint'
    ];
    
    provisionalSheet.appendRow(headers);
    provisionalSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
    
    // Build rows
    const rows = [];
    nameAggregates.forEach((agg, normalizedLower) => {
      const status = undiscoveredStatus.get(normalizedLower) || 'OPEN';
      const actionHint = status === 'OPEN' ? 'Create New Player' : (status === 'RESOLVED' ? 'Already Resolved' : 'Map to Existing');
      
      rows.push([
        agg.rawName,
        agg.normalizedName,
        agg.balance,
        agg.transactionCount,
        agg.firstSeen,
        agg.lastSeen,
        status,
        actionHint
      ]);
    });
    
    if (rows.length > 0) {
      provisionalSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      provisionalSheet.autoResizeColumns(1, headers.length);
    }
    
    return {
      success: true,
      message: `Generated ${rows.length} provisional ledger names`,
      provisionalCount: rows.length
    };
    
  } catch (e) {
    console.error('generateProvisionalLedgerNames error:', e);
    return {
      success: false,
      message: 'Error generating provisional names: ' + e.message
    };
  }
}

// ============================================================================
// PLAYERLOOKUP EXPANSION (Phase 5)
// ============================================================================

/**
 * Builds PlayerLookup sheet with transactional data rollups
 * ENFORCES hygiene: fails if Store_Credit_Ledger or Preorders have unresolved names
 * 
 * @return {Object} { success: boolean, message: string, playerCount: number }
 */
function SAFE_RUN_PLAYERLOOKUP_BUILD() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Step 1: Check hygiene status
    const ledgerStatus = getSheetHygieneStatus('Store_Credit_Ledger');
    const preordersStatus = getSheetHygieneStatus('Preorders');
    
    const errors = [];
    if (!ledgerStatus.clean && TRUTH_CONSUMERS.storeCreditLedger.blocksPlayerLookupRollupsWhenDirty) {
      errors.push('Store_Credit_Ledger has unresolved names');
    }
    if (!preordersStatus.clean && TRUTH_CONSUMERS.preorders.blocksPlayerLookupRollupsWhenDirty) {
      errors.push('Preorders has unresolved names');
    }
    
    if (errors.length > 0) {
      return {
        success: false,
        message: `PlayerLookup build blocked: ${errors.join(', ')}. Run SCAN functions first.`,
        errors: errors
      };
    }
    
    // Step 2: Load canonical names from PreferredNames
    const { canonicalList } = loadCanonicalNames_();
    
    if (canonicalList.length === 0) {
      return {
        success: false,
        message: 'PreferredNames sheet is empty'
      };
    }
    
    // Step 3: Build player lookup map
    const playerLookup = new Map();
    
    canonicalList.forEach(canonicalName => {
      playerLookup.set(canonicalName.toLowerCase(), {
        PreferredName: canonicalName,
        StoreCreditBalance: 0,
        OpenPreorderCount: 0,
        OpenPreorderQtyTotal: 0,
        DepositsTotal: 0,
        BalanceDueTotal: 0,
        LedgerClean: ledgerStatus.clean,
        PreordersClean: preordersStatus.clean,
        LastRefresh: new Date().toISOString()
      });
    });
    
    // Step 4: Aggregate Store Credit balances
    const ledgerSheet = ss.getSheetByName('Store_Credit_Ledger');
    if (ledgerSheet && ledgerSheet.getLastRow() > 1) {
      const ledgerData = ledgerSheet.getDataRange().getValues();
      const ledgerHeaders = ledgerData[0];
      const nameColIdx = findHeaderIndex_(ledgerHeaders, ['PreferredName', 'preferred_name_id', 'Player', 'Customer']);
      const balanceColIdx = findHeaderIndex_(ledgerHeaders, ['RunningBalance', 'Balance']);
      
      if (nameColIdx !== -1 && balanceColIdx !== -1) {
        // Get last balance for each player
        const lastBalances = new Map();
        for (let i = ledgerData.length - 1; i >= 1; i--) {
          const name = normalizePlayerName(ledgerData[i][nameColIdx]).toLowerCase();
          if (name && !lastBalances.has(name)) {
            const balance = parseFloat(ledgerData[i][balanceColIdx]) || 0;
            lastBalances.set(name, balance);
          }
        }
        
        // Apply to lookup
        lastBalances.forEach((balance, nameLower) => {
          if (playerLookup.has(nameLower)) {
            playerLookup.get(nameLower).StoreCreditBalance = balance;
          }
        });
      }
    }
    
    // Step 5: Aggregate Preorders
    const preordersSheet = ss.getSheetByName('Preorders_Sold') || ss.getSheetByName('Preorders');
    if (preordersSheet && preordersSheet.getLastRow() > 1) {
      const preordersData = preordersSheet.getDataRange().getValues();
      const preordersHeaders = preordersData[0];
      const nameColIdx = findHeaderIndex_(preordersHeaders, ['PreferredName', 'Customer', 'Player']);
      const qtyColIdx = findHeaderIndex_(preordersHeaders, ['Qty', 'Quantity']);
      const statusColIdx = findHeaderIndex_(preordersHeaders, ['Status']);
      const depositColIdx = findHeaderIndex_(preordersHeaders, ['Deposit_Paid', 'DepositPaid', 'Deposit']);
      const balanceDueColIdx = findHeaderIndex_(preordersHeaders, ['Balance_Due', 'BalanceDue', 'Balance']);
      
      if (nameColIdx !== -1) {
        for (let i = 1; i < preordersData.length; i++) {
          const name = normalizePlayerName(preordersData[i][nameColIdx]).toLowerCase();
          const status = statusColIdx !== -1 ? String(preordersData[i][statusColIdx] || '').toUpperCase() : '';
          
          // Only count open/pending preorders
          if (status && (status === 'CANCELLED' || status === 'COMPLETED' || status === 'FULFILLED')) {
            continue;
          }
          
          if (playerLookup.has(name)) {
            const player = playerLookup.get(name);
            player.OpenPreorderCount++;
            
            if (qtyColIdx !== -1) {
              player.OpenPreorderQtyTotal += parseInt(preordersData[i][qtyColIdx]) || 0;
            }
            
            if (depositColIdx !== -1) {
              player.DepositsTotal += parseFloat(preordersData[i][depositColIdx]) || 0;
            }
            
            if (balanceDueColIdx !== -1) {
              player.BalanceDueTotal += parseFloat(preordersData[i][balanceDueColIdx]) || 0;
            }
          }
        }
      }
    }
    
    // Step 6: Write to PlayerLookup sheet
    let lookupSheet = ss.getSheetByName('PlayerLookup');
    if (!lookupSheet) {
      lookupSheet = ss.insertSheet('PlayerLookup');
    }
    
    lookupSheet.clear();
    
    const headers = [
      'PreferredName',
      'StoreCreditBalance',
      'OpenPreorderCount',
      'OpenPreorderQtyTotal',
      'DepositsTotal',
      'BalanceDueTotal',
      'LedgerClean',
      'PreordersClean',
      'LastRefresh'
    ];
    
    lookupSheet.appendRow(headers);
    lookupSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
    
    // Build rows
    const rows = [];
    playerLookup.forEach(player => {
      rows.push([
        player.PreferredName,
        player.StoreCreditBalance,
        player.OpenPreorderCount,
        player.OpenPreorderQtyTotal,
        player.DepositsTotal,
        player.BalanceDueTotal,
        player.LedgerClean,
        player.PreordersClean,
        player.LastRefresh
      ]);
    });
    
    if (rows.length > 0) {
      lookupSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      lookupSheet.autoResizeColumns(1, headers.length);
    }
    
    // Step 7: Log action
    if (typeof logIntegrityAction === 'function') {
      try {
        logIntegrityAction('PLAYERLOOKUP_BUILD', {
          details: `Built PlayerLookup with ${rows.length} players (Ledger: ${ledgerStatus.clean ? 'CLEAN' : 'DIRTY'}, Preorders: ${preordersStatus.clean ? 'CLEAN' : 'DIRTY'})`,
          status: 'SUCCESS'
        });
      } catch (logError) {
        console.warn('Integrity logging failed:', logError);
      }
    }
    
    return {
      success: true,
      message: `PlayerLookup built successfully with ${rows.length} players`,
      playerCount: rows.length,
      ledgerClean: ledgerStatus.clean,
      preordersClean: preordersStatus.clean
    };
    
  } catch (e) {
    console.error('SAFE_RUN_PLAYERLOOKUP_BUILD error:', e);
    return {
      success: false,
      message: 'Error building PlayerLookup: ' + e.message
    };
  }
}

// ============================================================================
// UTILITY: Get canonical name from lookup sources
// ============================================================================

/**
 * Attempts to find the canonical name for a given raw name
 * Checks PreferredNames, UndiscoveredNames (resolved), and suggests matches
 * 
 * @param {string} rawName - Raw player name
 * @return {Object} { canonical: string|null, suggestions: Array<string>, needsResolution: boolean }
 */
function findCanonicalNameFor(rawName) {
  const normalized = normalizePlayerName(rawName);
  const normalizedLower = normalized.toLowerCase();
  
  // Check if already canonical
  const canonical = getCanonicalName(normalized);
  if (canonical) {
    return {
      canonical: canonical,
      suggestions: [],
      needsResolution: false
    };
  }
  
  // Check UndiscoveredNames for resolved mapping
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const undiscoveredSheet = ss.getSheetByName('UndiscoveredNames');
    
    if (undiscoveredSheet && undiscoveredSheet.getLastRow() > 1) {
      const data = undiscoveredSheet.getDataRange().getValues();
      const headers = data[0];
      const nameColIdx = headers.indexOf('NormalizedName') !== -1 ? headers.indexOf('NormalizedName') : 0;
      const statusColIdx = headers.indexOf('Status') !== -1 ? headers.indexOf('Status') : 8;
      const resolvedAsIdx = headers.indexOf('ResolvedAs') !== -1 ? headers.indexOf('ResolvedAs') : 9;
      
      for (let i = 1; i < data.length; i++) {
        const existingName = normalizePlayerName(data[i][nameColIdx]).toLowerCase();
        const status = String(data[i][statusColIdx] || 'OPEN').toUpperCase();
        const resolvedAs = String(data[i][resolvedAsIdx] || '').trim();
        
        if (existingName === normalizedLower && status === 'RESOLVED' && resolvedAs) {
          return {
            canonical: resolvedAs,
            suggestions: [],
            needsResolution: false
          };
        }
      }
    }
  } catch (e) {
    console.error('Error checking UndiscoveredNames:', e);
  }
  
  // Generate suggestions based on similarity
  const { canonicalList } = loadCanonicalNames_();
  const suggestions = [];
  
  // Simple substring matching for suggestions
  canonicalList.forEach(canonical => {
    if (canonical.toLowerCase().includes(normalizedLower) || normalizedLower.includes(canonical.toLowerCase())) {
      suggestions.push(canonical);
    }
  });
  
  return {
    canonical: null,
    suggestions: suggestions.slice(0, 5),
    needsResolution: true
  };
}
