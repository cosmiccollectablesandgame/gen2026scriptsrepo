/**
 * ============================================================================
 * BP HEADER RESOLVER v8.0.0
 * ============================================================================
 * 
 * @fileoverview Canonical header names and synonym resolution for BP system.
 * Eliminates raw indexOf() calls and provides:
 * - Single source of truth for column names
 * - Synonym/legacy name handling with warnings
 * - Required vs optional column validation
 * - Automatic logging to Integrity_Log
 * 
 * USAGE:
 *   const nameCol = resolveHeaderIndex(headers, BP_HEADERS.PREFERRED_NAME);
 *   const currentCol = resolveHeaderIndex(headers, BP_HEADERS.CURRENT_BP);
 * 
 * REPLACES:
 *   const nameCol = headers.indexOf('PreferredName');
 *   const currentCol = headers.indexOf('Current_BP');
 */

// ============================================================================
// CANONICAL HEADER NAMES (v8.0.0 Spec)
// ============================================================================

/**
 * Canonical BP_Total column names (use these EXACTLY in all new code)
 * @const {Object}
 */
const BP_HEADERS = {
  // Player Identity
  PREFERRED_NAME: 'PreferredName',
  
  // State Columns (Pipeline writes only)
  CURRENT_BP: 'Current_BP',           // Wallet (0-100 cap)
  HISTORICAL_BP: 'Historical_BP',     // Lifetime earned (monotonic ↑)
  REDEEMED_TOTAL: 'Redeemed_Total',   // Lifetime redeemed (monotonic ↑)
  PRESTIGE_BP: 'Prestige_BP',         // Spillover above cap (monotonic ↑)
  
  // Source Columns (Mission services write, Pipeline reads)
  ATTENDANCE_POINTS: 'Attendance Mission Points',
  FLAG_POINTS: 'Flag Mission Points',
  DICE_POINTS: 'Dice Roll Points',
  MANUAL_ADJUSTMENT: 'Manual_Adjustment_Points',
  
  // Metadata
  LAST_UPDATED: 'LastUpdated'
};

/**
 * Header synonyms map: deprecated/legacy names → canonical names
 * When found, logs a WARNING to Integrity_Log
 * @const {Object}
 */
const HEADER_SYNONYMS = {
  // PreferredName synonyms
  'Player': 'PreferredName',
  'Name': 'PreferredName',
  'Preferred_Name': 'PreferredName',
  'PlayerName': 'PreferredName',
  
  // Current_BP synonyms
  'BP_Current': 'Current_BP',
  'BP': 'Current_BP',
  'BonusPoints': 'Current_BP',
  'CurrentBP': 'Current_BP',
  
  // Historical_BP synonyms
  'BP_Historical': 'Historical_BP',
  'Lifetime_Earned': 'Historical_BP',
  'HistoricalBP': 'Historical_BP',
  'Total_BP_Earned_Lifetime': 'Historical_BP',
  
  // Redeemed_Total synonyms
  'Total_Redeemed': 'Redeemed_Total',
  'Redeemed_BP': 'Redeemed_Total',
  'BP_Redeemed': 'Redeemed_Total',
  'Total_BP_Redeemed_Lifetime': 'Redeemed_Total',
  
  // Prestige_BP synonyms
  'Prestige_Points': 'Prestige_BP',
  'PrestigePoints': 'Prestige_BP',
  'Overflow_BP': 'Prestige_BP',
  
  // Attendance Mission Points synonyms
  'Attendance Missions': 'Attendance Mission Points',
  'Attendance_Points': 'Attendance Mission Points',
  'Attendance_Missions': 'Attendance Mission Points',
  'AttendancePoints': 'Attendance Mission Points',
  
  // Flag Mission Points synonyms
  'Flag Missions': 'Flag Mission Points',
  'Flag_Mission_Points': 'Flag Mission Points',
  'Flag_Points': 'Flag Mission Points',
  'FlagPoints': 'Flag Mission Points',
  
  // Dice Roll Points synonyms
  'Dice_Points': 'Dice Roll Points',
  'DicePoints': 'Dice Roll Points',
  'Dice Points': 'Dice Roll Points',
  'Points_From_Dice_Rolls': 'Dice Roll Points',
  
  // Manual_Adjustment_Points synonyms
  'Manual_Adjustments': 'Manual_Adjustment_Points',
  'ManualPoints': 'Manual_Adjustment_Points',
  'Adjustment_Points': 'Manual_Adjustment_Points',
  
  // LastUpdated synonyms
  'Last_Updated': 'LastUpdated',
  'Updated': 'LastUpdated',
  'Timestamp': 'LastUpdated'
};

// ============================================================================
// HEADER RESOLVER FUNCTIONS
// ============================================================================

/**
 * Resolves a canonical header name to its column index with synonym support
 * 
 * @param {Array<string>} headers - Header row from sheet
 * @param {string} canonicalName - Canonical name from BP_HEADERS
 * @param {boolean} required - If true, throws error when missing
 * @return {number} Column index (0-based) or -1 if not found and not required
 * @throws {Error} If required header is missing
 */
function resolveHeaderIndex(headers, canonicalName, required = true) {
  // Try canonical name first (fast path)
  let idx = headers.indexOf(canonicalName);
  
  if (idx !== -1) {
    return idx;
  }
  
  // Try synonyms
  for (const [synonym, canonical] of Object.entries(HEADER_SYNONYMS)) {
    if (canonical === canonicalName) {
      idx = headers.indexOf(synonym);
      if (idx !== -1) {
        // Found via synonym - log warning
        if (typeof logIntegrityAction === 'function') {
          logIntegrityAction('HEADER_SYNONYM_WARNING', {
            dfTags: ['DF-091'],
            details: `Using deprecated header "${synonym}" instead of "${canonicalName}". Please update schema.`,
            status: 'WARNING'
          });
        }
        console.warn(`[BP Header Resolver] Found "${synonym}" instead of "${canonicalName}". Update schema to use canonical name.`);
        return idx;
      }
    }
  }
  
  // Not found
  if (required) {
    const errorMsg = `Required header "${canonicalName}" not found in sheet. Available: [${headers.join(', ')}]`;
    
    if (typeof logIntegrityAction === 'function') {
      logIntegrityAction('HEADER_MISSING_ERROR', {
        dfTags: ['DF-090'],
        details: errorMsg,
        status: 'ERROR'
      });
    }
    
    throw new Error(errorMsg);
  }
  
  return -1;
}

/**
 * Validates BP_Total sheet has all required headers
 * 
 * @param {Array<string>} headers - Header row from BP_Total
 * @return {Object} {valid: boolean, missing: Array<string>, synonyms: Array<Object>}
 */
function validateBPTotalHeaders(headers) {
  const requiredHeaders = [
    BP_HEADERS.PREFERRED_NAME,
    BP_HEADERS.CURRENT_BP,
    BP_HEADERS.HISTORICAL_BP,
    BP_HEADERS.ATTENDANCE_POINTS,
    BP_HEADERS.FLAG_POINTS,
    BP_HEADERS.DICE_POINTS,
    BP_HEADERS.LAST_UPDATED
  ];
  
  // Optional headers for v8.0.0 (will be added incrementally)
  const optionalHeaders = [
    BP_HEADERS.REDEEMED_TOTAL,
    BP_HEADERS.PRESTIGE_BP,
    BP_HEADERS.MANUAL_ADJUSTMENT
  ];
  
  const missing = [];
  const synonymsFound = [];
  
  for (const canonical of requiredHeaders) {
    // Check if canonical exists
    if (headers.indexOf(canonical) !== -1) {
      continue;
    }
    
    // Check if synonym exists
    let foundSynonym = false;
    for (const [synonym, canonicalName] of Object.entries(HEADER_SYNONYMS)) {
      if (canonicalName === canonical && headers.indexOf(synonym) !== -1) {
        synonymsFound.push({ synonym, canonical });
        foundSynonym = true;
        break;
      }
    }
    
    if (!foundSynonym) {
      missing.push(canonical);
    }
  }
  
  return {
    valid: missing.length === 0,
    missing,
    synonyms: synonymsFound
  };
}

/**
 * Bulk resolve multiple headers at once
 * Returns a map of canonical name → column index
 * 
 * @param {Array<string>} headers - Header row
 * @param {Array<string>} canonicalNames - Array of canonical names to resolve
 * @param {boolean} allRequired - If true, all headers must exist
 * @return {Object} Map of canonicalName → index
 */
function resolveHeaderIndices(headers, canonicalNames, allRequired = false) {
  const result = {};
  
  for (const canonical of canonicalNames) {
    try {
      result[canonical] = resolveHeaderIndex(headers, canonical, allRequired);
    } catch (e) {
      if (allRequired) {
        throw e;
      }
      result[canonical] = -1;
    }
  }
  
  return result;
}

/**
 * Gets all BP_Total required headers as an array (for schema creation)
 * @return {Array<string>} Array of canonical header names
 */
function getBPTotalRequiredHeaders() {
  return [
    BP_HEADERS.PREFERRED_NAME,
    BP_HEADERS.CURRENT_BP,
    BP_HEADERS.ATTENDANCE_POINTS,
    BP_HEADERS.FLAG_POINTS,
    BP_HEADERS.DICE_POINTS,
    BP_HEADERS.LAST_UPDATED,
    BP_HEADERS.HISTORICAL_BP
  ];
}

/**
 * Gets v8.0.0 full schema headers (including new columns)
 * @return {Array<string>} Array of canonical header names
 */
function getBPTotalV8Headers() {
  return [
    BP_HEADERS.PREFERRED_NAME,
    BP_HEADERS.CURRENT_BP,
    BP_HEADERS.HISTORICAL_BP,
    BP_HEADERS.REDEEMED_TOTAL,
    BP_HEADERS.PRESTIGE_BP,
    BP_HEADERS.ATTENDANCE_POINTS,
    BP_HEADERS.FLAG_POINTS,
    BP_HEADERS.DICE_POINTS,
    BP_HEADERS.MANUAL_ADJUSTMENT,
    BP_HEADERS.LAST_UPDATED
  ];
}
