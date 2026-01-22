/**
 * ============================================================================
 * SCHEMA REGISTRY v1.0.0
 * ============================================================================
 * 
 * @fileoverview Centralized schema definitions for all sheets in the system.
 * This file contains:
 * - SCHEMA_REGISTRY: Complete registry of all sheets with headers and metadata
 * - BP_HEADERS: Canonical header names for BP system
 * - validateSheetExists(): Check if a sheet exists
 * - validateAllSheets(): Validate all required sheets exist
 * 
 * Extracted from MegaBootstrap.js to provide a single source of truth for
 * sheet schemas across all services.
 */

// ============================================================================
// CANONICAL HEADER NAMES (from bpHeaderResolver.js)
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

// ============================================================================
// SCHEMA REGISTRY
// ============================================================================

/**
 * Complete schema registry for all sheets in the system
 * Extracted from MegaBootstrap.js
 * 
 * @const {Object}
 */
const SCHEMA_REGISTRY = {
  Prize_Catalog: {
    name: 'Prize_Catalog',
    headers: [
      'Code',
      'Name',
      'Level',
      'Rarity',
      'COGS',
      'EV_Cost',
      'Qty',
      'Eligible_Rounds',
      'Eligible_End',
      'Player_Threshold',
      'InStock',
      'EV_Explanation',
      'Round_Weight',
      'PV_Multiplier',
      'Projected_Qty'
    ],
    keyColumn: 'Code',
    required: true
  },

  Prize_Throttle: {
    name: 'Prize_Throttle',
    headers: [
      'Parameter',
      'Value'
    ],
    keyColumn: 'Parameter',
    required: true
  },

  Integrity_Log: {
    name: 'Integrity_Log',
    headers: [
      'Timestamp',
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
    ],
    keyColumn: null,
    required: true
  },

  Spent_Pool: {
    name: 'Spent_Pool',
    headers: [
      'Event_ID',
      'Item_Code',
      'Item_Name',
      'Level',
      'Qty',
      'COGS',
      'Total',
      'Timestamp',
      'Batch_ID',
      'Reverted',
      'Event_Type'
    ],
    keyColumn: null,
    required: true
  },

  Key_Tracker: {
    name: 'Key_Tracker',
    headers: [
      'PreferredName',
      'Red',
      'Blue',
      'Green',
      'Yellow',
      'Purple',
      'RainbowEligible',
      'LastUpdated'
    ],
    keyColumn: 'PreferredName',
    required: true
  },

  BP_Total: {
    name: 'BP_Total',
    headers: [
      'PreferredName',
      'Current_BP',
      'Attendance Mission Points',
      'Flag Mission Points',
      'Dice Roll Points',
      'LastUpdated',
      'BP_Historical'
    ],
    keyColumn: 'PreferredName',
    required: true
  },

  Attendance_Missions: {
    name: 'Attendance_Missions',
    headers: [
      'PreferredName',
      'Attendance Mission Points',
      'First Contact',
      'Stellar Explorer',
      'Deck Diver',
      'Lunar Loyalty',
      'Meteor Shower',
      'Sealed Voyager',
      'Draft Navigator',
      'Stellar Scholar',
      'cEDH Events',
      'Limited Events',
      'Academy Events',
      'Outreach Events',
      'Free Play Events',
      'Interstellar Strategist',
      'Black Hole Survivor',
      'LastUpdated'
    ],
    keyColumn: 'PreferredName',
    required: true
  },

  Flag_Missions: {
    name: 'Flag_Missions',
    headers: [
      'PreferredName',
      'Flag Mission Points',
      'Cosmic_Selfie',
      'Review_Writer',
      'Social_Media_Star',
      'App_Explorer',
      'Cosmic_Merchant',
      'Precon_Pioneer',
      'Gravitational_Pull',
      'Rogue_Planet',
      'Quantum_Collector',
      'LastUpdated'
    ],
    keyColumn: 'PreferredName',
    required: true
  },

  Dice_Roll_Points: {
    name: 'Dice Roll Points',
    headers: [
      'PreferredName',
      'Dice Roll Points',
      'LastUpdated'
    ],
    keyColumn: 'PreferredName',
    required: true
  },

  Players_Prize_Wall_Points: {
    name: 'Players_Prize-Wall-Points',
    headers: [
      'PreferredName',
      'Dice_Points_Available',
      'Dice_Points_Spent',
      'Last_Event',
      'LastUpdated'
    ],
    keyColumn: 'PreferredName',
    required: true
  },

  Preview_Artifacts: {
    name: 'Preview_Artifacts',
    headers: [
      'Artifact_ID',
      'Event_ID',
      'Seed',
      'Preview_Hash',
      'Created_At',
      'Expires_At'
    ],
    keyColumn: 'Artifact_ID',
    required: true
  },

  Event_Outcomes: {
    name: 'Event_Outcomes',
    headers: [
      'PreferredName',
      'R1_Result',
      'R2_Result',
      'R3_Result',
      'Notes'
    ],
    keyColumn: 'PreferredName',
    required: false
  },

  Prestige_Overflow: {
    name: 'Prestige_Overflow',
    headers: [
      'PreferredName',
      'Total_Overflow',
      'Last_Updated',
      'Prestige_Tier'
    ],
    keyColumn: 'PreferredName',
    required: false
  }
};

// ============================================================================
// VALIDATION FUNCTIONS
// ============================================================================

/**
 * Validates that a specific sheet exists in the spreadsheet
 * 
 * @param {string} sheetName - Name of the sheet to check
 * @return {Object} {valid: boolean, error: string|null}
 */
function validateSheetExists(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  return {
    valid: sheet !== null,
    error: sheet ? null : `Sheet "${sheetName}" not found`
  };
}

/**
 * Validates that all required sheets exist in the spreadsheet
 * 
 * @return {Object} {valid: boolean, missing: Array<string>, optional: Array<string>}
 */
function validateAllSheets() {
  const results = {
    valid: true,
    missing: [],
    optional: []
  };

  for (const [key, schema] of Object.entries(SCHEMA_REGISTRY)) {
    const validation = validateSheetExists(schema.name);
    
    if (!validation.valid) {
      if (schema.required) {
        results.valid = false;
        results.missing.push(schema.name);
      } else {
        results.optional.push(schema.name);
      }
    }
  }

  return results;
}

/**
 * Gets schema for a specific sheet
 * 
 * @param {string} sheetName - Name of the sheet
 * @return {Object|null} Schema object or null if not found
 */
function getSchemaForSheet(sheetName) {
  // Direct lookup by name
  for (const [key, schema] of Object.entries(SCHEMA_REGISTRY)) {
    if (schema.name === sheetName) {
      return schema;
    }
  }
  return null;
}

/**
 * Gets all required sheet names
 * 
 * @return {Array<string>} Array of required sheet names
 */
function getRequiredSheetNames() {
  return Object.values(SCHEMA_REGISTRY)
    .filter(schema => schema.required)
    .map(schema => schema.name);
}

/**
 * Gets all optional sheet names
 * 
 * @return {Array<string>} Array of optional sheet names
 */
function getOptionalSheetNames() {
  return Object.values(SCHEMA_REGISTRY)
    .filter(schema => !schema.required)
    .map(schema => schema.name);
}
