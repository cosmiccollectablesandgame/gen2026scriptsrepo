/**
 * Utility Functions for Cosmic Event Manager
 * @fileoverview PRNG, hashing, validation, and data transformation utilities
 */

// ============================================================================
// PRNG & SEEDING
// ============================================================================

/**
 * Generates a 10-character base62 alphanumeric seed
 * @return {string} 10-char seed
 */
function generateSeed() {
  const chars = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
  let seed = '';
  for (let i = 0; i < 10; i++) {
    seed += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return seed;
}

/**
 * Seeded PRNG using simple LCG (Linear Congruential Generator)
 * @param {string} seed - Alphanumeric seed string
 * @return {Function} PRNG function that returns 0-1
 */
function createSeededRandom(seed) {
  // Convert seed to numeric hash
  let hash = 0;
  for (let i = 0; i < seed.length; i++) {
    hash = ((hash << 5) - hash) + seed.charCodeAt(i);
    hash = hash & hash; // Convert to 32-bit integer
  }

  // LCG parameters
  let state = Math.abs(hash);

  return function() {
    state = (state * 1664525 + 1013904223) % 4294967296;
    return state / 4294967296;
  };
}

/**
 * Shuffles array using seeded random
 * @param {Array} array - Array to shuffle
 * @param {Function} rng - Seeded random function
 * @return {Array} Shuffled array
 */
function seededShuffle(array, rng) {
  const arr = [...array];
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(rng() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

// ============================================================================
// HASHING & CHECKSUMS
// ============================================================================

/**
 * Computes SHA-256 hash of string
 * @param {string} input - String to hash
 * @return {string} Hex hash
 */
function sha256(input) {
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    input,
    Utilities.Charset.UTF_8
  );

  return rawHash.map(byte => {
    const v = (byte < 0) ? 256 + byte : byte;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

/**
 * Computes checksum for range data
 * @param {Array<Array>} rangeValues - 2D array from getValues()
 * @return {string} Checksum hash
 */
function computeChecksum(rangeValues) {
  const serialized = JSON.stringify(rangeValues);
  return sha256(serialized).substring(0, 12); // First 12 chars
}

/**
 * Computes hash for preview object (deterministic)
 * @param {Object} obj - Object to hash
 * @return {string} Hash string
 */
function computeHash(obj) {
  // Stringify with sorted keys for determinism
  const serialized = JSON.stringify(obj, Object.keys(obj).sort());
  return sha256(serialized).substring(0, 16);
}

// ============================================================================
// DATE & TIME
// ============================================================================

/**
 * Returns current timestamp in ISO format
 * @return {string} ISO timestamp
 */
function dateISO() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd'T'HH:mm:ss'Z'"
  );
}

/**
 * Formats date as MM-DD-YYYY
 * @param {Date} date - Date object
 * @return {string} Formatted date
 */
function formatEventDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM-dd-yyyy');
}

/**
 * Parses ISO date string
 * @param {string} isoString - ISO date string
 * @return {Date} Date object
 */
function parseISODate(isoString) {
  return new Date(isoString);
}

// ============================================================================
// DATA PARSING & COERCION
// ============================================================================

/**
 * Parses CSV text
 * @param {string} csvText - CSV content
 * @param {string} delimiter - Delimiter (default: ',')
 * @return {Array<Array<string>>} Parsed rows
 */
function parseCSV(csvText, delimiter = ',') {
  const rows = [];
  const lines = csvText.split(/\r?\n/);

  for (const line of lines) {
    if (!line.trim()) continue;

    const row = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      const next = line[i + 1];

      if (char === '"') {
        if (inQuotes && next === '"') {
          current += '"';
          i++; // Skip next quote
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === delimiter && !inQuotes) {
        row.push(current);
        current = '';
      } else {
        current += char;
      }
    }
    row.push(current);
    rows.push(row);
  }

  return rows;
}

/**
 * Converts 2D array to key-value object (2-column format)
 * @param {Array<Array>} data - 2D array [[key, value], ...]
 * @return {Object} KV object
 */
function toKV(data) {
  const kv = {};
  for (let i = 1; i < data.length; i++) { // Skip header
    const [key, value] = data[i];
    if (key) kv[key] = value;
  }
  return kv;
}

/**
 * Converts 2D array to array of objects using first row as headers
 * @param {Array<Array>} data - 2D array with header row
 * @return {Array<Object>} Array of row objects
 */
function toObjects(data) {
  if (data.length === 0) return [];
  const headers = data[0];
  const objects = [];

  for (let i = 1; i < data.length; i++) {
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    objects.push(obj);
  }

  return objects;
}

/**
 * Converts array of objects to Map keyed by specified field
 * @param {Array<Object>} objects - Array of objects
 * @param {string} keyField - Field to use as key
 * @return {Map} Map of objects
 */
function toMapByKey(objects, keyField) {
  const map = new Map();
  objects.forEach(obj => {
    if (obj[keyField] !== undefined) {
      map.set(obj[keyField], obj);
    }
  });
  return map;
}

// ============================================================================
// VALIDATION & COERCION
// ============================================================================

/**
 * Coerces value to number, returns default if invalid
 * @param {*} value - Value to coerce
 * @param {number} defaultValue - Default if invalid
 * @return {number} Coerced number
 */
function coerceNumber(value, defaultValue = 0) {
  const num = Number(value);
  return isNaN(num) ? defaultValue : num;
}

/**
 * Coerces value to boolean
 * @param {*} value - Value to coerce
 * @return {boolean} Boolean value
 */
function coerceBoolean(value) {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'string') {
    const upper = value.toUpperCase();
    return upper === 'TRUE' || upper === 'YES' || upper === '1';
  }
  return Boolean(value);
}

/**
 * Validates and parses ratio string (e.g., "3:1")
 * @param {string} ratioStr - Ratio string "X:Y"
 * @return {Object} {num, den} or null if invalid
 */
function parseRatio(ratioStr) {
  if (!ratioStr || typeof ratioStr !== 'string') return null;

  const match = ratioStr.match(/^(\d+):(\d+)$/);
  if (!match) return null;

  const num = parseInt(match[1], 10);
  const den = parseInt(match[2], 10);

  if (den === 0) return null;

  return { num, den };
}

/**
 * Clamps number between min and max
 * @param {number} value - Value to clamp
 * @param {number} min - Minimum
 * @param {number} max - Maximum
 * @return {number} Clamped value
 */
function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

// ============================================================================
// SYNONYM MAPPING
// ============================================================================

/**
 * Synonym map for canonical header names
 * @return {Object} Map of synonym -> canonical
 */
function getHeaderSynonyms() {
  return {
    // Prize_Catalog
    'SKU': 'Code',
    'Product_Name': 'Name',
    'Product': 'Name',
    'Stock_Qty': 'Qty',
    'Quantity': 'Qty',
    'Retail_Value': 'EV_Cost',
    'Cost_Basis': 'COGS',
    'Cost': 'COGS',
    'L_Bucket': 'Level',
    'Tier': 'Level',
    'In_Stock': 'InStock',

    // Event tabs
    'Preferred_Name': 'PreferredName',
    'Player_Name': 'PreferredName',
    'Name': 'PreferredName',
    'Player': 'PreferredName',

    // Key_Tracker
    'Preferred_Name': 'PreferredName',
    'Rainbow_Eligible': 'RainbowEligible',
    'Last_Updated': 'LastUpdated',

    // BP_Total
    'BP': 'Current_BP',
    'Bonus_Points': 'Current_BP',
    'BP_Current': 'Current_BP'  // Legacy mapping
  };
}

/**
 * Normalizes header using synonym map
 * @param {string} header - Original header
 * @return {string} Canonical header
 */
function normalizeHeader(header) {
  const synonyms = getHeaderSynonyms();
  return synonyms[header] || header;
}

// ============================================================================
// ERROR HANDLING
// ============================================================================

/**
 * Throws formatted error with code
 * @param {string} message - Error message
 * @param {string} code - Error code (e.g., 'INVALID_INPUT')
 * @param {string} remediation - Suggested fix
 * @throws {Error} Formatted error
 */
function throwError(message, code, remediation = '') {
  const fullMessage = `[${code}] ${message}${remediation ? '\n\nFix: ' + remediation : ''}`;
  throw new Error(fullMessage);
}

// ============================================================================
// FORMATTING
// ============================================================================

/**
 * Formats number as currency
 * @param {number} value - Numeric value
 * @return {string} Formatted currency
 */
function formatCurrency(value) {
  return '$' + value.toFixed(2);
}

/**
 * Formats percentage
 * @param {number} value - Decimal value (0.95 = 95%)
 * @param {number} decimals - Decimal places
 * @return {string} Formatted percentage
 */
function formatPercent(value, decimals = 0) {
  return (value * 100).toFixed(decimals) + '%';
}

/**
 * Gets current user email
 * @return {string} User email
 */
function currentUser() {
  return Session.getActiveUser().getEmail() || 'unknown';
}

// ============================================================================
// ARRAY UTILITIES
// ============================================================================

/**
 * Groups array by key function
 * @param {Array} array - Array to group
 * @param {Function} keyFn - Function that returns group key
 * @return {Map} Map of key -> array of items
 */
function groupBy(array, keyFn) {
  const groups = new Map();
  array.forEach(item => {
    const key = keyFn(item);
    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key).push(item);
  });
  return groups;
}

/**
 * Sums array by value function
 * @param {Array} array - Array to sum
 * @param {Function} valueFn - Function that returns numeric value
 * @return {number} Sum
 */
function sumBy(array, valueFn) {
  return array.reduce((sum, item) => sum + valueFn(item), 0);
}

/**
 * Filters unique values
 * @param {Array} array - Array with duplicates
 * @return {Array} Unique values
 */
function unique(array) {
  return [...new Set(array)];
}

/**
 * Deep clones object/array
 * @param {*} obj - Object to clone
 * @return {*} Cloned object
 */
function deepClone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

// ============================================================================
// SHEET ALIAS MAPPING (Case-Insensitive)
// ============================================================================

/**
 * Sheet alias map - maps logical names to possible real sheet names
 * @return {Object} Map of logical name -> array of aliases
 */
function getSheetAliases_() {
  return {
    'DICE_POINTS': ['Dice Roll Points', 'Dice_Points'],
    'FLAG_MISSIONS': ['Flag_Missions', 'Flag_Points'],
    'ATTENDANCE_MISSIONS': ['Attendance_Missions', 'Attendance_Points'],
    'BP_TOTAL': ['BP_Total'],
    'KEY_TRACKER': ['Key_Tracker'],
    'PREORDERS': ['Preorders_Sold', 'Preorders'],
    'STORE_CREDIT_LEDGER': ['Store_Credit_Ledger'],
    'REDEEMED_BP': ['Redeemed_BP'],
    'PREFERRED_NAMES': ['PreferredNames', 'Preferred_Names']
  };
}

/**
 * Gets a sheet by name (case-insensitive)
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {string} name - Sheet name to find
 * @return {Sheet|null} Found sheet or null
 */
function getSheetByNameCI_(ss, name) {
  if (!ss || !name) return null;
  const nameLower = String(name).toLowerCase();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase() === nameLower) {
      return sheets[i];
    }
  }
  return null;
}

/**
 * Gets a sheet by trying multiple alias names (case-insensitive)
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {string[]} aliases - Array of possible sheet names to try
 * @return {Sheet|null} First found sheet or null
 */
function getSheetByAliasesCI_(ss, aliases) {
  if (!ss || !aliases || !aliases.length) return null;
  for (let i = 0; i < aliases.length; i++) {
    const sheet = getSheetByNameCI_(ss, aliases[i]);
    if (sheet) return sheet;
  }
  return null;
}

/**
 * Gets a sheet by logical name using alias map
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {string} logicalName - Logical name (e.g., 'DICE_POINTS')
 * @return {Sheet|null} Found sheet or null
 */
function getSheetByLogicalName_(ss, logicalName) {
  const aliasMap = getSheetAliases_();
  const aliases = aliasMap[logicalName];
  if (!aliases) {
    // Fallback: try the logical name itself
    return getSheetByNameCI_(ss, logicalName);
  }
  return getSheetByAliasesCI_(ss, aliases);
}

// ============================================================================
// COLUMN SYNONYM MAPPING (Header Resolution)
// ============================================================================

/**
 * Column synonym groups - each group represents the same logical field
 * @return {Object} Map of logical field -> array of possible header names
 */
function getColumnSynonyms_() {
  return {
    // Player name column
    'NAME': ['PreferredName', 'Preferred_Name', 'preferred_name_id', 'Player', 'Name', 'Customer_Name'],

    // BP columns
    'BP_CURRENT': ['BP_Current', 'Current_BP', 'Capped_BP', 'BP'],
    'BP_HISTORICAL': ['BP_Historical', 'Historical_BP', 'Historical', 'Lifetime_BP', 'Total_Earned'],
    'BP_REDEEMED': ['BP_Redeemed', 'Redeemed_BP', 'Total_Redeemed', 'Spent'],

    // BP source breakdown columns (EXACT names from real sheets)
    'ATTENDANCE_POINTS': ['Attendance Mission Points', 'Attendance_Mission_Points', 'AttendancePoints', 'Points'],
    'FLAG_POINTS': ['Flag Mission Points', 'Flag_Mission_Points', 'FlagPoints', 'Flag_Points'],
    'DICE_POINTS': ['Dice Roll Points', 'Dice_Points', 'DicePoints', 'Points'],

    // Last updated
    'LAST_UPDATED': ['LastUpdated', 'Last_Updated', 'LastUpdate', 'Updated_At'],

    // Store credit
    'AMOUNT': ['Amount', 'Balance', 'Credit', 'Total'],
    'INOUT': ['InOut', 'Type', 'Direction'],
    'RUNNING_BALANCE': ['RunningBalance', 'Running_Balance', 'Balance']
  };
}

/**
 * Finds column index using synonym mapping
 * @param {Array} headers - Header row array
 * @param {string|string[]} synonyms - Logical name or array of possible names
 * @return {number} Column index (0-based) or -1 if not found
 */
function findColumnBySynonym_(headers, synonyms) {
  if (!headers) return -1;

  // If string, look up in synonym map
  let namesToTry = synonyms;
  if (typeof synonyms === 'string') {
    const synonymMap = getColumnSynonyms_();
    namesToTry = synonymMap[synonyms] || [synonyms];
  }

  // Try each possible name
  for (let i = 0; i < namesToTry.length; i++) {
    const idx = headers.indexOf(namesToTry[i]);
    if (idx !== -1) return idx;
  }

  return -1;
}

/**
 * Creates a header map for a sheet's headers
 * Maps logical column names to actual indices
 * @param {Array} headers - Header row array
 * @return {Object} Map of logical name -> column index
 */
function createHeaderMap_(headers) {
  const synonymMap = getColumnSynonyms_();
  const result = {};

  // Map each logical name to its column index
  for (const logicalName in synonymMap) {
    result[logicalName] = findColumnBySynonym_(headers, synonymMap[logicalName]);
  }

  // Also store raw header positions
  result._raw = {};
  for (let i = 0; i < headers.length; i++) {
    result._raw[headers[i]] = i;
  }

  return result;
}

/**
 * Gets column value from a row using header map
 * @param {Array} row - Data row
 * @param {Object} headerMap - Header map from createHeaderMap_
 * @param {string} logicalName - Logical column name
 * @param {*} defaultValue - Default value if column not found
 * @return {*} Cell value or default
 */
function getColumnValue_(row, headerMap, logicalName, defaultValue) {
  const idx = headerMap[logicalName];
  if (idx === -1 || idx === undefined) return defaultValue;
  return row[idx] !== undefined ? row[idx] : defaultValue;
}

// ============================================================================
// PREORDERS ACCESS LAYER (Canonical)
// ============================================================================

/**
 * Gets the Preorders sheet using alias resolution (case-insensitive)
 * @param {Spreadsheet} ss - Spreadsheet (defaults to active)
 * @return {Sheet|null} The preorders sheet or null
 */
function getPreordersSheet_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  return getSheetByAliasesCI_(ss, ['Preorders_Sold', 'Preorders Sold', 'Preorders']);
}

/**
 * Resolves preorder column indices from headers with synonym support
 * @param {Array} headers - Header row
 * @return {Object} Column index map with all preorder fields
 */
function resolvePreordersCols_(headers) {
  // Define synonyms for each field
  const synonyms = {
    name: ['PreferredName', 'Preferred_Name', 'preferred_name_id', 'Customer_Name', 'Name', 'Player'],
    status: ['Status', 'status'],
    pickedUp: ['Picked_Up?', 'Picked Up?', 'Picked_Up', 'PickedUp?', 'PickedUp', 'Picked Up'],
    balanceDue: ['Balance_Due', 'BalanceDue', 'Balance Due'],
    deposit: ['Deposit_Paid', 'DepositPaid', 'Deposit Paid', 'Deposit'],
    totalDue: ['Total_Due', 'TotalDue', 'Total Due', 'Total'],
    setName: ['Set_Name', 'SetName', 'Set Name', 'Set'],
    itemName: ['Item_Name', 'ItemName', 'Item Name', 'Item'],
    qty: ['Qty', 'Quantity', 'qty'],
    preorderId: ['Preorder_ID', 'PreorderId', 'Preorder ID', 'ID'],
    itemCode: ['Item_Code', 'ItemCode', 'Item Code', 'Code'],
    unitPrice: ['Unit_Price', 'UnitPrice', 'Unit Price', 'Price'],
    lineTotal: ['Line_Total', 'LineTotal', 'Line Total'],
    targetPayoff: ['Target_Payoff', 'TargetPayoff', 'Target Payoff'],
    targetPayoffDate: ['Target_Payoff_Date', 'TargetPayoffDate', 'Target Payoff Date'],
    notes: ['Notes', 'Note', 'notes'],
    contactInfo: ['Contact_Info', 'ContactInfo', 'Contact Info', 'Contact'],
    createdAt: ['Created_At', 'CreatedAt', 'Created At', 'Created'],
    createdBy: ['Created_By', 'CreatedBy', 'Created By'],
    customerName: ['Customer_Name', 'CustomerName', 'Customer Name'],
    updatedAt: ['Updated_At', 'UpdatedAt', 'Updated At']
  };

  const result = {};
  for (const field in synonyms) {
    result[field + 'Col'] = findColumnIndex_(headers, synonyms[field]);
  }

  // Store raw header positions for direct access
  result._headers = headers;
  result._raw = {};
  for (let i = 0; i < headers.length; i++) {
    result._raw[headers[i]] = i;
  }

  return result;
}

/**
 * Checks if a preorder is considered "picked up" (closed)
 * @param {*} pickedUpValue - Value from Picked_Up? column
 * @return {boolean} True if picked up
 */
function isPreorderPickedUp_(pickedUpValue) {
  if (pickedUpValue === true) return true;
  if (!pickedUpValue) return false;

  const strVal = String(pickedUpValue).toLowerCase().trim();
  return strVal === 'true' || strVal === 'yes' || strVal === 'y' ||
         strVal === '1' || strVal === 'picked up' || strVal === 'picked';
}

/**
 * Normalizes preorder status for comparison
 * @param {*} status - Raw status value
 * @return {string} Lowercase trimmed status
 */
function normalizePreorderStatus_(status) {
  if (!status) return 'active';
  return String(status).toLowerCase().trim().replace(/_/g, ' ');
}

/**
 * Checks if a preorder status indicates it's closed/completed
 * @param {string} normalizedStatus - Normalized status string
 * @return {boolean} True if status indicates closed
 */
function isPreorderStatusClosed_(normalizedStatus) {
  const closedStatuses = [
    'completed', 'cancelled', 'canceled', 'picked up',
    'fulfilled', 'closed', 'refunded'
  ];
  return closedStatuses.includes(normalizedStatus);
}

/**
 * Determines if a preorder is OPEN (not picked up AND status not closed)
 * @param {*} pickedUpValue - Value from Picked_Up? column
 * @param {*} statusValue - Value from Status column
 * @return {boolean} True if preorder is open
 */
function isPreorderOpen_(pickedUpValue, statusValue) {
  // If picked up, it's closed
  if (isPreorderPickedUp_(pickedUpValue)) {
    return false;
  }

  // If status indicates closed, it's closed
  const normalizedStatus = normalizePreorderStatus_(statusValue);
  if (isPreorderStatusClosed_(normalizedStatus)) {
    return false;
  }

  // Otherwise, it's open (including deposit_paid, paid_in_full, active, etc.)
  return true;
}

// ============================================================================
// EVENT TAB DETECTION (CANONICAL - Case-Insensitive)
// ============================================================================

/**
 * Normalizes a tab name by trimming whitespace.
 * @param {*} s - Raw tab name (any type)
 * @return {string} Trimmed string
 */
function normalizeTabName_(s) {
  return String(s || '').trim();
}

/**
 * Normalizes a tab name to lowercase for case-insensitive matching.
 * @param {*} s - Raw tab name (any type)
 * @return {string} Trimmed lowercase string
 */
function normalizeTabNameLower_(s) {
  return normalizeTabName_(s).toLowerCase();
}

/**
 * Checks if a sheet name matches event tab patterns (CASE-INSENSITIVE).
 *
 * Patterns matched:
 * - MM-DD-YYYY (e.g., "11-26-2025")
 * - MM-DDx-YYYY with 1+ letter suffix (e.g., "11-26C-2025", "05-03c-2025", "07-26ZZ-2025")
 * - Legacy ABCnn_ prefix (e.g., "MTG01_", "CMD12_")
 *
 * @param {string} name - Sheet name to check
 * @return {boolean} True if matches event tab pattern
 */
function isEventTabName_(name) {
  var n = normalizeTabNameLower_(name);
  if (!n) return false;

  // Pattern 1: MM-DD-YYYY or MM-DDx-YYYY (x = one or more letters, case-insensitive)
  // Examples: 11-26-2025, 11-26c-2025, 05-03C-2025, 07-26zz-2025
  if (/^\d{2}-\d{2}([a-z]+)?-\d{4}$/.test(n)) {
    return true;
  }

  // Pattern 2: Legacy ABC12_ format (case-insensitive)
  // Examples: MTG01_, CMD12_, mtg01_
  if (/^[a-z]{3}\d{2}_/.test(n)) {
    return true;
  }

  return false;
}

/**
 * Lists all event tabs in a spreadsheet (CASE-INSENSITIVE matching).
 * Returns tab names in their original casing, sorted alphabetically.
 *
 * @param {Spreadsheet} ss - Spreadsheet to scan (defaults to active)
 * @return {Array<string>} Event tab names, sorted
 */
function listEventTabsCI_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()
    .map(function(s) { return s.getName(); })
    .filter(isEventTabName_)
    .sort();
}

/**
 * Debug function to test event tab detection.
 * Logs all detected event tabs to the Apps Script log.
 * Run this from the Script Editor to verify tab detection.
 */
function debug_listEventTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allTabs = ss.getSheets().map(function(s) { return s.getName(); });
  var eventTabs = listEventTabsCI_(ss);

  Logger.log('=== EVENT TAB DETECTION DEBUG ===');
  Logger.log('Total sheets: ' + allTabs.length);
  Logger.log('Event tabs detected: ' + eventTabs.length);
  Logger.log('');
  Logger.log('--- All sheets ---');
  allTabs.forEach(function(name) {
    var isEvent = isEventTabName_(name);
    Logger.log((isEvent ? '[EVENT] ' : '[     ] ') + name);
  });
  Logger.log('');
  Logger.log('--- Event tabs only ---');
  eventTabs.forEach(function(name) {
    Logger.log('  ' + name);
  });
  Logger.log('');

  // Test with known examples
  var testCases = ['04-19Z-2025', '05-03C-2025', '07-26q-2025', '11-29C-2025',
                   '11-26-2025', 'PreferredNames', 'BP_Total', 'MTG01_Legacy'];
  Logger.log('--- Test cases ---');
  testCases.forEach(function(name) {
    Logger.log((isEventTabName_(name) ? '[MATCH] ' : '[SKIP ] ') + name);
  });

  return eventTabs;
}