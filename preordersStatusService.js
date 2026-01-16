/**
 * Preorder Status Service
 * @fileoverview Backend service for View Preorder Status dashboard
 */

// ============================================================================
// PREORDER STATUS SERVICE
// ============================================================================

/**
 * Gets open preorders for the UI sidebar (called by preorder_status.html)
 * @param {string} searchFilter - Optional search text (name, set, item)
 * @return {Object} { pendingPayment, readyForPickup, summary }
 */
function getOpenPreorders(searchFilter) {
  return getPreordersByStatusCanonical_({ search: searchFilter || '' });
}

/**
 * Returns preorders grouped by status buckets for the status sidebar.
 * Uses canonical helpers from utils.js for robust column/sheet resolution.
 * @param {Object} filters (optional)
 *   - search: string | null   (search text; matches name, set, item)
 *   - includeCancelled: boolean (default false)
 *   - sortBy: string | null ('Created_At' | 'Target_Payoff' | 'Balance_Due')
 *   - sortDir: 'asc' | 'desc' (default 'asc')
 * @return {Object} {
 *   pendingPayment: Array<Object>,
 *   readyForPickup: Array<Object>,
 *   completed: Array<Object>,
 *   summary: {
 *     pendingPaymentCount: number,
 *     readyForPickupCount: number,
 *     completedCount: number,
 *     totalBalanceDue: number
 *   }
 * }
 */
function getPreordersByStatusCanonical_(filters) {
  filters = filters || {};

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Use canonical sheet lookup with aliases
  const sheet = getPreordersSheet_(ss);

  if (!sheet) {
    return buildEmptyPreorderResult_();
  }

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return buildEmptyPreorderResult_();
  }

  // Use canonical column resolver
  const headers = data[0];
  const cols = resolvePreordersCols_(headers);

  // Parse all rows into preorder objects
  const allPreorders = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const preorder = parsePreorderRowCanonical_(row, i + 1, cols, headers);

    // Skip empty rows
    if (!preorder.Preorder_ID && !preorder.PreferredName) {
      continue;
    }

    allPreorders.push(preorder);
  }

  // Apply search filter
  let filteredPreorders = allPreorders;
  if (filters.search && filters.search.trim()) {
    const searchLower = filters.search.trim().toLowerCase();
    filteredPreorders = allPreorders.filter(p => {
      const name = (p.PreferredName || '').toLowerCase();
      const setName = (p.Set_Name || '').toLowerCase();
      const itemName = (p.Item_Name || '').toLowerCase();
      return name.includes(searchLower) ||
             setName.includes(searchLower) ||
             itemName.includes(searchLower);
    });
  }

  // Group into buckets using canonical open/closed detection
  const pendingPayment = [];
  const readyForPickup = [];
  const completed = [];
  let totalBalanceDue = 0;

  for (const preorder of filteredPreorders) {
    const balance = coerceNumber(preorder.Balance_Due, 0);
    const isOpen = isPreorderOpen_(preorder.Picked_Up, preorder.Status);

    if (!isOpen) {
      // Closed preorders
      if (filters.includeCancelled) {
        completed.push(preorder);
      }
      continue;
    }

    // Track total balance due for open preorders
    if (balance > 0) {
      totalBalanceDue += balance;
      pendingPayment.push(preorder);
    } else {
      // Balance == 0 and open → ready for pickup
      readyForPickup.push(preorder);
    }
  }

  // Sort each bucket
  const sortBy = filters.sortBy || 'Created_At';
  const sortDir = filters.sortDir || 'asc';

  sortPreorderBucket_(pendingPayment, sortBy, sortDir);
  sortPreorderBucket_(readyForPickup, sortBy, sortDir);
  sortPreorderBucket_(completed, sortBy, sortDir);

  return {
    pendingPayment: pendingPayment,
    readyForPickup: readyForPickup,
    completed: completed,
    summary: {
      pendingPaymentCount: pendingPayment.length,
      readyForPickupCount: readyForPickup.length,
      completedCount: completed.length,
      totalBalanceDue: totalBalanceDue
    }
  };
}

/**
 * Legacy function - redirects to canonical implementation
 */
function getPreordersByStatus(filters) {
  return getPreordersByStatusCanonical_(filters);
}

// ============================================================================
// PRIVATE HELPERS
// ============================================================================

/**
 * Builds empty result object for preorders
 * @private
 */
function buildEmptyPreorderResult_() {
  return {
    pendingPayment: [],
    readyForPickup: [],
    completed: [],
    summary: {
      pendingPaymentCount: 0,
      readyForPickupCount: 0,
      completedCount: 0,
      totalBalanceDue: 0
    }
  };
}

/**
 * Parses a single row into a preorder object using canonical column resolver
 * @param {Array} row - Row data
 * @param {number} rowNumber - Sheet row number (1-based)
 * @param {Object} cols - Column index map from resolvePreordersCols_
 * @param {Array} headers - Header row for raw access
 * @return {Object} Preorder object
 * @private
 */
function parsePreorderRowCanonical_(row, rowNumber, cols, headers) {
  const getValue = (colKey) => {
    const idx = cols[colKey + 'Col'];
    return idx !== undefined && idx !== -1 ? row[idx] : undefined;
  };

  // Format dates for display
  const formatDate = (val) => {
    if (!val) return '';
    if (val instanceof Date) {
      try {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } catch (e) {
        return String(val);
      }
    }
    return String(val);
  };

  return {
    rowNumber: rowNumber,
    Preorder_ID: getValue('preorderId'),
    PreferredName: getValue('name'),  // Use canonical name field
    Contact_Info: getValue('contactInfo'),
    Set_Name: getValue('setName'),
    Item_Name: getValue('itemName'),
    Item_Code: getValue('itemCode'),
    Qty: getValue('qty'),
    Unit_Price: getValue('unitPrice'),
    Total_Due: getValue('totalDue'),
    Deposit_Paid: getValue('deposit'),
    Balance_Due: getValue('balanceDue'),
    Target_Payoff: formatDate(getValue('targetPayoff')),
    Status: getValue('status'),
    Picked_Up: getValue('pickedUp'),  // Include picked up status
    Notes: getValue('notes'),
    Created_At: formatDate(getValue('createdAt')),
    Created_By: getValue('createdBy')
  };
}

/**
 * Sorts a preorder bucket array by specified field
 * @param {Array} bucket - Array of preorder objects
 * @param {string} sortBy - Field to sort by
 * @param {string} sortDir - 'asc' or 'desc'
 * @private
 */
function sortPreorderBucket_(bucket, sortBy, sortDir) {
  const multiplier = sortDir === 'desc' ? -1 : 1;

  bucket.sort((a, b) => {
    let aVal = a[sortBy];
    let bVal = b[sortBy];

    // Handle missing values - push to end
    if (aVal === undefined || aVal === null || aVal === '') {
      return 1;
    }
    if (bVal === undefined || bVal === null || bVal === '') {
      return -1;
    }

    // Handle numeric fields
    if (sortBy === 'Balance_Due' || sortBy === 'Qty' || sortBy === 'Total_Due') {
      aVal = Number(aVal) || 0;
      bVal = Number(bVal) || 0;
      return (aVal - bVal) * multiplier;
    }

    // Handle date fields
    if (sortBy === 'Created_At' || sortBy === 'Target_Payoff') {
      const aDate = parsePreorderDate_(aVal);
      const bDate = parsePreorderDate_(bVal);

      if (!aDate) return 1;
      if (!bDate) return -1;

      return (aDate.getTime() - bDate.getTime()) * multiplier;
    }

    // String comparison fallback
    return String(aVal).localeCompare(String(bVal)) * multiplier;
  });
}

/**
 * Parses various date formats for preorders
 * @param {*} value - Date value (string or Date)
 * @return {Date|null} Parsed date or null
 * @private
 */
function parsePreorderDate_(value) {
  if (!value) return null;

  if (value instanceof Date) {
    return value;
  }

  // Try parsing string
  const parsed = new Date(value);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }

  return null;
}