/**
 * Preorder Status Service (Canonical)
 * @fileoverview Backend service for preorder tracking and status UI
 *
 * This is the CANONICAL preorder service. All preorder queries should go through here.
 * Uses helpers from utils.js: getPreordersSheet_, resolvePreordersCols_, isPreorderOpen_
 */

// ============================================================================
// CANONICAL API FUNCTION (Called by UI)
// ============================================================================

/**
 * Canonical API for preorder status queries
 * @param {Object} payload - Query parameters
 *   { mode: 'player', preferredName: 'Cy Diskin' }
 *   { mode: 'all_open' }
 *   { mode: 'search', search: 'search text' }
 * @return {Object} API response with orders grouped by Preorder_ID
 */
function api_getPreorderStatus(payload) {
  payload = payload || {};
  const mode = payload.mode || 'all_open';

  const response = {
    ok: true,
    mode: mode,
    generatedAt: new Date().toISOString(),
    sourceSheet: null,
    totals: {
      openOrders: 0,
      closedOrders: 0,
      openBalanceDue: 0,
      openTotalDue: 0,
      openDepositPaid: 0
    },
    orders: [],
    errors: []
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getPreordersSheet_(ss);

    if (!sheet) {
      response.ok = false;
      response.errors.push('Preorders_Sold sheet not found. Tried aliases: Preorders_Sold, Preorders Sold, Preorders');
      return response;
    }

    response.sourceSheet = sheet.getName();

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      response.errors.push('Sheet is empty (no data rows)');
      return response;
    }

    const headers = data[0];
    const cols = resolvePreordersCols_(headers);

    // Validate required columns
    if (cols.nameCol === -1) {
      response.errors.push('PreferredName column not found in headers: ' + headers.join(', '));
    }
    if (cols.preorderIdCol === -1) {
      response.errors.push('Preorder_ID column not found');
    }

    // Build grouped orders from rows
    const orderMap = {}; // Preorder_ID -> order object

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const lineItem = parsePreorderLineItem_(row, cols);

      // Skip empty rows
      if (!lineItem.preorderId && !lineItem.preferredName) {
        continue;
      }

      // Apply mode filter
      if (mode === 'player' && payload.preferredName) {
        const searchName = normalizeStr_(payload.preferredName);
        const rowName = normalizeStr_(lineItem.preferredName);
        if (rowName !== searchName) {
          continue;
        }
      } else if (mode === 'search' && payload.search) {
        const searchLower = normalizeStr_(payload.search);
        const matchesName = normalizeStr_(lineItem.preferredName).includes(searchLower);
        const matchesSet = normalizeStr_(lineItem.setName).includes(searchLower);
        const matchesItem = normalizeStr_(lineItem.itemName).includes(searchLower);
        if (!matchesName && !matchesSet && !matchesItem) {
          continue;
        }
      }

      // Get or create order in map
      const orderId = lineItem.preorderId || 'NO_ID_ROW_' + i;
      if (!orderMap[orderId]) {
        orderMap[orderId] = {
          preorderId: lineItem.preorderId,
          preferredName: lineItem.preferredName,
          contactInfo: lineItem.contactInfo,
          status: lineItem.status,
          pickedUp: lineItem.pickedUp,
          createdAt: lineItem.createdAt,
          updatedAt: lineItem.updatedAt,
          targetPayoff: lineItem.targetPayoff,
          notes: lineItem.notes,
          totalDue: 0,
          depositPaid: 0,
          balanceDue: 0,
          items: [],
          isOpen: isPreorderOpen_(lineItem.pickedUp, lineItem.status),
          rowNumbers: []
        };
      }

      const order = orderMap[orderId];

      // Add line item to order
      order.items.push({
        setName: lineItem.setName,
        itemName: lineItem.itemName,
        itemCode: lineItem.itemCode,
        qty: lineItem.qty,
        unitPrice: lineItem.unitPrice,
        lineTotal: lineItem.lineTotal,
        rowNumber: i + 1
      });

      order.rowNumbers.push(i + 1);

      // Aggregate totals (only count once per order, not per line)
      // For multi-line orders, the Balance_Due might be duplicated or at the order level
      // We'll take the max values to be safe
      order.totalDue = Math.max(order.totalDue, coerceNum_(lineItem.totalDue));
      order.depositPaid = Math.max(order.depositPaid, coerceNum_(lineItem.depositPaid));
      order.balanceDue = Math.max(order.balanceDue, coerceNum_(lineItem.balanceDue));
    }

    // Convert map to array and filter by open/closed for all_open mode
    const allOrders = Object.values(orderMap);

    for (const order of allOrders) {
      if (mode === 'all_open' && !order.isOpen) {
        response.totals.closedOrders++;
        continue;
      }

      if (order.isOpen) {
        response.totals.openOrders++;
        response.totals.openBalanceDue += order.balanceDue;
        response.totals.openTotalDue += order.totalDue;
        response.totals.openDepositPaid += order.depositPaid;
      } else {
        response.totals.closedOrders++;
      }

      response.orders.push(order);
    }

    // Sort orders by created date (newest first)
    response.orders.sort((a, b) => {
      if (!a.createdAt) return 1;
      if (!b.createdAt) return -1;
      return String(b.createdAt).localeCompare(String(a.createdAt));
    });

  } catch (e) {
    response.ok = false;
    response.errors.push('Error: ' + e.message);
    Logger.log('api_getPreorderStatus error: ' + e.message + '\n' + e.stack);
  }

  return response;
}

/**
 * Gets preorders for a specific player (wrapper for api_getPreorderStatus)
 * @param {string} preferredName - Player name to lookup
 * @return {Object} { open: [...], closed: [...], totals: {...} }
 */
function getPreordersForPlayer_(preferredName) {
  const result = {
    open: [],
    closed: [],
    totals: {
      openCount: 0,
      closedCount: 0,
      openBalanceDue: 0,
      openTotalDue: 0,
      openDepositPaid: 0
    },
    _debug: {
      sourceSheet: null,
      searchedName: preferredName,
      errors: []
    }
  };

  if (!preferredName) {
    result._debug.errors.push('No preferredName provided');
    return result;
  }

  const response = api_getPreorderStatus({ mode: 'player', preferredName: preferredName });

  result._debug.sourceSheet = response.sourceSheet;
  result._debug.errors = response.errors;

  for (const order of response.orders) {
    if (order.isOpen) {
      result.open.push(order);
      result.totals.openCount++;
      result.totals.openBalanceDue += order.balanceDue;
      result.totals.openTotalDue += order.totalDue;
      result.totals.openDepositPaid += order.depositPaid;
    } else {
      result.closed.push(order);
      result.totals.closedCount++;
    }
  }

  return result;
}

// ============================================================================
// LEGACY API (for backward compatibility)
// ============================================================================

/**
 * Gets open preorders for the UI sidebar (called by preorder_status.html)
 * This is the LEGACY function - use api_getPreorderStatus for new code
 * @param {string} searchFilter - Optional search text
 * @return {Object} { pendingPayment, readyForPickup, summary }
 */
function getOpenPreorders(searchFilter) {
  const mode = searchFilter ? 'search' : 'all_open';
  const response = api_getPreorderStatus({ mode: mode, search: searchFilter });

  // Convert to legacy format expected by old UI
  const pendingPayment = [];
  const readyForPickup = [];

  for (const order of response.orders) {
    if (!order.isOpen) continue;

    // Flatten order to legacy row format for each item (or just the order)
    const legacyRow = {
      rowNumber: order.rowNumbers[0],
      Preorder_ID: order.preorderId,
      PreferredName: order.preferredName,
      Set_Name: order.items[0]?.setName || '',
      Item_Name: order.items[0]?.itemName || '',
      Item_Code: order.items[0]?.itemCode || '',
      Qty: order.items.reduce((sum, item) => sum + (item.qty || 1), 0),
      Total_Due: order.totalDue,
      Deposit_Paid: order.depositPaid,
      Balance_Due: order.balanceDue,
      Status: order.status,
      Picked_Up: order.pickedUp,
      Created_At: order.createdAt,
      itemCount: order.items.length,
      items: order.items
    };

    if (order.balanceDue > 0) {
      pendingPayment.push(legacyRow);
    } else {
      readyForPickup.push(legacyRow);
    }
  }

  return {
    pendingPayment: pendingPayment,
    readyForPickup: readyForPickup,
    completed: [],
    summary: {
      pendingPaymentCount: pendingPayment.length,
      readyForPickupCount: readyForPickup.length,
      completedCount: 0,
      totalBalanceDue: response.totals.openBalanceDue
    }
  };
}

/**
 * Legacy function - redirects to getOpenPreorders
 */
function getPreordersByStatus(filters) {
  filters = filters || {};
  return getOpenPreorders(filters.search || '');
}

// ============================================================================
// PRIVATE HELPERS
// ============================================================================

/**
 * Parses a single row into a line item object
 * @param {Array} row - Row data
 * @param {Object} cols - Column index map from resolvePreordersCols_
 * @return {Object} Line item object
 * @private
 */
function parsePreorderLineItem_(row, cols) {
  const getValue = (colKey) => {
    const idx = cols[colKey + 'Col'];
    return idx !== undefined && idx !== -1 ? row[idx] : undefined;
  };

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
    preorderId: getValue('preorderId'),
    preferredName: getValue('name'),
    contactInfo: getValue('contactInfo'),
    setName: getValue('setName'),
    itemName: getValue('itemName'),
    itemCode: getValue('itemCode'),
    qty: coerceNum_(getValue('qty'), 1),
    unitPrice: coerceNum_(getValue('unitPrice'), 0),
    lineTotal: coerceNum_(getValue('lineTotal'), 0),
    totalDue: getValue('totalDue'),
    depositPaid: getValue('deposit'),
    balanceDue: getValue('balanceDue'),
    targetPayoff: formatDate(getValue('targetPayoff') || getValue('targetPayoffDate')),
    status: getValue('status') || 'Active',
    pickedUp: getValue('pickedUp'),
    notes: getValue('notes'),
    createdAt: formatDate(getValue('createdAt')),
    updatedAt: formatDate(getValue('updatedAt'))
  };
}

/**
 * Normalizes string for comparison (trim + lowercase)
 * @param {*} s - Value to normalize
 * @return {string} Normalized string
 * @private
 */
function normalizeStr_(s) {
  return String(s || '').toLowerCase().trim();
}

/**
 * Coerces value to number with default
 * @param {*} v - Value to coerce
 * @param {number} def - Default value
 * @return {number} Numeric value
 * @private
 */
function coerceNum_(v, def) {
  if (def === undefined) def = 0;
  if (v === undefined || v === null || v === '') return def;
  const num = Number(v);
  return isNaN(num) ? def : num;
}

/**
 * Coerces value to boolean
 * @param {*} v - Value to coerce
 * @return {boolean} Boolean value
 * @private
 */
function coerceBool_(v) {
  if (v === true) return true;
  if (!v) return false;
  const s = String(v).toLowerCase().trim();
  return s === 'true' || s === 'yes' || s === 'y' || s === '1';
}

// ============================================================================
// DEBUG FUNCTIONS
// ============================================================================

/**
 * Debug function for preorder status UI
 * Tests api_getPreorderStatus and logs results
 * @param {string} preferredName - Player name (default: "Cy Diskin")
 * @return {Object} API response
 */
function debug_preorderStatusUI(preferredName) {
  preferredName = preferredName || 'Cy Diskin';

  Logger.log('=== DEBUG: Preorder Status UI ===');
  Logger.log('Testing api_getPreorderStatus for: ' + preferredName);
  Logger.log('');

  // Test player mode
  const playerResult = api_getPreorderStatus({ mode: 'player', preferredName: preferredName });

  Logger.log('--- Player Mode Result ---');
  Logger.log('ok: ' + playerResult.ok);
  Logger.log('sourceSheet: ' + playerResult.sourceSheet);
  Logger.log('errors: ' + JSON.stringify(playerResult.errors));
  Logger.log('totals: ' + JSON.stringify(playerResult.totals));
  Logger.log('orders count: ' + playerResult.orders.length);
  Logger.log('');

  if (playerResult.orders.length > 0) {
    Logger.log('--- First Order ---');
    Logger.log(JSON.stringify(playerResult.orders[0], null, 2));
  }

  Logger.log('');
  Logger.log('--- Full Response ---');
  Logger.log(JSON.stringify(playerResult, null, 2));

  return playerResult;
}

/**
 * Debug function: Shows all preorders for a player name
 * @param {string} name - Player name (default: "Cy Diskin")
 * @return {Object} Debug info
 */
function debug_preorders(name) {
  name = name || 'Cy Diskin';

  Logger.log('=== DEBUG: Preorders for "' + name + '" ===');
  Logger.log('');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nameLower = normalizeStr_(name);

  const result = {
    query: name,
    sheet: null,
    columns: {},
    matchingRows: [],
    openCount: 0,
    closedCount: 0,
    groupedOrders: []
  };

  // Use canonical sheet lookup
  const sheet = getPreordersSheet_(ss);
  if (!sheet) {
    Logger.log('ERROR: Preorders sheet NOT FOUND');
    Logger.log('Tried aliases: Preorders_Sold, Preorders Sold, Preorders');
    return result;
  }

  result.sheet = sheet.getName();
  Logger.log('Sheet found: ' + result.sheet);
  Logger.log('Row count: ' + sheet.getLastRow());

  if (sheet.getLastRow() <= 1) {
    Logger.log('Sheet is empty (no data rows)');
    return result;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  Logger.log('Headers: ' + headers.join(', '));
  Logger.log('');

  // Use canonical column resolver
  const cols = resolvePreordersCols_(headers);

  // Log resolved columns
  Logger.log('--- Resolved Column Indices ---');
  const colKeys = ['name', 'status', 'pickedUp', 'preorderId', 'setName', 'itemName', 'balanceDue'];
  for (const key of colKeys) {
    const idx = cols[key + 'Col'];
    Logger.log(key + 'Col: ' + idx + ' → ' + (idx !== -1 ? headers[idx] : 'NOT FOUND'));
  }
  Logger.log('');

  result.columns = {
    name: cols.nameCol !== -1 ? headers[cols.nameCol] : 'NOT FOUND',
    status: cols.statusCol !== -1 ? headers[cols.statusCol] : 'NOT FOUND',
    pickedUp: cols.pickedUpCol !== -1 ? headers[cols.pickedUpCol] : 'NOT FOUND',
    preorderId: cols.preorderIdCol !== -1 ? headers[cols.preorderIdCol] : 'NOT FOUND'
  };

  Logger.log('--- Matching Rows ---');

  // Scan for matches
  let logCount = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowName = cols.nameCol !== -1 ? normalizeStr_(row[cols.nameCol]) : '';

    if (rowName !== nameLower) continue;

    const preorderId = cols.preorderIdCol !== -1 ? row[cols.preorderIdCol] : '';
    const statusValue = cols.statusCol !== -1 ? row[cols.statusCol] : '';
    const pickedUpValue = cols.pickedUpCol !== -1 ? row[cols.pickedUpCol] : '';
    const balanceDue = cols.balanceDueCol !== -1 ? coerceNum_(row[cols.balanceDueCol], 0) : 0;
    const setName = cols.setNameCol !== -1 ? row[cols.setNameCol] : '';
    const itemName = cols.itemNameCol !== -1 ? row[cols.itemNameCol] : '';

    const isOpen = isPreorderOpen_(pickedUpValue, statusValue);
    const normalizedStatus = normalizePreorderStatus_(statusValue);

    const rowInfo = {
      row: i + 1,
      preorderId: preorderId,
      setName: setName,
      itemName: itemName,
      status: statusValue,
      statusNormalized: normalizedStatus,
      pickedUpRaw: pickedUpValue,
      pickedUpParsed: isPreorderPickedUp_(pickedUpValue),
      isOpen: isOpen,
      balanceDue: balanceDue
    };

    result.matchingRows.push(rowInfo);

    if (isOpen) {
      result.openCount++;
    } else {
      result.closedCount++;
    }

    // Log first 10 rows
    if (logCount < 10) {
      Logger.log('Row ' + (i + 1) + ': ' + preorderId);
      Logger.log('  Set/Item: ' + setName + ' / ' + itemName);
      Logger.log('  Status: "' + statusValue + '" → "' + normalizedStatus + '"');
      Logger.log('  Picked_Up?: ' + JSON.stringify(pickedUpValue) + ' → ' + isPreorderPickedUp_(pickedUpValue));
      Logger.log('  Balance Due: $' + balanceDue);
      Logger.log('  → ' + (isOpen ? 'OPEN' : 'CLOSED'));
      Logger.log('');
      logCount++;
    }
  }

  Logger.log('=== SUMMARY ===');
  Logger.log('Total matching rows: ' + result.matchingRows.length);
  Logger.log('Open: ' + result.openCount);
  Logger.log('Closed: ' + result.closedCount);

  // Get grouped orders
  const playerData = getPreordersForPlayer_(name);
  result.groupedOrders = playerData.open;

  Logger.log('');
  Logger.log('Grouped open orders: ' + playerData.open.length);
  if (playerData.open.length > 0) {
    Logger.log('First grouped order:');
    Logger.log(JSON.stringify(playerData.open[0], null, 2));
  }

  return result;
}

/**
 * Debug function: Summary of all open preorders
 * @return {Object} Summary info
 */
function debug_preordersOpenSummary() {
  Logger.log('=== DEBUG: Open Preorders Summary ===');
  Logger.log('');

  const response = api_getPreorderStatus({ mode: 'all_open' });

  Logger.log('sourceSheet: ' + response.sourceSheet);
  Logger.log('ok: ' + response.ok);
  Logger.log('errors: ' + JSON.stringify(response.errors));
  Logger.log('');
  Logger.log('Totals:');
  Logger.log('  Open Orders: ' + response.totals.openOrders);
  Logger.log('  Closed Orders: ' + response.totals.closedOrders);
  Logger.log('  Open Balance Due: $' + response.totals.openBalanceDue.toFixed(2));
  Logger.log('');

  // Group by player
  const byPlayer = {};
  for (const order of response.orders) {
    const name = order.preferredName || 'Unknown';
    if (!byPlayer[name]) {
      byPlayer[name] = { count: 0, balance: 0 };
    }
    byPlayer[name].count++;
    byPlayer[name].balance += order.balanceDue;
  }

  Logger.log('Players with open orders: ' + Object.keys(byPlayer).length);

  // Log top 10
  const sorted = Object.entries(byPlayer)
    .sort((a, b) => b[1].count - a[1].count)
    .slice(0, 10);

  Logger.log('');
  Logger.log('--- Top 10 Players by Open Order Count ---');
  for (const [player, info] of sorted) {
    Logger.log(player + ': ' + info.count + ' orders ($' + info.balance.toFixed(2) + ' due)');
  }

  return {
    sourceSheet: response.sourceSheet,
    totals: response.totals,
    orderCount: response.orders.length,
    byPlayer: byPlayer
  };
}
