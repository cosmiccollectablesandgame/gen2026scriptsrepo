/**
 * Cosmic Games Tournament Manager
 * Main Entry Point - Code.gs
 *
 * CRITICAL: This is the ONLY file that contains onOpen() and onEdit()
 *
 * Requirements Met:
 * [x] Exactly one onOpen(e) and one onEdit(e)
 * [x] Ultra-lightweight onEdit (no BP sync, no heavy ops)
 * [x] Safe menu items via addMenuItemOrStub_()
 * [x] Case-insensitive sheet lookups everywhere
 * [x] Comprehensive Build/Repair with Ship Gates
 * [x] Player provisioning from event tabs
 * [x] BP_Total with historical/redeemed/current
 * [x] Player Lookup unified view
 */

// ============================================================================
// VERSION CONSTANTS
// ============================================================================

const ENGINE_VERSION = '7.9.9';
const RECIPE_VERSION = '1.0.2';

// ============================================================================
// COLUMN SYNONYM MAPS
// ============================================================================

const PLAYER_NAME_SYNONYMS = [
  'preferredname', 'preferred_name_id', 'player', 'name', 'playername', 'player_name'
];

const BP_COLUMN_SYNONYMS = {
  historical: ['bp_historical', 'historical_bp', 'total_earned', 'bp_earned', 'earned'],
  redeemed: ['bp_redeemed', 'redeemed_bp', 'spent', 'redeemed'],
  current: ['bp_current', 'current_bp', 'capped_bp', 'available', 'current']
};

// ============================================================================
// REQUIRED SHEET SCHEMAS
// ============================================================================

const REQUIRED_SHEETS = {
  PreferredNames: ['preferred_name_id', 'display_name', 'created_at', 'last_active', 'status'],
  BP_Total: ['preferred_name_id', 'BP_Historical', 'BP_Redeemed', 'BP_Current', 'Flag_Points', 'Attendance_Points', 'Dice_Points', 'LastUpdated'],
  Attendance_Missions: ['preferred_name_id', 'Total_Events', 'Attendance_Points', 'LastUpdated'],
  Flag_Missions: ['preferred_name_id', 'Cosmic_Selfie', 'Review_Writer', 'Social_Media_Star', 'App_Explorer', 'Cosmic_Merchant', 'Flag_Points', 'LastUpdated'],
  Dice_Points: ['preferred_name_id', 'Points', 'LastUpdated'],
  Attendance_Calendar: ['EventDate', 'EventName', 'Format', 'PlayerCount', 'Status'],
  Integrity_Log: ['Timestamp', 'Action', 'Details', 'Operator', 'Status'],
  Key_Tracker: ['PreferredName', 'White', 'Blue', 'Black', 'Red', 'Green', 'Total', 'Rainbow']
};

// ============================================================================
// MENU HELPER: Safe Item Addition
// ============================================================================

function addMenuItemOrStub_(menu, label, fnName, ui) {
  if (typeof globalThis[fnName] === 'function') {
    menu.addItem(label, fnName);
  } else {
    var stubName = '__stub_' + fnName;
    if (typeof globalThis[stubName] !== 'function') {
      globalThis[stubName] = function() {
        SpreadsheetApp.getUi().alert(
          'Feature Not Available',
          'Menu item "' + label + '" requires function "' + fnName + '" which is not defined.\n\n' +
          'Possible fixes:\n' +
          '• Search the project for "function ' + fnName + '"\n' +
          '• Ensure exactly one canonical definition exists\n' +
          '• This feature may not be implemented yet',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      };
    }
    menu.addItem(label + ' ⚠️', stubName);
  }
}

function addSep_(menu) {
  menu.addSeparator();
}

// ============================================================================
// CASE-INSENSITIVE HELPERS
// ============================================================================

function getSheetByNameCI_(ss, name) {
  var sheets = ss.getSheets();
  var lowerName = name.toLowerCase();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase() === lowerName) {
      return sheets[i];
    }
  }
  return null;
}

function listEventTabsCI_(ss) {
  var sheets = ss.getSheets();
  var eventTabs = [];
  var datePattern = /^\d{1,2}-\d{1,2}[a-zA-Z]*-\d{4}$/i;
  var legacyPattern = /^[A-Z]{2,4}\d{2}_/i;

  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (datePattern.test(name) || legacyPattern.test(name)) {
      eventTabs.push({ name: name, sheet: sheets[i] });
    }
  }
  return eventTabs;
}

function normalizePlayerName_(s) {
  if (!s) return '';
  return String(s).trim().replace(/\s+/g, ' ');
}

function playerNamesMatch_(a, b) {
  return normalizePlayerName_(a).toLowerCase() === normalizePlayerName_(b).toLowerCase();
}

// ============================================================================
// HEADER MAPPING HELPERS
// ============================================================================

function getHeaderMap_(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return {};
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).toLowerCase().replace(/[_\s-]/g, '');
    map[h] = i;
    map[String(headers[i])] = i;
  }
  return map;
}

function findPlayerColumn_(headerMap) {
  for (var i = 0; i < PLAYER_NAME_SYNONYMS.length; i++) {
    var syn = PLAYER_NAME_SYNONYMS[i].toLowerCase().replace(/[_\s-]/g, '');
    if (headerMap[syn] !== undefined) return headerMap[syn];
  }
  return 0;
}

function findColumnBySynonyms_(headerMap, synonyms) {
  for (var i = 0; i < synonyms.length; i++) {
    var syn = synonyms[i].toLowerCase().replace(/[_\s-]/g, '');
    if (headerMap[syn] !== undefined) return headerMap[syn];
  }
  return -1;
}

// ============================================================================
// SHEET MANAGEMENT HELPERS
// ============================================================================

function ensureSheetWithHeaders_(ss, sheetName, headers) {
  var sheet = getSheetByNameCI_(ss, sheetName);
  var created = false;
  var headersAdded = false;

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    created = true;
  }

  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    headersAdded = true;
  } else {
    var existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var existingLower = existing.map(function(h) { return String(h).toLowerCase(); });
    var missing = [];
    for (var i = 0; i < headers.length; i++) {
      if (existingLower.indexOf(headers[i].toLowerCase()) === -1) {
        missing.push(headers[i]);
      }
    }
    if (missing.length > 0) {
      var startCol = sheet.getLastColumn() + 1;
      for (var j = 0; j < missing.length; j++) {
        sheet.getRange(1, startCol + j).setValue(missing[j]).setFontWeight('bold');
      }
      headersAdded = true;
    }
  }

  return { sheet: sheet, created: created, headersAdded: headersAdded };
}

// ============================================================================
// PLAYER MANAGEMENT HELPERS
// ============================================================================

function getAllCanonicalPlayers_(ss) {
  var sheet = getSheetByNameCI_(ss, 'PreferredNames');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var headerMap = {};
  for (var i = 0; i < data[0].length; i++) {
    headerMap[String(data[0][i]).toLowerCase()] = i;
  }

  var nameCol = findPlayerColumn_(headerMap);
  var players = [];

  for (var r = 1; r < data.length; r++) {
    var name = normalizePlayerName_(data[r][nameCol]);
    if (name) players.push(name);
  }
  return players;
}

function playerExistsInSheet_(sheet, playerName) {
  if (!sheet || sheet.getLastRow() <= 1) return false;

  var data = sheet.getDataRange().getValues();
  var headerMap = getHeaderMap_(sheet);
  var nameCol = findPlayerColumn_(headerMap);

  var normalizedTarget = normalizePlayerName_(playerName).toLowerCase();

  for (var r = 1; r < data.length; r++) {
    if (normalizePlayerName_(data[r][nameCol]).toLowerCase() === normalizedTarget) {
      return true;
    }
  }
  return false;
}

function addPlayerToSheet_(sheet, playerName, defaultValues) {
  if (!sheet) return false;
  if (playerExistsInSheet_(sheet, playerName)) return false;

  var headerMap = getHeaderMap_(sheet);
  var numCols = sheet.getLastColumn() || 1;
  var row = new Array(numCols).fill('');

  var nameCol = findPlayerColumn_(headerMap);
  row[nameCol] = playerName;

  if (defaultValues) {
    for (var key in defaultValues) {
      var keyLower = key.toLowerCase().replace(/[_\s-]/g, '');
      if (headerMap[keyLower] !== undefined) {
        row[headerMap[keyLower]] = defaultValues[key];
      }
    }
  }

  var lastUpdatedCol = headerMap['lastupdated'];
  if (lastUpdatedCol !== undefined) {
    row[lastUpdatedCol] = new Date().toISOString();
  }

  sheet.appendRow(row);
  return true;
}

function provisionPlayerEverywhere_(ss, canonicalName) {
  var results = { provisioned: [], alreadyExisted: [], failed: [] };

  var sheetsToProvision = [
    { name: 'PreferredNames', defaults: { status: 'ACTIVE', created_at: new Date().toISOString() } },
    { name: 'BP_Total', defaults: { BP_Historical: 0, BP_Redeemed: 0, BP_Current: 0, Flag_Points: 0, Attendance_Points: 0, Dice_Points: 0 } },
    { name: 'Attendance_Missions', defaults: { Total_Events: 0, Attendance_Points: 0 } },
    { name: 'Flag_Missions', defaults: { Cosmic_Selfie: false, Review_Writer: false, Social_Media_Star: false, Flag_Points: 0 } },
    { name: 'Dice_Points', defaults: { Points: 0 } },
    { name: 'Key_Tracker', defaults: { White: 0, Blue: 0, Black: 0, Red: 0, Green: 0, Total: 0, Rainbow: false } }
  ];

  for (var i = 0; i < sheetsToProvision.length; i++) {
    var config = sheetsToProvision[i];
    var sheet = getSheetByNameCI_(ss, config.name);

    if (!sheet) {
      results.failed.push(config.name + ' (sheet missing)');
      continue;
    }

    if (playerExistsInSheet_(sheet, canonicalName)) {
      results.alreadyExisted.push(config.name);
    } else {
      if (addPlayerToSheet_(sheet, canonicalName, config.defaults)) {
        results.provisioned.push(config.name);
      } else {
        results.failed.push(config.name);
      }
    }
  }

  return results;
}

// ============================================================================
// SHIP GATES SYSTEM
// ============================================================================

function runShipGates_(mode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var report = {
    mode: mode,
    timestamp: new Date().toISOString(),
    gates: [],
    summary: { passed: 0, failed: 0, fixed: 0 },
    categories: {
      Tabs: [],
      Headers: [],
      DataIntegrity: [],
      Players: [],
      Indexing: [],
      UIWiring: []
    }
  };

  var gates = [
    { id: 'A', name: 'Required Sheets Exist', category: 'Tabs', check: checkRequiredSheets_, fix: fixRequiredSheets_ },
    { id: 'B', name: 'Sheet Headers Valid', category: 'Headers', check: checkSheetHeaders_, fix: fixSheetHeaders_ },
    { id: 'C', name: 'Event Tabs Detectable', category: 'Tabs', check: checkEventTabs_, fix: null },
    { id: 'D', name: 'Players Provisioned', category: 'Players', check: checkPlayersProvisioned_, fix: fixPlayersProvisioned_ },
    { id: 'E', name: 'BP_Total Schema Valid', category: 'DataIntegrity', check: checkBPTotalSchema_, fix: fixBPTotalSchema_ },
    { id: 'F', name: 'Attendance Calendar Built', category: 'Indexing', check: checkAttendanceCalendar_, fix: fixAttendanceCalendar_ },
    { id: 'G', name: 'No Orphan Players', category: 'DataIntegrity', check: checkOrphanPlayers_, fix: null },
    { id: 'H', name: 'Integrity Log Exists', category: 'UIWiring', check: checkIntegrityLog_, fix: fixIntegrityLog_ }
  ];

  for (var i = 0; i < gates.length; i++) {
    var gate = gates[i];
    var result = { id: gate.id, name: gate.name, category: gate.category, status: 'UNKNOWN', details: [], fixed: false };

    try {
      var checkResult = gate.check(ss);
      result.status = checkResult.passed ? 'PASS' : 'FAIL';
      result.details = checkResult.details || [];

      if (!checkResult.passed && mode === 'FIX' && gate.fix) {
        var fixResult = gate.fix(ss);
        result.fixed = fixResult.fixed;
        result.fixDetails = fixResult.details || [];
        if (fixResult.fixed) {
          report.summary.fixed++;
          var recheck = gate.check(ss);
          result.status = recheck.passed ? 'PASS' : 'FAIL';
        }
      }
    } catch (e) {
      result.status = 'ERROR';
      result.details = [e.message];
    }

    if (result.status === 'PASS') report.summary.passed++;
    else report.summary.failed++;

    report.gates.push(result);
    report.categories[gate.category].push(result);
  }

  return report;
}

function checkRequiredSheets_(ss) {
  var missing = [];
  for (var name in REQUIRED_SHEETS) {
    if (!getSheetByNameCI_(ss, name)) {
      missing.push(name);
    }
  }
  return { passed: missing.length === 0, details: missing.length > 0 ? ['Missing: ' + missing.join(', ')] : ['All required sheets present'] };
}

function fixRequiredSheets_(ss) {
  var fixed = [];
  for (var name in REQUIRED_SHEETS) {
    if (!getSheetByNameCI_(ss, name)) {
      ensureSheetWithHeaders_(ss, name, REQUIRED_SHEETS[name]);
      fixed.push(name);
    }
  }
  return { fixed: fixed.length > 0, details: fixed.length > 0 ? ['Created: ' + fixed.join(', ')] : [] };
}

function checkSheetHeaders_(ss) {
  var issues = [];
  for (var name in REQUIRED_SHEETS) {
    var sheet = getSheetByNameCI_(ss, name);
    if (!sheet) continue;

    var headers = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
    var headersLower = headers.map(function(h) { return String(h).toLowerCase(); });
    var required = REQUIRED_SHEETS[name];
    var missing = [];

    for (var i = 0; i < required.length; i++) {
      if (headersLower.indexOf(required[i].toLowerCase()) === -1) {
        missing.push(required[i]);
      }
    }

    if (missing.length > 0) {
      issues.push(name + ' missing: ' + missing.join(', '));
    }
  }
  return { passed: issues.length === 0, details: issues.length > 0 ? issues : ['All headers valid'] };
}

function fixSheetHeaders_(ss) {
  var fixed = [];
  for (var name in REQUIRED_SHEETS) {
    var result = ensureSheetWithHeaders_(ss, name, REQUIRED_SHEETS[name]);
    if (result.headersAdded) {
      fixed.push(name);
    }
  }
  return { fixed: fixed.length > 0, details: fixed.length > 0 ? ['Fixed headers in: ' + fixed.join(', ')] : [] };
}

function checkEventTabs_(ss) {
  var tabs = listEventTabsCI_(ss);
  return {
    passed: tabs.length > 0,
    details: ['Found ' + tabs.length + ' event tab(s)']
  };
}

function checkPlayersProvisioned_(ss) {
  var canonicalPlayers = getAllCanonicalPlayers_(ss);
  if (canonicalPlayers.length === 0) {
    return { passed: true, details: ['No players registered yet'] };
  }

  var sheetsToCheck = ['BP_Total', 'Attendance_Missions', 'Flag_Missions', 'Dice_Points'];
  var issues = [];

  for (var i = 0; i < sheetsToCheck.length; i++) {
    var sheet = getSheetByNameCI_(ss, sheetsToCheck[i]);
    if (!sheet) continue;

    var missing = 0;
    for (var j = 0; j < canonicalPlayers.length; j++) {
      if (!playerExistsInSheet_(sheet, canonicalPlayers[j])) {
        missing++;
      }
    }

    if (missing > 0) {
      issues.push(sheetsToCheck[i] + ': ' + missing + ' players not provisioned');
    }
  }

  return { passed: issues.length === 0, details: issues.length > 0 ? issues : ['All players provisioned in all sheets'] };
}

function fixPlayersProvisioned_(ss) {
  var canonicalPlayers = getAllCanonicalPlayers_(ss);
  var totalProvisioned = 0;

  for (var i = 0; i < canonicalPlayers.length; i++) {
    var result = provisionPlayerEverywhere_(ss, canonicalPlayers[i]);
    totalProvisioned += result.provisioned.length;
  }

  return { fixed: totalProvisioned > 0, details: ['Provisioned ' + totalProvisioned + ' missing player entries'] };
}

function checkBPTotalSchema_(ss) {
  var sheet = getSheetByNameCI_(ss, 'BP_Total');
  if (!sheet) return { passed: false, details: ['BP_Total sheet missing'] };

  var headerMap = getHeaderMap_(sheet);
  var required = ['BP_Historical', 'BP_Redeemed', 'BP_Current'];
  var missing = [];

  for (var i = 0; i < required.length; i++) {
    var found = false;
    var synonyms = BP_COLUMN_SYNONYMS[required[i].replace('BP_', '').toLowerCase()] || [required[i].toLowerCase()];
    for (var j = 0; j < synonyms.length; j++) {
      if (headerMap[synonyms[j].toLowerCase().replace(/[_\s-]/g, '')] !== undefined) {
        found = true;
        break;
      }
    }
    if (!found) missing.push(required[i]);
  }

  return { passed: missing.length === 0, details: missing.length > 0 ? ['Missing BP columns: ' + missing.join(', ')] : ['BP_Total schema valid'] };
}

function fixBPTotalSchema_(ss) {
  var result = ensureSheetWithHeaders_(ss, 'BP_Total', REQUIRED_SHEETS.BP_Total);
  return { fixed: result.headersAdded || result.created, details: result.created ? ['Created BP_Total'] : (result.headersAdded ? ['Added missing headers'] : []) };
}

function checkAttendanceCalendar_(ss) {
  var sheet = getSheetByNameCI_(ss, 'Attendance_Calendar');
  if (!sheet) return { passed: false, details: ['Attendance_Calendar sheet missing'] };

  var eventTabs = listEventTabsCI_(ss);
  var calendarRows = sheet.getLastRow() - 1;

  if (eventTabs.length > 0 && calendarRows < eventTabs.length / 2) {
    return { passed: false, details: ['Calendar appears stale: ' + calendarRows + ' entries but ' + eventTabs.length + ' event tabs'] };
  }

  return { passed: true, details: ['Attendance_Calendar has ' + calendarRows + ' entries'] };
}

function fixAttendanceCalendar_(ss) {
  var result = rebuildAttendanceCalendarInternal_(ss);
  return { fixed: result.rebuilt, details: ['Rebuilt calendar with ' + result.events + ' events'] };
}

function checkOrphanPlayers_(ss) {
  var canonical = getAllCanonicalPlayers_(ss);
  var canonicalLower = canonical.map(function(n) { return n.toLowerCase(); });

  var eventTabs = listEventTabsCI_(ss);
  var orphans = [];

  for (var i = 0; i < eventTabs.length; i++) {
    var sheet = eventTabs[i].sheet;
    if (sheet.getLastRow() <= 1) continue;

    var data = sheet.getDataRange().getValues();
    var headerMap = getHeaderMap_(sheet);
    var nameCol = findPlayerColumn_(headerMap);

    for (var r = 1; r < data.length; r++) {
      var name = normalizePlayerName_(data[r][nameCol]);
      if (name && canonicalLower.indexOf(name.toLowerCase()) === -1) {
        if (orphans.indexOf(name) === -1) orphans.push(name);
      }
    }
  }

  return {
    passed: orphans.length === 0,
    details: orphans.length > 0 ? ['Found ' + orphans.length + ' unregistered player(s): ' + orphans.slice(0, 5).join(', ') + (orphans.length > 5 ? '...' : '')] : ['No orphan players']
  };
}

function checkIntegrityLog_(ss) {
  var sheet = getSheetByNameCI_(ss, 'Integrity_Log');
  return { passed: sheet !== null, details: sheet ? ['Integrity_Log exists'] : ['Integrity_Log missing'] };
}

function fixIntegrityLog_(ss) {
  var result = ensureSheetWithHeaders_(ss, 'Integrity_Log', REQUIRED_SHEETS.Integrity_Log);
  return { fixed: result.created, details: result.created ? ['Created Integrity_Log'] : [] };
}

function rebuildAttendanceCalendarInternal_(ss) {
  var sheet = getSheetByNameCI_(ss, 'Attendance_Calendar');
  if (!sheet) {
    var result = ensureSheetWithHeaders_(ss, 'Attendance_Calendar', REQUIRED_SHEETS.Attendance_Calendar);
    sheet = result.sheet;
  }

  var eventTabs = listEventTabsCI_(ss);
  var calendarData = [REQUIRED_SHEETS.Attendance_Calendar];

  for (var i = 0; i < eventTabs.length; i++) {
    var tabName = eventTabs[i].name;
    var eventSheet = eventTabs[i].sheet;
    var playerCount = Math.max(0, eventSheet.getLastRow() - 1);

    var format = 'Unknown';
    var suffixMatch = tabName.match(/\d{4}([A-Za-z]+)?$/);
    if (suffixMatch && suffixMatch[1]) {
      var suffix = suffixMatch[1].toUpperCase();
      if (suffix === 'C') format = 'Commander';
      else if (suffix === 'T') format = 'cEDH';
      else if (suffix === 'M') format = 'Modern';
      else if (suffix === 'L') format = 'Legacy';
      else if (suffix === 'P') format = 'Pioneer';
      else if (suffix === 'S') format = 'Standard';
      else format = suffix;
    }

    calendarData.push([tabName, tabName, format, playerCount, 'COMPLETE']);
  }

  sheet.clearContents();
  if (calendarData.length > 0) {
    sheet.getRange(1, 1, calendarData.length, calendarData[0].length).setValues(calendarData);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, calendarData[0].length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  }

  return { rebuilt: true, events: eventTabs.length };
}

function formatShipGatesReport_(report) {
  var lines = [];
  lines.push('═══════════════════════════════════════════════════════');
  lines.push('       COSMIC ENGINE SHIP GATES HEALTH REPORT');
  lines.push('═══════════════════════════════════════════════════════');
  lines.push('Mode: ' + report.mode + '    Time: ' + report.timestamp);
  lines.push('');
  lines.push('SUMMARY: ' + report.summary.passed + ' PASSED | ' + report.summary.failed + ' FAILED | ' + report.summary.fixed + ' FIXED');
  lines.push('');

  lines.push('───────────────────────────────────────────────────────');
  lines.push('GATE RESULTS:');
  lines.push('───────────────────────────────────────────────────────');

  for (var i = 0; i < report.gates.length; i++) {
    var gate = report.gates[i];
    var icon = gate.status === 'PASS' ? '✓' : (gate.status === 'FAIL' ? '✗' : '⚠');
    lines.push('[' + gate.id + '] ' + icon + ' ' + gate.name + ' (' + gate.category + ')');

    for (var j = 0; j < gate.details.length; j++) {
      lines.push('    └─ ' + gate.details[j]);
    }

    if (gate.fixed && gate.fixDetails) {
      for (var k = 0; k < gate.fixDetails.length; k++) {
        lines.push('    └─ [FIXED] ' + gate.fixDetails[k]);
      }
    }
  }

  lines.push('');
  lines.push('───────────────────────────────────────────────────────');
  lines.push('FISHBONE BREAKDOWN BY CATEGORY:');
  lines.push('───────────────────────────────────────────────────────');

  var categories = ['Tabs', 'Headers', 'DataIntegrity', 'Players', 'Indexing', 'UIWiring'];
  for (var c = 0; c < categories.length; c++) {
    var cat = categories[c];
    var gates = report.categories[cat] || [];
    if (gates.length === 0) continue;

    var passed = gates.filter(function(g) { return g.status === 'PASS'; }).length;
    lines.push('  ' + cat + ': ' + passed + '/' + gates.length + ' passed');
  }

  lines.push('═══════════════════════════════════════════════════════');

  return lines.join('\n');
}

// ============================================================================
// PLAYER LOOKUP SYSTEM
// ============================================================================

function getPlayerLookupProfile_(playerName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var normalizedName = normalizePlayerName_(playerName);

  var profile = {
    name: normalizedName,
    found: false,
    canonicalName: null,
    bp: { historical: 0, redeemed: 0, current: 0, flag: 0, attendance: 0, dice: 0 },
    keys: { white: 0, blue: 0, black: 0, red: 0, green: 0, total: 0, rainbow: false },
    storeCredit: { balance: 0, transactions: [] },
    attendance: { totalEvents: 0, tier: 'None' },
    preorders: { hasOpen: false, list: [] },
    missing: []
  };

  var prefSheet = getSheetByNameCI_(ss, 'PreferredNames');
  if (prefSheet) {
    var data = prefSheet.getDataRange().getValues();
    var headerMap = getHeaderMap_(prefSheet);
    var nameCol = findPlayerColumn_(headerMap);

    for (var r = 1; r < data.length; r++) {
      if (playerNamesMatch_(data[r][nameCol], normalizedName)) {
        profile.found = true;
        profile.canonicalName = normalizePlayerName_(data[r][nameCol]);
        break;
      }
    }
  } else {
    profile.missing.push('PreferredNames');
  }

  if (!profile.found) return profile;

  var bpSheet = getSheetByNameCI_(ss, 'BP_Total');
  if (bpSheet) {
    var bpData = bpSheet.getDataRange().getValues();
    var bpHeaderMap = getHeaderMap_(bpSheet);
    var bpNameCol = findPlayerColumn_(bpHeaderMap);

    for (var r = 1; r < bpData.length; r++) {
      if (playerNamesMatch_(bpData[r][bpNameCol], profile.canonicalName)) {
        var histCol = findColumnBySynonyms_(bpHeaderMap, BP_COLUMN_SYNONYMS.historical);
        var redeemedCol = findColumnBySynonyms_(bpHeaderMap, BP_COLUMN_SYNONYMS.redeemed);
        var currentCol = findColumnBySynonyms_(bpHeaderMap, BP_COLUMN_SYNONYMS.current);
        var flagCol = bpHeaderMap['flagpoints'] !== undefined ? bpHeaderMap['flagpoints'] : bpHeaderMap['flag_points'];
        var attCol = bpHeaderMap['attendancepoints'] !== undefined ? bpHeaderMap['attendancepoints'] : bpHeaderMap['attendance_points'];
        var diceCol = bpHeaderMap['dicepoints'] !== undefined ? bpHeaderMap['dicepoints'] : bpHeaderMap['dice_points'];

        if (histCol !== -1) profile.bp.historical = Number(bpData[r][histCol]) || 0;
        if (redeemedCol !== -1) profile.bp.redeemed = Number(bpData[r][redeemedCol]) || 0;
        if (currentCol !== -1) profile.bp.current = Number(bpData[r][currentCol]) || 0;
        if (flagCol !== undefined) profile.bp.flag = Number(bpData[r][flagCol]) || 0;
        if (attCol !== undefined) profile.bp.attendance = Number(bpData[r][attCol]) || 0;
        if (diceCol !== undefined) profile.bp.dice = Number(bpData[r][diceCol]) || 0;
        break;
      }
    }
  } else {
    profile.missing.push('BP_Total');
  }

  var keySheet = getSheetByNameCI_(ss, 'Key_Tracker');
  if (keySheet) {
    var keyData = keySheet.getDataRange().getValues();
    var keyHeaderMap = getHeaderMap_(keySheet);
    var keyNameCol = findPlayerColumn_(keyHeaderMap);

    for (var r = 1; r < keyData.length; r++) {
      if (playerNamesMatch_(keyData[r][keyNameCol], profile.canonicalName)) {
        profile.keys.white = Number(keyData[r][keyHeaderMap['white']]) || 0;
        profile.keys.blue = Number(keyData[r][keyHeaderMap['blue']]) || 0;
        profile.keys.black = Number(keyData[r][keyHeaderMap['black']]) || 0;
        profile.keys.red = Number(keyData[r][keyHeaderMap['red']]) || 0;
        profile.keys.green = Number(keyData[r][keyHeaderMap['green']]) || 0;
        profile.keys.total = Number(keyData[r][keyHeaderMap['total']]) || 0;
        profile.keys.rainbow = Boolean(keyData[r][keyHeaderMap['rainbow']]);
        break;
      }
    }
  } else {
    profile.missing.push('Key_Tracker');
  }

  var attSheet = getSheetByNameCI_(ss, 'Attendance_Missions');
  if (attSheet) {
    var attData = attSheet.getDataRange().getValues();
    var attHeaderMap = getHeaderMap_(attSheet);
    var attNameCol = findPlayerColumn_(attHeaderMap);

    for (var r = 1; r < attData.length; r++) {
      if (playerNamesMatch_(attData[r][attNameCol], profile.canonicalName)) {
        var eventsCol = attHeaderMap['totalevents'] !== undefined ? attHeaderMap['totalevents'] : attHeaderMap['total_events'];
        if (eventsCol !== undefined) {
          profile.attendance.totalEvents = Number(attData[r][eventsCol]) || 0;
        }

        if (profile.attendance.totalEvents >= 100) profile.attendance.tier = 'Legend';
        else if (profile.attendance.totalEvents >= 50) profile.attendance.tier = 'Master';
        else if (profile.attendance.totalEvents >= 25) profile.attendance.tier = 'Expert';
        else if (profile.attendance.totalEvents >= 10) profile.attendance.tier = 'Regular';
        else if (profile.attendance.totalEvents >= 5) profile.attendance.tier = 'Newcomer';
        else profile.attendance.tier = 'New';
        break;
      }
    }
  } else {
    profile.missing.push('Attendance_Missions');
  }

  var scSheet = getSheetByNameCI_(ss, 'Store_Credit_Ledger') || getSheetByNameCI_(ss, 'Store_Credit_Log');
  if (scSheet) {
    var scData = scSheet.getDataRange().getValues();
    var scHeaderMap = getHeaderMap_(scSheet);
    var scNameCol = findPlayerColumn_(scHeaderMap);
    var amountCol = scHeaderMap['amount'];
    var directionCol = scHeaderMap['inout'] !== undefined ? scHeaderMap['inout'] : scHeaderMap['direction'];

    var balance = 0;
    for (var r = 1; r < scData.length; r++) {
      if (playerNamesMatch_(scData[r][scNameCol], profile.canonicalName)) {
        var amount = Number(scData[r][amountCol]) || 0;
        var dir = String(scData[r][directionCol]).toUpperCase();
        if (dir === 'IN' || dir === 'CREDIT') balance += amount;
        else if (dir === 'OUT' || dir === 'DEBIT') balance -= amount;

        if (profile.storeCredit.transactions.length < 5) {
          profile.storeCredit.transactions.push({ amount: amount, direction: dir });
        }
      }
    }
    profile.storeCredit.balance = balance;
  } else {
    profile.missing.push('Store_Credit');
  }

  var preorderSheet = getSheetByNameCI_(ss, 'Preorders') || getSheetByNameCI_(ss, 'Preorder_Log');
  if (preorderSheet) {
    var poData = preorderSheet.getDataRange().getValues();
    var poHeaderMap = getHeaderMap_(preorderSheet);
    var poNameCol = findPlayerColumn_(poHeaderMap);
    var statusCol = poHeaderMap['status'];

    for (var r = 1; r < poData.length; r++) {
      if (playerNamesMatch_(poData[r][poNameCol], profile.canonicalName)) {
        var status = String(poData[r][statusCol]).toUpperCase();
        if (status !== 'PICKED_UP' && status !== 'CANCELLED' && status !== 'COMPLETE') {
          profile.preorders.hasOpen = true;
          profile.preorders.list.push({ row: r + 1, status: status });
        }
      }
    }
  }

  return profile;
}

// ============================================================================
// TRIGGER: onOpen (ONLY ONE IN ENTIRE PROJECT)
// ============================================================================

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();

  var mainMenu = ui.createMenu('Cosmic Tournament v' + ENGINE_VERSION);

  var eventsMenu = ui.createMenu('Events');
  addMenuItemOrStub_(eventsMenu, 'Start New Event', 'onCreateEvent', ui);
  addMenuItemOrStub_(eventsMenu, 'Commander Event Wizard', 'onCommanderEventWizard', ui);
  addMenuItemOrStub_(eventsMenu, 'Import Player List (Roster)', 'onRosterImport', ui);
  addMenuItemOrStub_(eventsMenu, 'View Event Index', 'onViewEventIndex', ui);
  addSep_(eventsMenu);
  addMenuItemOrStub_(eventsMenu, 'Preview End Prizes', 'onPreviewEndPrizes', ui);
  addMenuItemOrStub_(eventsMenu, 'Lock In End Prizes', 'onGenerateEndPrizes', ui);
  addMenuItemOrStub_(eventsMenu, 'Commander Round Prizes', 'onCommanderRounds', ui);
  addSep_(eventsMenu);
  addMenuItemOrStub_(eventsMenu, 'Undo Last Prize Run', 'onRevertPrizes', ui);
  mainMenu.addSubMenu(eventsMenu);

  var playersMenu = ui.createMenu('Players');
  addMenuItemOrStub_(playersMenu, 'Add New Player', 'onAddNewPlayer', ui);
  addMenuItemOrStub_(playersMenu, 'Detect / Fix Player Names', 'onPlayerNameChecker', ui);
  addMenuItemOrStub_(playersMenu, 'Player Lookup', 'onPlayerLookup', ui);
  addSep_(playersMenu);
  addMenuItemOrStub_(playersMenu, 'Add Key', 'onAddKey', ui);
  mainMenu.addSubMenu(playersMenu);

  var bpMenu = ui.createMenu('Bonus Points');
  addMenuItemOrStub_(bpMenu, 'Award Bonus Points', 'onAwardBP', ui);
  addMenuItemOrStub_(bpMenu, 'Redeem Bonus Points', 'onRedeemBP', ui);
  addSep_(bpMenu);
  addMenuItemOrStub_(bpMenu, 'Sync BP from Sources (Canonical)', 'menuSyncBPFromSources', ui);
  addMenuItemOrStub_(bpMenu, 'Provision All Players', 'onProvisionAllPlayers', ui);
  mainMenu.addSubMenu(bpMenu);

  var missionsMenu = ui.createMenu('Missions & Attendance');
  addMenuItemOrStub_(missionsMenu, 'Scan Attendance / Missions (Canonical)', 'onScanAttendance', ui);
  addMenuItemOrStub_(missionsMenu, 'Rebuild Attendance Calendar', 'onRebuildAttendanceCalendar', ui);
  addSep_(missionsMenu);
  addMenuItemOrStub_(missionsMenu, 'Record Dice Roll Results', 'onDiceRollResults', ui);
  addMenuItemOrStub_(missionsMenu, 'Award Flag Mission', 'onAwardFlagMission', ui);
  addMenuItemOrStub_(missionsMenu, 'Record Attendance', 'onRecordAttendance', ui);
  addSep_(missionsMenu);
  addMenuItemOrStub_(missionsMenu, 'Validate Mission Points Integrity', 'onValidateMissionPoints', ui);
  mainMenu.addSubMenu(missionsMenu);

  var catalogMenu = ui.createMenu('Catalog');
  addMenuItemOrStub_(catalogMenu, 'Manage Prize Catalog', 'onCatalogManager', ui);
  addMenuItemOrStub_(catalogMenu, 'Prize Throttle (Switchboard)', 'onThrottle', ui);
  addSep_(catalogMenu);
  addMenuItemOrStub_(catalogMenu, 'Import Preorder Allocation', 'onPreorderImport', ui);
  mainMenu.addSubMenu(catalogMenu);

  var preordersMenu = ui.createMenu('Preorders');
  addMenuItemOrStub_(preordersMenu, 'Sell Preorder', 'onSellPreorder', ui);
  addMenuItemOrStub_(preordersMenu, 'View Preorder Status', 'onViewPreorderStatus', ui);
  addMenuItemOrStub_(preordersMenu, 'Mark Preorder Pickup', 'onMarkPreorderPickup', ui);
  addMenuItemOrStub_(preordersMenu, 'Cancel Preorder', 'onCancelPreorder', ui);
  addSep_(preordersMenu);
  addMenuItemOrStub_(preordersMenu, 'Manage Preorder Buckets', 'onPreorderBuckets', ui);
  addMenuItemOrStub_(preordersMenu, 'View Preorders Sold', 'onViewPreordersSold', ui);
  mainMenu.addSubMenu(preordersMenu);

  var opsMenu = ui.createMenu('Ops');
  addMenuItemOrStub_(opsMenu, 'Daily Close Checklist', 'onDailyCloseChecklist', ui);
  addSep_(opsMenu);
  addMenuItemOrStub_(opsMenu, '📊 Build Event Dashboard', 'buildEventDashboard', ui);
  addMenuItemOrStub_(opsMenu, '💰 Update Cost Per Player', 'updateCostPerPlayer', ui);
  addMenuItemOrStub_(opsMenu, '🔄 Refresh Dashboard', 'buildEventDashboard', ui);
  addSep_(opsMenu);
  addMenuItemOrStub_(opsMenu, 'Ship-Gates Health Check', 'onShipGates', ui);
  addMenuItemOrStub_(opsMenu, 'Build / Repair', 'onBuildRepair', ui);
  addMenuItemOrStub_(opsMenu, 'Organize Tabs', 'onOrganizeTabs', ui);
  addMenuItemOrStub_(opsMenu, 'Clean Old Previews', 'onCleanPreviews', ui);
  addSep_(opsMenu);
  addMenuItemOrStub_(opsMenu, 'View Integrity Log', 'onViewLog', ui);
  addMenuItemOrStub_(opsMenu, 'View Spent Pool', 'onViewSpent', ui);
  addMenuItemOrStub_(opsMenu, 'Export Reports', 'onExportReports', ui);
  addSep_(opsMenu);
  addMenuItemOrStub_(opsMenu, 'Force Unlock Event', 'onForceUnlock', ui);
  addMenuItemOrStub_(opsMenu, 'Emergency Revert', 'onEmergencyRevert', ui);
  mainMenu.addSubMenu(opsMenu);

  mainMenu.addToUi();

  var empMenu = ui.createMenu('Cosmic Employee Tools');
  addMenuItemOrStub_(empMenu, 'Employee Log', 'onOpenEmployeeLog', ui);
  addSep_(empMenu);
  addMenuItemOrStub_(empMenu, 'Create single assignment…', 'createSingleAssignment', ui);
  addMenuItemOrStub_(empMenu, 'View pending assignments', 'onViewPendingAssignments', ui);
  addSep_(empMenu);
  addMenuItemOrStub_(empMenu, 'This Needs... (New Task)', 'showThisNeedsSidebar', ui);
  addMenuItemOrStub_(empMenu, 'This Needs... Task Board', 'showThisNeedsTaskBoard', ui);
  empMenu.addToUi();

  var scMenu = ui.createMenu('Store Credit');
  addMenuItemOrStub_(scMenu, 'Spend Store Credit', 'onStoreCredit', ui);
  addMenuItemOrStub_(scMenu, 'View Store Credit', 'onStoreCreditViewer', ui);
  scMenu.addToUi();
}

// ============================================================================
// TRIGGER: onEdit (ONLY ONE IN ENTIRE PROJECT)
// ============================================================================

function onEdit(e) {
  if (!e || !e.range) return;

  if (typeof onDicePointCheckboxEdit === 'function') {
    try { onDicePointCheckboxEdit(e); } catch (err) { console.error('onEdit dice error:', err); }
  }

  if (typeof handleEmployeeLogEdit_ === 'function') {
    try { handleEmployeeLogEdit_(e); } catch (err) { console.error('onEdit employee error:', err); }
  }
}

// ============================================================================
// ERROR HELPER
// ============================================================================

function showError_(context, error) {
  var ui = SpreadsheetApp.getUi();
  ui.alert('Error', context + '\n\nError: ' + (error.message || error), ui.ButtonSet.OK);
  console.error(context, error);
}

// ============================================================================
// SHIP GATES & BUILD/REPAIR HANDLERS
// ============================================================================

function onShipGates() {
  try {
    var report = runShipGates_('CHECK');
    var text = formatShipGatesReport_(report);
    SpreadsheetApp.getUi().alert('Ship Gates Health Check', text, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { showError_('Ship Gates failed', e); }
}

function onBuildRepair() {
  try {
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert('Build / Repair',
      'This will run health checks and attempt to fix issues.\n\n' +
      'Options:\n' +
      '• YES = Run checks AND apply fixes\n' +
      '• NO = Run checks only (no changes)\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO_CANCEL);

    if (result === ui.Button.CANCEL) return;

    var mode = result === ui.Button.YES ? 'FIX' : 'CHECK';
    SpreadsheetApp.getActiveSpreadsheet().toast('Running Build/Repair in ' + mode + ' mode...', 'Please wait', -1);

    var report = runShipGates_(mode);
    var text = formatShipGatesReport_(report);

    SpreadsheetApp.getActiveSpreadsheet().toast('Complete!', 'Build/Repair', 3);
    ui.alert('Build / Repair Results', text, ui.ButtonSet.OK);
  } catch (e) { showError_('Build/Repair failed', e); }
}

// ============================================================================
// EVENT ROUTE HANDLERS
// ============================================================================

function onCreateEvent() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/create_event').setWidth(500).setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Create Event');
  } catch (e) { showError_('Failed to open Create Event', e); }
}

function onCommanderEventWizard() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/commander_wizard').setWidth(700).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Commander Event Wizard');
  } catch (e) { showError_('Failed to open Commander Wizard', e); }
}

function onRosterImport() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/roster_import').setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Roster Import', e); }
}

function onViewEventIndex() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/event_index').setTitle('Event Index').setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Event Index', e); }
}

function onPreviewEndPrizes() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/preview_end').setWidth(700).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Preview End Prizes');
  } catch (e) { showError_('Failed to open Preview End Prizes', e); }
}

function onGenerateEndPrizes() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/generate_end').setWidth(700).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Lock In End Prizes');
  } catch (e) { showError_('Failed to open Lock In End Prizes', e); }
}

function onCommanderRounds() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/round_prizes').setWidth(600).setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(html, 'Commander Round Prizes');
  } catch (e) { showError_('Failed to open Commander Rounds', e); }
}

function onRevertPrizes() {
  try {
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert('Undo Last Prize Run', 'This will revert the most recent prize distribution.\n\nContinue?', ui.ButtonSet.YES_NO);
    if (result === ui.Button.YES) {
      if (typeof revertLastPrizeRun === 'function') {
        var r = revertLastPrizeRun();
        ui.alert('Revert Complete', 'Reverted ' + (r.count || 0) + ' prize(s).', ui.ButtonSet.OK);
      } else {
        ui.alert('Not Available', 'Prize revert function not implemented.', ui.ButtonSet.OK);
      }
    }
  } catch (e) { showError_('Failed to revert prizes', e); }
}

// ============================================================================
// PLAYER ROUTE HANDLERS
// ============================================================================

function onAddNewPlayer() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/add_player').setWidth(450).setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Add New Player');
  } catch (e) { showError_('Failed to open Add New Player', e); }
}

function onPlayerNameChecker() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('PlayerNameChecker').setWidth(500).setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Player Name Checker');
  } catch (e) { showError_('Failed to open Player Name Checker', e); }
}

function onPlayerLookup() {
  try {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('Player Lookup', 'Enter player name:', ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() !== ui.Button.OK) return;

    var playerName = response.getResponseText().trim();
    if (!playerName) {
      ui.alert('No name entered.');
      return;
    }

    var profile = getPlayerLookupProfile_(playerName);

    if (!profile.found) {
      ui.alert('Player Not Found', 'No player matching "' + playerName + '" was found in PreferredNames.', ui.ButtonSet.OK);
      return;
    }

    var lines = [];
    lines.push('═══════════════════════════════════════════');
    lines.push('PLAYER: ' + profile.canonicalName);
    lines.push('═══════════════════════════════════════════');
    lines.push('');
    lines.push('BONUS POINTS:');
    lines.push('  Historical: ' + profile.bp.historical);
    lines.push('  Redeemed: ' + profile.bp.redeemed);
    lines.push('  Current: ' + profile.bp.current);
    lines.push('  (Flag: ' + profile.bp.flag + ' | Att: ' + profile.bp.attendance + ' | Dice: ' + profile.bp.dice + ')');
    lines.push('');
    lines.push('KEYS: W:' + profile.keys.white + ' U:' + profile.keys.blue + ' B:' + profile.keys.black + ' R:' + profile.keys.red + ' G:' + profile.keys.green);
    lines.push('  Total: ' + profile.keys.total + ' | Rainbow: ' + (profile.keys.rainbow ? 'YES' : 'No'));
    lines.push('');
    lines.push('ATTENDANCE: ' + profile.attendance.totalEvents + ' events (' + profile.attendance.tier + ')');
    lines.push('');
    lines.push('STORE CREDIT: $' + profile.storeCredit.balance.toFixed(2));
    lines.push('');
    lines.push('PREORDERS: ' + (profile.preorders.hasOpen ? profile.preorders.list.length + ' open' : 'None'));

    if (profile.missing.length > 0) {
      lines.push('');
      lines.push('Missing subsystems: ' + profile.missing.join(', '));
    }

    ui.alert('Player Profile', lines.join('\n'), ui.ButtonSet.OK);
  } catch (e) { showError_('Failed to lookup player', e); }
}

function onAddKey() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/add_key').setWidth(450).setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Add Key');
  } catch (e) { showError_('Failed to open Add Key', e); }
}

// ============================================================================
// BONUS POINTS ROUTE HANDLERS
// ============================================================================

function onAwardBP() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/award_bp').setWidth(500).setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Award Bonus Points');
  } catch (e) { showError_('Failed to open Award BP', e); }
}

function onRedeemBP() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/redeem_bp').setWidth(500).setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Redeem Bonus Points');
  } catch (e) { showError_('Failed to open Redeem BP', e); }
}

function menuSyncBPFromSources() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Syncing BP from source sheets...', 'Please Wait', -1);

  try {
    if (typeof updateBPTotalFromSources === 'function') {
      var count = updateBPTotalFromSources();
      ss.toast('Synced ' + count + ' player(s)', 'BP Sync Complete', 5);
    } else {
      ss.toast('BP sync function not available', 'Error', 5);
    }
  } catch (e) {
    ss.toast('Error: ' + e.message, 'Sync Failed', 5);
    console.error('BP Sync Error:', e);
  }
}

function onProvisionAllPlayers() {
  try {
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert('Provision All Players',
      'This will ensure all players from PreferredNames are provisioned in all tracking sheets.\n\nContinue?',
      ui.ButtonSet.YES_NO);

    if (result !== ui.Button.YES) return;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Provisioning players...', 'Please wait', -1);

    var players = getAllCanonicalPlayers_(ss);
    var totalProvisioned = 0;

    for (var i = 0; i < players.length; i++) {
      var r = provisionPlayerEverywhere_(ss, players[i]);
      totalProvisioned += r.provisioned.length;
    }

    ss.toast('Complete!', 'Provisioning', 3);
    ui.alert('Provisioning Complete', 'Provisioned ' + totalProvisioned + ' missing entries for ' + players.length + ' player(s).', ui.ButtonSet.OK);
  } catch (e) { showError_('Failed to provision players', e); }
}

// ============================================================================
// MISSIONS & ATTENDANCE ROUTE HANDLERS
// ============================================================================

function onScanAttendance() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    ss.toast('Scanning attendance...', 'Please wait', -1);

    if (typeof runMissionScan === 'function') {
      var result = runMissionScan();
      ui.alert('Scan Complete',
        'Events: ' + (result.eventsScanned || 0) + '\n' +
        'Players: ' + (result.playersTracked || 0) + '\n' +
        'Missions: ' + (result.missionsComputed || 0),
        ui.ButtonSet.OK);
    } else {
      ui.alert('Not Available', 'Mission scan function not implemented.', ui.ButtonSet.OK);
    }
  } catch (e) { showError_('Failed to scan attendance', e); }
}

function onRebuildAttendanceCalendar() {
  try {
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert('Rebuild Calendar', 'This will rebuild the Attendance Calendar from event tabs.\n\nContinue?', ui.ButtonSet.YES_NO);

    if (result !== ui.Button.YES) return;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var r = rebuildAttendanceCalendarInternal_(ss);
    ui.alert('Rebuild Complete', 'Processed ' + r.events + ' event(s).', ui.ButtonSet.OK);
  } catch (e) { showError_('Failed to rebuild calendar', e); }
}

function onDiceRollResults() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/dice_results').setWidth(500).setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Record Dice Roll Results');
  } catch (e) { showError_('Failed to open Dice Roll Results', e); }
}

function onAwardFlagMission() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/flag_mission').setWidth(500).setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Award Flag Mission');
  } catch (e) { showError_('Failed to open Award Flag Mission', e); }
}

function onRecordAttendance() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/record_attendance').setWidth(500).setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Record Attendance');
  } catch (e) { showError_('Failed to open Record Attendance', e); }
}

function onValidateMissionPoints() {
  try {
    var ui = SpreadsheetApp.getUi();
    if (typeof validateMissionPointsIntegrity === 'function') {
      var result = validateMissionPointsIntegrity();
      if (result.pass) {
        ui.alert('Validation Passed', 'All mission point data is valid.', ui.ButtonSet.OK);
      } else {
        ui.alert('Validation Issues', 'Found ' + result.issues.length + ' issue(s).', ui.ButtonSet.OK);
      }
    } else {
      ui.alert('Not Available', 'Validation function not implemented.', ui.ButtonSet.OK);
    }
  } catch (e) { showError_('Failed to validate', e); }
}

// ============================================================================
// CATALOG ROUTE HANDLERS
// ============================================================================

function onCatalogManager() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/catalog_manager').setWidth(800).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Manage Prize Catalog');
  } catch (e) { showError_('Failed to open Catalog Manager', e); }
}

function onThrottle() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/throttle').setWidth(500).setHeight(550);
    SpreadsheetApp.getUi().showModalDialog(html, 'Prize Throttle');
  } catch (e) { showError_('Failed to open Throttle', e); }
}

function onPreorderImport() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/preorder_import').setWidth(900).setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(html, 'Import Preorder Allocation');
  } catch (e) { showError_('Failed to open Preorder Import', e); }
}

// ============================================================================
// PREORDER ROUTE HANDLERS
// ============================================================================

function onSellPreorder() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/sell_preorder').setWidth(950).setHeight(900);
    SpreadsheetApp.getUi().showModalDialog(html, 'Sell Preorder');
  } catch (e) { showError_('Failed to open Sell Preorder', e); }
}

function onViewPreorderStatus() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/preorder_status').setTitle('Preorder Status').setWidth(900);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Preorder Status', e); }
}

function onMarkPreorderPickup() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/preorder_pickup').setWidth(900).setHeight(950);
    SpreadsheetApp.getUi().showModalDialog(html, 'Mark Preorder Pickup');
  } catch (e) { showError_('Failed to open Preorder Pickup', e); }
}

function onCancelPreorder() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/preorder_cancel').setWidth(900).setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(html, 'Cancel Preorder');
  } catch (e) { showError_('Failed to open Cancel Preorder', e); }
}

function onPreorderBuckets() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/Preorder_buckets').setWidth(900).setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(html, 'Manage Preorder Buckets');
  } catch (e) { showError_('Failed to open Preorder Buckets', e); }
}

function onViewPreordersSold() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/preorders_sold').setTitle('Preorders Sold').setWidth(800);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Preorders Sold', e); }
}

// ============================================================================
// OPS ROUTE HANDLERS
// ============================================================================

function onDailyCloseChecklist() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/daily_close').setWidth(600).setHeight(550);
    SpreadsheetApp.getUi().showModalDialog(html, 'Daily Close Checklist');
  } catch (e) { showError_('Failed to open Daily Close', e); }
}

function buildEventDashboard() {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Building dashboard...', 'Please wait', -1);
    if (typeof buildEventDashboard_ === 'function') {
      buildEventDashboard_();
      SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard built!', 'Complete', 3);
    } else {
      SpreadsheetApp.getUi().alert('Not Available', 'Dashboard builder not implemented.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (e) { showError_('Failed to build dashboard', e); }
}

function updateCostPerPlayer() {
  try {
    SpreadsheetApp.getUi().alert('Not Available', 'Cost per player update not implemented.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { showError_('Failed to update cost', e); }
}

function onOrganizeTabs() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/tab_organizer').setWidth(600).setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(html, 'Organize Tabs');
  } catch (e) { showError_('Failed to open Tab Organizer', e); }
}

function onCleanPreviews() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getSheetByNameCI_(ss, 'Preview_Artifacts');
    var count = 0;

    if (sheet && sheet.getLastRow() > 1) {
      var now = new Date().getTime();
      var cutoff = now - (24 * 60 * 60 * 1000);
      var data = sheet.getDataRange().getValues();

      for (var i = data.length - 1; i > 0; i--) {
        var exp = new Date(data[i][4]).getTime();
        if (exp < cutoff) {
          sheet.deleteRow(i + 1);
          count++;
        }
      }
    }

    SpreadsheetApp.getUi().alert('Clean Complete', 'Removed ' + count + ' stale preview(s).', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { showError_('Failed to clean previews', e); }
}

function onViewLog() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/log_viewer').setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Log Viewer', e); }
}

function onViewSpent() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/spent_pool_viewer').setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Spent Pool', e); }
}

function onExportReports() {
  SpreadsheetApp.getUi().alert('Export Reports',
    'To export, open the sheet and use:\nFile → Download → Comma Separated Values (.csv)',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function onForceUnlock() {
  try {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('Force Unlock', 'Enter Event ID (sheet name):', ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() !== ui.Button.OK) return;

    var eventId = response.getResponseText().trim();
    if (!eventId) return;

    if (typeof forceUnlockEvent === 'function') {
      forceUnlockEvent(eventId);
      ui.alert('Unlocked', 'Event "' + eventId + '" unlocked.', ui.ButtonSet.OK);
    } else {
      ui.alert('Not Available', 'Unlock function not implemented.', ui.ButtonSet.OK);
    }
  } catch (e) { showError_('Failed to force unlock', e); }
}

function onEmergencyRevert() {
  try {
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert('Emergency Revert', 'This will attempt to revert recent changes.\n\nAre you sure?', ui.ButtonSet.YES_NO);

    if (result === ui.Button.YES) {
      if (typeof emergencyRevert === 'function') {
        emergencyRevert();
        ui.alert('Revert Complete', 'Emergency revert executed.', ui.ButtonSet.OK);
      } else {
        ui.alert('Not Available', 'Emergency revert not implemented.', ui.ButtonSet.OK);
      }
    }
  } catch (e) { showError_('Failed emergency revert', e); }
}

// ============================================================================
// EMPLOYEE TOOLS ROUTE HANDLERS
// ============================================================================

function onOpenEmployeeLog() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/employee_log').setTitle('Employee Log').setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Employee Log', e); }
}

function createSingleAssignment() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/single_assignment').setWidth(500).setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Create Assignment');
  } catch (e) { showError_('Failed to open Create Assignment', e); }
}

function onViewPendingAssignments() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/pending_assignments').setTitle('Pending Assignments').setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Pending Assignments', e); }
}

function showThisNeedsSidebar() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/this_needs_dialog').setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open This Needs', e); }
}

function showThisNeedsTaskBoard() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/task_board').setWidth(800).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Task Board');
  } catch (e) { showError_('Failed to open Task Board', e); }
}

// ============================================================================
// STORE CREDIT ROUTE HANDLERS
// ============================================================================

function onStoreCredit() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/store_credit').setTitle('Spend Store Credit');
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Store Credit', e); }
}

function onStoreCreditViewer() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/store_credit_viewer').setTitle('View Store Credit').setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Store Credit Viewer', e); }
}

// ============================================================================
// BACKWARDS COMPATIBILITY ALIASES
// ============================================================================

function onScanMissions() { onScanAttendance(); }
function onSyncBPTotals() { menuSyncBPFromSources(); }
function onRefreshBPTotalFromSources() { menuSyncBPFromSources(); }
function onPreviewEnd() { onPreviewEndPrizes(); }
function onDetectNewPlayers() { onPlayerNameChecker(); }
function onPlayerProfile() { onPlayerLookup(); }
function onThisNeeds() { showThisNeedsSidebar(); }
function onTaskBoard() { showThisNeedsTaskBoard(); }
function onSingleAssignment() { createSingleAssignment(); }
function onPendingAssignments() { onViewPendingAssignments(); }
function rebuildAttendanceCalendar() { return rebuildAttendanceCalendarInternal_(SpreadsheetApp.getActiveSpreadsheet()); }
function provisionAllPlayers() { var ss = SpreadsheetApp.getActiveSpreadsheet(); var players = getAllCanonicalPlayers_(ss); var total = 0; for (var i = 0; i < players.length; i++) { var r = provisionPlayerEverywhere_(ss, players[i]); total += r.provisioned.length; } return { provisioned: total, alreadyExisted: players.length - total }; }

// ============================================================================
// LIGHTWEIGHT EDIT HANDLERS (called by onEdit)
// ============================================================================

function handleEmployeeLogEdit_(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName().toLowerCase() !== 'employee_log') return;

  var row = e.range.getRow();
  if (row <= 1) return;

  var col = e.range.getColumn();
  if (col > 1 && e.value && !sheet.getRange(row, 1).getValue()) {
    sheet.getRange(row, 1).setValue(new Date());
  }
}

// ============================================================================
// END OF FILE
// ============================================================================
