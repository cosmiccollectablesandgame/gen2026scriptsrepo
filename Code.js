/**
 * Cosmic Games Tournament Manager
 * Main Entry Point - Menu Creation and Triggers
 *
 * CRITICAL: This is the ONLY file that contains onOpen() and onEdit()
 *
 * @fileoverview Main menu wiring and route handlers for the Cosmic Engine
 *
 * REQUIREMENTS CHECKLIST:
 * [x] 1. Exactly one onOpen(e) and one onEdit(e)
 * [x] 2. onOpen(e) compiles with no broken submenu chains
 * [x] 3. onEdit(e) is ultra-lightweight, NO BP aggregation
 * [x] 4. onEdit(e) only calls lightweight handlers in try/catch
 * [x] 5. All menu items use addMenuItemOrStub_() for safety
 * [x] 6. Menu title uses ENGINE_VERSION dynamically
 * [x] 7. Includes showError_(context, error) helper
 * [x] 8. No top-level executable statements except constants
 * [x] 9. All braces properly matched
 * [x] 10. Backwards compatibility aliases included
 */

// ============================================================================
// VERSION CONSTANTS (only allowed top-level statements)
// ============================================================================

const ENGINE_VERSION = '7.9.8';
const RECIPE_VERSION = '1.0.1';

// ============================================================================
// MENU HELPER: Safe Item Addition
// ============================================================================

/**
 * Safely adds a menu item, stubbing missing handlers with a friendly alert.
 * @param {GoogleAppsScript.Base.Menu} menu - The menu to add to
 * @param {string} label - Menu item label
 * @param {string} fnName - Handler function name
 * @param {GoogleAppsScript.Base.Ui} ui - UI instance for creating stubs
 */
function addMenuItemOrStub_(menu, label, fnName, ui) {
  if (typeof globalThis[fnName] === 'function') {
    menu.addItem(label, fnName);
  } else {
    // Create a dynamic stub function name
    var stubName = '__stub_' + fnName;

    // Only create stub if it doesn't exist yet
    if (typeof globalThis[stubName] !== 'function') {
      globalThis[stubName] = function() {
        SpreadsheetApp.getUi().alert(
          'Feature Not Available',
          'Menu item "' + label + '" requires function "' + fnName + '" which is not defined.\n\n' +
          'Possible fixes:\n' +
          '• Search the project for "function ' + fnName + '"\n' +
          '• Check if the file containing this function was imported\n' +
          '• This feature may not be implemented yet',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      };
    }
    menu.addItem(label + ' ⚠️', stubName);
  }
}

/**
 * Adds a separator to the menu (convenience wrapper).
 * @param {GoogleAppsScript.Base.Menu} menu - The menu
 */
function addSep_(menu) {
  menu.addSeparator();
}

// ============================================================================
// TRIGGER: onOpen (ONLY ONE IN ENTIRE PROJECT)
// ============================================================================

/**
 * Creates custom menus on spreadsheet open.
 * This is the ONLY onOpen function - do NOT define another one elsewhere.
 * @param {Object} e - The onOpen event object
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();

  // =========================================================================
  // MENU 1: Cosmic Tournament (Main Menu)
  // =========================================================================
  var mainMenu = ui.createMenu('Cosmic Tournament v' + ENGINE_VERSION);

  // --- Events Submenu ---
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

  // --- Players Submenu ---
  var playersMenu = ui.createMenu('Players');
  addMenuItemOrStub_(playersMenu, 'Add New Player', 'onAddNewPlayer', ui);
  addMenuItemOrStub_(playersMenu, 'Detect / Fix Player Names', 'onPlayerNameChecker', ui);
  addMenuItemOrStub_(playersMenu, 'Player Lookup', 'onPlayerLookup', ui);
  addSep_(playersMenu);
  addMenuItemOrStub_(playersMenu, 'Add Key', 'onAddKey', ui);
  mainMenu.addSubMenu(playersMenu);

  // --- Bonus Points Submenu ---
  var bpMenu = ui.createMenu('Bonus Points');
  addMenuItemOrStub_(bpMenu, 'Award Bonus Points', 'onAwardBP', ui);
  addMenuItemOrStub_(bpMenu, 'Redeem Bonus Points', 'onRedeemBP', ui);
  addSep_(bpMenu);
  addMenuItemOrStub_(bpMenu, 'Sync BP from Sources (Canonical)', 'menuSyncBPFromSources', ui);
  addMenuItemOrStub_(bpMenu, 'Provision All Players', 'onProvisionAllPlayers', ui);
  mainMenu.addSubMenu(bpMenu);

  // --- Missions & Attendance Submenu ---
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

  // --- Catalog Submenu ---
  var catalogMenu = ui.createMenu('Catalog');
  addMenuItemOrStub_(catalogMenu, 'Manage Prize Catalog', 'onCatalogManager', ui);
  addMenuItemOrStub_(catalogMenu, 'Prize Throttle (Switchboard)', 'onThrottle', ui);
  addSep_(catalogMenu);
  addMenuItemOrStub_(catalogMenu, 'Import Preorder Allocation', 'onPreorderImport', ui);
  mainMenu.addSubMenu(catalogMenu);

  // --- Preorders Submenu ---
  var preordersMenu = ui.createMenu('Preorders');
  addMenuItemOrStub_(preordersMenu, 'Sell Preorder', 'onSellPreorder', ui);
  addMenuItemOrStub_(preordersMenu, 'View Preorder Status', 'onViewPreorderStatus', ui);
  addMenuItemOrStub_(preordersMenu, 'Mark Preorder Pickup', 'onMarkPreorderPickup', ui);
  addMenuItemOrStub_(preordersMenu, 'Cancel Preorder', 'onCancelPreorder', ui);
  addSep_(preordersMenu);
  addMenuItemOrStub_(preordersMenu, 'Manage Preorder Buckets', 'onPreorderBuckets', ui);
  addMenuItemOrStub_(preordersMenu, 'View Preorders Sold', 'onViewPreordersSold', ui);
  mainMenu.addSubMenu(preordersMenu);

  // --- Ops Submenu ---
  var opsMenu = ui.createMenu('Ops');
  addMenuItemOrStub_(opsMenu, 'Daily Close Checklist', 'onDailyCloseChecklist', ui);
  addSep_(opsMenu);
  addMenuItemOrStub_(opsMenu, 'Build Event Dashboard', 'buildEventDashboard', ui);
  addMenuItemOrStub_(opsMenu, 'Update Cost Per Player', 'updateCostPerPlayer', ui);
  addMenuItemOrStub_(opsMenu, 'Refresh Dashboard', 'buildEventDashboard', ui);
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

  // =========================================================================
  // MENU 2: Cosmic Employee Tools
  // =========================================================================
  var empMenu = ui.createMenu('Cosmic Employee Tools');
  addMenuItemOrStub_(empMenu, 'Employee Log', 'onOpenEmployeeLog', ui);
  addSep_(empMenu);
  addMenuItemOrStub_(empMenu, 'Create single assignment…', 'createSingleAssignment', ui);
  addMenuItemOrStub_(empMenu, 'View pending assignments', 'onViewPendingAssignments', ui);
  addSep_(empMenu);
  addMenuItemOrStub_(empMenu, 'This Needs... (New Task)', 'showThisNeedsSidebar', ui);
  addMenuItemOrStub_(empMenu, 'This Needs... Task Board', 'showThisNeedsTaskBoard', ui);
  empMenu.addToUi();

  // =========================================================================
  // MENU 3: Store Credit
  // =========================================================================
  var scMenu = ui.createMenu('Store Credit');
  addMenuItemOrStub_(scMenu, 'Spend Store Credit', 'onStoreCredit', ui);
  addMenuItemOrStub_(scMenu, 'View Store Credit', 'onStoreCreditViewer', ui);
  scMenu.addToUi();
}

// ============================================================================
// TRIGGER: onEdit (ONLY ONE IN ENTIRE PROJECT)
// ============================================================================

/**
 * Handles edit events for the spreadsheet.
 * This is the ONLY onEdit function - do NOT define another one elsewhere.
 *
 * CRITICAL: This function MUST be ultra-lightweight.
 * - NO BP aggregation or sync
 * - NO mission scans
 * - NO heavy computation
 * - BP sync is MENU-DRIVEN ONLY via menuSyncBPFromSources()
 *
 * @param {Object} e - The onEdit event object
 */
function onEdit(e) {
  if (!e || !e.range) return;

  // Dice point checkbox handler (lightweight)
  if (typeof onDicePointCheckboxEdit === 'function') {
    try {
      onDicePointCheckboxEdit(e);
    } catch (err) {
      console.error('onEdit - Dice checkbox error:', err);
    }
  }

  // Employee Log edit handler (lightweight)
  if (typeof handleEmployeeLogEdit_ === 'function') {
    try {
      handleEmployeeLogEdit_(e);
    } catch (err) {
      console.error('onEdit - Employee log error:', err);
    }
  }

  // NOTE: DO NOT add BP sync, mission scans, or heavy operations here.
  // Those are triggered via menu items only.
}

// ============================================================================
// ERROR HELPER
// ============================================================================

/**
 * Shows error dialog with friendly message for menu handler failures.
 * @param {string} context - User-friendly context description
 * @param {Error} error - The error object
 */
function showError_(context, error) {
  var ui = SpreadsheetApp.getUi();
  var message = context + '\n\nError: ' + (error.message || error) +
                '\n\nPlease check Build/Repair for health issues.';
  ui.alert('Error', message, ui.ButtonSet.OK);
  console.error(context, error);
}

// ============================================================================
// EVENT ROUTE HANDLERS
// ============================================================================

function onCreateEvent() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/create_event')
      .setWidth(500).setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Create Event');
  } catch (e) { showError_('Failed to open Create Event dialog', e); }
}

function onCommanderEventWizard() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/commander_wizard')
      .setWidth(700).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Commander Event Wizard');
  } catch (e) { showError_('Failed to open Commander Event Wizard', e); }
}

function onRosterImport() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/roster_import')
      .setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Roster Import', e); }
}

function onViewEventIndex() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/event_index')
      .setTitle('Event Index').setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Event Index', e); }
}

function onPreviewEndPrizes() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/preview_end')
      .setWidth(700).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Preview End Prizes');
  } catch (e) { showError_('Failed to open Preview End Prizes', e); }
}

function onGenerateEndPrizes() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/generate_end')
      .setWidth(700).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Lock In End Prizes');
  } catch (e) { showError_('Failed to open End Prizes generator', e); }
}

function onCommanderRounds() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/round_prizes')
      .setWidth(600).setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(html, 'Commander Round Prizes');
  } catch (e) { showError_('Failed to open Commander Rounds', e); }
}

function onRevertPrizes() {
  try {
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert('Undo Last Prize Run',
      'This will revert the most recent prize distribution.\nThis action cannot be undone. Continue?',
      ui.ButtonSet.YES_NO);
    if (result === ui.Button.YES) {
      if (typeof revertLastPrizeRun === 'function') {
        var revertResult = revertLastPrizeRun();
        ui.alert('Revert Complete', 'Reverted ' + (revertResult.count || 0) + ' prize(s).', ui.ButtonSet.OK);
      } else {
        ui.alert('Not Available', 'Prize revert function not available.', ui.ButtonSet.OK);
      }
    }
  } catch (e) { showError_('Failed to revert prizes', e); }
}

// ============================================================================
// PLAYER ROUTE HANDLERS
// ============================================================================

function onAddNewPlayer() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/add_player')
      .setWidth(450).setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Add New Player');
  } catch (e) { showError_('Failed to open Add New Player', e); }
}

function onPlayerNameChecker() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('PlayerNameChecker')
      .setWidth(500).setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Player Name Checker');
  } catch (e) { showError_('Failed to open Player Name Checker', e); }
}

function onPlayerLookup() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/player_lookup')
      .setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Player Lookup', e); }
}

function onAddKey() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/add_key')
      .setWidth(450).setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Add Key');
  } catch (e) { showError_('Failed to open Add Key dialog', e); }
}

// ============================================================================
// BONUS POINTS ROUTE HANDLERS
// ============================================================================

function onAwardBP() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/award_bp')
      .setWidth(500).setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Award Bonus Points');
  } catch (e) { showError_('Failed to open Award BP dialog', e); }
}

function onRedeemBP() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/redeem_bp')
      .setWidth(500).setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Redeem Bonus Points');
  } catch (e) { showError_('Failed to open BP Redemption', e); }
}

/**
 * Canonical menu handler for BP sync operation.
 * Provides user feedback via toast notifications.
 */
function menuSyncBPFromSources() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Syncing BP from source sheets...', 'Please Wait', -1);

  try {
    if (typeof updateBPTotalFromSources === 'function') {
      var updatedCount = updateBPTotalFromSources();
      if (updatedCount > 0) {
        ss.toast('Successfully synced ' + updatedCount + ' player(s)', 'BP Sync Complete', 5);
      } else {
        ss.toast('All players already up to date', 'BP Sync Complete', 3);
      }
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
      'This will ensure all players from PreferredNames are provisioned in:\n' +
      '- Flag_Missions\n- Attendance_Missions\n- Dice_Points\n- BP_Total\n\nContinue?',
      ui.ButtonSet.YES_NO);
    if (result === ui.Button.YES) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Provisioning players...', 'Provision', 5);
      if (typeof provisionAllPlayers === 'function') {
        var provResult = provisionAllPlayers();
        ui.alert('Provisioning Complete',
          'Provisioned: ' + provResult.provisioned + '\nAlready existed: ' + provResult.alreadyExisted,
          ui.ButtonSet.OK);
      } else {
        ui.alert('Not Available', 'Provisioning function not available.', ui.ButtonSet.OK);
      }
    }
  } catch (e) { showError_('Failed to provision players', e); }
}

// ============================================================================
// MISSIONS & ATTENDANCE ROUTE HANDLERS
// ============================================================================

function onScanAttendance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  try {
    ss.toast('Scanning attendance & missions...', 'Mission Scan', 10);
    if (typeof runMissionScan === 'function') {
      var result = runMissionScan();
      ui.alert('Mission Scan Complete',
        'Events Scanned: ' + (result.eventsScanned || 0) + '\n' +
        'Players Tracked: ' + (result.playersTracked || 0) + '\n' +
        'Missions Evaluated: ' + (result.missionsComputed || 0),
        ui.ButtonSet.OK);
    } else {
      ui.alert('Not Available', 'Mission scan function not available.', ui.ButtonSet.OK);
    }
  } catch (e) { showError_('Failed to scan attendance/missions', e); }
}

function onRebuildAttendanceCalendar() {
  try {
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert('Rebuild Attendance Calendar',
      'This will rebuild the Attendance_Calendar sheet from all event tabs.\n\n' +
      'Existing data will be replaced. Continue?',
      ui.ButtonSet.YES_NO);
    if (result === ui.Button.YES) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Rebuilding calendar...', 'Rebuild', 10);
      if (typeof rebuildAttendanceCalendar === 'function') {
        var rebuildResult = rebuildAttendanceCalendar();
        ui.alert('Rebuild Complete', 'Processed ' + (rebuildResult.events || 0) + ' event(s).', ui.ButtonSet.OK);
      } else {
        ui.alert('Not Available', 'Rebuild function not available.', ui.ButtonSet.OK);
      }
    }
  } catch (e) { showError_('Failed to rebuild attendance calendar', e); }
}

function onDiceRollResults() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/dice_results')
      .setWidth(500).setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Record Dice Roll Results');
  } catch (e) { showError_('Failed to open Dice Results dialog', e); }
}

function onAwardFlagMission() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/flag_mission')
      .setWidth(500).setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Award Flag Mission');
  } catch (e) { showError_('Failed to open Flag Mission dialog', e); }
}

function onRecordAttendance() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/record_attendance')
      .setWidth(500).setHeight(400);
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
        var issueList = result.issues.slice(0, 10).map(function(i) {
          return i.sheet + ' row ' + i.row + ': ' + i.issue;
        }).join('\n');
        ui.alert('Validation Issues Found',
          'Found ' + result.issues.length + ' issue(s):\n\n' + issueList +
          (result.issues.length > 10 ? '\n\n... and more' : ''),
          ui.ButtonSet.OK);
      }
    } else {
      ui.alert('Not Available', 'Validation function not available.', ui.ButtonSet.OK);
    }
  } catch (e) { showError_('Failed to validate mission points', e); }
}

// ============================================================================
// CATALOG ROUTE HANDLERS
// ============================================================================

function onCatalogManager() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/catalog_manager')
      .setWidth(800).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Manage Prize Catalog');
  } catch (e) { showError_('Failed to open Catalog Manager', e); }
}

function onThrottle() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/throttle')
      .setWidth(500).setHeight(550);
    SpreadsheetApp.getUi().showModalDialog(html, 'Prize Throttle (Switchboard)');
  } catch (e) { showError_('Failed to open Throttle Switchboard', e); }
}

function onPreorderImport() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/preorder_import')
      .setWidth(900).setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(html, 'Import Preorder Allocation');
  } catch (e) { showError_('Failed to open Preorder Import', e); }
}

// ============================================================================
// PREORDER ROUTE HANDLERS
// ============================================================================

function onSellPreorder() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/sell_preorder')
      .setWidth(950).setHeight(900);
    SpreadsheetApp.getUi().showModalDialog(html, 'Sell Preorder');
  } catch (e) { showError_('Failed to open Sell Preorder dialog', e); }
}

function onViewPreorderStatus() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/preorder_status')
      .setTitle('Preorder Status').setWidth(900);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Preorder Status', e); }
}

function onMarkPreorderPickup() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/preorder_pickup')
      .setWidth(900).setHeight(950);
    SpreadsheetApp.getUi().showModalDialog(html, 'Mark Preorder Pickup');
  } catch (e) { showError_('Failed to open Preorder Pickup', e); }
}

function onCancelPreorder() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/preorder_cancel')
      .setWidth(900).setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(html, 'Cancel Preorder');
  } catch (e) { showError_('Failed to open Cancel Preorder', e); }
}

function onPreorderBuckets() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/Preorder_buckets')
      .setWidth(900).setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(html, 'Manage Preorder Buckets');
  } catch (e) { showError_('Failed to open Preorder Buckets', e); }
}

function onViewPreordersSold() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/preorders_sold')
      .setTitle('Preorders Sold').setWidth(800);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Preorders Sold viewer', e); }
}

// ============================================================================
// OPS ROUTE HANDLERS
// ============================================================================

function onDailyCloseChecklist() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/daily_close')
      .setWidth(600).setHeight(550);
    SpreadsheetApp.getUi().showModalDialog(html, 'Daily Close Checklist');
  } catch (e) { showError_('Failed to open Daily Close Checklist', e); }
}

function onShipGates() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/ship_gates')
      .setWidth(600).setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(html, 'Ship-Gates Health Check');
  } catch (e) { showError_('Failed to open Ship-Gates Health Check', e); }
}

function onBuildRepair() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/health')
      .setWidth(600).setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(html, 'Build / Repair - Health Dashboard');
  } catch (e) { showError_('Failed to open Health Dashboard', e); }
}

function onOrganizeTabs() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/tab_organizer')
      .setWidth(600).setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(html, 'Tab Organizer');
  } catch (e) { showError_('Failed to open Tab Organizer', e); }
}

function onCleanPreviews() {
  try {
    var count = typeof cleanOldPreviews_ === 'function' ? cleanOldPreviews_(24) : 0;
    SpreadsheetApp.getUi().alert('Clean Previews Complete',
      'Removed ' + count + ' stale preview artifact(s).',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { showError_('Failed to clean previews', e); }
}

function onViewLog() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/log_viewer')
      .setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Log Viewer', e); }
}

function onViewSpent() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/spent_pool_viewer')
      .setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Spent Pool Viewer', e); }
}

function onExportReports() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Export Reports',
      'To export, open the Integrity_Log or Spent_Pool sheet and use:\n' +
      'File → Download → Comma Separated Values (.csv)',
      ui.ButtonSet.OK);
  } catch (e) { showError_('Failed to show export instructions', e); }
}

function onForceUnlock() {
  try {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('Force Unlock Event',
      'Enter the Event ID (sheet name) to unlock:',
      ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() === ui.Button.OK) {
      var eventId = response.getResponseText().trim();
      if (eventId && typeof forceUnlockEvent === 'function') {
        forceUnlockEvent(eventId);
        ui.alert('Unlocked', 'Event "' + eventId + '" has been unlocked.', ui.ButtonSet.OK);
      } else if (!eventId) {
        ui.alert('Cancelled', 'No event ID provided.', ui.ButtonSet.OK);
      } else {
        ui.alert('Not Available', 'Force unlock function not available.', ui.ButtonSet.OK);
      }
    }
  } catch (e) { showError_('Failed to force unlock event', e); }
}

function onEmergencyRevert() {
  try {
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert('Emergency Revert',
      'WARNING: This will attempt to revert recent changes.\n\n' +
      'This is a last-resort option. Are you sure?',
      ui.ButtonSet.YES_NO);
    if (result === ui.Button.YES) {
      if (typeof emergencyRevert === 'function') {
        emergencyRevert();
        ui.alert('Revert Complete', 'Emergency revert has been executed.', ui.ButtonSet.OK);
      } else {
        ui.alert('Not Available', 'Emergency revert function not available.', ui.ButtonSet.OK);
      }
    }
  } catch (e) { showError_('Failed to execute emergency revert', e); }
}

// ============================================================================
// EMPLOYEE TOOLS ROUTE HANDLERS
// ============================================================================

function onOpenEmployeeLog() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/employee_log')
      .setTitle('Employee Log').setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Employee Log', e); }
}

function createSingleAssignment() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/single_assignment')
      .setWidth(500).setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Create Single Assignment');
  } catch (e) { showError_('Failed to open Single Assignment', e); }
}

function onViewPendingAssignments() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/pending_assignments')
      .setTitle('Pending Assignments').setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Pending Assignments', e); }
}

function showThisNeedsSidebar() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/this_needs_dialog')
      .setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open This Needs sidebar', e); }
}

function showThisNeedsTaskBoard() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/task_board')
      .setWidth(800).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'This Needs... Task Board');
  } catch (e) { showError_('Failed to open Task Board', e); }
}

// ============================================================================
// STORE CREDIT ROUTE HANDLERS
// ============================================================================

function onStoreCredit() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/store_credit')
      .setTitle('Spend Store Credit');
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Store Credit sidebar', e); }
}

function onStoreCreditViewer() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ui/store_credit_viewer')
      .setTitle('View Store Credit').setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) { showError_('Failed to open Store Credit Viewer', e); }
}

// ============================================================================
// BACKWARDS COMPATIBILITY ALIASES
// These ensure old menu configurations or external callers still work
// ============================================================================

/** @deprecated Use onScanAttendance instead */
function onScanMissions() {
  console.log('onScanMissions called - redirecting to onScanAttendance');
  onScanAttendance();
}

/** @deprecated Use menuSyncBPFromSources instead */
function onSyncBPTotals() {
  console.log('onSyncBPTotals called - redirecting to menuSyncBPFromSources');
  menuSyncBPFromSources();
}

/** @deprecated Use menuSyncBPFromSources instead */
function onRefreshBPTotalFromSources() {
  console.log('onRefreshBPTotalFromSources called - redirecting to menuSyncBPFromSources');
  menuSyncBPFromSources();
}

/** @deprecated Use onPreviewEndPrizes instead */
function onPreviewEnd() {
  onPreviewEndPrizes();
}

/** @deprecated Use onPlayerNameChecker instead */
function onDetectNewPlayers() {
  onPlayerNameChecker();
}

/** @deprecated Use onViewPlayerProfile instead */
function onPlayerProfile() {
  onPlayerLookup();
}

/** @deprecated Use showThisNeedsSidebar instead */
function onThisNeeds() {
  showThisNeedsSidebar();
}

/** @deprecated Use showThisNeedsTaskBoard instead */
function onTaskBoard() {
  showThisNeedsTaskBoard();
}

/** @deprecated Use createSingleAssignment instead */
function onSingleAssignment() {
  createSingleAssignment();
}

/** @deprecated Use onViewPendingAssignments instead */
function onPendingAssignments() {
  onViewPendingAssignments();
}

// ============================================================================
// UTILITY FUNCTIONS USED BY HANDLERS
// ============================================================================

/**
 * Gets a list of event tab names.
 * @return {Array<string>} Array of event tab names
 */
function listEventTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var datePattern = /^\d{1,2}-\d{1,2}[A-Za-z]?-\d{4}/;

  return sheets
    .map(function(s) { return s.getName(); })
    .filter(function(name) { return datePattern.test(name); })
    .sort();
}

/**
 * Gets event properties from a sheet.
 * @param {string|Sheet} eventIdOrSheet - Event ID or Sheet object
 * @return {Object|null} Event properties
 */
function getEventProps(eventIdOrSheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = typeof eventIdOrSheet === 'string'
    ? ss.getSheetByName(eventIdOrSheet)
    : eventIdOrSheet;

  if (!sheet) return null;

  return {
    eventId: sheet.getName(),
    eventType: 'UNKNOWN',
    playerCount: 0,
    status: 'DRAFT'
  };
}

/**
 * Gets the current user's email.
 * @return {string} User email
 */
function currentUser() {
  try {
    return Session.getActiveUser().getEmail() || 'Unknown';
  } catch (e) {
    return 'anonymous';
  }
}

/**
 * Cleans old preview artifacts.
 * @param {number} hoursOld - Hours threshold
 * @return {number} Count of removed artifacts
 */
function cleanOldPreviews_(hoursOld) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Preview_Artifacts');
  if (!sheet) return 0;

  var data = sheet.getDataRange().getValues();
  var now = new Date().getTime();
  var cutoff = now - (hoursOld * 60 * 60 * 1000);
  var removeCount = 0;

  for (var i = data.length - 1; i > 0; i--) {
    var expiresAt = new Date(data[i][4]).getTime();
    if (expiresAt < cutoff) {
      sheet.deleteRow(i + 1);
      removeCount++;
    }
  }

  return removeCount;
}

/**
 * Handles edits to the Employee Log sheet.
 * @param {Object} e - Edit event
 */
function handleEmployeeLogEdit_(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== 'Employee_Log') return;

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
