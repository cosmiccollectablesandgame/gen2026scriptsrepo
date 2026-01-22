/**
 * ════════════════════════════════════════════════════════════════════════════
 * ATTENDANCE MISSION SCANNER - FULL IMPLEMENTATION
 * ════════════════════════════════════════════════════════════════════════════
 *
 * @fileoverview Scans event tabs, computes player stats, and awards mission points
 * 
 * This system:
 *   - Detects event sheets matching pattern MM-DD<suffix>-YYYY
 *   - Extracts standings from Column B (preferred_name_id)
 *   - Computes attendance stats (events, suffixes, weeks, months, placements)
 *   - Awards mission points based on defined criteria
 *   - Updates Attendance_Missions sheet with results
 *
 * Compatible with: Engine v7.9.6+
 * Version: 1.0.0
 * ════════════════════════════════════════════════════════════════════════════
 */

// ════════════════════════════════════════════════════════════════════════════
// PHASE 1: CORE FUNCTIONS
// ════════════════════════════════════════════════════════════════════════════

/**
 * Returns all event sheets from the spreadsheet
 * Event tabs match pattern: MM-DD<suffix>-YYYY
 * @return {Array<Sheet>} Array of event sheets
 */
function getEventSheets() {
  // Regex: 2 digits - 2 digits + any suffix + dash + 4 digits
  // Example matches: 05-10C-2025, 12-01Draft-2025, 07-26q-2025
  const EVENT_TAB_REGEX = /^\d{2}-\d{2}[A-Za-z]+-\d{4}$/;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  
  return allSheets.filter(sheet => EVENT_TAB_REGEX.test(sheet.getName()));
}

/**
 * Parses event sheet name into components
 * @param {string} name - Sheet name like "05-10C-2025"
 * @return {Object} {date: Date, suffix: string, monthKey: string, isoWeekKey: string}
 */
function parseEventSheetName(name) {
  // Pattern: MM-DD<suffix>-YYYY
  const match = name.match(/^(\d{2})-(\d{2})([A-Za-z]+)-(\d{4})$/);
  
  if (!match) {
    return null;
  }
  
  const month = parseInt(match[1], 10);
  const day = parseInt(match[2], 10);
  const suffix = match[3].toUpperCase(); // Normalize to uppercase
  const year = parseInt(match[4], 10);
  
  const date = new Date(year, month - 1, day);
  const monthKey = year + '-' + String(month).padStart(2, '0'); // "2025-05"
  const isoWeekKey = getISOWeekKey(date); // "2025-W19"
  
  return {
    date: date,
    suffix: suffix,
    monthKey: monthKey,
    isoWeekKey: isoWeekKey
  };
}

/**
 * Returns ISO week key for a date
 * Uses existing getWeekNumber() from MissionHelpers.js if available
 * @param {Date} date
 * @return {string} "YYYY-Www" format
 */
function getISOWeekKey(date) {
  // Calculate ISO week using the same logic as MissionHelpers.getWeekNumber
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return d.getUTCFullYear() + '-W' + String(weekNo).padStart(2, '0');
}

/**
 * Maps event suffix to category
 * @param {string} suffix - Uppercase suffix like "C", "D", "S"
 * @return {string} Category name
 * 
 * COSMIC GAMES SUFFIX CODES (Authoritative):
 *   A = Academy / Learn to Play
 *   B = Casual Commander (Brackets 1–2)
 *   C = Transitional Commander (Brackets 3–4)
 *   D = Booster Draft
 *   E = External / Outreach
 *   F = Commander Free Play
 *   G = Gundam / Gunpla
 *   H = Helped Out
 *   I = Yu-Gi-Oh TCG
 *   J = Riftbound Skirmish
 *   K = Kill Team
 *   L = Commander League
 *   M = Modern
 *   N = Pokémon TCG
 *   O = One Piece TCG
 *   P = Proxy / Cube Draft
 *   Q = Precon Event
 *   R = Prerelease
 *   S = Sealed
 *   T = Two-Headed Giant
 *   U = cEDH / High-Power Commander (Bracket 5)
 *   V = Riftbound Nexus Nights
 *   W = Workshop
 *   X = Multi-Event
 *   Y = Lorcana
 *   Z = Staff / Internal
 */
function getEventCategoryFromSuffix(suffix) {
  const SUFFIX_MAP = {
    // Commander variants
    'B': 'CASUAL_COMMANDER',
    'C': 'TRANSITIONAL_COMMANDER',
    'U': 'CEDH',
    'F': 'FREE_PLAY',
    'L': 'COMMANDER_LEAGUE',
    'Q': 'PRECON_EVENT',
    
    // Limited formats (count toward LIMITED missions)
    'D': 'DRAFT',
    'P': 'DRAFT',        // Proxy/Cube Draft counts as draft
    'S': 'SEALED',
    'R': 'PRERELEASE',
    
    // Academy / Learning
    'A': 'ACADEMY',
    'W': 'WORKSHOP',
    
    // Outreach
    'E': 'OUTREACH',
    
    // Other TCGs
    'G': 'GUNDAM',
    'I': 'YUGIOH',
    'J': 'RIFTBOUND',
    'K': 'KILL_TEAM',
    'N': 'POKEMON',
    'O': 'ONE_PIECE',
    'V': 'RIFTBOUND',
    'Y': 'LORCANA',
    
    // Constructed formats
    'M': 'MODERN',
    'T': 'TWO_HEADED_GIANT',
    
    // Internal / Special
    'H': 'HELPED_OUT',
    'X': 'MULTI_EVENT',
    'Z': 'STAFF_INTERNAL'
  };
  
  return SUFFIX_MAP[suffix.toUpperCase()] || 'OTHER';
}

/**
 * Checks if category is a Limited format
 * @param {string} category
 * @return {boolean}
 */
function isLimitedCategory(category) {
  return ['DRAFT', 'SEALED', 'PRERELEASE'].includes(category);
}

/**
 * Checks if category is a Commander format
 * @param {string} category
 * @return {boolean}
 */
function isCommanderCategory(category) {
  return ['CASUAL_COMMANDER', 'TRANSITIONAL_COMMANDER', 'CEDH', 'FREE_PLAY', 'COMMANDER_LEAGUE', 'PRECON_EVENT'].includes(category);
}

/**
 * Gets ordered standings from an event sheet
 * @param {Sheet} sheet - Event sheet
 * @return {Array<string>} Ordered list of preferred_name_id (non-empty, deduplicated)
 */
function getStandings(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  // Column B = preferred_name_id
  const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  
  const seen = new Set();
  const standings = [];
  
  for (let i = 0; i < data.length; i++) {
    const name = String(data[i][0] || '').trim();
    if (name && !seen.has(name)) {
      seen.add(name);
      standings.push(name);
    }
  }
  
  return standings;
}

// ════════════════════════════════════════════════════════════════════════════
// PHASE 2: PLAYER STATS COMPUTATION
// ════════════════════════════════════════════════════════════════════════════

/**
 * Scans all event tabs and builds attendance data
 * @return {Object} {
 *   events: [{name, date, suffix, category, monthKey, isoWeekKey, standings}],
 *   playerEvents: Map<playerId, [{eventName, position, isTop4, isLast}]>
 * }
 */
function scanAllEvents() {
  const eventSheets = getEventSheets();
  const events = [];
  const playerEvents = new Map();
  
  for (const sheet of eventSheets) {
    const name = sheet.getName();
    const parsed = parseEventSheetName(name);
    
    if (!parsed) continue;
    
    const category = getEventCategoryFromSuffix(parsed.suffix);
    const standings = getStandings(sheet);
    
    const eventInfo = {
      name: name,
      date: parsed.date,
      suffix: parsed.suffix,
      category: category,
      monthKey: parsed.monthKey,
      isoWeekKey: parsed.isoWeekKey,
      standings: standings
    };
    
    events.push(eventInfo);
    
    // Record each player's participation
    for (let pos = 0; pos < standings.length; pos++) {
      const playerId = standings[pos];
      
      if (!playerEvents.has(playerId)) {
        playerEvents.set(playerId, []);
      }
      
      playerEvents.get(playerId).push({
        eventName: name,
        suffix: parsed.suffix,
        category: category,
        monthKey: parsed.monthKey,
        isoWeekKey: parsed.isoWeekKey,
        position: pos + 1, // 1-based
        isTop4: pos < 4,
        isLast: pos === standings.length - 1
      });
    }
  }
  
  return { events, playerEvents };
}

/**
 * Computes all mission-relevant stats for a player
 * @param {string} playerId - preferred_name_id
 * @param {Array} playerEventList - List of events player attended
 * @return {Object} Stats object
 */
function computePlayerStats(playerId, playerEventList) {
  const stats = {
    playerId: playerId,
    totalEvents: 0,
    uniqueSuffixes: 0,
    loyalMonths: 0,
    limitedEvents: 0,
    draftEvents: 0,
    academyEvents: 0,
    meteorWeeks: 0,
    top4Finishes: 0,
    freePlayEvents: 0,
    lastPlaceFinishes: 0,
    // Category-specific counters (Phase 5)
    casualCommanderEvents: 0,
    transitionalCommanderEvents: 0,
    cedhEvents: 0,
    outreachEvents: 0,
    commanderLeagueEvents: 0,
    preconEvents: 0
  };
  
  if (!playerEventList || playerEventList.length === 0) {
    return stats;
  }
  
  stats.totalEvents = playerEventList.length;
  
  // Unique suffixes
  const suffixes = new Set();
  const monthCounts = new Map();
  const weekCounts = new Map();
  
  for (const event of playerEventList) {
    // Track suffix
    suffixes.add(event.suffix);
    
    // Track month attendance
    monthCounts.set(event.monthKey, (monthCounts.get(event.monthKey) || 0) + 1);
    
    // Track week attendance
    weekCounts.set(event.isoWeekKey, (weekCounts.get(event.isoWeekKey) || 0) + 1);
    
    // Category-based stats
    if (isLimitedCategory(event.category)) {
      stats.limitedEvents++;
    }
    if (event.category === 'DRAFT') {
      stats.draftEvents++;
    }
    if (event.category === 'ACADEMY') {
      stats.academyEvents++;
    }
    if (event.category === 'FREE_PLAY') {
      stats.freePlayEvents++;
    }
    
    // Phase 5: Category-specific counters
    if (event.category === 'CASUAL_COMMANDER') stats.casualCommanderEvents++;
    if (event.category === 'TRANSITIONAL_COMMANDER') stats.transitionalCommanderEvents++;
    if (event.category === 'CEDH') stats.cedhEvents++;
    if (event.category === 'OUTREACH') stats.outreachEvents++;
    if (event.category === 'COMMANDER_LEAGUE') stats.commanderLeagueEvents++;
    if (event.category === 'PRECON_EVENT') stats.preconEvents++;
    
    // Position-based stats
    if (event.isTop4) {
      stats.top4Finishes++;
    }
    if (event.isLast) {
      stats.lastPlaceFinishes++;
    }
  }
  
  stats.uniqueSuffixes = suffixes.size;
  
  // Loyal months: months with >= 4 events
  for (const count of monthCounts.values()) {
    if (count >= 4) {
      stats.loyalMonths++;
    }
  }
  
  // Meteor weeks: weeks with >= 2 events
  for (const count of weekCounts.values()) {
    if (count >= 2) {
      stats.meteorWeeks++;
    }
  }
  
  return stats;
}

// ════════════════════════════════════════════════════════════════════════════
// PHASE 3: MISSION POINTS CALCULATION
// ════════════════════════════════════════════════════════════════════════════

/**
 * Computes attendance mission points from stats
 * @param {Object} stats - Player stats object
 * @return {Object} Mission awards + total
 */
function computeAttendanceMissionPoints(stats) {
  // Mission columns that count toward Points total
  const missionColumns = {
    // One-time missions
    'First Contact': stats.totalEvents >= 1 ? 1 : 0,           // ATT-001
    'Stellar Explorer': stats.totalEvents >= 5 ? 2 : 0,        // ATT-002
    'Sealed Voyager': stats.limitedEvents >= 1 ? 1 : 0,        // ATT-005
    'Draft Navigator': stats.draftEvents >= 1 ? 1 : 0,         // ATT-006
    'Stellar Scholar': stats.academyEvents >= 1 ? 1 : 0,       // ATT-007
    
    // Scaling missions
    'Deck Diver': stats.uniqueSuffixes,                        // ATT-003
    'Lunar Loyalty': stats.loyalMonths,                        // ATT-004
    'Meteor Shower': stats.meteorWeeks,                        // ATT-008
    'Interstellar Strategist': stats.top4Finishes,             // ATT-009
    'Free Play Events': stats.freePlayEvents,                  // ATT-010
    'Black Hole Survivor': stats.lastPlaceFinishes             // ATT-011
  };
  
  // Category tracking columns (do NOT count toward Points)
  const categoryColumns = {
    'Casual Commander Events': stats.casualCommanderEvents,
    'Transitional Commander Events': stats.transitionalCommanderEvents,
    'cEDH Events': stats.cedhEvents,
    'Limited Events': stats.limitedEvents,
    'Academy Events': stats.academyEvents,
    'Outreach Events': stats.outreachEvents
  };
  
  // Combine all awards
  const awards = { ...missionColumns, ...categoryColumns };
  
  // Calculate Points total from mission columns only
  let total = 0;
  for (const key in missionColumns) {
    total += missionColumns[key];
  }
  awards['Points'] = total;
  
  return awards;
}

// ════════════════════════════════════════════════════════════════════════════
// PHASE 4: UPDATE ATTENDANCE_MISSIONS SHEET
// ════════════════════════════════════════════════════════════════════════════

/**
 * Main function: Scans events and updates Attendance_Missions sheet
 * @return {number} Number of players updated
 */
function syncAttendanceMissions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Attendance_Missions');
    
    if (!sheet) {
      throw new Error('Attendance_Missions sheet not found');
    }
    
    // Scan all events
    Logger.log('Scanning event tabs...');
    const { events, playerEvents } = scanAllEvents();
    Logger.log('Found ' + events.length + ' events, ' + playerEvents.size + ' players');
    
    // Get current sheet data
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Build column map
    const colMap = {};
    for (let i = 0; i < headers.length; i++) {
      colMap[headers[i]] = i;
    }
    
    // Verify required columns exist
    const requiredCols = [
      'PreferredName', 'First Contact', 'Stellar Explorer', 'Deck Diver',
      'Lunar Loyalty', 'Meteor Shower', 'Sealed Voyager', 'Draft Navigator',
      'Stellar Scholar', 'Interstellar Strategist', 'Black Hole Survivor',
      'Free Play Events', 'Points'
    ];
    
    for (const col of requiredCols) {
      if (colMap[col] === undefined) {
        throw new Error('Missing required column: ' + col);
      }
    }
    
    // Build existing rows map
    const existingRows = new Map();
    for (let i = 1; i < data.length; i++) {
      const name = String(data[i][colMap['PreferredName']] || '').trim();
      if (name) {
        existingRows.set(name, i + 1); // 1-based row number
      }
    }
    
    let updatedCount = 0;
    
    // Process each player
    for (const [playerId, eventList] of playerEvents) {
      const stats = computePlayerStats(playerId, eventList);
      const awards = computeAttendanceMissionPoints(stats);
      
      if (existingRows.has(playerId)) {
        // Update existing row
        const rowNum = existingRows.get(playerId);
        
        for (const [missionName, value] of Object.entries(awards)) {
          if (colMap[missionName] !== undefined) {
            sheet.getRange(rowNum, colMap[missionName] + 1).setValue(value);
          }
        }
        
        updatedCount++;
      } else {
        // New player - append row
        const newRow = new Array(headers.length).fill('');
        newRow[colMap['PreferredName']] = playerId;
        
        for (const [missionName, value] of Object.entries(awards)) {
          if (colMap[missionName] !== undefined) {
            newRow[colMap[missionName]] = value;
          }
        }
        
        sheet.appendRow(newRow);
        updatedCount++;
      }
    }
    
    Logger.log('Updated ' + updatedCount + ' players in Attendance_Missions');
    
    // Log success to Integrity_Log if available
    if (typeof logIntegrityAction === 'function') {
      logIntegrityAction('ATTENDANCE_MISSIONS_SYNC', {
        details: 'Scanned ' + events.length + ' events, updated ' + updatedCount + ' players',
        status: 'SUCCESS'
      });
    }
    
    // ════════════════════════════════════════════════════════════════════════
    // AUTO-PROVISION: Discover and provision new players found during scan
    // ════════════════════════════════════════════════════════════════════════
    if (typeof discoverAndProvisionNewPlayers === 'function') {
      try {
        const provisionResult = discoverAndProvisionNewPlayers();
        if (provisionResult.newPlayersFound > 0) {
          Logger.log('Auto-provisioned ' + provisionResult.provisioned + ' new player(s)');
          
          // Re-scan the newly provisioned players so their missions are computed
          if (provisionResult.provisioned > 0) {
            Logger.log('Re-scanning to compute missions for new players...');
            // Note: We don't recursively call syncAttendanceMissions to avoid infinite loops
            // The new players will be properly computed on the next scan
          }
        }
      } catch (provisionError) {
        Logger.log('Warning: Auto-provisioning failed: ' + provisionError.message);
        // Don't throw - provisioning failure shouldn't break the mission scan
      }
    }
    
    return updatedCount;
    
  } catch (error) {
    // Log error to Integrity_Log if available
    Logger.log('Error in syncAttendanceMissions: ' + error.message);
    
    if (typeof logIntegrityAction === 'function') {
      logIntegrityAction('ATTENDANCE_MISSIONS_SYNC', {
        details: 'Error: ' + error.message + '\n\nStack: ' + (error.stack || 'N/A'),
        status: 'FAILURE'
      });
    }
    
    throw error; // Re-throw to let caller handle
  }
}

/**
 * Adds menu item for manual sync
 */
function addAttendanceMissionMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 Missions')
    .addItem('Sync Attendance Missions', 'syncAttendanceMissions')
    .addItem('Test: Show Event Count', 'showEventCount')
    .addToUi();
}

function showEventCount() {
  const events = getEventSheets();
  SpreadsheetApp.getUi().alert('Found ' + events.length + ' event tabs');
}

// ════════════════════════════════════════════════════════════════════════════
// PHASE 6: TESTING
// ════════════════════════════════════════════════════════════════════════════

/**
 * Test function for attendance scanner
 */
function testAttendanceScanner() {
  Logger.log('=== ATTENDANCE SCANNER TEST ===');
  
  // Test event detection
  const events = getEventSheets();
  Logger.log('Event tabs found: ' + events.length);
  
  if (events.length > 0) {
    Logger.log('Sample event: ' + events[0].getName());
    const parsed = parseEventSheetName(events[0].getName());
    Logger.log('  Parsed: ' + JSON.stringify(parsed));
    
    const standings = getStandings(events[0]);
    Logger.log('  Standings count: ' + standings.length);
    Logger.log('  Top 4: ' + standings.slice(0, 4).join(', '));
  }
  
  // Test full scan
  const { events: allEvents, playerEvents } = scanAllEvents();
  Logger.log('');
  Logger.log('Full scan results:');
  Logger.log('  Total events: ' + allEvents.length);
  Logger.log('  Total players: ' + playerEvents.size);
  
  // Sample player stats
  if (playerEvents.size > 0) {
    const samplePlayer = playerEvents.keys().next().value;
    const stats = computePlayerStats(samplePlayer, playerEvents.get(samplePlayer));
    Logger.log('');
    Logger.log('Sample player: ' + samplePlayer);
    Logger.log('  Stats: ' + JSON.stringify(stats));
    
    const awards = computeAttendanceMissionPoints(stats);
    Logger.log('  Awards: ' + JSON.stringify(awards));
  }
  
  Logger.log('');
  Logger.log('=== TEST COMPLETE ===');
}
