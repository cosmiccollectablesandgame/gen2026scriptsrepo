# 🔧 Cosmic Engine Rebuild Checklist

> Execute these changes in order. Each phase is safe to deploy independently.

---

## Phase 1: Stabilize (Safe Deletions)

### 1.1 Delete Redundant Files

```bash
# These files are safe to delete - no dependencies
rm RunningBalance.js        # Incomplete stub
rm uiHandlersLegacy.js      # Marked DEPRECATED
rm PlayerLookupTypes.js     # Exact duplicate of playerLookupService.js
```

- [ ] `RunningBalance.js` - DELETE
- [ ] `uiHandlersLegacy.js` - DELETE
- [ ] `PlayerLookupTypes.js` - DELETE

### 1.2 Remove Dead Code from Code.js

Remove these deprecated wrapper functions (they're never called):

| Lines | Function | Action |
|-------|----------|--------|
| 1878-1880 | `onScanMissions_deprecated()` | DELETE |
| 1884-1886 | `onRecordAttendance_deprecated()` | DELETE |
| 1890-1893 | `onRefreshBPTotalFromSources_deprecated()` | DELETE |

- [ ] Remove `onScanMissions_deprecated()`
- [ ] Remove `onRecordAttendance_deprecated()`
- [ ] Remove `onRefreshBPTotalFromSources_deprecated()`

### 1.3 Remove Deprecated Functions from bpTotalPipeline.js

| Lines | Function | Action |
|-------|----------|--------|
| 462-473 | `syncBPTotals()` | DELETE (duplicate) |
| 475-482 | `ensureBPTotalConsolidatedSchema()` | DELETE (duplicate) |
| 484-487 | `migrateBPTotalSchema_()` | DELETE (duplicate) |

- [ ] Remove `syncBPTotals()` from bpTotalPipeline.js
- [ ] Remove `ensureBPTotalConsolidatedSchema()` from bpTotalPipeline.js
- [ ] Remove `migrateBPTotalSchema_()` from bpTotalPipeline.js

### 1.4 Remove Stub from playerPipelineService.js

| Lines | Function | Action |
|-------|----------|--------|
| 527-542 | `completeMission()` | DELETE (stub, never called) |

- [ ] Remove `completeMission()` stub

---

## Phase 2: Canonicalize Pipelines

### 2.1 Clean MissionPointsService.js (Heavy Refactor)

Remove these duplicated functions (keep originals in canonical files):

| Function | Line | Keep In | Action |
|----------|------|---------|--------|
| `syncBPTotals()` | 924 | bpTotalPipeline.js | DELETE |
| `ensureBPTotalConsolidatedSchema()` | 800 | bpTotalPipeline.js | DELETE |
| `migrateBPTotalSchema_()` | 855 | bpTotalPipeline.js | DELETE |
| `provisionAllPlayers()` | 1252 | PlayerProvisioning.js | DELETE |
| `ensurePreferredNamesSchema()` | 59 | PlayerProvisioning.js | DELETE |
| `getAllPreferredNames()` | 103 | PlayerProvisioning.js | DELETE |
| `awardDicePoints()` | 666 | dicePointsBackEndService.js | DELETE |
| `awardFlagMission()` | 317 | flagMissionService.js | DELETE |
| `getCanonicalNames()` | 1304 | Keep ONE in utils.js | DELETE |

- [ ] Remove `syncBPTotals()` from MissionPointsService.js
- [ ] Remove `ensureBPTotalConsolidatedSchema()` from MissionPointsService.js
- [ ] Remove `migrateBPTotalSchema_()` from MissionPointsService.js
- [ ] Remove `provisionAllPlayers()` from MissionPointsService.js
- [ ] Remove `ensurePreferredNamesSchema()` from MissionPointsService.js
- [ ] Remove `getAllPreferredNames()` from MissionPointsService.js
- [ ] Remove `awardDicePoints()` from MissionPointsService.js
- [ ] Remove `awardFlagMission()` from MissionPointsService.js
- [ ] Remove `getCanonicalNames()` from MissionPointsService.js

### 2.2 Consolidate logIntegrityAction()

Keep ONLY in `integrityService.js`, remove from:

| File | Line | Action |
|------|------|--------|
| Code.js | 1482 | DELETE (or delegate to integrityService) |
| eventService.js | 740 | DELETE |
| PlayerNameService.js | 805 | DELETE |
| MissionHelpers.js | 137 | DELETE |

- [ ] Remove/delegate `logIntegrityAction()` from Code.js
- [ ] Remove `logIntegrityAction()` from eventService.js
- [ ] Remove `logIntegrityAction()` from PlayerNameService.js
- [ ] Remove `logIntegrityAction()` from MissionHelpers.js

### 2.3 Consolidate Utility Functions to utils.js

Remove duplicates, keep only in `utils.js`:

| Function | Remove From |
|----------|-------------|
| `formatEventDate()` | eventService.js:711 |
| `generateSeed()` | eventService.js:722 |
| `formatCurrency()` | eventService.js:731 |
| `coerceNumber()` | MissionHelpers.js:22 |
| `currentUser()` | Code.js:1836 |

- [ ] Consolidate `formatEventDate()` to utils.js
- [ ] Consolidate `generateSeed()` to utils.js
- [ ] Consolidate `formatCurrency()` to utils.js
- [ ] Consolidate `coerceNumber()` to utils.js
- [ ] Consolidate `currentUser()` to utils.js

### 2.4 Wire Sync Triggers

Add `updateBPTotalFromSources()` call after:

**In flagMissionService.js after `syncFlagMissionsRow()`:**
```javascript
// At end of syncFlagMissionsRow() function
if (typeof updateBPTotalFromSources === 'function') {
  updateBPTotalFromSources();
}
```

**In dicePointsBackEndService.js after `awardDicePoints()`:**
```javascript
// At end of awardDicePoints() function
if (typeof updateBPTotalFromSources === 'function') {
  updateBPTotalFromSources();
}
```

- [ ] Add BP sync after flag mission award
- [ ] Add BP sync after dice points award

---

## Phase 3: Fix Attendance Visibility

### 3.1 Standardize Event Regex

Update to flexible pattern in ALL files:

```javascript
// NEW: Supports MM-DD-YYYY, M-D-YYYY, MM-DDx-YYYY (any case suffix)
const EVENT_PATTERN = /^(\d{1,2})-(\d{1,2})([A-Za-z])?-(\d{4})$/;
```

Files to update:
- [ ] `attendanceCallendarService.js` line 135
- [ ] `eventService.js` line 20
- [ ] `attendanceConfig.js` line 297
- [ ] `OmegaAttendanceSystem.js` (any pattern references)

### 3.2 Create Missing scanAllEventSheets()

Add to `OmegaAttendanceSystem.js` (or delegate to MissionScanService):

```javascript
/**
 * Scans all event sheets and returns structured event data.
 * Required by runOmegaAttendanceScan() at line 62.
 * @param {Spreadsheet} ss - The spreadsheet
 * @return {Object} eventData with events, players, playerEventHistory
 */
function scanAllEventSheets(ss) {
  // Delegate to MissionScanService if available
  if (typeof scanAllEvents_ === 'function') {
    return scanAllEvents_(ss);
  }

  // Fallback implementation
  const sheets = ss.getSheets();
  const eventPattern = /^(\d{1,2})-(\d{1,2})([A-Za-z])?-(\d{4})$/;

  const events = [];
  const allPlayers = new Set();
  const playerEventHistory = new Map();

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (!eventPattern.test(name)) return;

    // Extract event data...
    // (full implementation needed)
  });

  return { events, players: allPlayers, playerEventHistory };
}
```

- [ ] Create `scanAllEventSheets()` function

### 3.3 Create Missing BP Functions in attendaceService.js

Add these functions (or delegate to BonusPointsService):

```javascript
/**
 * Gets current BP for a player.
 * @param {string} playerName
 * @return {number}
 */
function getPlayerBP(playerName) {
  if (typeof getPlayerBP_ === 'function') {
    return getPlayerBP_(playerName);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BP_Total');
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const bpCol = headers.indexOf('Current_BP');

  if (nameCol === -1 || bpCol === -1) return 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === playerName) {
      return Number(data[i][bpCol]) || 0;
    }
  }
  return 0;
}

/**
 * Sets BP for a player.
 * @param {string} playerName
 * @param {number} newBP
 */
function setPlayerBP(playerName, newBP) {
  if (typeof setPlayerBP_ === 'function') {
    return setPlayerBP_(playerName, newBP);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BP_Total');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const bpCol = headers.indexOf('Current_BP');

  if (nameCol === -1 || bpCol === -1) return;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === playerName) {
      sheet.getRange(i + 1, bpCol + 1).setValue(newBP);
      return;
    }
  }
}
```

- [ ] Add `getPlayerBP()` to attendaceService.js
- [ ] Add `setPlayerBP()` to attendaceService.js

---

## Phase 4: Unify Provisioning

### 4.1 Ensure Single Provisioning Engine

Verify `PlayerProvisioning.js` has:
- [ ] `addNewPlayer(name)` - canonical entry point
- [ ] `provisionAllPlayers()` - bulk provisioning
- [ ] `provisionSinglePlayerProfile_(name)` - internal helper

### 4.2 Wire Provisioning to Mission Scan

In `MissionScanService.js`, after discovering unknown players:

```javascript
// After resolving names, provision any new players
if (typeof provisionSinglePlayerProfile_ === 'function') {
  newPlayers.forEach(name => provisionSinglePlayerProfile_(name));
}
```

- [ ] Add provisioning call after mission scan

### 4.3 Fix UndiscoveredNamesService Column Detection

Change line 82 from:
```javascript
const rawName = values[r][1]; // column B (0-based index)
```

To header-based detection:
```javascript
const headers = values[0];
const nameCol = headers.findIndex(h =>
  ['preferredname', 'preferred_name_id', 'player', 'name']
    .includes(String(h).toLowerCase().replace(/[_\s]/g, ''))
);
const rawName = values[r][nameCol !== -1 ? nameCol : 1];
```

- [ ] Fix column detection in UndiscoveredNamesService.js

---

## Phase 5: Legacy Cleanup

### 5.1 Archive Unused Functions

Rename any remaining legacy functions to `_LEGACY_*` pattern:

```javascript
// Example
function syncBPTotals_LEGACY() {
  console.warn('LEGACY: Use updateBPTotalFromSources() instead');
  return updateBPTotalFromSources();
}
```

- [ ] Rename remaining legacy functions

### 5.2 Add Deprecation Headers

Add to top of legacy files:

```javascript
/**
 * @deprecated This file is scheduled for removal.
 * Functions have been migrated to canonical services.
 * DO NOT add new code here.
 */
```

- [ ] Add deprecation headers to legacy files

### 5.3 Final Validation

- [ ] Run `menuSyncBPFromSources()` - verify BP sync works
- [ ] Run `onScanAttendance()` - verify mission scan works
- [ ] Run `onProvisionAllPlayers()` - verify provisioning works
- [ ] Run `onRebuildAttendanceCalendar()` - verify calendar rebuild works
- [ ] Check for console errors in Apps Script editor

---

## Quick Reference: What Stays vs Goes

### ✅ KEEP (Canonical)
```
Code.js                      # Control plane
bpTotalPipeline.js           # BP aggregation
BonusPointsService.js        # BP operations
PlayerProvisioning.js        # Player provisioning
MissionScanService.js        # Mission scanning
attendanceCallendarService.js # Attendance calendar
CommanderWizardService.js    # Commander wizard
integrityService.js          # Audit logging
playerLookupService.js       # Player search
storeCreditService.js        # Store credit
utils.js                     # Utilities
flagMissionService.js        # Flag missions
dicePointsBackEndService.js  # Dice points
eventService.js              # Event management
```

### ❌ DELETE
```
RunningBalance.js
uiHandlersLegacy.js
PlayerLookupTypes.js
```

### ⚠️ HEAVY REFACTOR
```
MissionPointsService.js      # Remove ~500 lines of duplicates
OmegaAttendanceSystem.js     # Add scanAllEventSheets()
attendaceService.js          # Add getPlayerBP(), setPlayerBP()
UndiscoveredNamesService.js  # Fix column detection
```

---

*Checklist Version: 1.0*
*Created: 2026-01-11*
