# Player Provisioning System

## Overview
The Player Provisioning System ensures every player in `PreferredNames` has corresponding rows in all tracking sheets. It maintains data integrity by automatically provisioning new players discovered in event tabs.

## Files

### PlayerProvisioning.js
Main implementation with 14 functions:

**Core Functions:**
- `normalizePlayerName(name)` - Normalizes player names for consistent comparison
- `playerNamesMatch(a, b)` - Case-insensitive name matching
- `getAllPreferredNames()` - Gets all canonical player names
- `playerExists(name)` - Checks if player exists in PreferredNames
- `getSheetCI_(name)` - Case-insensitive sheet lookup
- `getHeaderMap_(sheet)` - Gets column name to index mapping

**Provisioning Functions:**
- `provisionToSheet_(sheetName, playerName, keyColumn, defaults)` - Provisions single player to one sheet
- `provisionPlayerProfile(playerName)` - Provisions player to ALL tracking sheets
- `addNewPlayer(name)` - Adds new player to PreferredNames and provisions everywhere
- `provisionMultiplePlayers(names)` - Batch provisions multiple players

**Discovery Functions:**
- `getUnprovisionedPlayers()` - Finds players in event tabs but not in PreferredNames
- `discoverAndProvisionNewPlayers()` - Scans and provisions unprovisioned players

**Menu Handler:**
- `runFullProvisioning()` - Menu handler for "Provision All Players"

**Helper:**
- `logProvisioningAction_(action, details)` - Logs to Integrity_Log if available

### PlayerProvisioning.test.js
Comprehensive test suite with 7 test functions:

- `testPlayerProvisioning()` - Main test runner
- `testNameNormalization_()` - Tests name normalization and matching
- `testPlayerExists_()` - Tests player existence check
- `testGetAllPreferredNames_()` - Tests getting all player names
- `testProvisionTargetsConfig_()` - Tests configuration structure
- `testSheetLookup_()` - Tests case-insensitive sheet lookup
- `quickValidateProvisioning()` - Quick function existence validator

### attendanceMissionScanner.js (modified)
Added auto-provisioning integration that:
- Discovers new players during attendance scan
- Automatically provisions them to all tracking sheets
- Logs provisioning results
- Handles errors gracefully without breaking the scan

## Provision Targets

The system provisions players to these sheets:

1. **BP_Total** - Bonus Points tracking
   - Key: `preferred_name_id`
   - Defaults: BP_Historical, BP_Redeemed, BP_Current, Flag_Points, Attendance_Points, Dice_Points, LastUpdated

2. **Attendance_Missions** - Mission progress tracking
   - Key: `PreferredName`
   - Defaults: All mission columns (First Contact, Stellar Explorer, etc.), Points

3. **Flag_Missions** - Flag achievement tracking
   - Key: `preferred_name_id`
   - Defaults: Cosmic_Selfie, Review_Writer, Social_Media_Star, App_Explorer, Cosmic_Merchant, Flag_Points, LastUpdated

4. **Dice_Points** - Dice roll points
   - Key: `preferred_name_id`
   - Defaults: Points, LastUpdated

5. **Key_Tracker** - MTG Key tracking
   - Key: `PreferredName`
   - Defaults: White, Blue, Black, Red, Green, Total, Rainbow

## Usage

### Manual Provisioning
```javascript
// Add a single new player
addNewPlayer('PlayerName');

// Provision multiple players
provisionMultiplePlayers(['Player1', 'Player2', 'Player3']);

// Provision all players from PreferredNames
runFullProvisioning();
```

### Automatic Provisioning
The system automatically provisions new players when:
- Running attendance mission scan via `syncAttendanceMissions()`
- New players are found in event tabs (pattern: MM-DD<suffix>-YYYY)

### Testing
```javascript
// Quick validation - checks all functions exist
quickValidateProvisioning();

// Full test suite
testPlayerProvisioning();
```

## Event Detection Pattern
Event tabs are detected using the pattern: `/^\d{2}-\d{2}[A-Za-z]+-\d{4}$/`

Examples:
- `05-10C-2025` ✓
- `12-01Draft-2025` ✓
- `07-26q-2025` ✓
- `InvalidFormat` ✗

## Features

### Case-Insensitive Matching
All player name comparisons are case-insensitive:
- "John Doe" matches "john doe" and "JOHN DOE"

### Name Normalization
Player names are normalized by:
1. Trimming whitespace
2. Collapsing multiple spaces to single spaces
3. Converting to consistent casing for comparison

### Error Handling
- Missing sheets are handled gracefully
- Missing columns are logged but don't crash
- Function-based defaults are evaluated at provision time
- All errors are logged for debugging

### Integration with Integrity_Log
If `logIntegrityAction()` is available, all provisioning actions are logged with:
- Action type
- Player names
- Results (created/existed counts)
- Errors

## Architecture

### PreferredNames as Source of Truth
- PreferredNames sheet is the canonical player registry
- All provisioning operations reference this sheet
- New players must be added to PreferredNames first

### Idempotent Operations
- Provisioning the same player multiple times is safe
- Already-existing rows are detected and skipped
- No duplicate entries are created

### Header-Driven Configuration
- Column names are looked up dynamically
- Multiple header name synonyms supported
- Case-insensitive and whitespace-insensitive matching

## Return Values

### provisionToSheet_
```javascript
{
  created: boolean,   // true if new row created
  existed: boolean,   // true if row already existed
  error: string|null  // error message if failed
}
```

### provisionPlayerProfile
```javascript
{
  createdBySheet: {},    // { sheetName: boolean }
  existedBySheet: {},    // { sheetName: boolean }
  skippedSheets: []      // sheet names that were skipped
}
```

### addNewPlayer
```javascript
{
  success: boolean,           // operation succeeded
  alreadyExisted: boolean,    // player was already in PreferredNames
  profileResult: Object       // result from provisionPlayerProfile
}
```

### discoverAndProvisionNewPlayers
```javascript
{
  newPlayersFound: number,  // count of unprovisioned players found
  provisioned: number,      // count of players provisioned
  errors: []                // array of error messages
}
```

## Deployment Checklist

- [x] PlayerProvisioning.js created with all functions
- [x] PlayerProvisioning.test.js created with test suite
- [x] attendanceMissionScanner.js modified with integration
- [x] All syntax checks passed
- [x] All functions verified present
- [ ] Deploy to Google Apps Script environment
- [ ] Run quickValidateProvisioning()
- [ ] Run testPlayerProvisioning()
- [ ] Test runFullProvisioning() menu item
- [ ] Verify auto-provisioning during attendance scan

## Version History

### 1.0.0 (2026-01-22)
- Initial implementation
- 14 core functions
- 5 provision targets configured
- Comprehensive test suite
- Integration with attendance scanner
- Auto-discovery and provisioning
