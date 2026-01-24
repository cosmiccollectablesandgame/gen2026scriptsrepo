# Phase 5: Transactional Truth Consumers - Implementation Guide

> **Status**: ✅ COMPLETE  
> **Version**: 1.0.0  
> **Last Updated**: 2026-01-24

---

## Table of Contents
1. [Overview](#overview)
2. [Core Concepts](#core-concepts)
3. [Architecture](#architecture)
4. [Getting Started](#getting-started)
5. [Public Functions](#public-functions)
6. [Workflows](#workflows)
7. [Configuration](#configuration)
8. [Testing](#testing)
9. [Troubleshooting](#troubleshooting)

---

## Overview

Phase 5 extends the identity hygiene system (Phases 1-4) into transactional systems, ensuring that **Store Credit** and **Preorders** use canonical player names from **PreferredNames**.

### Key Principles

```
PreferredNames is the registrar.
Transactions must name real people.
PlayerLookup is the unified read model.
```

### What Phase 5 Adds

1. **Retail-Friendly Store Credit** - Allows unknown names but queues them for onboarding
2. **Strict Preorder Enforcement** - Requires canonical names to prevent mismatches
3. **Expanded PlayerLookup** - Aggregates Store Credit balances and Preorder totals
4. **Staff Dashboards** - Resolve Unknowns Dashboard and Provisional Ledger Names
5. **Onboarding Workflow** - `CREATE_NEW_PLAYER_FROM_UNDISCOVERED()` for retail staff
6. **Comprehensive Auditing** - Phase 5 audit reports and hygiene tracking

---

## Core Concepts

### 1. Canonical Names (PreferredNames)

**PreferredNames** is the single source of truth for player identity.

- Lives in `PreferredNames` sheet, column A
- All transactional and derived systems reference this registry
- Case-insensitive comparison, whitespace normalized
- Exact spelling preserved for display

### 2. Unknown Names (UndiscoveredNames)

**UndiscoveredNames** is the intake queue for names not yet in PreferredNames.

Enhanced schema in Phase 5:
```
NormalizedName    - Normalized form
RawName          - Original spelling
SourceType       - EVENT_SCAN | STORE_CREDIT_LEDGER | PREORDERS
SourceSheets     - Where discovered
SourceRefs       - Row numbers or IDs
SeenCount        - How many times seen
FirstSeen        - First discovery timestamp
LastSeen         - Last discovery timestamp
Status           - OPEN | RESOLVED
ResolvedAs       - Canonical name (when resolved)
ResolutionType   - ADD_CANONICAL | MAP_EXISTING | IGNORE
ResolvedBy       - Who resolved it
ResolvedAt       - When resolved
```

### 3. Enforcement Modes

**Store_Credit_Ledger**: RETAIL_FRIENDLY
- ✅ Allows unknown names on write
- ✅ Queues unknowns to UndiscoveredNames
- ✅ Transactions proceed normally
- ❌ Does NOT block PlayerLookup rollups

**Preorders**: STRICT
- ❌ Blocks unknown names on write
- ✅ Queues unknowns to UndiscoveredNames
- ✅ Returns clear error message
- ✅ Blocks PlayerLookup rollups when dirty

### 4. Hygiene Status

Each transactional sheet has a hygiene status:
- **CLEAN**: All names are canonical
- **DIRTY**: Has unresolved unknown names

Status stored in Document Properties:
```javascript
{ 
  clean: true/false,
  lastChecked: "ISO timestamp"
}
```

---

## Architecture

### Data Flow

```
┌─────────────────┐
│ PreferredNames  │ ← Canonical Registry
└────────┬────────┘
         │
         ├─────────────────────┬──────────────────┐
         │                     │                  │
┌────────▼────────┐   ┌───────▼──────┐   ┌──────▼────────┐
│ Store_Credit    │   │  Preorders   │   │  Events       │
│ Ledger          │   │  (Strict)    │   │  (Legacy)     │
│ (Retail-Friend) │   └──────┬───────┘   └───────────────┘
└────────┬────────┘          │
         │                   │
         │    ┌──────────────┴────────────┐
         │    │                           │
         │    │   Unknown Names?          │
         │    │                           │
         │    └──────────┬────────────────┘
         │               │
         │       ┌───────▼──────────┐
         │       │ UndiscoveredNames│
         │       │    (Queue)       │
         │       └───────┬──────────┘
         │               │
         │               │ Staff resolves
         │               │
         │       ┌───────▼──────────┐
         │       │  Resolve Unknowns│
         │       │   Dashboard      │
         │       └──────────────────┘
         │
         │
┌────────▼────────────────────────────┐
│         PlayerLookup                │
│  (Unified Read Model)               │
│  - StoreCreditBalance               │
│  - OpenPreorderCount                │
│  - OpenPreorderQtyTotal             │
│  - DepositsTotal                    │
│  - BalanceDueTotal                  │
│  - LedgerClean / PreordersClean     │
└─────────────────────────────────────┘
```

### File Structure

```
Phase5Service.js              - Core service (all Phase 5 functions)
Phase5Service.test.js         - Test suite
storeCreditService.js         - Updated with hygiene queueing
preordersService.js           - Updated with strict enforcement
playerLookupService.js        - Updated to read Phase 5 sheets
NameHygineService.js          - Updated to exclude retired sheets
```

---

## Getting Started

### Prerequisites

1. **PreferredNames sheet exists** with canonical player names in column A
2. **UndiscoveredNames sheet** will be auto-created on first scan
3. **Store_Credit_Ledger** and/or **Preorders_Sold** sheets exist

### Initial Setup

#### Step 1: Scan for Unknown Names

Run these functions to discover any unknown names:

```javascript
// Scan Store Credit Ledger
SCAN_STORE_CREDIT_LEDGER_FOR_UNKNOWN_NAMES();

// Scan Preorders
SCAN_PREORDERS_FOR_UNKNOWN_NAMES();
```

**Result**: Unknown names queued to `UndiscoveredNames` sheet

#### Step 2: Review Unknown Names

Generate the dashboard to see what needs resolution:

```javascript
REFRESH_RESOLVE_UNKNOWN_DASHBOARD();
```

Open the `Resolve_Unknowns_Dashboard` sheet to see categorized unknowns.

#### Step 3: Resolve Unknowns

For each unknown name, decide:

**Option A: Create New Player**
```javascript
CREATE_NEW_PLAYER_FROM_UNDISCOVERED('john smith', 'John Smith');
```

**Option B: Manual mapping** (if name is misspelled)
- Use existing `mapPlayerName()` function from Phase 1-4

#### Step 4: Build PlayerLookup

Once unknowns are resolved:

```javascript
SAFE_RUN_PLAYERLOOKUP_BUILD();
```

**Result**: `PlayerLookup` sheet with Store Credit balances and Preorder totals

---

## Public Functions

### Identity & Normalization

#### `normalizePlayerName(name)`
Normalizes a player name for consistent comparison.

**Parameters**:
- `name` (string): Raw player name

**Returns**: (string) Normalized name

**Example**:
```javascript
normalizePlayerName('  John  Doe  ');  // "John Doe"
normalizePlayerName('...Mary-Jane...'); // "Mary-Jane"
```

---

#### `isCanonicalName(name)`
Checks if a name exists in PreferredNames.

**Parameters**:
- `name` (string): Name to check

**Returns**: (boolean) True if canonical

**Example**:
```javascript
isCanonicalName('John Smith');  // true/false
```

---

#### `getCanonicalName(name)`
Gets the exact canonical spelling for a name.

**Parameters**:
- `name` (string): Name to look up

**Returns**: (string|null) Canonical name or null

**Example**:
```javascript
getCanonicalName('john smith');  // "John Smith"
```

---

### Scanning Functions

#### `SCAN_STORE_CREDIT_LEDGER_FOR_UNKNOWN_NAMES([scanWindowRows])`
Scans Store_Credit_Ledger for unknown names.

**Parameters**:
- `scanWindowRows` (number, optional): Recent rows to scan (default: 1000)

**Returns**: (Object)
```javascript
{
  scannedRows: 150,
  unknownNames: ['Jane Doe', 'Bob Smith'],
  queuedCount: 2,
  message: "Found 2 unknown names (queued 2)"
}
```

**Example**:
```javascript
const result = SCAN_STORE_CREDIT_LEDGER_FOR_UNKNOWN_NAMES();
console.log(result.message);
```

---

#### `SCAN_PREORDERS_FOR_UNKNOWN_NAMES([scanWindowRows])`
Scans Preorders for unknown names.

**Parameters**:
- `scanWindowRows` (number, optional): Recent rows to scan (default: 1000)

**Returns**: (Object) - Same structure as ledger scan

---

### Onboarding Workflow

#### `CREATE_NEW_PLAYER_FROM_UNDISCOVERED(normalizedName, [canonicalName])`
Creates a new player from an undiscovered name.

**Parameters**:
- `normalizedName` (string): Normalized name from UndiscoveredNames
- `canonicalName` (string, optional): Override canonical spelling

**Returns**: (Object)
```javascript
{
  success: true,
  message: "Successfully created player \"John Smith\"",
  canonicalName: "John Smith"
}
```

**Example**:
```javascript
// Create with normalized name
CREATE_NEW_PLAYER_FROM_UNDISCOVERED('john smith');

// Create with specific spelling
CREATE_NEW_PLAYER_FROM_UNDISCOVERED('john smith', 'John Smith');
```

**What it does**:
1. Adds canonical name to PreferredNames
2. Marks UndiscoveredNames entry as RESOLVED
3. Triggers provisioning (if available)
4. Logs to Integrity_Log

---

### Dashboard Functions

#### `REFRESH_RESOLVE_UNKNOWN_DASHBOARD()`
Generates staff-friendly dashboard of open unknowns.

**Returns**: (Object)
```javascript
{
  success: true,
  message: "Dashboard refreshed: 12 open items",
  sections: {
    storeCreditLedger: 5,
    events: 4,
    preorders: 2,
    other: 1
  }
}
```

**Creates sheet**: `Resolve_Unknowns_Dashboard` with sections:
1. **Retail-first (Store Credit)** - "Create New Player"
2. **Event-first (Events/Scans)** - "Spellcheck + Canonicalize"
3. **Preorders** - "Must Resolve Before Processing"
4. **Other** - "Review Manually"

---

#### `generateProvisionalLedgerNames()`
Shows non-canonical Store Credit customers awaiting onboarding.

**Returns**: (Object)
```javascript
{
  success: true,
  message: "Generated 8 provisional ledger names",
  provisionalCount: 8
}
```

**Creates sheet**: `Provisional_Ledger_Names`

Columns:
- RawName
- NormalizedName
- LedgerBalance
- TransactionCount
- FirstSeen / LastSeen
- Status (OPEN/RESOLVED)
- ActionHint

---

### PlayerLookup

#### `SAFE_RUN_PLAYERLOOKUP_BUILD()`
Builds PlayerLookup with transactional rollups (enforced).

**Returns**: (Object)
```javascript
{
  success: true,
  message: "PlayerLookup built successfully with 245 players",
  playerCount: 245,
  ledgerClean: true,
  preordersClean: true
}
```

**Creates/Updates sheet**: `PlayerLookup`

Columns:
- PreferredName
- StoreCreditBalance
- OpenPreorderCount
- OpenPreorderQtyTotal
- DepositsTotal
- BalanceDueTotal
- LedgerClean (TRUE/FALSE)
- PreordersClean (TRUE/FALSE)
- LastRefresh

**Enforcement**:
- Fails if Preorders have unresolved names (when `blocksPlayerLookupRollupsWhenDirty: true`)
- Store Credit unknowns do NOT block (retail-friendly)

---

### Auditing

#### `GENERATE_PHASE5_AUDIT_REPORT()`
Generates comprehensive audit report.

**Returns**: (string) Formatted report text

**Example output**:
```
═══════════════════════════════════════════════════════════
PHASE 5 TRUTH CONSUMERS AUDIT REPORT
═══════════════════════════════════════════════════════════

Timestamp: 2026-01-24T12:30:00.000Z

─── CANONICAL NAMES (PreferredNames) ───
  Count: 245
  Exists: true

─── UNDISCOVERED NAMES (Identity Queue) ───
  Total: 12
  Open: 5
  Resolved: 7

─── TRANSACTIONAL TRUTH CONSUMERS ───
  Store_Credit_Ledger:
    Exists: true
    Rows: 1823
    Hygiene: CLEAN
    Enforcement: RETAIL_FRIENDLY
    Allows Unknown: true
    
  Preorders:
    Exists: true
    Sheet: Preorders_Sold
    Rows: 234
    Hygiene: CLEAN
    Enforcement: STRICT
    Allows Unknown: false

─── DERIVED TRUTH CONSUMERS ───
  PlayerLookup:
    Exists: true
    Rows: 245
    Status: Has data
    
  Provisional_Ledger_Names: 3 rows
  Resolve_Unknowns_Dashboard: 5 rows

─── RETIRED SHEETS (Legacy - Read Only) ───
  Players_Prize-Wall-Points: 0 rows

═══════════════════════════════════════════════════════════
```

---

## Workflows

### Workflow 1: Retail Customer with Store Credit

**Scenario**: Walk-in customer "Jane Doe" wants to earn store credit, but she's not in the system yet.

**Steps**:
1. **Staff grants store credit** (via Store Credit UI)
   ```javascript
   logStoreCreditTransaction({
     preferred_name_id: 'Jane Doe',
     direction: 'IN',
     amount: 10,
     reason: 'Facebook share'
   });
   ```
   - Transaction succeeds ✅
   - Name queued to UndiscoveredNames ✅

2. **Later: Staff scans ledger**
   ```javascript
   SCAN_STORE_CREDIT_LEDGER_FOR_UNKNOWN_NAMES();
   ```
   - Confirms "Jane Doe" is unknown
   - Status: DIRTY

3. **Staff reviews dashboard**
   ```javascript
   REFRESH_RESOLVE_UNKNOWN_DASHBOARD();
   ```
   - "Jane Doe" appears in Retail-first section
   - Action: "Create New Player"

4. **Staff creates new player**
   ```javascript
   CREATE_NEW_PLAYER_FROM_UNDISCOVERED('Jane Doe');
   ```
   - Added to PreferredNames ✅
   - UndiscoveredNames marked RESOLVED ✅
   - Provisioned to tracking sheets ✅

5. **Staff rebuilds PlayerLookup**
   ```javascript
   SAFE_RUN_PLAYERLOOKUP_BUILD();
   ```
   - Jane Doe now has a complete profile with Store Credit balance

---

### Workflow 2: Preorder with Unknown Customer

**Scenario**: Customer "Bob Smith" tries to place a preorder, but he's not in PreferredNames.

**Steps**:
1. **Staff attempts to create preorder**
   ```javascript
   sellPreorder({
     customerName: 'Bob Smith',
     basket: [{ itemName: 'Booster Box', qty: 1, unitPrice: 120 }],
     totalDue: 120,
     depositAmount: 60
   });
   ```
   - **BLOCKED** ❌
   - Error: "Customer name \"Bob Smith\" is not in PreferredNames"
   - Name queued to UndiscoveredNames

2. **Staff reviews error** and has two options:

   **Option A: Create new player**
   ```javascript
   CREATE_NEW_PLAYER_FROM_UNDISCOVERED('Bob Smith');
   ```
   Then retry preorder

   **Option B: Fix spelling** (if typo)
   - Check PreferredNames for similar names
   - Use correct spelling in preorder

3. **After resolution, retry preorder**
   - Succeeds with canonical name ✅

---

### Workflow 3: Weekly Hygiene Maintenance

**Scenario**: Weekly cleanup to keep system tidy.

**Steps**:
1. **Run scans**
   ```javascript
   SCAN_STORE_CREDIT_LEDGER_FOR_UNKNOWN_NAMES();
   SCAN_PREORDERS_FOR_UNKNOWN_NAMES();
   ```

2. **Review dashboard**
   ```javascript
   REFRESH_RESOLVE_UNKNOWN_DASHBOARD();
   ```

3. **Resolve unknowns** (batch or individual)

4. **Rebuild PlayerLookup**
   ```javascript
   SAFE_RUN_PLAYERLOOKUP_BUILD();
   ```

5. **Generate audit**
   ```javascript
   GENERATE_PHASE5_AUDIT_REPORT();
   ```

6. **Review provisional names** (if needed)
   ```javascript
   generateProvisionalLedgerNames();
   ```

---

## Configuration

### Enforcement Modes

Edit `TRUTH_CONSUMERS` in `Phase5Service.js`:

```javascript
const TRUTH_CONSUMERS = {
  storeCreditLedger: {
    allowUnknownNamesOnWrite: true,        // Retail-friendly
    mustQueueUnknowns: true,
    blocksPlayerLookupRollupsWhenDirty: false,  // Don't block
    sheetNames: ['Store_Credit_Ledger'],
    scanWindowRows: 1000
  },
  preorders: {
    allowUnknownNamesOnWrite: false,       // Strict
    mustQueueUnknowns: true,
    blocksPlayerLookupRollupsWhenDirty: true,   // Block if dirty
    sheetNames: ['Preorders', 'PreOrders', 'Preorder_Requests', 'Preorders_Sold'],
    scanWindowRows: 1000
  }
};
```

### Retired Sheets

Edit `RETIRED_SHEETS` in `Phase5Service.js`:

```javascript
const RETIRED_SHEETS = [
  'Players_Prize-Wall-Points',
  'Player\'s Prize-Wall-Points'
];
```

---

## Testing

### Run Unit Tests

```javascript
RUN_PHASE5_TESTS();
```

Tests:
- ✓ normalizePlayerName()
- ✓ namesMatch()
- ✓ isCanonicalName()
- ✓ queueUnknownName()
- ✓ Hygiene status functions
- ✓ isRetiredSheet()

### Run Integration Test

```javascript
TEST_PHASE5_INTEGRATION();
```

Workflow:
1. Scans Store_Credit_Ledger
2. Scans Preorders
3. Refreshes dashboard
4. Generates provisional names
5. Runs audit

---

## Troubleshooting

### Problem: "PreferredNames sheet not found"

**Solution**: Create PreferredNames sheet with names in column A

```javascript
// Manual fix
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.insertSheet('PreferredNames');
sheet.appendRow(['PreferredName']);
sheet.appendRow(['John Smith']);
sheet.appendRow(['Jane Doe']);
```

---

### Problem: Preorder blocked with "not in PreferredNames"

**Cause**: Strict enforcement mode

**Solution**:
1. Check spelling
2. Create player if new:
   ```javascript
   CREATE_NEW_PLAYER_FROM_UNDISCOVERED('customer name');
   ```
3. Retry preorder

---

### Problem: PlayerLookup build fails with "blocked"

**Cause**: Preorders have unresolved names

**Solution**:
1. Run scan:
   ```javascript
   SCAN_PREORDERS_FOR_UNKNOWN_NAMES();
   ```
2. Resolve unknowns
3. Retry build

---

### Problem: Unknown names not appearing in dashboard

**Cause**: Status is RESOLVED or scan not run recently

**Solution**:
1. Re-run scans
2. Check UndiscoveredNames Status column (should be OPEN)
3. Refresh dashboard

---

### Problem: Store Credit balance is zero in PlayerLookup

**Possible causes**:
1. Name mismatch (different spelling)
2. Ledger uses different column name

**Solution**:
1. Check Store_Credit_Ledger headers (should have 'preferred_name_id' or 'PreferredName')
2. Verify exact name spelling matches PreferredNames
3. Check RunningBalance column exists

---

## Support

For issues or questions:
1. Check Integrity_Log for error details
2. Run audit: `GENERATE_PHASE5_AUDIT_REPORT()`
3. Review UndiscoveredNames Status column
4. Check hygiene status: `getSheetHygieneStatus('Store_Credit_Ledger')`

---

## Version History

**v1.0.0** (2026-01-24)
- Initial Phase 5 implementation
- Store Credit retail-friendly mode
- Preorder strict enforcement
- PlayerLookup expansion
- Dashboards and reporting
- Test suite

---

**End of Phase 5 Implementation Guide**
