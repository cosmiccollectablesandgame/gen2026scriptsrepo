# 🌌 Cosmic Engine — Architecture Map & Rebuild Plan (v7.9.x)

> **Purpose**
> This document defines the authoritative architecture of the Cosmic Engine (Google Sheets + Apps Script), designates canonical vs legacy components, and outlines a rebuild plan to close known gaps in Bonus Points, Attendance, Missions, and Player Provisioning.

---

## 1️⃣ High-Level System Architecture

```
┌──────────────────────────────────────────┐
│            CONTROL PLANE                 │
│  (menus, routing, triggers, UI entry)    │
└───────────────┬──────────────────────────┘
                │
                ▼
┌──────────────────────────────────────────┐
│           APPLICATION SERVICES            │
│  (Commander, Attendance, Missions, BP)   │
└───────────────┬──────────────────────────┘
                │
                ▼
┌──────────────────────────────────────────┐
│            DATA PIPELINES (RIVERS)        │
│  Attendance / Dice / Flags → BP_Total    │
└───────────────┬──────────────────────────┘
                │
                ▼
┌──────────────────────────────────────────┐
│         LEDGERS & HISTORICAL LOGS         │
│  Integrity_Log, Spent_Pool, BP History   │
└──────────────────────────────────────────┘
```

---

## 2️⃣ Control Plane (Authoritative)

### ✅ CANONICAL

| File          | Role                                                         |
| ------------- | ------------------------------------------------------------ |
| **`Code.js`** | **ONLY file allowed to contain `onOpen(e)` and `onEdit(e)`** |

**Responsibilities**

* Build menus
* Route menu clicks to services
* Minimal, safe `onEdit` logic only

**Rules**

* ❌ No heavy computation in `onEdit`
* ❌ No trigger creation here
* ❌ No data aggregation logic here

---

### ⚠️ LEGACY (to eliminate or quarantine)

| Pattern                          | Action             |
| -------------------------------- | ------------------ |
| Any other `function onOpen()`    | ❌ Rename or delete |
| Any other `function onEdit()`    | ❌ Rename or delete |
| Trigger creation inside services | ❌ Remove           |

---

## 3️⃣ Menu Architecture (Canonical)

```
Cosmic Engine v7.9.8
├── Events
│   ├── Start New Event
│   ├── Commander Event Wizard
│   ├── Import Player List (Roster)
│   ├── View Event Index
│   ├── Preview End Prizes
│   ├── Lock In End Prizes
│   ├── Commander Round Prizes
│   └── Undo Last Prize Run
│
├── Players
│   ├── Add New Player
│   ├── Detect / Fix Player Names
│   ├── Player Lookup
│   └── Add Key
│
├── Bonus Points
│   ├── Award Bonus Points
│   ├── Redeem Bonus Points
│   ├── Sync BP from Sources (Canonical)  ← SINGLE ENTRY POINT
│   └── Provision All Players
│
├── Missions & Attendance
│   ├── Scan Attendance / Missions (Canonical)
│   ├── Rebuild Attendance Calendar
│   ├── Record Dice Roll Results
│   ├── Award Flag Mission
│   ├── Record Attendance
│   └── Validate Mission Points Integrity
│
├── Catalog
│   ├── Manage Prize Catalog
│   ├── Prize Throttle (Switchboard)
│   └── Import Preorder Allocation
│
├── Preorders
│   ├── Sell Preorder
│   ├── View Preorder Status
│   ├── Mark Preorder Pickup
│   ├── Cancel Preorder
│   ├── Manage Preorder Buckets
│   └── View Preorders Sold
│
├── Ops
│   ├── Daily Close Checklist
│   ├── Build Event Dashboard
│   ├── Ship-Gates Health Check
│   ├── Build / Repair
│   ├── Organize Tabs
│   ├── Clean Old Previews
│   ├── View Integrity Log
│   ├── View Spent Pool
│   ├── Export Reports
│   ├── Force Unlock Event
│   └── Emergency Revert
│
└── Admin / Diagnostics
    ├── BP Diagnostics
    ├── Attendance Diagnostics
    ├── Mission Diagnostics
    └── First Run Setup
```

**Design Rule**

> If staff can trigger it, it should be a **menu button**, not an `onEdit`.

---

## 4️⃣ Bonus Points (BP River) — CRITICAL SYSTEM

### ❌ CURRENT PROBLEM

Multiple competing implementations of:

```js
updateBPTotalFromSources()
```

Apps Script uses a **single global namespace**, so:

> whichever file loads last silently overwrites the others

---

### ✅ CANONICAL BP PIPELINE

| Layer        | File                      | Status      |
| ------------ | ------------------------- | ----------- |
| Entry Point  | `menuSyncBPFromSources()` | ✅ Canonical |
| Aggregator   | `bpTotalPipeline.js`      | ✅ Canonical |
| Ledger Write | `BonusPointsService.js`   | ✅ Canonical |

**Canonical Flow**

```
Attendance_Missions
Flag_Missions
Dice_Points
        ↓
updateBPTotalFromSources()  [bpTotalPipeline.js]
        ↓
BP_Total
        ↓
Spent_Pool / History
```

---

### ⚠️ LEGACY BP Functions (Action Required)

| File                     | Function                           | Action                    |
| ------------------------ | ---------------------------------- | ------------------------- |
| `MissionPointsService.js`| `syncBPTotals()`                   | ❌ Remove or rename to `_LEGACY_` |
| `MissionPointsService.js`| `ensureBPTotalConsolidatedSchema()`| ❌ Remove (duplicate)      |
| `MissionPointsService.js`| `migrateBPTotalSchema_()`          | ❌ Remove (duplicate)      |
| `MissionPointsService.js`| `provisionAllPlayers()`            | ❌ Remove (use PlayerProvisioning.js) |
| `MissionPointsService.js`| `getAllPreferredNames()`           | ❌ Remove (use PlayerProvisioning.js) |
| `MissionPointsService.js`| `ensurePreferredNamesSchema()`     | ❌ Remove (duplicate)      |

---

## 5️⃣ Attendance Calendar

### ❌ CURRENT HOLE

Calendar "misses events" due to **over-strict name matching**.

### ✅ CANONICAL

| File                            | Role                       |
| ------------------------------- | -------------------------- |
| `attendanceCallendarService.js` | Builds Attendance_Calendar |

**Current Regex**

```js
/^(\d{2})-(\d{2})([A-Z])?-(\d{4})$/
```

### 🔧 FIX REQUIRED

Update regex to support:

* lowercase suffixes: `[A-Za-z]?`
* single-digit dates: `\d{1,2}`
* Consistent pattern across all files

**Canonical Event Source**

> Event sheet name = event identity
> No metadata fallback.

---

## 6️⃣ Mission Scan & Mission Log

### ✅ CANONICAL COMPONENTS

| File                    | Role                    |
| ----------------------- | ----------------------- |
| `MissionScanService.js` | Event scan + mission evaluation |
| `OmegaAttendanceSystem.js` | Attendance aggregation (NEEDS FIX) |
| `MissionLog` sheet      | Historical record       |

### ❌ CRITICAL BUG

`OmegaAttendanceSystem.js` line 62 calls `scanAllEventSheets(ss)` which **does not exist**.

**Fix Required:**
```js
// Option A: Create the missing function
function scanAllEventSheets(ss) {
  // Use logic from MissionScanService.scanAllEvents_()
}

// Option B: Replace call with existing function
var eventData = scanAllEvents_(ss);  // from MissionScanService.js
```

**Canonical Flow**

```
Event Sheets
   ↓
Scan Attendance (MissionScanService.runMissionScan)
   ↓
Resolve Players (PreferredNames lookup)
   ↓
Award Missions
   ↓
MissionLog + Attendance_Missions
```

---

## 7️⃣ Player Identity & Provisioning

### ✅ CANONICAL IDENTITY

| Asset            | Role                   |
| ---------------- | ---------------------- |
| `PreferredNames` | Single source of truth |

### Name Resolution

| Case            | Action                   |
| --------------- | ------------------------ |
| Known name      | Normalize                |
| Nickname / typo | Flag for review          |
| Unknown         | Log to UndiscoveredNames |

### ✅ CANONICAL Provisioning

| File                   | Role                           |
| ---------------------- | ------------------------------ |
| `PlayerProvisioning.js`| Single provisioning engine     |

**Target Sheets:**

| Sheet               | Action  |
| ------------------- | ------- |
| PreferredNames      | Add row (source of truth) |
| Attendance_Missions | Add row |
| Dice_Points         | Add row |
| Flag_Missions       | Add row |
| BP_Total            | Add row |
| Key_Tracker         | Add row (optional) |

---

## 8️⃣ Commander Wizard

### ✅ CANONICAL

| File                        | Role                |
| --------------------------- | ------------------- |
| `CommanderWizardService.js` | Guided Commander UI |

**Relies On**

* Event metadata
* Integrity_Log
* Prize state inference

---

## 9️⃣ Files to DELETE

| File | Reason |
| ---- | ------ |
| `RunningBalance.js` | Incomplete stub, `getCurrentBalance()` duplicated elsewhere |
| `uiHandlersLegacy.js` | 790 lines marked DEPRECATED |
| `PlayerLookupTypes.js` | Exact duplicate of `playerLookupService.js` |

---

## 🔟 Canonical vs Legacy Summary Table

| Area         | Canonical                       | Legacy (Remove/Rename)           |
| ------------ | ------------------------------- | -------------------------------- |
| Triggers     | `Code.js` only                  | Any other file                   |
| BP Sync      | `bpTotalPipeline.js`            | MissionPointsService BP functions|
| Provisioning | `PlayerProvisioning.js`         | MissionPointsService provisioning|
| Attendance   | `attendanceCallendarService.js` | Ad-hoc scans                     |
| Missions     | `MissionScanService.js`         | Deprecated routes                |
| Integrity    | `integrityService.js`           | Duplicates in other files        |
| Store Credit | `storeCreditService.js`         | Duplicates in Code.js            |
| Player Lookup| `playerLookupService.js`        | PlayerLookupTypes.js (delete)    |

---

# 🔧 REBUILD PLAN (SAFE, SEQUENTIAL)

## Phase 1 — Stabilize (No Behavior Change)

- [ ] Delete `RunningBalance.js`
- [ ] Delete `uiHandlersLegacy.js`
- [ ] Delete `PlayerLookupTypes.js`
- [ ] Remove deprecated functions from `Code.js` (lines 1878-1893)
- [ ] Remove deprecated functions from `bpTotalPipeline.js` (lines 462-487)
- [ ] Remove BP sync from `onEdit` (already done in new Code.js)

## Phase 2 — Canonicalize Pipelines

- [ ] Remove duplicate BP functions from `MissionPointsService.js`
- [ ] Remove duplicate provisioning from `MissionPointsService.js`
- [ ] Consolidate `logIntegrityAction()` to `integrityService.js` only
- [ ] Consolidate utility functions to `utils.js` only
- [ ] Add sync triggers after flag/dice awards → call `updateBPTotalFromSources()`

## Phase 3 — Fix Attendance Visibility

- [ ] Update event regex to: `/^(\d{1,2})-(\d{1,2})([A-Za-z])?-(\d{4})$/i`
- [ ] Apply consistent regex in all files:
  - `attendanceCallendarService.js`
  - `eventService.js`
  - `attendanceConfig.js`
- [ ] Create `scanAllEventSheets()` function OR wire to `MissionScanService.scanAllEvents_()`
- [ ] Define missing `getPlayerBP()` and `setPlayerBP()` in `attendaceService.js`

## Phase 4 — Unify Provisioning

- [ ] Ensure `PlayerProvisioning.js` is the single provisioning engine
- [ ] Add provisioning call after mission scan discovers new players
- [ ] Add user confirmation before auto-creating players

## Phase 5 — Legacy Cleanup

- [ ] Remove unused menu item handlers
- [ ] Archive deprecated routes (rename to `_LEGACY_*`)
- [ ] Freeze legacy files with header comments
- [ ] Run full integration test

---

## 📝 Final Design Principles

> **Controls before convenience.**
> **One canonical path per system.**
> **Menus over magic.**
> **Logs over guesses.**

---

## 📊 File Inventory

### Core (Keep)
- `Code.js` - Control plane
- `bpTotalPipeline.js` - BP aggregation
- `BonusPointsService.js` - BP operations
- `PlayerProvisioning.js` - Player provisioning
- `MissionScanService.js` - Mission scanning
- `attendanceCallendarService.js` - Attendance calendar
- `CommanderWizardService.js` - Commander wizard
- `integrityService.js` - Audit logging
- `playerLookupService.js` - Player search
- `storeCreditService.js` - Store credit
- `utils.js` - Shared utilities

### Delete
- `RunningBalance.js`
- `uiHandlersLegacy.js`
- `PlayerLookupTypes.js`

### Heavy Refactor
- `MissionPointsService.js` - Remove duplicated functions
- `OmegaAttendanceSystem.js` - Add missing `scanAllEventSheets()`
- `attendaceService.js` - Add missing `getPlayerBP()`, `setPlayerBP()`

---

*Document Version: 1.0*
*Last Updated: 2026-01-11*
*Engine Version: 7.9.8*
