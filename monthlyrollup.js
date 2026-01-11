/**
 * ╔══════════════════════════════════════════════════════════════════════════════╗
 * ║          COSMIC GAMES - EVENT ANALYTICS DASHBOARD v2.0                       ║
 * ╠══════════════════════════════════════════════════════════════════════════════╣
 * ║  Features:                                                                    ║
 * ║  • Per-event + Monthly rollups with cost tracking                            ║
 * ║  • Player retention analysis (new vs returning)                              ║
 * ║  • Day-of-week patterns                                                      ║
 * ║  • Event type breakdown (by suffix: A, B, C, etc.)                           ║
 * ║  • Player frequency distribution & top players                               ║
 * ║  • Sparklines + embedded charts                                              ║
 * ║  • Anomaly detection (attendance/cost outliers)                              ║
 * ║  • Growth rate calculations                                                  ║
 * ╚══════════════════════════════════════════════════════════════════════════════╝
 */

// =============================================================================
// CONFIG (namespaced to avoid collisions with other scripts)
// =============================================================================
const DASHBOARD_CONFIG = {
  DEFAULT_AVG_COST_PER_PLAYER: 8.00,
  REPORT_SHEET_NAME: 'Event_Dashboard',
  INPUT_CELL: 'E7',              // Cell where user can edit cost per player
  INPUT_CELL_ROW: 7,
  INPUT_CELL_COL: 5,             // Column E = 5
  ANOMALY_THRESHOLD_STDEV: 2.0,  // Flag if > 2 std devs from mean
  TOP_PLAYERS_COUNT: 15,
  FREQUENCY_BUCKETS: [1, 2, 3, 5, 10], // "1 event", "2 events", "3-4", "5-9", "10+"
  
  // Colors for conditional formatting
  COLORS: {
    HEADER_BG: '#1a237e',      // Deep indigo
    HEADER_TEXT: '#ffffff',
    SECTION_BG: '#3949ab',     // Lighter indigo
    SECTION_TEXT: '#ffffff',
    KPI_GOOD: '#c8e6c9',       // Light green
    KPI_WARNING: '#fff9c4',    // Light yellow
    KPI_BAD: '#ffcdd2',        // Light red
    CHART_PRIMARY: '#3949ab',
    CHART_SECONDARY: '#7986cb',
    CHART_ACCENT: '#ff7043',
    ALTERNATING_ROW: '#e8eaf6',
    SPARKLINE_UP: '#4caf50',
    SPARKLINE_DOWN: '#f44336',
    INPUT_BG: '#fff3e0',       // Light orange for input cells
    INPUT_BORDER: '#ff9800'    // Orange border for input cells
  }
};

// =============================================================================
// UPDATE COST PER PLAYER - Prompts user and rebuilds dashboard
// =============================================================================
function updateCostPerPlayer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Check if dashboard exists
  const dashboardSheet = ss.getSheetByName(DASHBOARD_CONFIG.REPORT_SHEET_NAME);
  if (!dashboardSheet) {
    ui.alert('No Dashboard Found', 
      'Please run "Build Event Dashboard" first to create the dashboard.', 
      ui.ButtonSet.OK);
    return;
  }
  
  // Get current value
  const currentCost = getCostPerPlayerFromSheet_(ss);
  
  // Prompt for new value
  const response = ui.prompt(
    '💰 Update Cost Per Player',
    `Current value: $${currentCost.toFixed(2)}\n\nEnter new cost per player (just the number, e.g., 8.50):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  // Check if user cancelled
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  // Parse and validate input
  const inputText = response.getResponseText().trim();
  const newValue = parseFloat(inputText.replace(/[$,]/g, '')); // Strip $ and commas
  
  if (isNaN(newValue) || newValue < 0) {
    ui.alert('Invalid Input', 
      `"${inputText}" is not a valid cost. Please enter a positive number.`, 
      ui.ButtonSet.OK);
    return;
  }
  
  // Confirm the change
  const confirmResponse = ui.alert(
    'Confirm Update',
    `Update cost per player from $${currentCost.toFixed(2)} to $${newValue.toFixed(2)}?\n\nThis will recalculate all cost estimates.`,
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResponse !== ui.Button.YES) {
    return;
  }
  
  // Rebuild the dashboard with the new value
  rebuildDashboardWithNewCost_(ss, newValue);
  
  ui.alert('✅ Dashboard Updated!', 
    `Cost per player changed to $${newValue.toFixed(2)}.\n\nAll cost calculations have been updated.`,
    ui.ButtonSet.OK);
}

// =============================================================================
// REBUILD DASHBOARD - Called when cost per player changes
// =============================================================================
function rebuildDashboardWithNewCost_(ss, newCostPerPlayer) {
  const events = getEventSheetsFlexible_(ss);
  if (!events.length) return;
  
  const spentPoolCostMap = getSpentPoolCostMap_(ss);
  const hasSpentPool = spentPoolCostMap !== null;
  
  // Build analytics with the NEW cost per player value
  const analytics = buildAnalyticsWithCost_(ss, events, spentPoolCostMap, hasSpentPool, newCostPerPlayer);
  
  // Rewrite the dashboard (preserving the input cell value)
  writeDashboard_(ss, analytics, hasSpentPool, newCostPerPlayer);
}

// =============================================================================
// READ COST FROM INPUT CELL - Gets current value or default
// =============================================================================
function getCostPerPlayerFromSheet_(ss) {
  const sheet = ss.getSheetByName(DASHBOARD_CONFIG.REPORT_SHEET_NAME);
  if (!sheet) return DASHBOARD_CONFIG.DEFAULT_AVG_COST_PER_PLAYER;
  
  const value = sheet.getRange(DASHBOARD_CONFIG.INPUT_CELL).getValue();
  if (typeof value === 'number' && value > 0) {
    return value;
  }
  return DASHBOARD_CONFIG.DEFAULT_AVG_COST_PER_PLAYER;
}

// =============================================================================
// ENTRY POINT
// =============================================================================
function buildEventDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Show progress
  ui.alert('Building Dashboard', 'Analyzing event data... This may take a moment.', ui.ButtonSet.OK);
  
  // Collect all data
  const events = getEventSheetsFlexible_(ss);
  if (!events.length) {
    ui.alert('No event sheets found. Make sure your event tabs have a date-based name (e.g., 11-26-2025 or 11-26C-2025).');
    return;
  }
  
  const spentPoolCostMap = getSpentPoolCostMap_(ss);
  const hasSpentPool = spentPoolCostMap !== null;
  
  // Get cost per player (from existing sheet if rebuilding, otherwise default)
  const costPerPlayer = getCostPerPlayerFromSheet_(ss);
  
  // Build comprehensive analytics
  const analytics = buildAnalyticsWithCost_(ss, events, spentPoolCostMap, hasSpentPool, costPerPlayer);
  
  // Write the dashboard
  writeDashboard_(ss, analytics, hasSpentPool, costPerPlayer);
  
  ui.alert('Dashboard Complete!', 
    `Analyzed ${events.length} events across ${analytics.monthly.length} months.\n` +
    `Tracked ${analytics.playerStats.totalUniquePlayers} unique players.\n` +
    `Dashboard written to "${DASHBOARD_CONFIG.REPORT_SHEET_NAME}".\n\n` +
    `💡 TIP: Use "💰 Update Cost Per Player" from the menu to change cost assumptions!`,
    ui.ButtonSet.OK
  );
}

// =============================================================================
// ANALYTICS ENGINE
// =============================================================================
function buildAnalyticsWithCost_(ss, events, spentPoolCostMap, hasSpentPool, costPerPlayer) {
  // Use provided cost per player for estimates
  const estimateCostPerPlayer = costPerPlayer || DASHBOARD_CONFIG.DEFAULT_AVG_COST_PER_PLAYER;
  // Master player tracking: playerName -> { firstSeen: Date, events: [sheetName], totalEvents: n }
  const playerHistory = new Map();
  
  // Per-event data collection
  const perEventData = [];
  
  // Monthly aggregation
  const monthMap = new Map();
  
  // Day-of-week aggregation (0=Sun, 1=Mon, ..., 6=Sat)
  const dayOfWeekMap = new Map([
    [0, { name: 'Sunday', events: 0, attendance: 0, cost: 0 }],
    [1, { name: 'Monday', events: 0, attendance: 0, cost: 0 }],
    [2, { name: 'Tuesday', events: 0, attendance: 0, cost: 0 }],
    [3, { name: 'Wednesday', events: 0, attendance: 0, cost: 0 }],
    [4, { name: 'Thursday', events: 0, attendance: 0, cost: 0 }],
    [5, { name: 'Friday', events: 0, attendance: 0, cost: 0 }],
    [6, { name: 'Saturday', events: 0, attendance: 0, cost: 0 }]
  ]);
  
  // Event type (suffix) aggregation
  const eventTypeMap = new Map();
  
  // Process each event
  events.forEach(ev => {
    const sheet = ss.getSheetByName(ev.sheetName);
    const rosterNames = getEventRosterArray_(sheet); // Array for order preservation
    const playerCount = rosterNames.length;
    
    // Date components
    const yyyy = ev.eventDate.getFullYear();
    const mm = String(ev.eventDate.getMonth() + 1).padStart(2, '0');
    const dd = String(ev.eventDate.getDate()).padStart(2, '0');
    const monthKey = `${yyyy}-${mm}`;
    const dayOfWeek = ev.eventDate.getDay();
    const eventType = ev.suffix || 'Standard';
    
    // Cost calculation (uses dynamic costPerPlayer for estimates)
    const actualCost = hasSpentPool ? (spentPoolCostMap.get(ev.sheetName) ?? null) : null;
    const estimatedCost = playerCount * estimateCostPerPlayer;
    const eventCost = (actualCost !== null) ? actualCost : estimatedCost;
    const costSource = (actualCost !== null) ? 'Spent_Pool' : 
                       (hasSpentPool ? 'Estimated' : 'No Spent_Pool');
    const avgCostPerPlayer = playerCount > 0 ? (eventCost / playerCount) : 0;
    
    // Track new vs returning players for this event
    let newPlayersThisEvent = 0;
    let returningPlayersThisEvent = 0;
    
    rosterNames.forEach(name => {
      if (!playerHistory.has(name)) {
        playerHistory.set(name, {
          firstSeen: ev.eventDate,
          firstSeenMonth: monthKey,
          events: [ev.sheetName],
          totalEvents: 1
        });
        newPlayersThisEvent++;
      } else {
        const ph = playerHistory.get(name);
        ph.events.push(ev.sheetName);
        ph.totalEvents++;
        returningPlayersThisEvent++;
      }
    });
    
    // Store per-event data
    perEventData.push({
      date: `${yyyy}-${mm}-${dd}`,
      monthKey,
      sheetName: ev.sheetName,
      eventDate: ev.eventDate,
      dayOfWeek,
      eventType,
      playerCount,
      newPlayers: newPlayersThisEvent,
      returningPlayers: returningPlayersThisEvent,
      avgCostPerPlayer,
      eventCost,
      costSource
    });
    
    // Monthly aggregation
    if (!monthMap.has(monthKey)) {
      monthMap.set(monthKey, {
        events: 0,
        totalAttendance: 0,
        uniquePlayers: new Set(),
        newPlayers: new Set(),
        returningPlayers: new Set(),
        totalCost: 0,
        eventCosts: [],
        eventAttendances: []
      });
    }
    const monthAgg = monthMap.get(monthKey);
    monthAgg.events++;
    monthAgg.totalAttendance += playerCount;
    monthAgg.totalCost += eventCost;
    monthAgg.eventCosts.push(eventCost);
    monthAgg.eventAttendances.push(playerCount);
    
    rosterNames.forEach(name => {
      monthAgg.uniquePlayers.add(name);
      const ph = playerHistory.get(name);
      if (ph.firstSeenMonth === monthKey) {
        monthAgg.newPlayers.add(name);
      } else {
        monthAgg.returningPlayers.add(name);
      }
    });
    
    // Day-of-week aggregation
    const dowAgg = dayOfWeekMap.get(dayOfWeek);
    dowAgg.events++;
    dowAgg.attendance += playerCount;
    dowAgg.cost += eventCost;
    
    // Event type aggregation
    if (!eventTypeMap.has(eventType)) {
      eventTypeMap.set(eventType, {
        events: 0,
        attendance: 0,
        cost: 0,
        uniquePlayers: new Set()
      });
    }
    const typeAgg = eventTypeMap.get(eventType);
    typeAgg.events++;
    typeAgg.attendance += playerCount;
    typeAgg.cost += eventCost;
    rosterNames.forEach(name => typeAgg.uniquePlayers.add(name));
  });
  
  // Build monthly rows with calculated metrics
  const monthKeysSorted = Array.from(monthMap.keys()).sort();
  const monthlyData = monthKeysSorted.map((monthKey, idx) => {
    const agg = monthMap.get(monthKey);
    const avgCostPerEvent = agg.events > 0 ? (agg.totalCost / agg.events) : 0;
    const weightedCostPerPlayer = agg.totalAttendance > 0 ? (agg.totalCost / agg.totalAttendance) : 0;
    const avgAttendancePerEvent = agg.events > 0 ? (agg.totalAttendance / agg.events) : 0;
    
    // Previous month for MoM calculations
    let momAttendanceChange = null;
    let momCostChange = null;
    if (idx > 0) {
      const prevAgg = monthMap.get(monthKeysSorted[idx - 1]);
      if (prevAgg.totalAttendance > 0) {
        momAttendanceChange = ((agg.totalAttendance - prevAgg.totalAttendance) / prevAgg.totalAttendance) * 100;
      }
      if (prevAgg.totalCost > 0) {
        momCostChange = ((agg.totalCost - prevAgg.totalCost) / prevAgg.totalCost) * 100;
      }
    }
    
    return {
      monthKey,
      events: agg.events,
      totalAttendance: agg.totalAttendance,
      uniquePlayers: agg.uniquePlayers.size,
      newPlayers: agg.newPlayers.size,
      returningPlayers: agg.returningPlayers.size,
      retentionRate: agg.uniquePlayers.size > 0 ? (agg.returningPlayers.size / agg.uniquePlayers.size) * 100 : 0,
      totalCost: agg.totalCost,
      avgCostPerEvent,
      weightedCostPerPlayer,
      avgAttendancePerEvent,
      momAttendanceChange,
      momCostChange,
      eventCosts: agg.eventCosts,
      eventAttendances: agg.eventAttendances
    };
  });
  
  // Calculate trends
  const xs = monthKeysSorted.map((_, i) => i);
  const monthlyAttendance = monthlyData.map(m => m.totalAttendance);
  const monthlyCost = monthlyData.map(m => m.totalCost);
  const monthlyNewPlayers = monthlyData.map(m => m.newPlayers);
  
  const attendanceSlope = linearRegressionSlope_(xs, monthlyAttendance);
  const costSlope = linearRegressionSlope_(xs, monthlyCost);
  const newPlayerSlope = linearRegressionSlope_(xs, monthlyNewPlayers);
  
  // Anomaly detection
  const attendanceMean = mean_(perEventData.map(e => e.playerCount));
  const attendanceStdev = stdev_(perEventData.map(e => e.playerCount));
  const costMean = mean_(perEventData.map(e => e.eventCost));
  const costStdev = stdev_(perEventData.map(e => e.eventCost));
  
  perEventData.forEach(ev => {
    ev.attendanceZScore = attendanceStdev > 0 ? (ev.playerCount - attendanceMean) / attendanceStdev : 0;
    ev.costZScore = costStdev > 0 ? (ev.eventCost - costMean) / costStdev : 0;
    ev.isAttendanceAnomaly = Math.abs(ev.attendanceZScore) > DASHBOARD_CONFIG.ANOMALY_THRESHOLD_STDEV;
    ev.isCostAnomaly = Math.abs(ev.costZScore) > DASHBOARD_CONFIG.ANOMALY_THRESHOLD_STDEV;
  });
  
  // Player frequency distribution
  const frequencyDist = buildFrequencyDistribution_(playerHistory);
  
  // Top players
  const topPlayers = buildTopPlayersList_(playerHistory, DASHBOARD_CONFIG.TOP_PLAYERS_COUNT);
  
  // Day-of-week summary
  const dayOfWeekData = Array.from(dayOfWeekMap.entries())
    .map(([day, data]) => ({
      dayNum: day,
      dayName: data.name,
      events: data.events,
      attendance: data.attendance,
      cost: data.cost,
      avgAttendance: data.events > 0 ? data.attendance / data.events : 0,
      avgCost: data.events > 0 ? data.cost / data.events : 0
    }))
    .filter(d => d.events > 0);
  
  // Event type summary
  const eventTypeData = Array.from(eventTypeMap.entries())
    .map(([type, data]) => ({
      type,
      events: data.events,
      attendance: data.attendance,
      cost: data.cost,
      uniquePlayers: data.uniquePlayers.size,
      avgAttendance: data.events > 0 ? data.attendance / data.events : 0,
      avgCost: data.events > 0 ? data.cost / data.events : 0
    }))
    .sort((a, b) => b.events - a.events);
  
  // Summary KPIs
  const totalEvents = perEventData.length;
  const totalAttendance = perEventData.reduce((sum, e) => sum + e.playerCount, 0);
  const totalCost = perEventData.reduce((sum, e) => sum + e.eventCost, 0);
  const totalNewPlayers = monthlyData.reduce((sum, m) => sum + m.newPlayers, 0);
  
  return {
    perEvent: perEventData,
    monthly: monthlyData,
    dayOfWeek: dayOfWeekData,
    eventType: eventTypeData,
    frequencyDist,
    topPlayers,
    playerStats: {
      totalUniquePlayers: playerHistory.size,
      totalNewPlayers,
      avgEventsPerPlayer: playerHistory.size > 0 ? totalAttendance / playerHistory.size : 0
    },
    trends: {
      attendanceSlope,
      costSlope,
      newPlayerSlope
    },
    anomalyStats: {
      attendanceMean,
      attendanceStdev,
      costMean,
      costStdev
    },
    summary: {
      totalEvents,
      totalAttendance,
      totalCost,
      blendedCostPerPlayer: totalAttendance > 0 ? totalCost / totalAttendance : 0,
      avgAttendancePerEvent: totalEvents > 0 ? totalAttendance / totalEvents : 0,
      avgCostPerEvent: totalEvents > 0 ? totalCost / totalEvents : 0
    }
  };
}

// =============================================================================
// DASHBOARD WRITER
// =============================================================================
function writeDashboard_(ss, analytics, hasSpentPool, costPerPlayer) {
  let sheet = ss.getSheetByName(DASHBOARD_CONFIG.REPORT_SHEET_NAME);
  if (sheet) {
    sheet.clear();
    sheet.clearFormats();
    sheet.clearConditionalFormatRules();
  } else {
    sheet = ss.insertSheet(DASHBOARD_CONFIG.REPORT_SHEET_NAME);
  }
  
  let currentRow = 1;
  
  // =========================================================================
  // TITLE
  // =========================================================================
  sheet.getRange(currentRow, 1, 1, 10).merge();
  sheet.getRange(currentRow, 1)
    .setValue('🎮 COSMIC GAMES - EVENT ANALYTICS DASHBOARD')
    .setFontSize(18)
    .setFontWeight('bold')
    .setBackground(DASHBOARD_CONFIG.COLORS.HEADER_BG)
    .setFontColor(DASHBOARD_CONFIG.COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');
  currentRow += 2;
  
  // Generated timestamp
  sheet.getRange(currentRow, 1)
    .setValue(`Generated: ${new Date().toLocaleString()}`)
    .setFontStyle('italic')
    .setFontColor('#666666');
  currentRow += 2;
  
  // =========================================================================
  // INPUT CELL FOR COST PER PLAYER (Row 5-7 area) - Display only now
  // =========================================================================
  // Label for input
  sheet.getRange(6, 4) // D6
    .setValue('💰 Cost per Player:')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('right');
  
  // The display cell at E7 (shows current value)
  const inputCell = sheet.getRange(DASHBOARD_CONFIG.INPUT_CELL_ROW, DASHBOARD_CONFIG.INPUT_CELL_COL); // E7
  inputCell
    .setValue(costPerPlayer)
    .setNumberFormat('$#,##0.00')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(DASHBOARD_CONFIG.COLORS.INPUT_BG)
    .setBorder(true, true, true, true, false, false, 
               DASHBOARD_CONFIG.COLORS.INPUT_BORDER, 
               SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setHorizontalAlignment('center');
  
  // Instruction note
  sheet.getRange(6, 6) // F6
    .setValue('Use menu: 🎮 Cosmic Ops → 💰 Update Cost Per Player')
    .setFontStyle('italic')
    .setFontSize(9)
    .setFontColor('#666666');
  
  sheet.getRange(7, 6) // F7
    .setValue('to change this value and recalculate costs')
    .setFontStyle('italic')
    .setFontSize(9)
    .setFontColor('#888888');
  
  // =========================================================================
  // KPI CARDS (Summary Stats)
  // =========================================================================
  currentRow = writeKPISection_(sheet, currentRow, analytics, hasSpentPool);
  currentRow += 2;
  
  // =========================================================================
  // MONTHLY ROLLUP WITH SPARKLINES
  // =========================================================================
  currentRow = writeMonthlySection_(sheet, currentRow, analytics);
  currentRow += 2;
  
  // =========================================================================
  // DAY-OF-WEEK ANALYSIS
  // =========================================================================
  currentRow = writeDayOfWeekSection_(sheet, currentRow, analytics);
  currentRow += 2;
  
  // =========================================================================
  // EVENT TYPE BREAKDOWN
  // =========================================================================
  if (analytics.eventType.length > 1) { // Only show if multiple types exist
    currentRow = writeEventTypeSection_(sheet, currentRow, analytics);
    currentRow += 2;
  }
  
  // =========================================================================
  // PLAYER ANALYTICS (Frequency + Top Players)
  // =========================================================================
  currentRow = writePlayerAnalyticsSection_(sheet, currentRow, analytics);
  currentRow += 2;
  
  // =========================================================================
  // ANOMALIES / FLAGS
  // =========================================================================
  const anomalies = analytics.perEvent.filter(e => e.isAttendanceAnomaly || e.isCostAnomaly);
  if (anomalies.length > 0) {
    currentRow = writeAnomalySection_(sheet, currentRow, anomalies);
    currentRow += 2;
  }
  
  // =========================================================================
  // PER-EVENT DETAIL
  // =========================================================================
  currentRow = writePerEventSection_(sheet, currentRow, analytics);
  
  // =========================================================================
  // CHARTS
  // =========================================================================
  createCharts_(sheet, analytics);
  
  // =========================================================================
  // FINAL FORMATTING
  // =========================================================================
  sheet.autoResizeColumns(1, 12);
  sheet.setFrozenRows(5); // Freeze through KPI section
}

// =============================================================================
// SECTION WRITERS
// =============================================================================

function writeKPISection_(sheet, startRow, analytics, hasSpentPool) {
  let row = startRow;
  
  // Section header
  sheet.getRange(row, 1, 1, 10).merge();
  sheet.getRange(row, 1)
    .setValue('📊 KEY PERFORMANCE INDICATORS')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(DASHBOARD_CONFIG.COLORS.SECTION_BG)
    .setFontColor(DASHBOARD_CONFIG.COLORS.SECTION_TEXT);
  row++;
  
  // KPI Grid (2 rows x 5 columns of KPI cards)
  const kpis = [
    { label: 'Total Events', value: analytics.summary.totalEvents, format: '#,##0' },
    { label: 'Total Attendance', value: analytics.summary.totalAttendance, format: '#,##0' },
    { label: 'Unique Players', value: analytics.playerStats.totalUniquePlayers, format: '#,##0' },
    { label: 'Total Cost', value: analytics.summary.totalCost, format: '$#,##0.00' },
    { label: 'Blended $/Player', value: analytics.summary.blendedCostPerPlayer, format: '$#,##0.00' },
    { label: 'Avg Attendance/Event', value: analytics.summary.avgAttendancePerEvent, format: '#,##0.0' },
    { label: 'Avg Cost/Event', value: analytics.summary.avgCostPerEvent, format: '$#,##0.00' },
    { label: 'Avg Events/Player', value: analytics.playerStats.avgEventsPerPlayer, format: '#,##0.0' },
    { label: 'Attendance Trend', value: analytics.trends.attendanceSlope, format: '+#,##0.0;-#,##0.0', isTrend: true },
    { label: 'Cost Trend ($/mo)', value: analytics.trends.costSlope, format: '+$#,##0.00;-$#,##0.00', isTrend: true }
  ];
  
  // Write KPIs in pairs (label above value)
  for (let i = 0; i < kpis.length; i++) {
    const col = (i % 5) + 1;
    const kpiRow = row + Math.floor(i / 5) * 3;
    
    // Label
    sheet.getRange(kpiRow, col)
      .setValue(kpis[i].label)
      .setFontWeight('bold')
      .setFontSize(9)
      .setFontColor('#666666');
    
    // Value
    const valueCell = sheet.getRange(kpiRow + 1, col);
    valueCell.setValue(kpis[i].value)
      .setFontSize(16)
      .setFontWeight('bold')
      .setNumberFormat(kpis[i].format);
    
    // Color trends
    if (kpis[i].isTrend) {
      if (kpis[i].label.includes('Attendance') && kpis[i].value > 0) {
        valueCell.setFontColor(DASHBOARD_CONFIG.COLORS.SPARKLINE_UP);
      } else if (kpis[i].label.includes('Attendance') && kpis[i].value < 0) {
        valueCell.setFontColor(DASHBOARD_CONFIG.COLORS.SPARKLINE_DOWN);
      } else if (kpis[i].label.includes('Cost') && kpis[i].value < 0) {
        valueCell.setFontColor(DASHBOARD_CONFIG.COLORS.SPARKLINE_UP); // Lower cost = good
      } else if (kpis[i].label.includes('Cost') && kpis[i].value > 0) {
        valueCell.setFontColor(DASHBOARD_CONFIG.COLORS.SPARKLINE_DOWN);
      }
    }
  }
  
  row += 7; // Move past KPI grid
  
  // Cost source note - now references the input cell
  sheet.getRange(row, 1)
    .setValue(`Cost Source: ${hasSpentPool ? 'Spent_Pool (with estimates from E7 where missing)' : 'Estimated using E7 value (Spent_Pool not found)'}`)
    .setFontStyle('italic')
    .setFontSize(9)
    .setFontColor('#888888');
  
  return row + 1;
}

function writeMonthlySection_(sheet, startRow, analytics) {
  let row = startRow;
  
  // Section header
  sheet.getRange(row, 1, 1, 12).merge();
  sheet.getRange(row, 1)
    .setValue('📅 MONTHLY ROLLUP')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(DASHBOARD_CONFIG.COLORS.SECTION_BG)
    .setFontColor(DASHBOARD_CONFIG.COLORS.SECTION_TEXT);
  row++;
  
  // Column headers
  const headers = [
    'Month', 'Events', 'Attendance', 'Unique', 'New', 'Returning', 
    'Retention %', 'Total Cost', 'Avg $/Event', '$/Player', 
    'MoM Attend %', 'Attend Trend'
  ];
  
  sheet.getRange(row, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground(DASHBOARD_CONFIG.COLORS.HEADER_BG)
    .setFontColor(DASHBOARD_CONFIG.COLORS.HEADER_TEXT);
  row++;
  
  // Data rows
  analytics.monthly.forEach((m, idx) => {
    const sparklineData = m.eventAttendances.join(',');
    const sparklineFormula = m.eventAttendances.length > 1 
      ? `=SPARKLINE({${sparklineData}}, {"charttype","column";"color","${DASHBOARD_CONFIG.COLORS.CHART_PRIMARY}"})` 
      : '';
    
    const rowData = [
      m.monthKey,
      m.events,
      m.totalAttendance,
      m.uniquePlayers,
      m.newPlayers,
      m.returningPlayers,
      m.retentionRate / 100, // Will format as %
      m.totalCost,
      m.avgCostPerEvent,
      m.weightedCostPerPlayer,
      m.momAttendanceChange !== null ? m.momAttendanceChange / 100 : '',
      '' // Sparkline placeholder
    ];
    
    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
    
    // Add sparkline formula
    if (sparklineFormula) {
      sheet.getRange(row, 12).setFormula(sparklineFormula);
    }
    
    // Alternating row colors
    if (idx % 2 === 1) {
      sheet.getRange(row, 1, 1, headers.length).setBackground(DASHBOARD_CONFIG.COLORS.ALTERNATING_ROW);
    }
    
    row++;
  });
  
  // Format columns
  const dataStartRow = startRow + 2;
  const dataRows = analytics.monthly.length;
  
  if (dataRows > 0) {
    // Retention %
    sheet.getRange(dataStartRow + 1, 7, dataRows, 1).setNumberFormat('0.0%');
    // Cost columns
    sheet.getRange(dataStartRow + 1, 8, dataRows, 3).setNumberFormat('$#,##0.00');
    // MoM %
    sheet.getRange(dataStartRow + 1, 11, dataRows, 1).setNumberFormat('+0.0%;-0.0%');
    
    // Conditional formatting for retention rate
    const retentionRange = sheet.getRange(dataStartRow + 1, 7, dataRows, 1);
    const retentionRules = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue(DASHBOARD_CONFIG.COLORS.KPI_GOOD, SpreadsheetApp.InterpolationType.NUMBER, '0.8')
      .setGradientMidpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '0.5')
      .setGradientMinpointWithValue(DASHBOARD_CONFIG.COLORS.KPI_BAD, SpreadsheetApp.InterpolationType.NUMBER, '0.2')
      .setRanges([retentionRange])
      .build();
    
    const rules = sheet.getConditionalFormatRules();
    rules.push(retentionRules);
    sheet.setConditionalFormatRules(rules);
  }
  
  return row;
}

function writeDayOfWeekSection_(sheet, startRow, analytics) {
  let row = startRow;
  
  // Section header
  sheet.getRange(row, 1, 1, 6).merge();
  sheet.getRange(row, 1)
    .setValue('📆 DAY-OF-WEEK PATTERNS')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(DASHBOARD_CONFIG.COLORS.SECTION_BG)
    .setFontColor(DASHBOARD_CONFIG.COLORS.SECTION_TEXT);
  row++;
  
  const headers = ['Day', 'Events', 'Total Attend', 'Avg Attend', 'Total Cost', 'Avg Cost'];
  sheet.getRange(row, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground(DASHBOARD_CONFIG.COLORS.HEADER_BG)
    .setFontColor(DASHBOARD_CONFIG.COLORS.HEADER_TEXT);
  row++;
  
  analytics.dayOfWeek.forEach((d, idx) => {
    sheet.getRange(row, 1, 1, 6).setValues([[
      d.dayName,
      d.events,
      d.attendance,
      d.avgAttendance,
      d.cost,
      d.avgCost
    ]]);
    
    if (idx % 2 === 1) {
      sheet.getRange(row, 1, 1, 6).setBackground(DASHBOARD_CONFIG.COLORS.ALTERNATING_ROW);
    }
    row++;
  });
  
  // Format
  const dataRows = analytics.dayOfWeek.length;
  if (dataRows > 0) {
    sheet.getRange(startRow + 3, 4, dataRows, 1).setNumberFormat('#,##0.0');
    sheet.getRange(startRow + 3, 5, dataRows, 2).setNumberFormat('$#,##0.00');
  }
  
  return row;
}

function writeEventTypeSection_(sheet, startRow, analytics) {
  let row = startRow;
  
  sheet.getRange(row, 1, 1, 7).merge();
  sheet.getRange(row, 1)
    .setValue('🎯 EVENT TYPE BREAKDOWN')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(DASHBOARD_CONFIG.COLORS.SECTION_BG)
    .setFontColor(DASHBOARD_CONFIG.COLORS.SECTION_TEXT);
  row++;
  
  const headers = ['Type', 'Events', 'Total Attend', 'Unique Players', 'Avg Attend', 'Total Cost', 'Avg Cost'];
  sheet.getRange(row, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground(DASHBOARD_CONFIG.COLORS.HEADER_BG)
    .setFontColor(DASHBOARD_CONFIG.COLORS.HEADER_TEXT);
  row++;
  
  analytics.eventType.forEach((t, idx) => {
    sheet.getRange(row, 1, 1, 7).setValues([[
      t.type,
      t.events,
      t.attendance,
      t.uniquePlayers,
      t.avgAttendance,
      t.cost,
      t.avgCost
    ]]);
    
    if (idx % 2 === 1) {
      sheet.getRange(row, 1, 1, 7).setBackground(DASHBOARD_CONFIG.COLORS.ALTERNATING_ROW);
    }
    row++;
  });
  
  const dataRows = analytics.eventType.length;
  if (dataRows > 0) {
    sheet.getRange(startRow + 3, 5, dataRows, 1).setNumberFormat('#,##0.0');
    sheet.getRange(startRow + 3, 6, dataRows, 2).setNumberFormat('$#,##0.00');
  }
  
  return row;
}

function writePlayerAnalyticsSection_(sheet, startRow, analytics) {
  let row = startRow;
  
  sheet.getRange(row, 1, 1, 10).merge();
  sheet.getRange(row, 1)
    .setValue('👥 PLAYER ANALYTICS')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(DASHBOARD_CONFIG.COLORS.SECTION_BG)
    .setFontColor(DASHBOARD_CONFIG.COLORS.SECTION_TEXT);
  row++;
  
  // Two side-by-side tables: Frequency Distribution (left) and Top Players (right)
  
  // -- Frequency Distribution --
  sheet.getRange(row, 1)
    .setValue('Event Frequency Distribution')
    .setFontWeight('bold');
  sheet.getRange(row, 6)
    .setValue('Top Players by Attendance')
    .setFontWeight('bold');
  row++;
  
  // Frequency headers
  sheet.getRange(row, 1, 1, 3)
    .setValues([['# Events', 'Players', '% of Total']])
    .setFontWeight('bold')
    .setBackground(DASHBOARD_CONFIG.COLORS.HEADER_BG)
    .setFontColor(DASHBOARD_CONFIG.COLORS.HEADER_TEXT);
  
  // Top players headers
  sheet.getRange(row, 6, 1, 3)
    .setValues([['Rank', 'Player', 'Events']])
    .setFontWeight('bold')
    .setBackground(DASHBOARD_CONFIG.COLORS.HEADER_BG)
    .setFontColor(DASHBOARD_CONFIG.COLORS.HEADER_TEXT);
  row++;
  
  // Write frequency data
  const freqStartRow = row;
  analytics.frequencyDist.forEach((f, idx) => {
    sheet.getRange(row, 1, 1, 3).setValues([[
      f.bucket,
      f.count,
      f.count / analytics.playerStats.totalUniquePlayers
    ]]);
    if (idx % 2 === 1) {
      sheet.getRange(row, 1, 1, 3).setBackground(DASHBOARD_CONFIG.COLORS.ALTERNATING_ROW);
    }
    row++;
  });
  
  // Format frequency %
  sheet.getRange(freqStartRow, 3, analytics.frequencyDist.length, 1).setNumberFormat('0.0%');
  
  // Write top players
  row = freqStartRow;
  analytics.topPlayers.forEach((p, idx) => {
    sheet.getRange(row, 6, 1, 3).setValues([[
      idx + 1,
      p.name,
      p.events
    ]]);
    if (idx % 2 === 1) {
      sheet.getRange(row, 6, 1, 3).setBackground(DASHBOARD_CONFIG.COLORS.ALTERNATING_ROW);
    }
    row++;
  });
  
  return Math.max(freqStartRow + analytics.frequencyDist.length, freqStartRow + analytics.topPlayers.length);
}

function writeAnomalySection_(sheet, startRow, anomalies) {
  let row = startRow;
  
  sheet.getRange(row, 1, 1, 8).merge();
  sheet.getRange(row, 1)
    .setValue('⚠️ ANOMALIES DETECTED')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#ff7043')
    .setFontColor(DASHBOARD_CONFIG.COLORS.HEADER_TEXT);
  row++;
  
  const headers = ['Date', 'Sheet', 'Attendance', 'Attend Z-Score', 'Cost', 'Cost Z-Score', 'Flag'];
  sheet.getRange(row, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground(DASHBOARD_CONFIG.COLORS.HEADER_BG)
    .setFontColor(DASHBOARD_CONFIG.COLORS.HEADER_TEXT);
  row++;
  
  anomalies.forEach((a, idx) => {
    const flags = [];
    if (a.isAttendanceAnomaly) flags.push(a.attendanceZScore > 0 ? '📈 High Attend' : '📉 Low Attend');
    if (a.isCostAnomaly) flags.push(a.costZScore > 0 ? '💸 High Cost' : '💰 Low Cost');
    
    sheet.getRange(row, 1, 1, 7).setValues([[
      a.date,
      a.sheetName,
      a.playerCount,
      a.attendanceZScore,
      a.eventCost,
      a.costZScore,
      flags.join(', ')
    ]]);
    
    // Highlight anomaly rows
    sheet.getRange(row, 1, 1, 7).setBackground(DASHBOARD_CONFIG.COLORS.KPI_WARNING);
    row++;
  });
  
  // Format
  const dataRows = anomalies.length;
  if (dataRows > 0) {
    sheet.getRange(startRow + 3, 4, dataRows, 1).setNumberFormat('+0.00;-0.00');
    sheet.getRange(startRow + 3, 5, dataRows, 1).setNumberFormat('$#,##0.00');
    sheet.getRange(startRow + 3, 6, dataRows, 1).setNumberFormat('+0.00;-0.00');
  }
  
  return row;
}

function writePerEventSection_(sheet, startRow, analytics) {
  let row = startRow;
  
  sheet.getRange(row, 1, 1, 10).merge();
  sheet.getRange(row, 1)
    .setValue('📋 PER-EVENT DETAIL')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(DASHBOARD_CONFIG.COLORS.SECTION_BG)
    .setFontColor(DASHBOARD_CONFIG.COLORS.SECTION_TEXT);
  row++;
  
  const headers = [
    'Date', 'Month', 'Sheet', 'Day', 'Type', 
    'Players', 'New', 'Returning', 'Cost', 'Source'
  ];
  
  sheet.getRange(row, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground(DASHBOARD_CONFIG.COLORS.HEADER_BG)
    .setFontColor(DASHBOARD_CONFIG.COLORS.HEADER_TEXT);
  row++;
  
  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  
  analytics.perEvent.forEach((e, idx) => {
    sheet.getRange(row, 1, 1, 10).setValues([[
      e.date,
      e.monthKey,
      e.sheetName,
      dayNames[e.dayOfWeek],
      e.eventType,
      e.playerCount,
      e.newPlayers,
      e.returningPlayers,
      e.eventCost,
      e.costSource
    ]]);
    
    // Highlight anomalies
    if (e.isAttendanceAnomaly || e.isCostAnomaly) {
      sheet.getRange(row, 1, 1, 10).setBackground(DASHBOARD_CONFIG.COLORS.KPI_WARNING);
    } else if (idx % 2 === 1) {
      sheet.getRange(row, 1, 1, 10).setBackground(DASHBOARD_CONFIG.COLORS.ALTERNATING_ROW);
    }
    row++;
  });
  
  // Format cost column
  const dataRows = analytics.perEvent.length;
  if (dataRows > 0) {
    sheet.getRange(startRow + 3, 9, dataRows, 1).setNumberFormat('$#,##0.00');
  }
  
  return row;
}

// =============================================================================
// CHART CREATION
// =============================================================================
function createCharts_(sheet, analytics) {
  // Find a good position for charts (to the right of the data)
  const chartCol = 14;
  
  // Chart 1: Monthly Attendance + Cost Trend (Combo Chart)
  if (analytics.monthly.length >= 2) {
    const chartDataRange = buildChartDataRange_(sheet, analytics.monthly, chartCol);
    
    const attendanceCostChart = sheet.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(chartDataRange)
      .setPosition(3, chartCol, 0, 0)
      .setOption('title', 'Monthly Attendance & Cost Trends')
      .setOption('series', {
        0: { type: 'bars', targetAxisIndex: 0, color: DASHBOARD_CONFIG.COLORS.CHART_PRIMARY },
        1: { type: 'line', targetAxisIndex: 1, color: DASHBOARD_CONFIG.COLORS.CHART_ACCENT, lineWidth: 3 }
      })
      .setOption('vAxes', {
        0: { title: 'Attendance', minValue: 0 },
        1: { title: 'Cost ($)', minValue: 0, format: '$#,##0' }
      })
      .setOption('hAxis', { title: 'Month', slantedText: true })
      .setOption('legend', { position: 'top' })
      .setOption('width', 600)
      .setOption('height', 350)
      .build();
    
    sheet.insertChart(attendanceCostChart);
  }
  
  // Chart 2: New vs Returning Players (Stacked Bar)
  if (analytics.monthly.length >= 2) {
    const retentionDataRange = buildRetentionChartData_(sheet, analytics.monthly, chartCol + 8);
    
    const retentionChart = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(retentionDataRange)
      .setPosition(22, chartCol, 0, 0)
      .setOption('title', 'New vs Returning Players by Month')
      .setOption('isStacked', true)
      .setOption('colors', [DASHBOARD_CONFIG.COLORS.CHART_ACCENT, DASHBOARD_CONFIG.COLORS.CHART_PRIMARY])
      .setOption('hAxis', { title: 'Players' })
      .setOption('legend', { position: 'top' })
      .setOption('width', 600)
      .setOption('height', 300)
      .build();
    
    sheet.insertChart(retentionChart);
  }
  
  // Chart 3: Day of Week Distribution (Pie)
  if (analytics.dayOfWeek.length >= 2) {
    const dowDataRange = buildDOWChartData_(sheet, analytics.dayOfWeek, chartCol + 8);
    
    const dowChart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dowDataRange)
      .setPosition(40, chartCol, 0, 0)
      .setOption('title', 'Events by Day of Week')
      .setOption('legend', { position: 'right' })
      .setOption('pieSliceText', 'percentage')
      .setOption('width', 400)
      .setOption('height', 300)
      .build();
    
    sheet.insertChart(dowChart);
  }
}

function buildChartDataRange_(sheet, monthlyData, startCol) {
  const data = [['Month', 'Attendance', 'Cost']];
  monthlyData.forEach(m => {
    data.push([m.monthKey, m.totalAttendance, m.totalCost]);
  });
  
  sheet.getRange(1, startCol, data.length, 3).setValues(data);
  // Hide the data columns
  sheet.hideColumns(startCol, 3);
  
  return sheet.getRange(1, startCol, data.length, 3);
}

function buildRetentionChartData_(sheet, monthlyData, startCol) {
  const data = [['Month', 'New Players', 'Returning Players']];
  monthlyData.forEach(m => {
    data.push([m.monthKey, m.newPlayers, m.returningPlayers]);
  });
  
  sheet.getRange(1, startCol, data.length, 3).setValues(data);
  sheet.hideColumns(startCol, 3);
  
  return sheet.getRange(1, startCol, data.length, 3);
}

function buildDOWChartData_(sheet, dowData, startCol) {
  const data = [['Day', 'Events']];
  dowData.forEach(d => {
    data.push([d.dayName, d.events]);
  });
  
  sheet.getRange(1, startCol, data.length, 2).setValues(data);
  sheet.hideColumns(startCol, 2);
  
  return sheet.getRange(1, startCol, data.length, 2);
}

// =============================================================================
// HELPER FUNCTIONS
// =============================================================================

function buildFrequencyDistribution_(playerHistory) {
  const buckets = [
    { label: '1 event', min: 1, max: 1, count: 0 },
    { label: '2 events', min: 2, max: 2, count: 0 },
    { label: '3-4 events', min: 3, max: 4, count: 0 },
    { label: '5-9 events', min: 5, max: 9, count: 0 },
    { label: '10+ events', min: 10, max: Infinity, count: 0 }
  ];
  
  playerHistory.forEach(player => {
    const events = player.totalEvents;
    for (const bucket of buckets) {
      if (events >= bucket.min && events <= bucket.max) {
        bucket.count++;
        break;
      }
    }
  });
  
  return buckets.map(b => ({ bucket: b.label, count: b.count }));
}

function buildTopPlayersList_(playerHistory, limit) {
  return Array.from(playerHistory.entries())
    .map(([name, data]) => ({ name, events: data.totalEvents }))
    .sort((a, b) => b.events - a.events)
    .slice(0, limit);
}

function mean_(arr) {
  if (!arr || arr.length === 0) return 0;
  return arr.reduce((sum, v) => sum + v, 0) / arr.length;
}

function stdev_(arr) {
  if (!arr || arr.length < 2) return 0;
  const m = mean_(arr);
  const variance = arr.reduce((sum, v) => sum + Math.pow(v - m, 2), 0) / arr.length;
  return Math.sqrt(variance);
}

function linearRegressionSlope_(xs, ys) {
  if (!xs || !ys || xs.length !== ys.length || xs.length < 2) return 0;
  
  const n = xs.length;
  let sumX = 0, sumY = 0, sumXY = 0, sumXX = 0;
  
  for (let i = 0; i < n; i++) {
    sumX += xs[i];
    sumY += ys[i];
    sumXY += xs[i] * ys[i];
    sumXX += xs[i] * xs[i];
  }
  
  const denom = (n * sumXX) - (sumX * sumX);
  if (denom === 0) return 0;
  
  return ((n * sumXY) - (sumX * sumY)) / denom;
}

// =============================================================================
// EVENT SHEET DISCOVERY
// =============================================================================
function getEventSheetsFlexible_(ss) {
  const sheets = ss.getSheets();
  const re = /^(\d{1,2})-(\d{1,2})([A-Za-z])?-(\d{4})$/;
  
  const events = [];
  
  sheets.forEach(sh => {
    const name = sh.getName().trim();
    const m = name.match(re);
    if (!m) return;
    
    const month = parseInt(m[1], 10);
    const day = parseInt(m[2], 10);
    const suffix = (m[3] || '').toUpperCase();
    const year = parseInt(m[4], 10);
    
    if (month < 1 || month > 12) return;
    if (day < 1 || day > 31) return;
    
    const d = new Date(year, month - 1, day);
    if (d.getFullYear() !== year || d.getMonth() !== (month - 1) || d.getDate() !== day) return;
    
    events.push({ sheetName: name, eventDate: d, month, day, year, suffix });
  });
  
  events.sort((a, b) => {
    const t = a.eventDate.getTime() - b.eventDate.getTime();
    if (t !== 0) return t;
    return a.sheetName.localeCompare(b.sheetName);
  });
  
  return events;
}

// =============================================================================
// ROSTER EXTRACTION
// =============================================================================
function getEventRosterArray_(sheet) {
  const roster = [];
  const seen = new Set();
  
  if (!sheet) return roster;
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol < 1) return roster;
  
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(h => String(h).trim());
  
  const nameColIndex = findNameColumnIndex_(headers);
  const col = nameColIndex !== -1 ? nameColIndex : 1;
  
  for (let r = 1; r < values.length; r++) {
    const name = String(values[r][col] ?? '').trim();
    if (!name || seen.has(name)) continue;
    seen.add(name);
    roster.push(name);
  }
  
  return roster;
}

function findNameColumnIndex_(headers) {
  const normalized = headers.map(h => normalizeHeader_(h));
  const targets = new Set([
    'preferredname', 'preferred_name', 'preferred_name_id',
    'player', 'playername', 'name', 'player_name'
  ]);
  
  for (let i = 0; i < normalized.length; i++) {
    if (targets.has(normalized[i])) return i;
  }
  return -1;
}

function normalizeHeader_(s) {
  return String(s).trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
}

// =============================================================================
// SPENT_POOL COST MAP
// =============================================================================
function getSpentPoolCostMap_(ss) {
  const sh = ss.getSheetByName('Spent_Pool');
  if (!sh) return null;
  
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= 1 || lastCol < 1) return new Map();
  
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(h => String(h).trim());
  const norm = headers.map(h => normalizeHeader_(h));
  
  const eventIdCol = findFirstCol_(norm, ['event_id', 'eventid', 'event', 'sheet', 'sheet_name', 'event_sheet']);
  const totalCostCol = findFirstCol_(norm, ['total_cost', 'totalcost', 'total_cogs', 'totalcogs', 'cost']);
  const revertedCol = findFirstCol_(norm, ['reverted', 'is_reverted', 'was_reverted']);
  
  if (eventIdCol === -1 || totalCostCol === -1) return new Map();
  
  const map = new Map();
  for (let r = 1; r < values.length; r++) {
    const eventId = String(values[r][eventIdCol] ?? '').trim();
    if (!eventId) continue;
    
    if (revertedCol !== -1) {
      const rv = values[r][revertedCol];
      const isReverted = (rv === true) || 
                         (String(rv).toLowerCase() === 'true') || 
                         (String(rv).toLowerCase() === 'yes') || 
                         (String(rv).toLowerCase() === '1');
      if (isReverted) continue;
    }
    
    const costVal = values[r][totalCostCol];
    const cost = (typeof costVal === 'number') ? costVal : parseFloat(String(costVal).replace(/[^0-9.\-]/g, ''));
    if (isNaN(cost)) continue;
    
    map.set(eventId, (map.get(eventId) || 0) + cost);
  }
  return map;
}

function findFirstCol_(normalizedHeaders, targets) {
  const set = new Set(targets);
  for (let i = 0; i < normalizedHeaders.length; i++) {
    if (set.has(normalizedHeaders[i])) return i;
  }
  return -1;
}

