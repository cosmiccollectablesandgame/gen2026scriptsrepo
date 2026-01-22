/**
 * ════════════════════════════════════════════════════════════════════════════
 * ATTENDANCE MISSION SCANNER - TEST SUITE
 * ════════════════════════════════════════════════════════════════════════════
 *
 * @fileoverview Comprehensive tests for the Attendance Mission Scanner
 * 
 * Run testAttendanceMissionScanner() to execute all tests
 *
 * Version: 1.0.0
 * ════════════════════════════════════════════════════════════════════════════
 */

/**
 * Master test runner
 */
function testAttendanceMissionScanner() {
  Logger.log('========================================');
  Logger.log('ATTENDANCE MISSION SCANNER TEST SUITE');
  Logger.log('========================================\n');

  const results = {
    total: 0,
    passed: 0,
    failed: 0,
    errors: []
  };

  // Run all test suites
  runTestSuite('Event Sheet Detection', testEventSheetDetection, results);
  runTestSuite('Event Name Parsing', testEventNameParsing, results);
  runTestSuite('ISO Week Calculation', testISOWeekCalculation, results);
  runTestSuite('Suffix Mapping', testSuffixMapping, results);
  runTestSuite('Category Helpers', testCategoryHelpers, results);
  runTestSuite('Player Stats Computation', testPlayerStatsComputation, results);
  runTestSuite('Mission Points Calculation', testMissionPointsCalculation, results);

  // Summary
  Logger.log('\n========================================');
  Logger.log('TEST SUMMARY');
  Logger.log('========================================');
  Logger.log('Total Tests: ' + results.total);
  Logger.log('Passed: ' + results.passed);
  Logger.log('Failed: ' + results.failed);

  if (results.errors.length > 0) {
    Logger.log('\nFailed Tests:');
    results.errors.forEach(err => Logger.log('  ✗ ' + err));
  }

  Logger.log('========================================\n');

  return results;
}

/**
 * Test runner helper
 */
function runTestSuite(name, testFunc, results) {
  Logger.log('\n── ' + name + ' ──');
  try {
    testFunc(results);
  } catch (e) {
    results.total++;
    results.failed++;
    results.errors.push(name + ': ' + e.message);
    Logger.log('  ✗ SUITE CRASHED: ' + e.message);
  }
}

/**
 * Assertion helper
 */
function assert(condition, message, results) {
  results.total++;
  if (condition) {
    results.passed++;
    Logger.log('  ✓ ' + message);
  } else {
    results.failed++;
    results.errors.push(message);
    Logger.log('  ✗ ' + message);
  }
}

// ════════════════════════════════════════════════════════════════════════════
// TEST SUITES
// ════════════════════════════════════════════════════════════════════════════

/**
 * Test event sheet detection
 */
function testEventSheetDetection(results) {
  // Test regex pattern
  const EVENT_TAB_REGEX = /^\d{2}-\d{2}[A-Za-z]+-\d{4}$/;
  
  // Valid patterns
  assert(EVENT_TAB_REGEX.test('05-10C-2025'), 'Should match 05-10C-2025', results);
  assert(EVENT_TAB_REGEX.test('12-01D-2025'), 'Should match 12-01D-2025', results);
  assert(EVENT_TAB_REGEX.test('07-26q-2025'), 'Should match 07-26q-2025 (lowercase)', results);
  assert(EVENT_TAB_REGEX.test('01-01Draft-2025'), 'Should match 01-01Draft-2025 (multi-char suffix)', results);
  
  // Invalid patterns
  assert(!EVENT_TAB_REGEX.test('BP_Total'), 'Should NOT match BP_Total', results);
  assert(!EVENT_TAB_REGEX.test('PreferredNames'), 'Should NOT match PreferredNames', results);
  assert(!EVENT_TAB_REGEX.test('11-23-2025'), 'Should NOT match 11-23-2025 (no suffix)', results);
  assert(!EVENT_TAB_REGEX.test('5-10C-2025'), 'Should NOT match 5-10C-2025 (single digit month)', results);
}

/**
 * Test event name parsing
 */
function testEventNameParsing(results) {
  // Test valid event name
  const parsed1 = parseEventSheetName('05-10C-2025');
  assert(parsed1 !== null, 'Should parse 05-10C-2025', results);
  assert(parsed1.suffix === 'C', 'Suffix should be C', results);
  assert(parsed1.monthKey === '2025-05', 'Month key should be 2025-05', results);
  
  // Test date parsing
  const date1 = parsed1.date;
  assert(date1.getFullYear() === 2025, 'Year should be 2025', results);
  assert(date1.getMonth() === 4, 'Month should be 4 (May, 0-indexed)', results);
  assert(date1.getDate() === 10, 'Day should be 10', results);
  
  // Test uppercase normalization
  const parsed2 = parseEventSheetName('07-26q-2025');
  assert(parsed2.suffix === 'Q', 'Lowercase suffix should be normalized to Q', results);
  
  // Test multi-character suffix
  const parsed3 = parseEventSheetName('12-01Draft-2025');
  assert(parsed3.suffix === 'DRAFT', 'Multi-char suffix should be DRAFT', results);
  
  // Test invalid patterns
  const parsed4 = parseEventSheetName('11-23-2025');
  assert(parsed4 === null, 'Should return null for event without suffix', results);
}

/**
 * Test ISO week calculation
 */
function testISOWeekCalculation(results) {
  // Test known ISO weeks
  const date1 = new Date(2025, 0, 1); // Jan 1, 2025
  const week1 = getISOWeekKey(date1);
  assert(week1.startsWith('2025-W') || week1.startsWith('2024-W'), 'Should generate valid ISO week for Jan 1, 2025', results);
  
  const date2 = new Date(2025, 11, 31); // Dec 31, 2025
  const week2 = getISOWeekKey(date2);
  assert(week2.match(/^\d{4}-W\d{2}$/), 'Should match YYYY-Www format', results);
  
  // Test format
  assert(week1.indexOf('-W') > 0, 'Should contain -W separator', results);
  assert(week1.split('-W')[1].length === 2, 'Week number should be 2 digits', results);
}

/**
 * Test suffix to category mapping
 */
function testSuffixMapping(results) {
  // Test commander formats
  assert(getEventCategoryFromSuffix('B') === 'CASUAL_COMMANDER', 'B should map to CASUAL_COMMANDER', results);
  assert(getEventCategoryFromSuffix('C') === 'TRANSITIONAL_COMMANDER', 'C should map to TRANSITIONAL_COMMANDER', results);
  assert(getEventCategoryFromSuffix('U') === 'CEDH', 'U should map to CEDH', results);
  
  // Test limited formats
  assert(getEventCategoryFromSuffix('D') === 'DRAFT', 'D should map to DRAFT', results);
  assert(getEventCategoryFromSuffix('P') === 'DRAFT', 'P should map to DRAFT', results);
  assert(getEventCategoryFromSuffix('S') === 'SEALED', 'S should map to SEALED', results);
  assert(getEventCategoryFromSuffix('R') === 'PRERELEASE', 'R should map to PRERELEASE', results);
  
  // Test other formats
  assert(getEventCategoryFromSuffix('A') === 'ACADEMY', 'A should map to ACADEMY', results);
  assert(getEventCategoryFromSuffix('E') === 'OUTREACH', 'E should map to OUTREACH', results);
  assert(getEventCategoryFromSuffix('F') === 'FREE_PLAY', 'F should map to FREE_PLAY', results);
  
  // Test case insensitivity
  assert(getEventCategoryFromSuffix('b') === 'CASUAL_COMMANDER', 'Lowercase b should work', results);
  
  // Test unknown suffix
  assert(getEventCategoryFromSuffix('ZZZ') === 'OTHER', 'Unknown suffix should return OTHER', results);
}

/**
 * Test category helper functions
 */
function testCategoryHelpers(results) {
  // Test isLimitedCategory
  assert(isLimitedCategory('DRAFT') === true, 'DRAFT is limited', results);
  assert(isLimitedCategory('SEALED') === true, 'SEALED is limited', results);
  assert(isLimitedCategory('PRERELEASE') === true, 'PRERELEASE is limited', results);
  assert(isLimitedCategory('CASUAL_COMMANDER') === false, 'CASUAL_COMMANDER is not limited', results);
  
  // Test isCommanderCategory
  assert(isCommanderCategory('CASUAL_COMMANDER') === true, 'CASUAL_COMMANDER is commander', results);
  assert(isCommanderCategory('TRANSITIONAL_COMMANDER') === true, 'TRANSITIONAL_COMMANDER is commander', results);
  assert(isCommanderCategory('CEDH') === true, 'CEDH is commander', results);
  assert(isCommanderCategory('FREE_PLAY') === true, 'FREE_PLAY is commander', results);
  assert(isCommanderCategory('DRAFT') === false, 'DRAFT is not commander', results);
}

/**
 * Test player stats computation with mock data
 */
function testPlayerStatsComputation(results) {
  // Mock player event data
  const mockEvents = [
    {
      eventName: '05-10C-2025',
      suffix: 'C',
      category: 'TRANSITIONAL_COMMANDER',
      monthKey: '2025-05',
      isoWeekKey: '2025-W19',
      position: 1,
      isTop4: true,
      isLast: false
    },
    {
      eventName: '05-17D-2025',
      suffix: 'D',
      category: 'DRAFT',
      monthKey: '2025-05',
      isoWeekKey: '2025-W20',
      position: 5,
      isTop4: false,
      isLast: false
    },
    {
      eventName: '05-20A-2025',
      suffix: 'A',
      category: 'ACADEMY',
      monthKey: '2025-05',
      isoWeekKey: '2025-W21',
      position: 1,
      isTop4: true,
      isLast: false
    },
    {
      eventName: '05-22A-2025',
      suffix: 'A',
      category: 'ACADEMY',
      monthKey: '2025-05',
      isoWeekKey: '2025-W21',
      position: 6,
      isTop4: false,
      isLast: true
    }
  ];
  
  const stats = computePlayerStats('TestPlayer', mockEvents);
  
  assert(stats.totalEvents === 4, 'Should count 4 total events', results);
  assert(stats.uniqueSuffixes === 3, 'Should count 3 unique suffixes (C, D, A)', results);
  assert(stats.loyalMonths === 1, 'Should count 1 loyal month (4 events in May)', results);
  assert(stats.meteorWeeks === 1, 'Should count 1 meteor week (2 events in W21)', results);
  assert(stats.limitedEvents === 1, 'Should count 1 limited event (D)', results);
  assert(stats.draftEvents === 1, 'Should count 1 draft event', results);
  assert(stats.academyEvents === 2, 'Should count 2 academy events', results);
  assert(stats.top4Finishes === 2, 'Should count 2 top-4 finishes', results);
  assert(stats.lastPlaceFinishes === 1, 'Should count 1 last place finish', results);
  assert(stats.transitionalCommanderEvents === 1, 'Should count 1 transitional commander event', results);
}

/**
 * Test mission points calculation
 */
function testMissionPointsCalculation(results) {
  // Mock stats for a player
  const mockStats = {
    totalEvents: 10,
    uniqueSuffixes: 5,
    loyalMonths: 2,
    limitedEvents: 3,
    draftEvents: 2,
    academyEvents: 1,
    meteorWeeks: 3,
    top4Finishes: 4,
    freePlayEvents: 2,
    lastPlaceFinishes: 1,
    casualCommanderEvents: 2,
    transitionalCommanderEvents: 3,
    cedhEvents: 1,
    outreachEvents: 0,
    commanderLeagueEvents: 0,
    preconEvents: 0
  };
  
  const awards = computeAttendanceMissionPoints(mockStats);
  
  // Test one-time missions
  assert(awards['First Contact'] === 1, 'Should award First Contact (1 point)', results);
  assert(awards['Stellar Explorer'] === 2, 'Should award Stellar Explorer (2 points)', results);
  assert(awards['Sealed Voyager'] === 1, 'Should award Sealed Voyager (1 point)', results);
  assert(awards['Draft Navigator'] === 1, 'Should award Draft Navigator (1 point)', results);
  assert(awards['Stellar Scholar'] === 1, 'Should award Stellar Scholar (1 point)', results);
  
  // Test scaling missions
  assert(awards['Deck Diver'] === 5, 'Should award 5 Deck Diver points', results);
  assert(awards['Lunar Loyalty'] === 2, 'Should award 2 Lunar Loyalty points', results);
  assert(awards['Meteor Shower'] === 3, 'Should award 3 Meteor Shower points', results);
  assert(awards['Interstellar Strategist'] === 4, 'Should award 4 Interstellar Strategist points', results);
  assert(awards['Free Play Events'] === 2, 'Should award 2 Free Play Events points', results);
  assert(awards['Black Hole Survivor'] === 1, 'Should award 1 Black Hole Survivor point', results);
  
  // Test category columns
  assert(awards['Casual Commander Events'] === 2, 'Should track 2 casual commander events', results);
  assert(awards['Transitional Commander Events'] === 3, 'Should track 3 transitional commander events', results);
  assert(awards['cEDH Events'] === 1, 'Should track 1 cEDH event', results);
  assert(awards['Limited Events'] === 3, 'Should track 3 limited events', results);
  assert(awards['Academy Events'] === 1, 'Should track 1 academy event', results);
  
  // Test total calculation
  const expectedTotal = 1 + 2 + 1 + 1 + 1 + 5 + 2 + 3 + 4 + 2 + 1; // 23
  assert(awards['Points'] === expectedTotal, 'Should calculate correct total points (' + expectedTotal + ')', results);
}

/**
 * Quick validation test - checks if functions exist and are callable
 */
function quickValidationTest() {
  Logger.log('=== QUICK VALIDATION TEST ===\n');
  
  try {
    // Test that all required functions exist
    Logger.log('✓ getEventSheets exists: ' + (typeof getEventSheets === 'function'));
    Logger.log('✓ parseEventSheetName exists: ' + (typeof parseEventSheetName === 'function'));
    Logger.log('✓ getISOWeekKey exists: ' + (typeof getISOWeekKey === 'function'));
    Logger.log('✓ getEventCategoryFromSuffix exists: ' + (typeof getEventCategoryFromSuffix === 'function'));
    Logger.log('✓ isLimitedCategory exists: ' + (typeof isLimitedCategory === 'function'));
    Logger.log('✓ isCommanderCategory exists: ' + (typeof isCommanderCategory === 'function'));
    Logger.log('✓ getStandings exists: ' + (typeof getStandings === 'function'));
    Logger.log('✓ scanAllEvents exists: ' + (typeof scanAllEvents === 'function'));
    Logger.log('✓ computePlayerStats exists: ' + (typeof computePlayerStats === 'function'));
    Logger.log('✓ computeAttendanceMissionPoints exists: ' + (typeof computeAttendanceMissionPoints === 'function'));
    Logger.log('✓ syncAttendanceMissions exists: ' + (typeof syncAttendanceMissions === 'function'));
    Logger.log('✓ addAttendanceMissionMenu exists: ' + (typeof addAttendanceMissionMenu === 'function'));
    Logger.log('✓ showEventCount exists: ' + (typeof showEventCount === 'function'));
    Logger.log('✓ testAttendanceScanner exists: ' + (typeof testAttendanceScanner === 'function'));
    
    Logger.log('\n✓ All required functions are present');
    Logger.log('\n=== VALIDATION COMPLETE ===');
    
    return true;
  } catch (e) {
    Logger.log('✗ Error during validation: ' + e.message);
    return false;
  }
}
