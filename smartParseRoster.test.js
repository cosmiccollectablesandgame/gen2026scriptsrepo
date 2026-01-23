/**
 * ════════════════════════════════════════════════════════════════════════════
 * SMART PARSE ROSTER - TEST SUITE
 * ════════════════════════════════════════════════════════════════════════════
 *
 * @fileoverview Tests for the smartParseRoster function in eventService.js
 * Specifically testing the WER smushed format parsing fix
 * 
 * Run testSmartParseRoster() to execute all tests
 *
 * Version: 1.0.0
 * ════════════════════════════════════════════════════════════════════════════
 */

/**
 * Master test runner
 */
function testSmartParseRoster() {
  Logger.log('========================================');
  Logger.log('SMART PARSE ROSTER TEST SUITE');
  Logger.log('========================================\n');

  const results = {
    total: 0,
    passed: 0,
    failed: 0,
    errors: []
  };

  // Run all test suites
  runTestSuite('WER Smushed Format', testWERSmushedFormat, results);
  runTestSuite('Existing Format Compatibility', testExistingFormats, results);

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

  return results.failed === 0;
}

/**
 * Helper to run a test suite
 */
function runTestSuite(name, testFunc, results) {
  Logger.log('Running: ' + name);
  try {
    testFunc(results);
  } catch (e) {
    results.errors.push(name + ': ' + e.message);
    Logger.log('  ✗ Suite Error: ' + e.message);
  }
  Logger.log('');
}

/**
 * Helper to run individual test
 */
function assert(condition, testName, results) {
  results.total++;
  if (condition) {
    results.passed++;
    Logger.log('  ✓ ' + testName);
  } else {
    results.failed++;
    results.errors.push(testName);
    Logger.log('  ✗ ' + testName);
  }
}

/**
 * Helper to compare arrays
 */
function arraysEqual(a, b) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}

/**
 * Test WER Smushed Format Parsing
 */
function testWERSmushedFormat(results) {
  // Test case 1: Basic WER smushed format
  const input1 = `1Cy Diskin93/0/055.6%100.0%55.6%
2Justin Johnson93/0/055.6%85.7%51.0%
14walker beck62/1/044.4%71.4%42.7%
32mcquaid beck20/1/255.6%33.3%47.2%`;
  
  const result1 = smartParseRoster(input1);
  const expected1 = ["Cy Diskin", "Justin Johnson", "Walker Beck", "Mcquaid Beck"];
  
  assert(
    arraysEqual(result1, expected1),
    'WER Smushed Format - Basic parsing',
    results
  );
  
  // Test case 2: Single entry
  const input2 = "1John Smith93/0/055.6%100.0%55.6%";
  const result2 = smartParseRoster(input2);
  const expected2 = ["John Smith"];
  
  assert(
    arraysEqual(result2, expected2),
    'WER Smushed Format - Single entry',
    results
  );
  
  // Test case 3: Names with apostrophes
  const input3 = "5Patrick O'Brien82/1/044.4%71.4%42.7%";
  const result3 = smartParseRoster(input3);
  const expected3 = ["Patrick O'Brien"];
  
  assert(
    arraysEqual(result3, expected3),
    'WER Smushed Format - Name with apostrophe',
    results
  );
  
  // Test case 4: Names with hyphens
  const input4 = "12Jean-Luc Picard71/2/033.3%57.1%38.2%";
  const result4 = smartParseRoster(input4);
  const expected4 = ["Jean-Luc Picard"];
  
  assert(
    arraysEqual(result4, expected4),
    'WER Smushed Format - Name with hyphen',
    results
  );
  
  // Test case 5: Mixed with regular formats
  const input5 = `1Cy Diskin93/0/055.6%100.0%55.6%
1. Alice Williams
2 Bob Jones`;
  
  const result5 = smartParseRoster(input5);
  
  assert(
    result5.includes("Cy Diskin") && result5.includes("Alice Williams") && result5.includes("Bob Jones"),
    'WER Smushed Format - Mixed with regular formats',
    results
  );
  
  // Test case 6: Proper casing verification
  const input6 = "8lowercase name51/1/144.4%71.4%42.7%";
  const result6 = smartParseRoster(input6);
  const expected6 = ["Lowercase Name"];
  
  assert(
    arraysEqual(result6, expected6),
    'WER Smushed Format - Proper casing applied',
    results
  );
  
  // Test case 7: Single word name
  const input7 = "3Cher93/0/055.6%100.0%55.6%";
  const result7 = smartParseRoster(input7);
  const expected7 = ["Cher"];
  
  assert(
    arraysEqual(result7, expected7),
    'WER Smushed Format - Single word name',
    results
  );
}

/**
 * Test that existing formats still work
 */
function testExistingFormats(results) {
  // Test numbered list
  const input1 = `1. John Smith
2. Jane Doe
3. Bob Wilson`;
  const result1 = smartParseRoster(input1);
  
  assert(
    result1.length === 3 && result1.includes("John Smith") && result1.includes("Jane Doe"),
    'Existing Format - Numbered list',
    results
  );
  
  // Test with W-L-D records
  const input2 = `John Smith 2-1-0
Jane Doe 3-0-0`;
  const result2 = smartParseRoster(input2);
  
  assert(
    result2.length === 2 && result2.includes("John Smith") && result2.includes("Jane Doe"),
    'Existing Format - W-L-D records',
    results
  );
  
  // Test with trailing numbers
  const input3 = `John Smith 9 66 85 69
Jane Doe 12 88 92 71`;
  const result3 = smartParseRoster(input3);
  
  assert(
    result3.length === 2 && result3.includes("John Smith") && result3.includes("Jane Doe"),
    'Existing Format - Trailing numbers',
    results
  );
  
  // Test with parenthetical notes
  const input4 = `John Smith (Active)
Jane Doe (Dropped)`;
  const result4 = smartParseRoster(input4);
  
  assert(
    result4.length === 2 && result4.includes("John Smith") && result4.includes("Jane Doe"),
    'Existing Format - Parenthetical notes',
    results
  );
  
  // Test deduplication
  const input5 = `John Smith
john smith
JOHN SMITH`;
  const result5 = smartParseRoster(input5);
  
  assert(
    result5.length === 1 && result5[0] === "John Smith",
    'Existing Format - Deduplication',
    results
  );
}
