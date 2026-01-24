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
  // Test case 1: Actual WER smushed format - SINGLE LINE, space-separated
  // This is the REAL format from the problem statement
  const input1 = "Cy Diskin93/0/055.6%100.0%55.6% " +
                 "2Justin Johnson93/0/055.6%85.7%51.0% " +
                 "3Jake Martin93/0/050.0%85.7%47.3% " +
                 "4Michael Davis93/0/044.4%75.0%44.4%";
  const result1 = smartParseRoster(input1);
  const expected1 = ["Cy Diskin", "Justin Johnson", "Jake Martin", "Michael Davis"];
  
  assert(
    arraysEqual(result1, expected1),
    'WER Smushed Format - Actual single line format (Format A)',
    results
  );
  
  // Test case 2: Single player on single line
  const input2 = "Cy Diskin93/0/055.6%100.0%55.6%";
  const result2 = smartParseRoster(input2);
  const expected2 = ["Cy Diskin"];
  
  assert(
    arraysEqual(result2, expected2),
    'WER Smushed Format - Single player',
    results
  );
  
  // Test case 3: Names with apostrophes on single line
  const input3 = "Patrick O'Brien93/0/055.6%100.0%55.6% 2John O'Malley82/1/044.4%71.4%42.7%";
  const result3 = smartParseRoster(input3);
  const expected3 = ["Patrick O'Brien", "John O'Malley"];
  
  assert(
    arraysEqual(result3, expected3),
    'WER Smushed Format - Names with apostrophes',
    results
  );
  
  // Test case 4: Names with hyphens on single line
  const input4 = "Jean-Luc Picard93/0/055.6%100.0%55.6% 2Mary-Kate Olsen71/2/033.3%57.1%38.2%";
  const result4 = smartParseRoster(input4);
  const expected4 = ["Jean-Luc Picard", "Mary-Kate Olsen"];
  
  assert(
    arraysEqual(result4, expected4),
    'WER Smushed Format - Names with hyphens',
    results
  );
  
  // Test case 5: Single word names on single line
  const input5 = "Cher93/0/055.6%100.0%55.6% 2Madonna82/1/044.4%71.4%42.7% 3Prince71/2/033.3%57.1%38.2%";
  const result5 = smartParseRoster(input5);
  const expected5 = ["Cher", "Madonna", "Prince"];
  
  assert(
    arraysEqual(result5, expected5),
    'WER Smushed Format - Single word names',
    results
  );
  
  // Test case 6: Proper casing on single line
  const input6 = "lowercase name93/0/055.6%100.0%55.6% 2UPPERCASE NAME82/1/044.4%71.4%42.7%";
  const result6 = smartParseRoster(input6);
  const expected6 = ["Lowercase Name", "Uppercase Name"];
  
  assert(
    arraysEqual(result6, expected6),
    'WER Smushed Format - Proper casing applied',
    results
  );
  
  // Test case 7: Names with initials on single line
  const input7 = "Jeremy B93/0/055.6%100.0%55.6% 2Michael J Fox82/1/044.4%71.4%42.7%";
  const result7 = smartParseRoster(input7);
  const expected7 = ["Jeremy B", "Michael J Fox"];
  
  assert(
    arraysEqual(result7, expected7),
    'WER Smushed Format - Names with initials',
    results
  );
  
  // Test case 8: High rank numbers (double digits)
  const input8 = "John Smith93/0/055.6%100.0%55.6% 14Walker Beck62/1/044.4%71.4%42.7% 32Mcquaid Beck20/1/255.6%33.3%47.2%";
  const result8 = smartParseRoster(input8);
  const expected8 = ["John Smith", "Walker Beck", "Mcquaid Beck"];
  
  assert(
    arraysEqual(result8, expected8),
    'WER Smushed Format - High rank numbers',
    results
  );
  
  // Test case 9: Multiple percentages in sequence
  const input9 = "Alice Johnson93/0/055.6%100.0%55.6%62.3% 2Bob Williams82/1/044.4%71.4%42.7%89.2%";
  const result9 = smartParseRoster(input9);
  const expected9 = ["Alice Johnson", "Bob Williams"];
  
  assert(
    arraysEqual(result9, expected9),
    'WER Smushed Format - Multiple percentages',
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
  
  // Test EventLink format (Format B) - junk lines with newlines
  const input6 = `John Jones
None
Not Submitted
Manual
Jeremy Bookland
None
Not Submitted
Manual`;
  const result6 = smartParseRoster(input6);
  
  assert(
    result6.length === 2 && result6.includes("John Jones") && result6.includes("Jeremy Bookland"),
    'Existing Format - EventLink with junk lines (Format B)',
    results
  );
}
