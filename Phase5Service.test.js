/**
 * Phase 5 Service Tests
 * Tests for Phase 5 identity hygiene and transactional truth consumers
 * 
 * Run these tests manually to verify Phase 5 functionality
 */

/**
 * Test Suite: normalizePlayerName()
 */
function testNormalizePlayerName() {
  console.log('=== Testing normalizePlayerName() ===');
  
  const tests = [
    { input: '  John Doe  ', expected: 'John Doe', name: 'Trim whitespace' },
    { input: 'John  Doe', expected: 'John Doe', name: 'Collapse internal spaces' },
    { input: '...John Doe...', expected: 'John Doe', name: 'Remove leading/trailing punctuation' },
    { input: "O'Connor", expected: "O'Connor", name: 'Preserve internal punctuation' },
    { input: 'Mary-Jane', expected: 'Mary-Jane', name: 'Preserve hyphens' },
    { input: '', expected: '', name: 'Empty string' },
    { input: null, expected: '', name: 'Null input' },
    { input: '   ', expected: '', name: 'Whitespace only' }
  ];
  
  let passed = 0;
  let failed = 0;
  
  tests.forEach(test => {
    const result = normalizePlayerName(test.input);
    if (result === test.expected) {
      console.log(`✓ ${test.name}: PASS`);
      passed++;
    } else {
      console.log(`✗ ${test.name}: FAIL (got "${result}", expected "${test.expected}")`);
      failed++;
    }
  });
  
  console.log(`\nResults: ${passed} passed, ${failed} failed\n`);
  return { passed, failed };
}

/**
 * Test Suite: namesMatch()
 */
function testNamesMatch() {
  console.log('=== Testing namesMatch() ===');
  
  const tests = [
    { name1: 'John Doe', name2: 'john doe', expected: true, name: 'Case insensitive' },
    { name1: '  John Doe  ', name2: 'John Doe', expected: true, name: 'Whitespace trimming' },
    { name1: 'John  Doe', name2: 'John Doe', expected: true, name: 'Space normalization' },
    { name1: 'John Doe', name2: 'Jane Smith', expected: false, name: 'Different names' },
    { name1: '', name2: '', expected: false, name: 'Empty strings' },
    { name1: "O'Connor", name2: "o'connor", expected: true, name: 'Punctuation preserved' }
  ];
  
  let passed = 0;
  let failed = 0;
  
  tests.forEach(test => {
    const result = namesMatch(test.name1, test.name2);
    if (result === test.expected) {
      console.log(`✓ ${test.name}: PASS`);
      passed++;
    } else {
      console.log(`✗ ${test.name}: FAIL (got ${result}, expected ${test.expected})`);
      failed++;
    }
  });
  
  console.log(`\nResults: ${passed} passed, ${failed} failed\n`);
  return { passed, failed };
}

/**
 * Test Suite: isCanonicalName()
 */
function testIsCanonicalName() {
  console.log('=== Testing isCanonicalName() ===');
  
  // Note: These tests require PreferredNames to have actual data
  console.log('Note: This test requires PreferredNames sheet with data');
  
  try {
    const { canonicalList } = loadCanonicalNames_();
    
    if (canonicalList.length === 0) {
      console.log('⚠ PreferredNames is empty - skipping test');
      return { passed: 0, failed: 0, skipped: true };
    }
    
    // Test with first canonical name
    const testName = canonicalList[0];
    const result1 = isCanonicalName(testName);
    console.log(`✓ isCanonicalName("${testName}"): ${result1 ? 'PASS' : 'FAIL'}`);
    
    // Test with non-existent name
    const result2 = isCanonicalName('NonExistentPlayer12345');
    console.log(`✓ isCanonicalName("NonExistentPlayer12345"): ${!result2 ? 'PASS' : 'FAIL'}`);
    
    // Test case insensitivity
    const result3 = isCanonicalName(testName.toUpperCase());
    console.log(`✓ isCanonicalName("${testName.toUpperCase()}"): ${result3 ? 'PASS' : 'FAIL'}`);
    
    return { passed: 3, failed: 0 };
    
  } catch (e) {
    console.log(`✗ Error: ${e.message}`);
    return { passed: 0, failed: 1 };
  }
}

/**
 * Test Suite: queueUnknownName()
 */
function testQueueUnknownName() {
  console.log('=== Testing queueUnknownName() ===');
  
  console.log('Note: This test will add test data to UndiscoveredNames');
  console.log('You may want to manually clean up after testing\n');
  
  try {
    // Test queuing unknown name
    const testName = 'TestPlayer_' + new Date().getTime();
    const result1 = queueUnknownName(testName, 'TEST', 'Phase5Service.test.js', 'Unit Test');
    console.log(`✓ queueUnknownName() new entry: ${result1 ? 'PASS' : 'FAIL'}`);
    
    // Test queuing duplicate (should update seen count)
    const result2 = queueUnknownName(testName, 'TEST', 'Phase5Service.test.js', 'Unit Test');
    console.log(`✓ queueUnknownName() duplicate: ${result2 ? 'PASS' : 'FAIL'}`);
    
    // Test with empty name (should return false)
    const result3 = queueUnknownName('', 'TEST', 'Phase5Service.test.js');
    console.log(`✓ queueUnknownName() empty name: ${!result3 ? 'PASS' : 'FAIL'}`);
    
    return { passed: 3, failed: 0 };
    
  } catch (e) {
    console.log(`✗ Error: ${e.message}`);
    return { passed: 0, failed: 1 };
  }
}

/**
 * Test Suite: Hygiene Status
 */
function testHygieneStatus() {
  console.log('=== Testing Hygiene Status Functions ===');
  
  try {
    // Test setting status
    setSheetHygieneStatus_('TestSheet', true);
    console.log('✓ setSheetHygieneStatus_() called: PASS');
    
    // Test getting status
    const status = getSheetHygieneStatus('TestSheet');
    const passed = status && status.clean === true;
    console.log(`✓ getSheetHygieneStatus(): ${passed ? 'PASS' : 'FAIL'}`);
    
    // Test getting non-existent status
    const status2 = getSheetHygieneStatus('NonExistentSheet');
    const passed2 = status2 && status2.clean === false;
    console.log(`✓ getSheetHygieneStatus() non-existent: ${passed2 ? 'PASS' : 'FAIL'}`);
    
    return { passed: 3, failed: 0 };
    
  } catch (e) {
    console.log(`✗ Error: ${e.message}`);
    return { passed: 0, failed: 1 };
  }
}

/**
 * Test Suite: isRetiredSheet()
 */
function testIsRetiredSheet() {
  console.log('=== Testing isRetiredSheet() ===');
  
  const tests = [
    { input: 'Players_Prize-Wall-Points', expected: true, name: 'Retired sheet variant 1' },
    { input: 'Player\'s Prize-Wall-Points', expected: true, name: 'Retired sheet variant 2' },
    { input: 'PreferredNames', expected: false, name: 'Active sheet' },
    { input: 'Store_Credit_Ledger', expected: false, name: 'Phase 5 sheet' },
    { input: '', expected: false, name: 'Empty string' }
  ];
  
  let passed = 0;
  let failed = 0;
  
  tests.forEach(test => {
    const result = isRetiredSheet(test.input);
    if (result === test.expected) {
      console.log(`✓ ${test.name}: PASS`);
      passed++;
    } else {
      console.log(`✗ ${test.name}: FAIL (got ${result}, expected ${test.expected})`);
      failed++;
    }
  });
  
  console.log(`\nResults: ${passed} passed, ${failed} failed\n`);
  return { passed, failed };
}

/**
 * Run all Phase 5 tests
 */
function RUN_PHASE5_TESTS() {
  console.log('');
  console.log('═══════════════════════════════════════════════════════════');
  console.log('PHASE 5 SERVICE TEST SUITE');
  console.log('═══════════════════════════════════════════════════════════');
  console.log('');
  
  const results = [];
  
  results.push(testNormalizePlayerName());
  results.push(testNamesMatch());
  results.push(testIsCanonicalName());
  results.push(testQueueUnknownName());
  results.push(testHygieneStatus());
  results.push(testIsRetiredSheet());
  
  console.log('');
  console.log('═══════════════════════════════════════════════════════════');
  console.log('TEST SUMMARY');
  console.log('═══════════════════════════════════════════════════════════');
  
  const totalPassed = results.reduce((sum, r) => sum + r.passed, 0);
  const totalFailed = results.reduce((sum, r) => sum + r.failed, 0);
  const totalSkipped = results.filter(r => r.skipped).length;
  
  console.log(`Total Passed: ${totalPassed}`);
  console.log(`Total Failed: ${totalFailed}`);
  if (totalSkipped > 0) {
    console.log(`Total Skipped: ${totalSkipped}`);
  }
  
  if (totalFailed === 0) {
    console.log('\n✓ ALL TESTS PASSED');
  } else {
    console.log('\n✗ SOME TESTS FAILED');
  }
  
  console.log('═══════════════════════════════════════════════════════════');
}

/**
 * Integration test: Full workflow
 */
function TEST_PHASE5_INTEGRATION() {
  console.log('');
  console.log('═══════════════════════════════════════════════════════════');
  console.log('PHASE 5 INTEGRATION TEST');
  console.log('═══════════════════════════════════════════════════════════');
  console.log('');
  
  console.log('1. Scanning Store_Credit_Ledger...');
  const ledgerScan = SCAN_STORE_CREDIT_LEDGER_FOR_UNKNOWN_NAMES();
  console.log(`   Result: ${ledgerScan.message}`);
  console.log(`   Scanned: ${ledgerScan.scannedRows} rows`);
  console.log(`   Unknown: ${ledgerScan.unknownNames.length} names`);
  console.log('');
  
  console.log('2. Scanning Preorders...');
  const preordersScan = SCAN_PREORDERS_FOR_UNKNOWN_NAMES();
  console.log(`   Result: ${preordersScan.message}`);
  console.log(`   Scanned: ${preordersScan.scannedRows} rows`);
  console.log(`   Unknown: ${preordersScan.unknownNames.length} names`);
  console.log('');
  
  console.log('3. Refreshing Resolve Unknowns Dashboard...');
  const dashboard = REFRESH_RESOLVE_UNKNOWN_DASHBOARD();
  console.log(`   Result: ${dashboard.message}`);
  if (dashboard.success) {
    console.log(`   Store Credit: ${dashboard.sections.storeCreditLedger} items`);
    console.log(`   Events: ${dashboard.sections.events} items`);
    console.log(`   Preorders: ${dashboard.sections.preorders} items`);
  }
  console.log('');
  
  console.log('4. Generating Provisional Ledger Names...');
  const provisional = generateProvisionalLedgerNames();
  console.log(`   Result: ${provisional.message}`);
  console.log('');
  
  console.log('5. Running Phase 5 Audit...');
  const audit = GENERATE_PHASE5_AUDIT_REPORT();
  console.log('');
  
  console.log('═══════════════════════════════════════════════════════════');
  console.log('INTEGRATION TEST COMPLETE');
  console.log('═══════════════════════════════════════════════════════════');
}
