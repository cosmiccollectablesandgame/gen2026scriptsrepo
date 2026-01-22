/**
 * ════════════════════════════════════════════════════════════════════════════
 * PLAYER PROVISIONING - TEST SUITE
 * ════════════════════════════════════════════════════════════════════════════
 */

function testPlayerProvisioning() {
  Logger.log('========================================');
  Logger.log('PLAYER PROVISIONING TEST SUITE');
  Logger.log('========================================\n');
  
  const results = { total: 0, passed: 0, failed: 0, errors: [] };
  
  runTestSuite_('Name Normalization', testNameNormalization_, results);
  runTestSuite_('Player Exists Check', testPlayerExists_, results);
  runTestSuite_('Get All Preferred Names', testGetAllPreferredNames_, results);
  runTestSuite_('Provision Targets Config', testProvisionTargetsConfig_, results);
  runTestSuite_('Sheet Lookup', testSheetLookup_, results);
  
  // Summary
  Logger.log('\n========================================');
  Logger.log('RESULTS: ' + results.passed + '/' + results.total + ' passed');
  if (results.failed > 0) {
    Logger.log('FAILED: ' + results.errors.join(', '));
  }
  Logger.log('========================================');
  
  return results;
}

function runTestSuite_(name, fn, results) {
  Logger.log('\n── ' + name + ' ──');
  try {
    fn(results);
  } catch (e) {
    results.total++;
    results.failed++;
    results.errors.push(name + ': ' + e.message);
    Logger.log('  ✗ CRASHED: ' + e.message);
  }
}

function assert_(condition, message, results) {
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
// TEST CASES
// ════════════════════════════════════════════════════════════════════════════

function testNameNormalization_(results) {
  assert_(normalizePlayerName('  John Doe  ') === 'John Doe', 'Should trim whitespace', results);
  assert_(normalizePlayerName('John   Doe') === 'John Doe', 'Should collapse multiple spaces', results);
  assert_(normalizePlayerName('') === '', 'Should handle empty string', results);
  assert_(normalizePlayerName(null) === '', 'Should handle null', results);
  assert_(normalizePlayerName(undefined) === '', 'Should handle undefined', results);
  
  assert_(playerNamesMatch('John Doe', 'john doe'), 'Should match case-insensitive', results);
  assert_(playerNamesMatch('John Doe', '  JOHN DOE  '), 'Should match with whitespace differences', results);
  assert_(!playerNamesMatch('John Doe', 'Jane Doe'), 'Should not match different names', results);
}

function testPlayerExists_(results) {
  // This test requires PreferredNames sheet to exist
  const exists = typeof playerExists === 'function';
  assert_(exists, 'playerExists function should exist', results);
  
  if (exists) {
    // Test with a name that definitely doesn't exist
    const fakeResult = playerExists('ZZZZZ_FAKE_PLAYER_12345');
    assert_(fakeResult === false, 'Should return false for non-existent player', results);
  }
}

function testGetAllPreferredNames_(results) {
  const exists = typeof getAllPreferredNames === 'function';
  assert_(exists, 'getAllPreferredNames function should exist', results);
  
  if (exists) {
    const names = getAllPreferredNames();
    assert_(Array.isArray(names), 'Should return an array', results);
    Logger.log('    Found ' + names.length + ' players');
  }
}

function testProvisionTargetsConfig_(results) {
  assert_(typeof PROVISION_TARGETS === 'object', 'PROVISION_TARGETS should exist', results);
  assert_(PROVISION_TARGETS.BP_Total !== undefined, 'Should have BP_Total config', results);
  assert_(PROVISION_TARGETS.Attendance_Missions !== undefined, 'Should have Attendance_Missions config', results);
  assert_(PROVISION_TARGETS.Flag_Missions !== undefined, 'Should have Flag_Missions config', results);
  assert_(PROVISION_TARGETS.Dice_Points !== undefined, 'Should have Dice_Points config', results);
  assert_(PROVISION_TARGETS.Key_Tracker !== undefined, 'Should have Key_Tracker config', results);
  
  // Check structure
  assert_(PROVISION_TARGETS.BP_Total.sheetName === 'BP_Total', 'BP_Total sheetName should be correct', results);
  assert_(PROVISION_TARGETS.BP_Total.keyColumn === 'preferred_name_id', 'BP_Total keyColumn should be correct', results);
  assert_(typeof PROVISION_TARGETS.BP_Total.defaults === 'object', 'BP_Total should have defaults', results);
}

function testSheetLookup_(results) {
  const exists = typeof getSheetCI_ === 'function';
  assert_(exists, 'getSheetCI_ function should exist', results);
  
  if (exists) {
    // Test case-insensitive lookup
    const sheet1 = getSheetCI_('PreferredNames');
    const sheet2 = getSheetCI_('preferrednames');
    const sheet3 = getSheetCI_('PREFERREDNAMES');
    
    if (sheet1) {
      assert_(sheet1 !== null, 'Should find PreferredNames', results);
      assert_(sheet2 !== null, 'Should find preferrednames (lowercase)', results);
      assert_(sheet3 !== null, 'Should find PREFERREDNAMES (uppercase)', results);
    } else {
      Logger.log('    (PreferredNames sheet not found - skipping case tests)');
    }
    
    const fake = getSheetCI_('ZZZZZ_FAKE_SHEET');
    assert_(fake === null, 'Should return null for non-existent sheet', results);
  }
}

/**
 * Quick validation - checks all required functions exist
 */
function quickValidateProvisioning() {
  Logger.log('=== QUICK VALIDATION ===\n');
  
  const required = [
    'normalizePlayerName',
    'playerNamesMatch',
    'getAllPreferredNames',
    'playerExists',
    'getSheetCI_',
    'getHeaderMap_',
    'provisionToSheet_',
    'provisionPlayerProfile',
    'addNewPlayer',
    'provisionMultiplePlayers',
    'getUnprovisionedPlayers',
    'discoverAndProvisionNewPlayers',
    'runFullProvisioning'
  ];
  
  let allPresent = true;
  for (const fn of required) {
    const exists = typeof globalThis[fn] === 'function';
    Logger.log((exists ? '✓' : '✗') + ' ' + fn);
    if (!exists) allPresent = false;
  }
  
  Logger.log('\n' + (allPresent ? '✅ All functions present' : '❌ Missing functions'));
  return allPresent;
}
