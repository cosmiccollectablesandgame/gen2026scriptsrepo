/**
 * Test runner for Smart Parse Roster
 * This simulates the Google Apps Script environment for local testing
 */

// Mock Logger for Google Apps Script
global.Logger = {
  log: function(message) {
    console.log(message);
  }
};

// Load the implementation
eval(require('fs').readFileSync('./eventService.js', 'utf8'));

// Load the tests
eval(require('fs').readFileSync('./smartParseRoster.test.js', 'utf8'));

// Run tests
console.log('Starting Smart Parse Roster Tests...\n');

try {
  const results = testSmartParseRoster();
  
  // Exit with appropriate code
  if (results.failed > 0) {
    console.log('\n❌ Tests FAILED');
    process.exit(1);
  } else {
    console.log('\n✅ All tests PASSED');
    process.exit(0);
  }
} catch (e) {
  console.error('\n❌ Test execution failed:', e.message);
  console.error(e.stack);
  process.exit(1);
}
