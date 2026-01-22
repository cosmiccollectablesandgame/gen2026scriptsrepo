/**
 * Test runner for Attendance Mission Scanner
 * This simulates the Google Apps Script environment for local testing
 */

// Mock Logger for Google Apps Script
global.Logger = {
  log: function(message) {
    console.log(message);
  }
};

// Load the implementation
eval(require('fs').readFileSync('./attendanceMissionScanner.js', 'utf8'));

// Load the tests
eval(require('fs').readFileSync('./attendanceMissionScanner.test.js', 'utf8'));

// Run tests
console.log('Starting Attendance Mission Scanner Tests...\n');

try {
  // Run quick validation first
  console.log('Running quick validation...\n');
  quickValidationTest();
  
  console.log('\n\nRunning full test suite...\n');
  const results = testAttendanceMissionScanner();
  
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
