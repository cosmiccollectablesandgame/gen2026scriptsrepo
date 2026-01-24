/**
 * Manual verification of the fix with exact data from problem statement
 */

// Mock Logger
global.Logger = {
  log: function(message) {
    console.log(message);
  }
};

// Load the implementation
eval(require('fs').readFileSync('./eventService.js', 'utf8'));

console.log('========================================');
console.log('MANUAL VERIFICATION WITH PROBLEM DATA');
console.log('========================================\n');

// Test with exact Format A from problem statement
console.log('Test 1: Format A - WER Smushed (single line, space-separated)');
const formatA = "Cy Diskin93/0/055.6%100.0%55.6% 2Justin Johnson93/0/055.6%85.7%51.0% 3Jake Martin93/0/050.0%85.7%47.3% 4Michael Davis93/0/044.4%75.0%44.4%";
const resultA = smartParseRoster(formatA);
console.log('Input:', formatA);
console.log('Output:', JSON.stringify(resultA));
console.log('Expected: ["Cy Diskin", "Justin Johnson", "Jake Martin", "Michael Davis"]');
console.log('✓ Pass:', JSON.stringify(resultA) === JSON.stringify(["Cy Diskin", "Justin Johnson", "Jake Martin", "Michael Davis"]));
console.log('');

// Test with exact Format B from problem statement
console.log('Test 2: Format B - EventLink junk lines (newline-separated)');
const formatB = `John Jones
None
Not Submitted
Manual
Jeremy Bookland
None
Not Submitted
Manual`;
const resultB = smartParseRoster(formatB);
console.log('Input (truncated):', formatB.split('\n').slice(0, 4).join(' / ') + '...');
console.log('Output:', JSON.stringify(resultB));
console.log('Expected: ["John Jones", "Jeremy Bookland"]');
console.log('✓ Pass:', JSON.stringify(resultB) === JSON.stringify(["John Jones", "Jeremy Bookland"]));
console.log('');

// Test edge case: First player with no rank
console.log('Test 3: First player has no rank (as per problem statement)');
const noRankFirst = "Cy Diskin93/0/055.6%100.0%55.6% 2Justin Johnson93/0/055.6%85.7%51.0%";
const resultNoRank = smartParseRoster(noRankFirst);
console.log('Input:', noRankFirst);
console.log('Output:', JSON.stringify(resultNoRank));
console.log('Expected: ["Cy Diskin", "Justin Johnson"]');
console.log('✓ Pass:', JSON.stringify(resultNoRank) === JSON.stringify(["Cy Diskin", "Justin Johnson"]));
console.log('');

// Test with names with apostrophes
console.log("Test 4: Names with apostrophes (O'Brien)");
const withApostrophe = "Patrick O'Brien93/0/055.6%100.0%55.6% 2Mary O'Connor82/1/044.4%71.4%42.7%";
const resultApostrophe = smartParseRoster(withApostrophe);
console.log('Input:', withApostrophe);
console.log('Output:', JSON.stringify(resultApostrophe));
console.log('Expected: ["Patrick O\'Brien", "Mary O\'Connor"]');
console.log('✓ Pass:', JSON.stringify(resultApostrophe) === JSON.stringify(["Patrick O'Brien", "Mary O'Connor"]));
console.log('');

// Test with names with hyphens
console.log('Test 5: Names with hyphens (Jean-Luc)');
const withHyphen = "Jean-Luc Picard93/0/055.6%100.0%55.6% 2Mary-Kate Ashley82/1/044.4%71.4%42.7%";
const resultHyphen = smartParseRoster(withHyphen);
console.log('Input:', withHyphen);
console.log('Output:', JSON.stringify(resultHyphen));
console.log('Expected: ["Jean-Luc Picard", "Mary-Kate Ashley"]');
console.log('✓ Pass:', JSON.stringify(resultHyphen) === JSON.stringify(["Jean-Luc Picard", "Mary-Kate Ashley"]));
console.log('');

// Test with names with initials
console.log('Test 6: Names with initials (Jeremy B)');
const withInitial = "Jeremy B93/0/055.6%100.0%55.6% 2Michael J Fox82/1/044.4%71.4%42.7%";
const resultInitial = smartParseRoster(withInitial);
console.log('Input:', withInitial);
console.log('Output:', JSON.stringify(resultInitial));
console.log('Expected: ["Jeremy B", "Michael J Fox"]');
console.log('✓ Pass:', JSON.stringify(resultInitial) === JSON.stringify(["Jeremy B", "Michael J Fox"]));
console.log('');

console.log('========================================');
console.log('ALL MANUAL VERIFICATIONS COMPLETE');
console.log('========================================');
