/**
 * Player Lookup Service - Main service for Player Lookup Sidebar
 * @fileoverview Provides searchCustomers and getPlayerLookupProfile functions
 */

// ============================================================================
// SEARCH CUSTOMERS
// ============================================================================

/**
 * Searches for customers matching a query string
 * @param {string} query - Search query (name, email, partial)
 * @return {CustomerSearchResult[]} Array of matching customers
 */
function searchCustomers(query) {
  if (!query || query.trim().length < 2) {
    return [];
  }

  const normalizedQuery = query.trim().toLowerCase();
  const results = [];
  const seenNames = new Set();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Search Key_Tracker
  try {
    const keySheet = ss.getSheetByName('Key_Tracker');
    if (keySheet && keySheet.getLastRow() > 1) {
      const keyData = keySheet.getDataRange().getValues();
      const headers = keyData[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);

      if (nameCol !== -1) {
        for (let i = 1; i < keyData.length; i++) {
          const name = String(keyData[i][nameCol] || '').trim();
          if (!name || seenNames.has(name.toLowerCase())) continue;

          const score = computeMatchScore_(name, normalizedQuery);
          if (score > 0) {
            seenNames.add(name.toLowerCase());
            const keyCount = sumKeyCount_(keyData[i], headers);
            results.push({
              name: name,
              preferredName: name,
              email: '',
              source: 'Key_Tracker',
              hasPreorders: false,
              bp: 0,
              keyCount: keyCount,
              score: score
            });
          }
        }
      }
    }
  } catch (e) {
    Logger.log('Error searching Key_Tracker: ' + e.message);
  }

  // Search BP_Total
  try {
    const bpSheet = ss.getSheetByName('BP_Total');
    if (bpSheet && bpSheet.getLastRow() > 1) {
      const bpData = bpSheet.getDataRange().getValues();
      const headers = bpData[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
      const bpCol = findColumnIndex_(headers, ['BP_Current', 'BP', 'Bonus_Points']);

      if (nameCol !== -1) {
        for (let i = 1; i < bpData.length; i++) {
          const name = String(bpData[i][nameCol] || '').trim();
          if (!name) continue;

          const nameLower = name.toLowerCase();
          const score = computeMatchScore_(name, normalizedQuery);

          if (score > 0) {
            const bp = bpCol !== -1 ? coerceNumber(bpData[i][bpCol], 0) : 0;

            if (seenNames.has(nameLower)) {
              // Update existing result with BP info
              const existing = results.find(r => r.name.toLowerCase() === nameLower);
              if (existing) {
                existing.bp = bp;
                if (existing.source !== 'BP_Total') {
                  existing.source += ', BP_Total';
                }
              }
            } else {
              seenNames.add(nameLower);
              results.push({
                name: name,
                preferredName: name,
                email: '',
                source: 'BP_Total',
                hasPreorders: false,
                bp: bp,
                keyCount: 0,
                score: score
              });
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log('Error searching BP_Total: ' + e.message);
  }

  // Search Attendance_Missions
  try {
    const attendSheet = ss.getSheetByName('Attendance_Missions');
    if (attendSheet && attendSheet.getLastRow() > 1) {
      const attendData = attendSheet.getDataRange().getValues();
      const headers = attendData[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);

      if (nameCol !== -1) {
        for (let i = 1; i < attendData.length; i++) {
          const name = String(attendData[i][nameCol] || '').trim();
          if (!name) continue;

          const nameLower = name.toLowerCase();
          const score = computeMatchScore_(name, normalizedQuery);

          if (score > 0 && !seenNames.has(nameLower)) {
            seenNames.add(nameLower);
            results.push({
              name: name,
              preferredName: name,
              email: '',
              source: 'Attendance',
              hasPreorders: false,
              bp: 0,
              keyCount: 0,
              score: score
            });
          }
        }
      }
    }
  } catch (e) {
    Logger.log('Error searching Attendance_Missions: ' + e.message);
  }

  // Search PreferredNames sheet if it exists
  try {
    const prefSheet = ss.getSheetByName('PreferredNames');
    if (prefSheet && prefSheet.getLastRow() > 1) {
      const prefData = prefSheet.getDataRange().getValues();
      const headers = prefData[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
      const emailCol = findColumnIndex_(headers, ['Email', 'email']);
      const aliasCol = findColumnIndex_(headers, ['Alt_Names', 'Aliases', 'AltNames']);

      if (nameCol !== -1) {
        for (let i = 1; i < prefData.length; i++) {
          const name = String(prefData[i][nameCol] || '').trim();
          if (!name) continue;

          const nameLower = name.toLowerCase();
          const email = emailCol !== -1 ? String(prefData[i][emailCol] || '') : '';
          const aliases = aliasCol !== -1 ? String(prefData[i][aliasCol] || '') : '';

          // Score based on name, email, or aliases
          let score = computeMatchScore_(name, normalizedQuery);
          if (email && email.toLowerCase().includes(normalizedQuery)) {
            score = Math.max(score, 80);
          }
          if (aliases) {
            const aliasList = aliases.split(',').map(a => a.trim());
            for (const alias of aliasList) {
              const aliasScore = computeMatchScore_(alias, normalizedQuery);
              score = Math.max(score, aliasScore * 0.9); // Slight penalty for alias match
            }
          }

          if (score > 0) {
            if (seenNames.has(nameLower)) {
              // Update existing with PreferredNames source (higher priority)
              const existing = results.find(r => r.name.toLowerCase() === nameLower);
              if (existing) {
                existing.source = 'PreferredNames';
                existing.email = email || existing.email;
                existing.score = Math.max(existing.score, score + 10); // Bonus for PreferredNames
              }
            } else {
              seenNames.add(nameLower);
              results.push({
                name: name,
                preferredName: name,
                email: email,
                source: 'PreferredNames',
                hasPreorders: false,
                bp: 0,
                keyCount: 0,
                score: score + 10 // Bonus for PreferredNames
              });
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log('Error searching PreferredNames: ' + e.message);
  }

  // Sort by score descending, then by name
  results.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    return a.name.localeCompare(b.name);
  });

  // Limit to top 20 results
  return results.slice(0, 20);
}

/**
 * Computes a match score between a name and query
 * @param {string} name - Name to match
 * @param {string} query - Normalized query (lowercase)
 * @return {number} Score (0 = no match, 100 = exact match)
 * @private
 */
function computeMatchScore_(name, query) {
  const nameLower = name.toLowerCase();

  // Exact match
  if (nameLower === query) return 100;

  // Starts with query
  if (nameLower.startsWith(query)) return 90;

  // Contains query as a word
  const words = nameLower.split(/\s+/);
  for (const word of words) {
    if (word === query) return 85;
    if (word.startsWith(query)) return 75;
  }

  // Contains query anywhere
  if (nameLower.includes(query)) return 60;

  // Query contains name (partial name typed)
  if (query.includes(nameLower)) return 40;

  return 0;
}

/**
 * Finds column index from possible header names
 * @param {Array} headers - Header row
 * @param {string[]} possibleNames - Possible column names
 * @return {number} Column index or -1
 * @private
 */
function findColumnIndex_(headers, possibleNames) {
  for (const name of possibleNames) {
    const idx = headers.indexOf(name);
    if (idx !== -1) return idx;
  }
  return -1;
}

/**
 * Sums key count from a row
 * @param {Array} row - Data row
 * @param {Array} headers - Header row
 * @return {number} Total key count
 * @private
 */
function sumKeyCount_(row, headers) {
  const colors = ['Red', 'Blue', 'Green', 'Yellow', 'Purple'];
  let total = 0;

  for (const color of colors) {
    const col = headers.indexOf(color);
    if (col !== -1) {
      total += coerceNumber(row[col], 0);
    }
  }

  return total;
}

// ============================================================================
// GET PLAYER LOOKUP PROFILE
// ============================================================================

/**
 * Main entrypoint for Player Lookup sidebar.
 * @param {string} query - What staff typed (name, email, partial).
 * @return {PlayerLookupProfile}
 */
function getPlayerLookupProfile(query) {
  try {
    // Validate input
    if (!query || query.trim().length < 2) {
      return createErrorResponse_(query || '', 'Please type at least 2 characters.');
    }

    const trimmedQuery = query.trim();

    // Search for matching players
    const matches = searchCustomers(trimmedQuery);

    if (matches.length === 0) {
      return {
        status: 'NOT_FOUND',
        query: trimmedQuery,
        message: 'No player found matching "' + trimmedQuery + '".'
      };
    }

    // Resolve canonical name
    const canonicalName = resolveCanonicalName_(matches, trimmedQuery);

    // Build full profile
    return buildPlayerProfile_(canonicalName, trimmedQuery);

  } catch (e) {
    Logger.log('Error in getPlayerLookupProfile: ' + e.message);
    return createErrorResponse_(query || '', 'An error occurred: ' + e.message);
  }
}

/**
 * Creates an error response
 * @param {string} query - Original query
 * @param {string} message - Error message
 * @return {PlayerLookupProfile}
 * @private
 */
function createErrorResponse_(query, message) {
  return {
    status: 'ERROR',
    query: query,
    message: message,
    meta: {
      lastUpdated: dateISO(),
      source: 'v7.9.7 PlayerLookup',
      errors: [message]
    }
  };
}

/**
 * Resolves the best canonical name from matches
 * @param {CustomerSearchResult[]} matches - Search results
 * @param {string} query - Original query
 * @return {string} Canonical name
 * @private
 */
function resolveCanonicalName_(matches, query) {
  const normalizedQuery = query.toLowerCase();

  // Exact match (case-insensitive)
  for (const match of matches) {
    if (match.name.toLowerCase() === normalizedQuery) {
      return match.preferredName || match.name;
    }
  }

  // Prefer PreferredNames source
  const prefNameMatch = matches.find(m => m.source.includes('PreferredNames'));
  if (prefNameMatch) {
    return prefNameMatch.preferredName || prefNameMatch.name;
  }

  // Starts with query
  const startsWithMatch = matches.find(m => m.name.toLowerCase().startsWith(normalizedQuery));
  if (startsWithMatch) {
    return startsWithMatch.preferredName || startsWithMatch.name;
  }

  // Return highest scoring match
  return matches[0].preferredName || matches[0].name;
}

/**
 * Builds a complete player profile
 * @param {string} canonicalName - Resolved player name
 * @param {string} query - Original query
 * @return {PlayerLookupProfile}
 * @private
 */
function buildPlayerProfile_(canonicalName, query) {
  const errors = [];

  // Get identity info
  const identity = getIdentityInfo_(canonicalName, errors);

  // Get BP info
  const bonusPoints = getBonusPointsInfo_(canonicalName, errors);

  // Get keys info
  const keys = getKeysInfo_(canonicalName, errors);

  // Get store credit info
  const storeCredit = getStoreCreditInfo_(canonicalName, errors);

  // Get preorders info
  const preorders = getPreordersInfo_(canonicalName, errors);

  // Get attendance info
  const attendance = getAttendanceInfo_(canonicalName, errors);

  // Get missions info
  const missions = getMissionsInfo_(canonicalName, errors);

  // Get flags info
  const flags = getFlagsInfo_(canonicalName, errors);

  // Build summary line
  const summary = buildSummaryLine_(identity, bonusPoints, keys, storeCredit, attendance);

  return {
    status: 'OK',
    query: query,
    playerId: canonicalName,
    preferredName: canonicalName,
    displayName: identity.displayName || canonicalName,
    aliases: identity.aliases || [],

    contact: identity.contact,
    membership: identity.membership,
    summary: summary,
    bonusPoints: bonusPoints,
    keys: keys,
    storeCredit: storeCredit,
    preorders: preorders,
    attendance: attendance,
    missions: missions,
    leagues: [],
    flags: flags,

    meta: {
      lastUpdated: dateISO(),
      source: 'v7.9.7 PlayerLookup',
      errors: errors.length > 0 ? errors : undefined
    }
  };
}

/**
 * Gets identity and contact info
 * @param {string} name - Player name
 * @param {string[]} errors - Error array to append to
 * @return {Object} Identity info
 * @private
 */
function getIdentityInfo_(name, errors) {
  const result = {
    displayName: name,
    aliases: [],
    contact: {
      email: '',
      phone: '',
      discord: ''
    },
    membership: {
      isMember: false,
      tier: null,
      joinedAt: null,
      expiresAt: null
    }
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Try PreferredNames sheet
    const prefSheet = ss.getSheetByName('PreferredNames');
    if (prefSheet && prefSheet.getLastRow() > 1) {
      const data = prefSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
      const emailCol = findColumnIndex_(headers, ['Email', 'email']);
      const phoneCol = findColumnIndex_(headers, ['Phone', 'phone']);
      const discordCol = findColumnIndex_(headers, ['Discord', 'discord']);
      const aliasCol = findColumnIndex_(headers, ['Alt_Names', 'Aliases', 'AltNames']);
      const displayCol = findColumnIndex_(headers, ['DisplayName', 'Display_Name']);

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === name.toLowerCase()) {
            if (emailCol !== -1) result.contact.email = String(data[i][emailCol] || '');
            if (phoneCol !== -1) result.contact.phone = String(data[i][phoneCol] || '');
            if (discordCol !== -1) result.contact.discord = String(data[i][discordCol] || '');
            if (displayCol !== -1 && data[i][displayCol]) {
              result.displayName = String(data[i][displayCol]);
            }
            if (aliasCol !== -1 && data[i][aliasCol]) {
              result.aliases = String(data[i][aliasCol]).split(',').map(a => a.trim()).filter(a => a);
            }
            break;
          }
        }
      }
    }

    // Try Members sheet for membership info
    const membersSheet = ss.getSheetByName('Members');
    if (membersSheet && membersSheet.getLastRow() > 1) {
      const data = membersSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
      const tierCol = findColumnIndex_(headers, ['Tier', 'tier', 'Level']);
      const joinedCol = findColumnIndex_(headers, ['JoinedAt', 'Joined', 'Start_Date']);
      const expiresCol = findColumnIndex_(headers, ['ExpiresAt', 'Expires', 'End_Date']);

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === name.toLowerCase()) {
            result.membership.isMember = true;
            if (tierCol !== -1) result.membership.tier = String(data[i][tierCol] || '');
            if (joinedCol !== -1 && data[i][joinedCol]) {
              result.membership.joinedAt = formatDateSafe_(data[i][joinedCol]);
            }
            if (expiresCol !== -1 && data[i][expiresCol]) {
              result.membership.expiresAt = formatDateSafe_(data[i][expiresCol]);
            }
            break;
          }
        }
      }
    }
  } catch (e) {
    errors.push('Identity lookup error: ' + e.message);
    Logger.log('Identity lookup error: ' + e.message);
  }

  return result;
}

/**
 * Gets bonus points info with full breakdown
 * Uses alias mapping and case-insensitive lookups for robustness
 * @param {string} name - Player name
 * @param {string[]} errors - Error array
 * @return {BonusPointsInfo}
 * @private
 */
function getBonusPointsInfo_(name, errors) {
  const result = {
    current: 0,
    lifetimeEarned: 0,
    lifetimeSpent: 0,
    pending: 0,
    lastEarnedFrom: 'Unknown',
    lastUpdated: '',
    // Breakdown components
    breakdown: {
      historical: 0,    // Total ever earned
      redeemed: 0,      // Total spent/redeemed
      current: 0,       // Available balance
      attendancePoints: 0,
      flagPoints: 0,
      dicePoints: 0
    },
    // Debug info for troubleshooting
    _sources: {
      bpTotal: 'not found',
      redeemedBP: 'not found',
      attendance: 'not found',
      flag: 'not found',
      dice: 'not found'
    }
  };

  const nameLower = String(name).toLowerCase();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ─────────────────────────────────────────────────────────────────────
    // Query BP_Total for main BP values + breakdown columns
    // Real headers: PreferredName, BP_Current, Attendance Mission Points,
    //               Flag Mission Points, Dice Roll Points, LastUpdated,
    //               BP_Historical, BP_Redeemed
    // ─────────────────────────────────────────────────────────────────────
    const bpSheet = getSheetByAliasesCI_(ss, ['BP_Total']);
    if (bpSheet && bpSheet.getLastRow() > 1) {
      result._sources.bpTotal = bpSheet.getName();
      const data = bpSheet.getDataRange().getValues();
      const headers = data[0];

      // Name column
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name', 'preferred_name_id']);
      // BP columns
      const currentCol = findColumnIndex_(headers, ['BP_Current', 'Current_BP', 'Capped_BP', 'BP']);
      const histCol = findColumnIndex_(headers, ['BP_Historical', 'Historical_BP', 'Historical', 'Lifetime_BP']);
      const redeemedCol = findColumnIndex_(headers, ['BP_Redeemed', 'Redeemed_BP', 'Total_Redeemed']);
      const updatedCol = findColumnIndex_(headers, ['LastUpdated', 'Last_Updated', 'LastUpdate']);
      // Breakdown columns (EXACT names from real BP_Total sheet)
      const attPointsCol = findColumnIndex_(headers, ['Attendance Mission Points', 'Attendance_Mission_Points']);
      const flagPointsCol = findColumnIndex_(headers, ['Flag Mission Points', 'Flag_Mission_Points']);
      const dicePointsCol = findColumnIndex_(headers, ['Dice Roll Points', 'Dice_Points']);

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === nameLower) {
            // Main BP values
            result.current = currentCol !== -1 ? coerceNumber(data[i][currentCol], 0) : 0;
            result.breakdown.current = result.current;

            if (histCol !== -1) {
              result.lifetimeEarned = coerceNumber(data[i][histCol], 0);
              result.breakdown.historical = result.lifetimeEarned;
            }

            if (redeemedCol !== -1) {
              result.lifetimeSpent = coerceNumber(data[i][redeemedCol], 0);
              result.breakdown.redeemed = result.lifetimeSpent;
            }

            // Breakdown from BP_Total (preferred source)
            if (attPointsCol !== -1) {
              result.breakdown.attendancePoints = coerceNumber(data[i][attPointsCol], 0);
            }
            if (flagPointsCol !== -1) {
              result.breakdown.flagPoints = coerceNumber(data[i][flagPointsCol], 0);
            }
            if (dicePointsCol !== -1) {
              result.breakdown.dicePoints = coerceNumber(data[i][dicePointsCol], 0);
            }

            if (updatedCol !== -1 && data[i][updatedCol]) {
              result.lastUpdated = formatDateSafe_(data[i][updatedCol]);
            }
            break;
          }
        }
      }
    }

    // ─────────────────────────────────────────────────────────────────────
    // Query Redeemed_BP sheet for authoritative redemption history
    // Real headers: PreferredName, Total_Redeemed, Item_Redeemed, Notes,
    //               BP_Current, BP_Historical, LastUpdated
    // ─────────────────────────────────────────────────────────────────────
    const redeemedSheet = getSheetByAliasesCI_(ss, ['Redeemed_BP']);
    if (redeemedSheet && redeemedSheet.getLastRow() > 1) {
      result._sources.redeemedBP = redeemedSheet.getName();
      const data = redeemedSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
      const totalRedeemedCol = findColumnIndex_(headers, ['Total_Redeemed', 'TotalRedeemed', 'Redeemed']);

      if (nameCol !== -1 && totalRedeemedCol !== -1) {
        // Sum all redemptions for this player (could be multiple rows)
        let totalRedeemed = 0;
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === nameLower) {
            totalRedeemed += coerceNumber(data[i][totalRedeemedCol], 0);
          }
        }
        // Use Redeemed_BP as authoritative source if found
        if (totalRedeemed > 0 || result.breakdown.redeemed === 0) {
          result.lifetimeSpent = totalRedeemed;
          result.breakdown.redeemed = totalRedeemed;
        }
      }
    }

    // ─────────────────────────────────────────────────────────────────────
    // Fallback: Query source sheets if BP_Total breakdown is empty
    // ─────────────────────────────────────────────────────────────────────

    // Attendance_Missions (column "Points")
    if (result.breakdown.attendancePoints === 0) {
      const attendSheet = getSheetByAliasesCI_(ss, ['Attendance_Missions', 'Attendance_Points']);
      if (attendSheet && attendSheet.getLastRow() > 1) {
        result._sources.attendance = attendSheet.getName();
        const data = attendSheet.getDataRange().getValues();
        const headers = data[0];
        const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
        const pointsCol = findColumnIndex_(headers, ['Points', 'Attendance_Points', 'AttendancePoints', 'BP_Earned']);

        if (nameCol !== -1 && pointsCol !== -1) {
          for (let i = 1; i < data.length; i++) {
            if (String(data[i][nameCol]).toLowerCase() === nameLower) {
              result.breakdown.attendancePoints = coerceNumber(data[i][pointsCol], 0);
              break;
            }
          }
        }
      }
    }

    // Flag_Missions (column "Flag Mission Points")
    if (result.breakdown.flagPoints === 0) {
      const flagSheet = getSheetByAliasesCI_(ss, ['Flag_Missions', 'Flag_Points']);
      if (flagSheet && flagSheet.getLastRow() > 1) {
        result._sources.flag = flagSheet.getName();
        const data = flagSheet.getDataRange().getValues();
        const headers = data[0];
        const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
        const pointsCol = findColumnIndex_(headers, ['Flag Mission Points', 'Flag_Mission_Points', 'FlagPoints']);

        if (nameCol !== -1 && pointsCol !== -1) {
          for (let i = 1; i < data.length; i++) {
            if (String(data[i][nameCol]).toLowerCase() === nameLower) {
              result.breakdown.flagPoints = coerceNumber(data[i][pointsCol], 0);
              break;
            }
          }
        }
      }
    }

    // Dice Roll Points (column "Dice Roll Points")
    if (result.breakdown.dicePoints === 0) {
      const diceSheet = getSheetByAliasesCI_(ss, ['Dice Roll Points', 'Dice_Points']);
      if (diceSheet && diceSheet.getLastRow() > 1) {
        result._sources.dice = diceSheet.getName();
        const data = diceSheet.getDataRange().getValues();
        const headers = data[0];
        const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
        const pointsCol = findColumnIndex_(headers, ['Dice Roll Points', 'Dice_Points', 'DicePoints', 'Points']);

        if (nameCol !== -1 && pointsCol !== -1) {
          for (let i = 1; i < data.length; i++) {
            if (String(data[i][nameCol]).toLowerCase() === nameLower) {
              result.breakdown.dicePoints = coerceNumber(data[i][pointsCol], 0);
              break;
            }
          }
        }
      }
    }

    // If historical is 0 but we have breakdown values, compute it
    if (result.breakdown.historical === 0) {
      result.breakdown.historical = result.breakdown.attendancePoints +
                                    result.breakdown.flagPoints +
                                    result.breakdown.dicePoints;
      result.lifetimeEarned = result.breakdown.historical;
    }

  } catch (e) {
    errors.push('BP lookup error: ' + e.message);
    Logger.log('BP lookup error: ' + e.message);
  }

  return result;
}

/**
 * Gets keys info
 * @param {string} name - Player name
 * @param {string[]} errors - Error array
 * @return {KeyInfo}
 * @private
 */
function getKeysInfo_(name, errors) {
  const result = {
    Red: 0,
    Blue: 0,
    Green: 0,
    Yellow: 0,
    Purple: 0,
    RainbowEligible: 0,
    lifetimeUnlocked: 0,
    lastUnlockDate: ''
  };

  try {
    // Use existing getPlayerKeys function
    const keys = getPlayerKeys(name);
    if (keys) {
      result.Red = coerceNumber(keys.Red, 0);
      result.Blue = coerceNumber(keys.Blue, 0);
      result.Green = coerceNumber(keys.Green, 0);
      result.Yellow = coerceNumber(keys.Yellow, 0);
      result.Purple = coerceNumber(keys.Purple, 0);
      result.RainbowEligible = coerceNumber(keys.RainbowEligible, 0);

      if (keys.LastUpdated) {
        result.lastUnlockDate = formatDateSafe_(keys.LastUpdated);
      }
    }
  } catch (e) {
    errors.push('Keys lookup error: ' + e.message);
    Logger.log('Keys lookup error: ' + e.message);
  }

  return result;
}

/**
 * Gets store credit info
 * @param {string} name - Player name
 * @param {string[]} errors - Error array
 * @return {StoreCreditInfo}
 * @private
 */
function getStoreCreditInfo_(name, errors) {
  const result = {
    balance: 0,
    currency: 'USD',
    lastUpdated: '',
    lastTxSummary: '',
    transactionCount: 0,
    _source: 'not found',
    _nameColumn: 'not found'
  };

  const nameLower = String(name).toLowerCase();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ─────────────────────────────────────────────────────────────────────
    // Query Store_Credit_Ledger (ledger with transactions)
    // Real headers: Timestamp, preferred_name_id (OR PreferredName), InOut,
    //               Amount, Reason, Category, TenderType, Description,
    //               POSRefType, POSRefId, RunningBalance, RowId
    // ─────────────────────────────────────────────────────────────────────
    const ledgerSheet = getSheetByAliasesCI_(ss, ['Store_Credit_Ledger', 'Store_Credit']);
    if (ledgerSheet && ledgerSheet.getLastRow() > 1) {
      result._source = ledgerSheet.getName();
      const data = ledgerSheet.getDataRange().getValues();
      const headers = data[0];

      // Support both preferred_name_id and PreferredName (transition support)
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'preferred_name_id', 'Preferred_Name', 'Player', 'Name']);
      const amountCol = findColumnIndex_(headers, ['Amount', 'Balance', 'Credit']);
      const inOutCol = findColumnIndex_(headers, ['InOut', 'Type', 'Direction']);
      const runningBalanceCol = findColumnIndex_(headers, ['RunningBalance', 'Running_Balance', 'Balance']);
      const timestampCol = findColumnIndex_(headers, ['Timestamp', 'Date', 'Created_At']);
      const reasonCol = findColumnIndex_(headers, ['Reason', 'Description', 'Notes']);

      result._nameColumn = nameCol !== -1 ? headers[nameCol] : 'not found';

      if (nameCol !== -1) {
        // Find all transactions for this player and get latest running balance
        let latestTimestamp = null;
        let latestBalance = 0;
        let latestReason = '';

        for (let i = 1; i < data.length; i++) {
          const rowName = String(data[i][nameCol]).toLowerCase();
          if (rowName === nameLower) {
            result.transactionCount++;

            // Get running balance if available, otherwise compute from InOut + Amount
            if (runningBalanceCol !== -1) {
              const timestamp = timestampCol !== -1 ? data[i][timestampCol] : null;
              if (!latestTimestamp || (timestamp && new Date(timestamp) > new Date(latestTimestamp))) {
                latestTimestamp = timestamp;
                latestBalance = coerceNumber(data[i][runningBalanceCol], 0);
                latestReason = reasonCol !== -1 ? String(data[i][reasonCol] || '') : '';
              }
            } else if (amountCol !== -1 && inOutCol !== -1) {
              // Compute balance from InOut and Amount
              const inOut = String(data[i][inOutCol]).toUpperCase();
              const amount = coerceNumber(data[i][amountCol], 0);
              if (inOut === 'IN' || inOut === 'ADD' || inOut === 'CREDIT') {
                result.balance += amount;
              } else if (inOut === 'OUT' || inOut === 'SUBTRACT' || inOut === 'DEBIT' || inOut === 'SPEND') {
                result.balance -= amount;
              }
            }
          }
        }

        // Use running balance approach if we found it
        if (runningBalanceCol !== -1 && result.transactionCount > 0) {
          result.balance = latestBalance;
          result.lastTxSummary = latestReason;
          if (latestTimestamp) {
            result.lastUpdated = formatDateSafe_(latestTimestamp);
          }
        }
      } else {
        // Name column not found - report schema issue
        result._source = ledgerSheet.getName() + ' (schema mismatch - name column not found)';
      }
    }

    // ─────────────────────────────────────────────────────────────────────
    // Fallback: Try simple Store_Credit sheet (single row per player)
    // ─────────────────────────────────────────────────────────────────────
    if (result._source === 'not found' || result._nameColumn === 'not found') {
      const scSheet = getSheetByAliasesCI_(ss, ['Store_Credit']);
      if (scSheet && scSheet.getLastRow() > 1 && scSheet.getName() !== result._source) {
        result._source = scSheet.getName() + ' (fallback)';
        const data = scSheet.getDataRange().getValues();
        const headers = data[0];
        const nameCol = findColumnIndex_(headers, ['PreferredName', 'preferred_name_id', 'Preferred_Name', 'Name']);
        const balanceCol = findColumnIndex_(headers, ['Balance', 'Credit', 'Amount']);
        const updatedCol = findColumnIndex_(headers, ['LastUpdated', 'Last_Updated']);

        if (nameCol !== -1 && balanceCol !== -1) {
          for (let i = 1; i < data.length; i++) {
            if (String(data[i][nameCol]).toLowerCase() === nameLower) {
              result.balance = coerceNumber(data[i][balanceCol], 0);
              if (updatedCol !== -1 && data[i][updatedCol]) {
                result.lastUpdated = formatDateSafe_(data[i][updatedCol]);
              }
              break;
            }
          }
        }
      }
    }
  } catch (e) {
    errors.push('Store credit lookup error: ' + e.message);
    Logger.log('Store credit lookup error: ' + e.message);
  }

  return result;
}

/**
 * Gets preorders info from Preorders_Sold sheet
 * @param {string} name - Player name
 * @param {string[]} errors - Error array
 * @return {PreordersInfo}
 * @private
 */
function getPreordersInfo_(name, errors) {
  const result = {
    active: [],
    historyCount: 0,
    totalCount: 0,
    openCount: 0,
    totalBalanceDue: 0,
    _source: 'not found',
    _nameColumn: 'not found',
    _matchedBy: 'none'
  };

  const nameLower = String(name).toLowerCase().trim();
  if (!nameLower) return result;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Use canonical preorders sheet lookup (case-insensitive with aliases)
    const preorderSheet = getPreordersSheet_(ss);
    if (!preorderSheet) {
      result._source = 'Preorders sheet not found';
      return result;
    }

    result._source = preorderSheet.getName();

    if (preorderSheet.getLastRow() <= 1) {
      result._source += ' (empty)';
      return result;
    }

    const data = preorderSheet.getDataRange().getValues();
    const headers = data[0];

    // Use canonical column resolver with synonyms
    const cols = resolvePreordersCols_(headers);
    result._nameColumn = cols.nameCol !== -1 ? headers[cols.nameCol] : 'not found';

    // If no name column found, try Customer_Name fallback
    let primaryNameCol = cols.nameCol;
    let fallbackNameCol = cols.customerNameCol;

    if (primaryNameCol === -1 && fallbackNameCol !== -1) {
      primaryNameCol = fallbackNameCol;
      result._nameColumn = headers[fallbackNameCol] + ' (fallback)';
    }

    if (primaryNameCol === -1) {
      result._source += ' (no name column)';
      return result;
    }

    // Scan all rows for matching preorders
    const matchingRows = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Check primary name column (case-insensitive + trim)
      const primaryName = String(row[primaryNameCol] || '').toLowerCase().trim();
      let isMatch = (primaryName === nameLower);

      // If no match on primary, try fallback Customer_Name if different column
      if (!isMatch && fallbackNameCol !== -1 && fallbackNameCol !== primaryNameCol) {
        const fallbackName = String(row[fallbackNameCol] || '').toLowerCase().trim();
        isMatch = (fallbackName === nameLower);
        if (isMatch && result._matchedBy === 'none') {
          result._matchedBy = 'Customer_Name';
        }
      }

      if (!isMatch) continue;

      if (result._matchedBy === 'none') {
        result._matchedBy = 'PreferredName';
      }

      result.totalCount++;

      // Get status and picked up values
      const statusValue = cols.statusCol !== -1 ? row[cols.statusCol] : '';
      const pickedUpValue = cols.pickedUpCol !== -1 ? row[cols.pickedUpCol] : '';

      // Use canonical open/closed detection
      const isOpen = isPreorderOpen_(pickedUpValue, statusValue);
      const balanceDue = cols.balanceDueCol !== -1 ? coerceNumber(row[cols.balanceDueCol], 0) : 0;

      if (isOpen) {
        result.openCount++;
        result.historyCount++;

        result.active.push({
          rowIndex: i + 1,
          preorderId: cols.preorderIdCol !== -1 ? String(row[cols.preorderIdCol] || '') : '',
          setName: cols.setNameCol !== -1 ? String(row[cols.setNameCol] || '') : '',
          itemName: cols.itemNameCol !== -1 ? String(row[cols.itemNameCol] || '') : '',
          itemCode: cols.itemCodeCol !== -1 ? String(row[cols.itemCodeCol] || '') : '',
          qty: cols.qtyCol !== -1 ? coerceNumber(row[cols.qtyCol], 1) : 1,
          unitPrice: cols.unitPriceCol !== -1 ? coerceNumber(row[cols.unitPriceCol], 0) : 0,
          totalDue: cols.totalDueCol !== -1 ? coerceNumber(row[cols.totalDueCol], 0) : 0,
          depositPaid: cols.depositCol !== -1 ? coerceNumber(row[cols.depositCol], 0) : 0,
          balanceDue: balanceDue,
          targetPayoffDate: cols.targetPayoffDateCol !== -1 ? formatDateSafe_(row[cols.targetPayoffDateCol]) : '',
          status: String(statusValue || 'Active'),
          pickedUp: isPreorderPickedUp_(pickedUpValue),
          notes: cols.notesCol !== -1 ? String(row[cols.notesCol] || '') : ''
        });

        result.totalBalanceDue += balanceDue;
      } else {
        // Still count in history even if closed
        result.historyCount++;
      }
    }

    // Limit active list to top 5 for display (sorted by balance due desc)
    if (result.active.length > 5) {
      result.active.sort((a, b) => b.balanceDue - a.balanceDue);
      result.active = result.active.slice(0, 5);
    }

  } catch (e) {
    errors.push('Preorders lookup error: ' + e.message);
    Logger.log('Preorders lookup error: ' + e.message);
  }

  return result;
}

/**
 * Gets attendance info
 * @param {string} name - Player name
 * @param {string[]} errors - Error array
 * @return {AttendanceInfo}
 * @private
 */
function getAttendanceInfo_(name, errors) {
  const result = {
    totalEvents: 0,
    eventsLast90Days: 0,
    lastEventDate: '',
    recentEvents: [],
    primaryFormat: ''
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Search event tabs for player attendance
    const eventTabs = listEventTabs();
    const now = new Date();
    const ninetyDaysAgo = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);
    const recentEvents = [];

    for (const tabName of eventTabs.slice(-50)) { // Check last 50 events max
      try {
        const sheet = ss.getSheetByName(tabName);
        if (!sheet || sheet.getLastRow() <= 1) continue;

        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);

        if (nameCol === -1) continue;

        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === name.toLowerCase()) {
            result.totalEvents++;

            // Parse event date from tab name (MM-DD-YYYY)
            const dateMatch = tabName.match(/^(\d{2})-(\d{2})-(\d{4})/);
            if (dateMatch) {
              const eventDate = new Date(dateMatch[3], dateMatch[1] - 1, dateMatch[2]);

              if (eventDate >= ninetyDaysAgo) {
                result.eventsLast90Days++;
              }

              if (!result.lastEventDate || eventDate > new Date(result.lastEventDate)) {
                result.lastEventDate = formatDateSafe_(eventDate);
              }

              // Get event type from metadata
              const props = getEventProps(sheet);
              const eventType = props.event_type || 'Event';

              recentEvents.push({
                date: tabName,
                type: eventType
              });
            }
            break;
          }
        }
      } catch (e) {
        // Skip problematic tabs
      }
    }

    // Get last 5 recent events
    result.recentEvents = recentEvents
      .slice(-5)
      .reverse()
      .map(e => e.date + ' (' + e.type + ')');

    // Try Attendance_Missions for additional info
    const attendSheet = ss.getSheetByName('Attendance_Missions');
    if (attendSheet && attendSheet.getLastRow() > 1) {
      const data = attendSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === name.toLowerCase()) {
            // If we have total from missions sheet, use it if higher
            const cTotalCol = headers.indexOf('C_Total');
            if (cTotalCol !== -1) {
              const missionTotal = coerceNumber(data[i][cTotalCol], 0);
              result.totalEvents = Math.max(result.totalEvents, missionTotal);
            }
            break;
          }
        }
      }
    }
  } catch (e) {
    errors.push('Attendance lookup error: ' + e.message);
    Logger.log('Attendance lookup error: ' + e.message);
  }

  return result;
}

/**
 * Gets missions info with individual flag mission status
 * @param {string} name - Player name
 * @param {string[]} errors - Error array
 * @return {MissionsInfo}
 * @private
 */
function getMissionsInfo_(name, errors) {
  // Flag mission definitions with point values
  const FLAG_MISSIONS = {
    'Cosmic_Selfie': 1,
    'Review_Writer': 2,
    'Social_Media_Star': 2,
    'App_Explorer': 1,
    'Cosmic_Merchant': 3,
    'Precon_Pioneer': 2,
    'Gravitational_Pull': 5,
    'Rogue_Planet': 3,
    'Quantum_Collector': 5
  };

  const result = {
    completed: 0,
    inProgress: 0,
    badges: [],
    recentAwards: [],
    // Individual flag mission status
    flagMissions: {},
    flagMissionsCompleted: 0,
    flagMissionsTotal: Object.keys(FLAG_MISSIONS).length
  };

  // Initialize all flag missions as not completed
  for (const mission in FLAG_MISSIONS) {
    result.flagMissions[mission] = {
      completed: false,
      points: FLAG_MISSIONS[mission]
    };
  }

  const nameLower = String(name).toLowerCase();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ─────────────────────────────────────────────────────────────────────
    // Query Attendance_Missions for attendance badges
    // ─────────────────────────────────────────────────────────────────────
    const attendSheet = ss.getSheetByName('Attendance_Missions');
    if (attendSheet && attendSheet.getLastRow() > 1) {
      const data = attendSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
      const badgesCol = findColumnIndex_(headers, ['Badges', 'badges']);

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === nameLower) {
            if (badgesCol !== -1 && data[i][badgesCol]) {
              result.badges = String(data[i][badgesCol]).split(',').map(b => b.trim()).filter(b => b);
              result.completed = result.badges.length;
            }
            break;
          }
        }
      }
    }

    // ─────────────────────────────────────────────────────────────────────
    // Query Flag_Missions for individual flag mission completion
    // ─────────────────────────────────────────────────────────────────────
    const flagSheet = ss.getSheetByName('Flag_Missions');
    if (flagSheet && flagSheet.getLastRow() > 1) {
      const data = flagSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === nameLower) {
            // Check each flag mission column
            for (const mission in FLAG_MISSIONS) {
              const missionCol = headers.indexOf(mission);
              if (missionCol !== -1) {
                const value = data[i][missionCol];
                // Mission is completed if checkbox is TRUE or numeric value > 0
                const isCompleted = value === true || (typeof value === 'number' && value > 0);
                result.flagMissions[mission].completed = isCompleted;
                if (isCompleted) {
                  result.flagMissionsCompleted++;
                }
              }
            }
            break;
          }
        }
      }
    }

    // Add total completed missions count
    result.completed += result.flagMissionsCompleted;

  } catch (e) {
    errors.push('Missions lookup error: ' + e.message);
    Logger.log('Missions lookup error: ' + e.message);
  }

  return result;
}

/**
 * Gets flags and notes info
 * @param {string} name - Player name
 * @param {string[]} errors - Error array
 * @return {FlagsInfo}
 * @private
 */
function getFlagsInfo_(name, errors) {
  const result = {
    isBanned: false,
    warningsCount: 0,
    tags: [],
    internalNote: ''
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Try Player_Notes sheet
    const notesSheet = ss.getSheetByName('Player_Notes');
    if (notesSheet && notesSheet.getLastRow() > 1) {
      const data = notesSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
      const bannedCol = findColumnIndex_(headers, ['IsBanned', 'Banned', 'is_banned']);
      const warningsCol = findColumnIndex_(headers, ['Warnings', 'WarningCount', 'warnings_count']);
      const tagsCol = findColumnIndex_(headers, ['Tags', 'tags']);
      const noteCol = findColumnIndex_(headers, ['Note', 'Notes', 'InternalNote']);

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === name.toLowerCase()) {
            if (bannedCol !== -1) result.isBanned = coerceBoolean(data[i][bannedCol]);
            if (warningsCol !== -1) result.warningsCount = coerceNumber(data[i][warningsCol], 0);
            if (tagsCol !== -1 && data[i][tagsCol]) {
              result.tags = String(data[i][tagsCol]).split(',').map(t => t.trim()).filter(t => t);
            }
            if (noteCol !== -1) result.internalNote = String(data[i][noteCol] || '');
            break;
          }
        }
      }
    }
  } catch (e) {
    errors.push('Flags lookup error: ' + e.message);
    Logger.log('Flags lookup error: ' + e.message);
  }

  return result;
}

/**
 * Builds a summary line from profile data
 * @param {Object} identity - Identity info
 * @param {BonusPointsInfo} bonusPoints - BP info
 * @param {KeyInfo} keys - Keys info
 * @param {StoreCreditInfo} storeCredit - Store credit info
 * @param {AttendanceInfo} attendance - Attendance info
 * @return {string} Summary line
 * @private
 */
function buildSummaryLine_(identity, bonusPoints, keys, storeCredit, attendance) {
  const parts = [];

  // Membership
  if (identity.membership.isMember) {
    const tier = identity.membership.tier ? identity.membership.tier + ' ' : '';
    parts.push(tier + 'Member');
  }

  // Events
  if (attendance.totalEvents > 0) {
    parts.push(attendance.totalEvents + ' events');
  }

  // Keys
  const totalKeys = keys.Red + keys.Blue + keys.Green + keys.Yellow + keys.Purple;
  if (totalKeys > 0 || keys.RainbowEligible > 0) {
    const keyText = keys.RainbowEligible > 0
      ? keys.RainbowEligible + ' rainbow'
      : totalKeys + ' keys';
    parts.push(keyText);
  }

  // Store credit
  if (storeCredit.balance > 0) {
    parts.push(formatCurrency(storeCredit.balance) + ' credit');
  }

  // BP
  if (bonusPoints.current > 0) {
    parts.push(bonusPoints.current + ' BP');
  }

  if (parts.length === 0) {
    return 'New player';
  }

  return parts.join(' • ');
}

/**
 * Safely formats a date value
 * @param {*} value - Date value
 * @return {string} ISO string or empty string
 * @private
 */
function formatDateSafe_(value) {
  if (!value) return '';

  try {
    if (value instanceof Date) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
    }
    if (typeof value === 'string') {
      // Already a string, try to parse and reformat
      const date = new Date(value);
      if (!isNaN(date.getTime())) {
        return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
      }
      return value;
    }
    return String(value);
  } catch (e) {
    return String(value || '');
  }
}

// ============================================================================
// MISSING WRAPPER FUNCTIONS FOR UI COMPATIBILITY
// ============================================================================

/**
 * Wrapper for UI compatibility - UI calls getPlayerProfile()
 * @param {string} name - Player name to look up
 * @return {PlayerLookupProfile} Full player profile
 */
function getPlayerProfile(name) {
  return getPlayerLookupProfile(name);
}

/**
 * Gets all player names for UI dropdown
 * @return {Array<string>} Array of player names
 */
function getPlayerNames() {
  // Delegate to PlayerProvisioning.js canonical function
  return getAllPreferredNames();
}

/**
 * Gets player BP balance
 * @param {string} name - Player name
 * @return {number} Current BP balance
 */
function getPlayerBP(name) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bpSheet = ss.getSheetByName('BP_Total');
    if (!bpSheet || bpSheet.getLastRow() <= 1) return 0;

    const data = bpSheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name', 'preferred_name_id']);
    const bpCol = findColumnIndex_(headers, ['BP_Current', 'BP', 'Bonus_Points', 'Capped_BP']);

    if (nameCol === -1) return 0;

    const nameLower = String(name).toLowerCase();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nameCol]).toLowerCase() === nameLower) {
        return bpCol !== -1 ? coerceNumber(data[i][bpCol], 0) : 0;
      }
    }
    return 0;
  } catch (e) {
    Logger.log('getPlayerBP error: ' + e.message);
    return 0;
  }
}

// ============================================================================
// DEBUG TOOLS
// ============================================================================

/**
 * Debug function: Shows detailed player lookup data sources and values
 * Run from Script Editor to troubleshoot player lookup issues
 * @param {string} name - Player name to look up (default: "Cy Diskin")
 * @return {Object} Debug info object
 */
function debug_playerLookup(name) {
  name = name || 'Cy Diskin';

  Logger.log('=== DEBUG: Player Lookup for "' + name + '" ===');
  Logger.log('');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {
    query: name,
    sheets: {},
    profile: null
  };

  // Check each relevant sheet
  const sheetChecks = [
    { aliases: ['BP_Total'], desc: 'BP Total' },
    { aliases: ['Redeemed_BP'], desc: 'Redeemed BP' },
    { aliases: ['Attendance_Missions'], desc: 'Attendance Missions' },
    { aliases: ['Flag_Missions'], desc: 'Flag Missions' },
    { aliases: ['Dice Roll Points', 'Dice_Points'], desc: 'Dice Points' },
    { aliases: ['Key_Tracker'], desc: 'Key Tracker' },
    { aliases: ['Preorders_Sold'], desc: 'Preorders' },
    { aliases: ['Store_Credit_Ledger', 'Store_Credit'], desc: 'Store Credit' }
  ];

  const nameLower = name.toLowerCase();

  for (const check of sheetChecks) {
    const sheet = getSheetByAliasesCI_(ss, check.aliases);
    if (sheet) {
      const sheetInfo = {
        found: true,
        name: sheet.getName(),
        rowCount: sheet.getLastRow(),
        headers: [],
        playerFound: false,
        playerRow: null,
        playerData: {}
      };

      if (sheet.getLastRow() > 1) {
        const data = sheet.getDataRange().getValues();
        sheetInfo.headers = data[0];

        const nameCol = findColumnIndex_(data[0], ['PreferredName', 'Preferred_Name', 'preferred_name_id', 'Name', 'Player']);
        if (nameCol !== -1) {
          for (let i = 1; i < data.length; i++) {
            if (String(data[i][nameCol]).toLowerCase() === nameLower) {
              sheetInfo.playerFound = true;
              sheetInfo.playerRow = i + 1;
              // Store all column values
              for (let j = 0; j < data[0].length; j++) {
                sheetInfo.playerData[data[0][j]] = data[i][j];
              }
              break;
            }
          }
        }
      }

      result.sheets[check.desc] = sheetInfo;
      Logger.log('--- ' + check.desc + ' ---');
      Logger.log('Sheet: ' + sheet.getName());
      Logger.log('Player found: ' + sheetInfo.playerFound + (sheetInfo.playerRow ? ' (row ' + sheetInfo.playerRow + ')' : ''));
      if (sheetInfo.playerFound) {
        Logger.log('Data: ' + JSON.stringify(sheetInfo.playerData, null, 2));
      }
      Logger.log('');
    } else {
      result.sheets[check.desc] = { found: false, aliases: check.aliases };
      Logger.log('--- ' + check.desc + ' ---');
      Logger.log('Sheet NOT FOUND (tried: ' + check.aliases.join(', ') + ')');
      Logger.log('');
    }
  }

  // Get full profile
  Logger.log('=== FULL PROFILE ===');
  result.profile = getPlayerLookupProfile(name);
  Logger.log(JSON.stringify(result.profile, null, 2));

  return result;
}

/**
 * Debug function: Shows preorder matching logic for a player
 * Uses canonical helpers to match real sheet structure
 * @param {string} name - Player name (default: "Cy Diskin")
 * @return {Object} Debug info
 */
function debug_findPreordersForName(name) {
  name = name || 'Cy Diskin';

  Logger.log('=== DEBUG: Preorders for "' + name + '" ===');
  Logger.log('');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nameLower = String(name).toLowerCase().trim();

  const result = {
    query: name,
    sheet: null,
    columns: {},
    matchesByPreferredName: 0,
    matchesByCustomerName: 0,
    matchingRows: [],
    openCount: 0,
    closedCount: 0
  };

  // Use canonical sheet lookup
  const preorderSheet = getPreordersSheet_(ss);
  if (!preorderSheet) {
    Logger.log('Preorders sheet NOT FOUND (tried aliases: Preorders_Sold, Preorders Sold, Preorders)');
    return result;
  }

  result.sheet = preorderSheet.getName();
  Logger.log('Sheet found: ' + result.sheet);
  Logger.log('Row count: ' + preorderSheet.getLastRow());

  if (preorderSheet.getLastRow() <= 1) {
    Logger.log('Sheet is empty (no data rows)');
    return result;
  }

  const data = preorderSheet.getDataRange().getValues();
  const headers = data[0];
  Logger.log('Headers: ' + headers.join(', '));
  Logger.log('');

  // Use canonical column resolver
  const cols = resolvePreordersCols_(headers);

  // Log resolved columns
  Logger.log('--- Resolved Column Indices ---');
  Logger.log('nameCol: ' + cols.nameCol + ' → ' + (cols.nameCol !== -1 ? headers[cols.nameCol] : 'NOT FOUND'));
  Logger.log('customerNameCol: ' + cols.customerNameCol + ' → ' + (cols.customerNameCol !== -1 ? headers[cols.customerNameCol] : 'NOT FOUND'));
  Logger.log('statusCol: ' + cols.statusCol + ' → ' + (cols.statusCol !== -1 ? headers[cols.statusCol] : 'NOT FOUND'));
  Logger.log('pickedUpCol: ' + cols.pickedUpCol + ' → ' + (cols.pickedUpCol !== -1 ? headers[cols.pickedUpCol] : 'NOT FOUND'));
  Logger.log('preorderIdCol: ' + cols.preorderIdCol + ' → ' + (cols.preorderIdCol !== -1 ? headers[cols.preorderIdCol] : 'NOT FOUND'));
  Logger.log('setNameCol: ' + cols.setNameCol + ' → ' + (cols.setNameCol !== -1 ? headers[cols.setNameCol] : 'NOT FOUND'));
  Logger.log('itemNameCol: ' + cols.itemNameCol + ' → ' + (cols.itemNameCol !== -1 ? headers[cols.itemNameCol] : 'NOT FOUND'));
  Logger.log('balanceDueCol: ' + cols.balanceDueCol + ' → ' + (cols.balanceDueCol !== -1 ? headers[cols.balanceDueCol] : 'NOT FOUND'));
  Logger.log('');

  result.columns = {
    name: cols.nameCol !== -1 ? headers[cols.nameCol] : 'NOT FOUND',
    customerName: cols.customerNameCol !== -1 ? headers[cols.customerNameCol] : 'NOT FOUND',
    status: cols.statusCol !== -1 ? headers[cols.statusCol] : 'NOT FOUND',
    pickedUp: cols.pickedUpCol !== -1 ? headers[cols.pickedUpCol] : 'NOT FOUND',
    preorderId: cols.preorderIdCol !== -1 ? headers[cols.preorderIdCol] : 'NOT FOUND'
  };

  Logger.log('--- Matching Rows ---');

  // Scan for matches (limit to first 10 for logging)
  let logCount = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Check PreferredName match
    const prefName = cols.nameCol !== -1 ? String(row[cols.nameCol] || '').toLowerCase().trim() : '';
    const custName = cols.customerNameCol !== -1 ? String(row[cols.customerNameCol] || '').toLowerCase().trim() : '';

    let matchedBy = null;
    if (prefName === nameLower) {
      matchedBy = 'PreferredName';
      result.matchesByPreferredName++;
    } else if (custName === nameLower) {
      matchedBy = 'Customer_Name';
      result.matchesByCustomerName++;
    }

    if (!matchedBy) continue;

    // Get values
    const preorderId = cols.preorderIdCol !== -1 ? row[cols.preorderIdCol] : '';
    const prefNameValue = cols.nameCol !== -1 ? row[cols.nameCol] : '';
    const statusValue = cols.statusCol !== -1 ? row[cols.statusCol] : '';
    const pickedUpValue = cols.pickedUpCol !== -1 ? row[cols.pickedUpCol] : '';
    const balanceDue = cols.balanceDueCol !== -1 ? coerceNumber(row[cols.balanceDueCol], 0) : 0;
    const setName = cols.setNameCol !== -1 ? row[cols.setNameCol] : '';
    const itemName = cols.itemNameCol !== -1 ? row[cols.itemNameCol] : '';

    // Use canonical open/closed detection
    const isOpen = isPreorderOpen_(pickedUpValue, statusValue);
    const normalizedStatus = normalizePreorderStatus_(statusValue);

    const rowInfo = {
      row: i + 1,
      matchedBy: matchedBy,
      preorderId: preorderId,
      preferredName: prefNameValue,
      setName: setName,
      itemName: itemName,
      status: statusValue,
      statusNormalized: normalizedStatus,
      pickedUpRaw: pickedUpValue,
      pickedUpParsed: isPreorderPickedUp_(pickedUpValue),
      isOpen: isOpen,
      balanceDue: balanceDue
    };

    result.matchingRows.push(rowInfo);

    if (isOpen) {
      result.openCount++;
    } else {
      result.closedCount++;
    }

    // Log first 10 rows
    if (logCount < 10) {
      Logger.log('Row ' + (i + 1) + ' [' + matchedBy + ']: ' + preorderId);
      Logger.log('  Name: ' + prefNameValue);
      Logger.log('  Set/Item: ' + setName + ' / ' + itemName);
      Logger.log('  Status: "' + statusValue + '" → normalized: "' + normalizedStatus + '"');
      Logger.log('  Picked_Up?: ' + JSON.stringify(pickedUpValue) + ' → parsed: ' + isPreorderPickedUp_(pickedUpValue));
      Logger.log('  Balance Due: ' + balanceDue);
      Logger.log('  → ' + (isOpen ? 'OPEN' : 'CLOSED'));
      Logger.log('');
      logCount++;
    }
  }

  Logger.log('=== SUMMARY ===');
  Logger.log('Total matches: ' + result.matchingRows.length);
  Logger.log('  By PreferredName: ' + result.matchesByPreferredName);
  Logger.log('  By Customer_Name: ' + result.matchesByCustomerName);
  Logger.log('Open: ' + result.openCount);
  Logger.log('Closed: ' + result.closedCount);

  return result;
}

/**
 * Legacy alias for debug function
 */
function debug_preorders(name) {
  return debug_findPreordersForName(name);
}

/**
 * Debug function: Shows summary of all open preorders
 * @return {Object} Summary of open preorders
 */
function debug_preordersOpenSummary() {
  Logger.log('=== DEBUG: Open Preorders Summary ===');
  Logger.log('');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const preorderSheet = getPreordersSheet_(ss);

  if (!preorderSheet) {
    Logger.log('Preorders sheet NOT FOUND');
    return { error: 'Sheet not found', totalOpen: 0 };
  }

  const data = preorderSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log('Sheet is empty');
    return { error: 'Sheet empty', totalOpen: 0 };
  }

  const headers = data[0];
  const cols = resolvePreordersCols_(headers);

  // Group open preorders by player
  const openByPlayer = {};
  let totalOpen = 0;
  let totalBalance = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const statusValue = cols.statusCol !== -1 ? row[cols.statusCol] : '';
    const pickedUpValue = cols.pickedUpCol !== -1 ? row[cols.pickedUpCol] : '';

    if (!isPreorderOpen_(pickedUpValue, statusValue)) continue;

    totalOpen++;
    const playerName = cols.nameCol !== -1 ? String(row[cols.nameCol] || 'Unknown').trim() : 'Unknown';
    const balanceDue = cols.balanceDueCol !== -1 ? coerceNumber(row[cols.balanceDueCol], 0) : 0;

    if (!openByPlayer[playerName]) {
      openByPlayer[playerName] = { count: 0, totalBalance: 0, items: [] };
    }
    openByPlayer[playerName].count++;
    openByPlayer[playerName].totalBalance += balanceDue;
    totalBalance += balanceDue;

    if (openByPlayer[playerName].items.length < 3) {
      const setName = cols.setNameCol !== -1 ? row[cols.setNameCol] : '';
      const itemName = cols.itemNameCol !== -1 ? row[cols.itemNameCol] : '';
      openByPlayer[playerName].items.push(setName + ' / ' + itemName);
    }
  }

  Logger.log('Total open preorders: ' + totalOpen);
  Logger.log('Total balance due: $' + totalBalance.toFixed(2));
  Logger.log('Players with open preorders: ' + Object.keys(openByPlayer).length);
  Logger.log('');

  // Log top 10 players by count
  const sorted = Object.entries(openByPlayer)
    .sort((a, b) => b[1].count - a[1].count)
    .slice(0, 10);

  Logger.log('--- Top 10 Players by Open Count ---');
  for (const [player, info] of sorted) {
    Logger.log(player + ': ' + info.count + ' open ($' + info.totalBalance.toFixed(2) + ')');
    info.items.forEach(item => Logger.log('  - ' + item));
  }

  return {
    sheetName: preorderSheet.getName(),
    totalOpen: totalOpen,
    totalBalance: totalBalance,
    playerCount: Object.keys(openByPlayer).length,
    byPlayer: openByPlayer
  };
}

/**
 * Debug function: Shows BP river source sheets and sample data
 * @return {Object} Debug info about BP sources
 */
function debug_bpRiverPreview() {
  Logger.log('=== DEBUG: BP River Preview ===');
  Logger.log('');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {
    sources: {},
    bpTotal: null,
    samplePlayers: []
  };

  // Check source sheets
  const sources = [
    { key: 'attendance', aliases: ['Attendance_Missions'], pointsCol: ['Points', 'Attendance_Points'] },
    { key: 'flag', aliases: ['Flag_Missions'], pointsCol: ['Flag Mission Points', 'Flag_Mission_Points'] },
    { key: 'dice', aliases: ['Dice Roll Points', 'Dice_Points'], pointsCol: ['Dice Roll Points', 'Dice_Points', 'Points'] }
  ];

  for (const src of sources) {
    const sheet = getSheetByAliasesCI_(ss, src.aliases);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const pointsIdx = findColumnIndex_(headers, src.pointsCol);

      result.sources[src.key] = {
        found: true,
        sheetName: sheet.getName(),
        rowCount: sheet.getLastRow() - 1,
        headers: headers,
        pointsColumn: pointsIdx !== -1 ? headers[pointsIdx] : 'NOT FOUND'
      };

      Logger.log('--- ' + src.key.toUpperCase() + ' ---');
      Logger.log('Sheet: ' + sheet.getName());
      Logger.log('Rows: ' + (sheet.getLastRow() - 1));
      Logger.log('Points column: ' + (pointsIdx !== -1 ? headers[pointsIdx] : 'NOT FOUND'));
    } else {
      result.sources[src.key] = { found: false, tried: src.aliases };
      Logger.log('--- ' + src.key.toUpperCase() + ' ---');
      Logger.log('NOT FOUND (tried: ' + src.aliases.join(', ') + ')');
    }
    Logger.log('');
  }

  // Check BP_Total
  const bpSheet = getSheetByAliasesCI_(ss, ['BP_Total']);
  if (bpSheet && bpSheet.getLastRow() > 1) {
    const data = bpSheet.getDataRange().getValues();
    const headers = data[0];

    result.bpTotal = {
      found: true,
      sheetName: bpSheet.getName(),
      rowCount: bpSheet.getLastRow() - 1,
      headers: headers
    };

    Logger.log('--- BP_TOTAL ---');
    Logger.log('Sheet: ' + bpSheet.getName());
    Logger.log('Rows: ' + (bpSheet.getLastRow() - 1));
    Logger.log('Headers: ' + headers.join(', '));
    Logger.log('');

    // Sample first 5 players
    const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
    const currentCol = findColumnIndex_(headers, ['BP_Current', 'Current_BP']);
    const attCol = findColumnIndex_(headers, ['Attendance Mission Points']);
    const flagCol = findColumnIndex_(headers, ['Flag Mission Points']);
    const diceCol = findColumnIndex_(headers, ['Dice Roll Points']);

    Logger.log('--- SAMPLE PLAYERS (first 5) ---');
    for (let i = 1; i < Math.min(6, data.length); i++) {
      const player = {
        name: nameCol !== -1 ? data[i][nameCol] : 'N/A',
        current: currentCol !== -1 ? data[i][currentCol] : 'N/A',
        attendance: attCol !== -1 ? data[i][attCol] : 'N/A',
        flag: flagCol !== -1 ? data[i][flagCol] : 'N/A',
        dice: diceCol !== -1 ? data[i][diceCol] : 'N/A'
      };
      result.samplePlayers.push(player);
      Logger.log(player.name + ': Current=' + player.current +
                 ' (Att=' + player.attendance + ', Flag=' + player.flag + ', Dice=' + player.dice + ')');
    }
  } else {
    result.bpTotal = { found: false };
    Logger.log('--- BP_TOTAL ---');
    Logger.log('NOT FOUND');
  }

  return result;
}