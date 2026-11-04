/**
 * =========================================================================
 * VICE VERSE 2026 - VVCE MYSURU
 * 2 TEAMS PER PROBLEM STATEMENT SELECTOR
 * Columns: Team Name | Problem Statement
 * =========================================================================
 */

const MAX_TEAMS_PER_PS = 2; // Up to 2 teams can select each PS

/**
 * 1. Setup the 2-column sheet (Run once)
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Team Name", "Problem Statement"]);
    
    // Header style
    const header = sheet.getRange(1, 1, 1, 2);
    header.setBackground("#0F172A");
    header.setFontColor("#FFFFFF");
    header.setFontWeight("bold");
    sheet.setRowHeight(1, 35);
    sheet.setColumnWidth(1, 250);
    sheet.setColumnWidth(2, 250);
  }
}

/**
 * 2. GET API: Reads the 2-column sheet and returns claims & slot counts
 */
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();
  
  if (sheet.getLastRow() === 0) {
    setupSheet();
  }

  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getClaims";

  if (action === "ping") {
    return sendJson({ status: "ok", message: "2-Teams-per-PS API is Online" });
  }

  if (action === "claim") {
    return handleClaim(e.parameter.teamName, e.parameter.psCode);
  }

  // Read all rows
  const claims = getClaimedList(sheet);
  return sendJson({ success: true, claimed: claims, maxPerPs: MAX_TEAMS_PER_PS });
}

/**
 * 3. POST API: Saves Team Name and PS Code into the 2 columns
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (err) {
        payload = e.parameter;
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    return handleClaim(payload.teamName, payload.psCode);
  } catch (err) {
    return sendJson({ success: false, message: "Server error: " + err.toString() });
  }
}

/**
 * 4. ATOMIC CLAIM: Allows up to 2 TEAMS per Problem Statement
 */
function handleClaim(teamName, psCode) {
  teamName = (teamName || "").trim();
  psCode = (psCode || "").trim().toUpperCase();

  if (!teamName) {
    return sendJson({ success: false, message: "Please enter your Team Name!" });
  }
  if (!psCode) {
    return sendJson({ success: false, message: "Please select a Problem Statement!" });
  }

  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(10000);
  if (!hasLock) {
    return sendJson({ success: false, message: "Server is busy. Please try again in 2 seconds." });
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getActiveSheet();
    
    if (sheet.getLastRow() === 0) {
      setupSheet();
    }

    const lastRow = sheet.getLastRow();
    let psCount = 0;
    const existingTeamsForThisPs = [];

    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

      for (let i = 0; i < data.length; i++) {
        const rowTeam = (data[i][0] || "").toString().trim();
        const rowPs = (data[i][1] || "").toString().trim().toUpperCase();

        // 1. Strict check: Team cannot select multiple PS
        if (rowTeam.toLowerCase() === teamName.toLowerCase()) {
          return sendJson({
            success: false,
            message: 'Team "' + teamName + '" has already selected PS [' + rowPs + ']! Only 1 PS per team.'
          });
        }

        // Count how many teams have already chosen this PS
        if (rowPs === psCode) {
          psCount++;
          existingTeamsForThisPs.push(rowTeam);
        }
      }

      // 2. Strict check: Max 2 teams per PS
      if (psCount >= MAX_TEAMS_PER_PS) {
        return sendJson({
          success: false,
          message: 'Problem Statement [' + psCode + '] is ALREADY FULL! (Chosen by ' + existingTeamsForThisPs.join(' and ') + '). Maximum ' + MAX_TEAMS_PER_PS + ' teams allowed.'
        });
      }
    }

    // Append directly: Column A = Team Name, Column B = PS Code
    sheet.appendRow([teamName, psCode]);
    SpreadsheetApp.flush();

    const slotsLeft = MAX_TEAMS_PER_PS - (psCount + 1);

    return sendJson({
      success: true,
      message: 'Successfully assigned PS [' + psCode + '] to team "' + teamName + '"! (' + (slotsLeft === 0 ? 'Now FULL' : '1 slot remaining') + ')',
      teamName: teamName,
      psCode: psCode,
      slotsLeft: slotsLeft
    });

  } catch (err) {
    return sendJson({ success: false, message: "Error saving: " + err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper to read claimed rows
 */
function getClaimedList(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const list = [];
  for (let i = 0; i < data.length; i++) {
    const team = (data[i][0] || "").toString().trim();
    const ps = (data[i][1] || "").toString().trim().toUpperCase();
    if (team && ps) {
      list.push({ teamName: team, psCode: ps });
    }
  }
  return list;
}

function sendJson(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
