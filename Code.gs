/**
 * =========================================================================
 * IDEATHON 2025 - VVCE MYSURU
 * SIMPLE 2-COLUMN GOOGLE APPS SCRIPT
 * Columns: Team Name | Problem Statement
 * =========================================================================
 */

/**
 * 1. Setup the 2-column sheet (Run once or let it auto-initialize)
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
 * 2. GET API: Reads the 2-column sheet and returns which PS are taken
 */
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();
  
  // Auto-setup header if sheet is completely empty
  if (sheet.getLastRow() === 0) {
    setupSheet();
  }

  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getClaims";

  if (action === "ping") {
    return sendJson({ status: "ok", message: "2-Column Sheet API is Online" });
  }

  if (action === "claim") {
    return handleClaim(e.parameter.teamName, e.parameter.psCode);
  }

  // Return all claims from the 2-column sheet
  const claims = getClaimedList(sheet);
  return sendJson({ success: true, claimed: claims });
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
 * 4. ATOMIC CLAIM: Appends [Team Name, PS Code] with concurrency locking
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
    
    // If there are existing records, check for duplicates
    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

      for (let i = 0; i < data.length; i++) {
        const rowTeam = (data[i][0] || "").toString().trim().toLowerCase();
        const rowPs = (data[i][1] || "").toString().trim().toUpperCase();

        // Check if team already claimed a PS
        if (rowTeam === teamName.toLowerCase()) {
          return sendJson({
            success: false,
            message: 'Team "' + teamName + '" has already selected PS [' + rowPs + ']! Only 1 PS per team.'
          });
        }

        // Check if this PS is already taken by another team
        if (rowPs === psCode) {
          return sendJson({
            success: false,
            message: 'Problem Statement [' + psCode + '] is ALREADY TAKEN by team "' + data[i][0] + '"!'
          });
        }
      }
    }

    // Append to the 2 columns: Column A = Team Name, Column B = PS Code
    sheet.appendRow([teamName, psCode]);
    SpreadsheetApp.flush();

    return sendJson({
      success: true,
      message: 'Successfully assigned PS [' + psCode + '] to team "' + teamName + '"!',
      teamName: teamName,
      psCode: psCode
    });

  } catch (err) {
    return sendJson({ success: false, message: "Error saving: " + err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper to get list of claimed rows
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
