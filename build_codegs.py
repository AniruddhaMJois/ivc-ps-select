# build_codegs.py
import json

with open("problem_statements.json") as f:
    ps_data = json.load(f)

js_array = json.dumps(ps_data, indent=2, ensure_ascii=False)

template = """/**
 * =========================================================================
 * IDEATHON 2025 - VVCE, MYSURU
 * PROBLEM STATEMENT SELECTION & REAL-TIME RESERVATION SYSTEM
 * Google Apps Script Backend (Code.gs)
 * =========================================================================
 * 
 * Instructions:
 * 1. Open your Google Sheet -> Click Extensions -> Apps Script.
 * 2. Delete everything and paste this entire file.
 * 3. Click "Save" (Ctrl+S).
 * 4. In toolbar, select "setupSheet" from the function dropdown and click "Run" once.
 * 5. Deploy -> Manage deployments -> Edit (pencil) -> New version -> Deploy.
 */

const SHEET_NAME_PS = "Problem_Statements";
const SHEET_NAME_LOG = "Selection_Audit_Log";

/**
 * 1-CLICK SETUP FUNCTION
 * Run this function once from toolbar (Run -> setupSheet).
 * It creates the sheets, formats headers, and populates all 30 statements.
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Problem_Statements Sheet
  let sheet = ss.getSheetByName(SHEET_NAME_PS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_PS);
  } else {
    sheet.clear();
  }
  
  const headers = [
    "PS Code",
    "SIH PS No",
    "Domain",
    "Title",
    "Category",
    "Theme",
    "Organization / Ministry",
    "Status",
    "Claimed By Team",
    "Claimed Timestamp",
    "Description"
  ];
  
  sheet.appendRow(headers);
  
  // Header Styling
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#1E293B"); // Slate 800
  headerRange.setFontColor("#FFFFFF");
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment("center");
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);
  
  // Populate all 30 Problem Statements
  const psList = getProblemStatementsData();
  const rows = psList.map(ps => [
    ps.code,
    ps.sihNo,
    ps.domain,
    ps.title,
    ps.category,
    ps.theme,
    ps.org,
    "Available", // Status
    "",          // Claimed By Team
    "",          // Claimed Timestamp
    ps.description
  ]);
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    
    // Formatting
    sheet.getRange(2, 1, rows.length, 1).setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(2, 2, rows.length, 1).setHorizontalAlignment("center");
    sheet.getRange(2, 5, rows.length, 1).setHorizontalAlignment("center");
    sheet.getRange(2, 8, rows.length, 1).setHorizontalAlignment("center").setFontWeight("bold");
    
    // Conditional Formatting for Status column (H)
    applyStatusColors(sheet, rows.length);
  }
  
  sheet.autoResizeColumns(1, 10);
  sheet.setColumnWidth(4, 280); // Title column width
  sheet.setColumnWidth(11, 350); // Description column width
  
  // 2. Setup Selection_Audit_Log Sheet
  let logSheet = ss.getSheetByName(SHEET_NAME_LOG);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_NAME_LOG);
  } else {
    logSheet.clear();
  }
  
  const logHeaders = [
    "Timestamp",
    "Action",
    "Result",
    "PS Code",
    "Team Name",
    "Remarks"
  ];
  
  logSheet.appendRow(logHeaders);
  const logHeaderRange = logSheet.getRange(1, 1, 1, logHeaders.length);
  logHeaderRange.setBackground("#0F172A");
  logHeaderRange.setFontColor("#FFFFFF");
  logHeaderRange.setFontWeight("bold");
  logSheet.setFrozenRows(1);
  logSheet.autoResizeColumns(1, logHeaders.length);
  
  try {
    SpreadsheetApp.getUi().alert(
      "Setup Completed Successfully!",
      "Loaded 30 Problem Statements across CSE/AIML (15), ECE (10), and Mech/Civil (5).\\n\\n" +
      "Deploy this script as a Web App: Deploy > New deployment > Web App > Who has access: Anyone.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log("Setup completed in headless web context.");
  }
}

/**
 * Green for Available, Red for Claimed
 */
function applyStatusColors(sheet, numRows) {
  const range = sheet.getRange(2, 8, numRows, 1);
  const ruleAvailable = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Available")
    .setBackground("#DCFCE7")
    .setFontColor("#166534")
    .setRanges([range])
    .build();
    
  const ruleClaimed = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Claimed")
    .setBackground("#FEE2E2")
    .setFontColor("#991B1B")
    .setRanges([range])
    .build();
    
  sheet.setConditionalFormatRules([ruleAvailable, ruleClaimed]);
}

/**
 * doGet: Handles live status fetching and GET claim fallback
 */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getPS";
  
  if (action === "ping") {
    return createJsonResponse({ status: "ok", message: "VVCE Ideathon PS Selection API is Online" });
  }
  
  if (action === "claim") {
    return handleClaim({
      teamName: e.parameter.teamName,
      psCode: e.parameter.psCode
    });
  }
  
  // Default: Return all Problem Statements with live availability
  return fetchAllProblemStatements();
}

/**
 * doPost: Handles atomic claiming from web app
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
    
    const action = payload.action || "claim";
    if (action === "claim") {
      return handleClaim(payload);
    }
    
    return createJsonResponse({ success: false, message: "Unknown action: " + action });
  } catch (err) {
    return createJsonResponse({ success: false, message: "Server error: " + err.toString() });
  }
}

/**
 * CONCURRENCY-SAFE ATOMIC CLAIM HANDLER
 * Uses LockService.getScriptLock() to strictly guarantee only one team gets each PS.
 */
function handleClaim(data) {
  const teamName = (data.teamName || "").trim();
  const psCode = (data.psCode || "").trim().toUpperCase();
  
  if (!teamName) {
    return createJsonResponse({ success: false, message: "Please enter your Team Name." });
  }
  if (!psCode) {
    return createJsonResponse({ success: false, message: "Please select a Problem Statement." });
  }
  
  // Acquire Script Lock (wait up to 10 seconds for concurrent requests)
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(10000);
  
  if (!hasLock) {
    logAudit("CLAIM_ATTEMPT", "FAILED_LOCK", psCode, teamName, "Server busy under high load");
    return createJsonResponse({
      success: false,
      message: "Server is currently processing concurrent team requests. Please try again in 2 seconds."
    });
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME_PS);
    if (!sheet || sheet.getLastRow() < 2) {
      setupSheet();
      sheet = ss.getSheetByName(SHEET_NAME_PS);
    }
    
    const lastRow = sheet.getLastRow();
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 11);
    const values = dataRange.getValues();
    
    // 1. Check if this team has ALREADY claimed any PS (Strict 1 PS per team rule)
    for (let i = 0; i < values.length; i++) {
      const existingTeam = (values[i][8] || "").toString().trim().toLowerCase();
      if (existingTeam && existingTeam === teamName.toLowerCase()) {
        const alreadyClaimedCode = values[i][0];
        logAudit("CLAIM_ATTEMPT", "DUPLICATE_TEAM", psCode, teamName, "Team already holds " + alreadyClaimedCode);
        return createJsonResponse({
          success: false,
          code: "ALREADY_HAS_PS",
          message: 'Team "' + teamName + '" has already claimed problem statement [' + alreadyClaimedCode + ']. Only one selection is allowed per team.'
        });
      }
    }
    
    // 2. Find requested PS row
    let targetRowIndex = -1;
    let targetPS = null;
    
    for (let i = 0; i < values.length; i++) {
      if (values[i][0].toString().trim().toUpperCase() === psCode) {
        targetRowIndex = i + 2; // 1-based index including header
        targetPS = values[i];
        break;
      }
    }
    
    if (targetRowIndex === -1) {
      return createJsonResponse({ success: false, message: 'Problem Statement [' + psCode + '] not found.' });
    }
    
    const currentStatus = (targetPS[7] || "").toString().trim();
    const currentHolder = (targetPS[8] || "").toString().trim();
    
    // 3. Check if already claimed by someone else
    if (currentStatus.toLowerCase() === "claimed" || currentHolder !== "") {
      logAudit("CLAIM_ATTEMPT", "ALREADY_CLAIMED", psCode, teamName, "Already claimed by " + currentHolder);
      return createJsonResponse({
        success: false,
        code: "PS_TAKEN",
        message: 'Sorry! Problem Statement [' + psCode + '] was just claimed by team "' + currentHolder + '". Please choose another one.'
      });
    }
    
    // 4. ATOMIC CLAIM: Update sheet row
    const timestampStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
    
    sheet.getRange(targetRowIndex, 8).setValue("Claimed");
    sheet.getRange(targetRowIndex, 9).setValue(teamName);
    sheet.getRange(targetRowIndex, 10).setValue(timestampStr);
    
    // Flush changes immediately
    SpreadsheetApp.flush();
    
    logAudit("CLAIM_SUCCESS", "SUCCESS", psCode, teamName, "Locked successfully at " + timestampStr);
    
    return createJsonResponse({
      success: true,
      message: 'Congratulations! Problem Statement [' + psCode + '] is now officially locked for team "' + teamName + '".',
      data: {
        psCode: psCode,
        sihNo: targetPS[1],
        title: targetPS[3],
        domain: targetPS[2],
        category: targetPS[4],
        teamName: teamName,
        timestamp: timestampStr
      }
    });
    
  } catch (err) {
    return createJsonResponse({ success: false, message: "Error processing claim: " + err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Fetch all problem statements with real-time availability stats
 * (Auto-initializes the sheet if not yet configured)
 */
function fetchAllProblemStatements() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME_PS);
    
    // Auto-setup if sheet is missing or empty
    if (!sheet || sheet.getLastRow() < 2) {
      setupSheet();
      sheet = ss.getSheetByName(SHEET_NAME_PS);
    }
    
    const lastRow = sheet.getLastRow();
    const values = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    let availableCount = 0;
    let claimedCount = 0;
    
    const list = values.map(row => {
      const isClaimed = (row[7] || "").toString().trim().toLowerCase() === "claimed" || !!(row[8] || "").toString().trim();
      if (isClaimed) {
        claimedCount++;
      } else {
        availableCount++;
      }
      
      return {
        code: row[0],
        sihNo: row[1],
        domain: row[2],
        title: row[3],
        category: row[4],
        theme: row[5],
        org: row[6],
        status: isClaimed ? "Claimed" : "Available",
        claimedBy: isClaimed ? row[8] : "",
        claimedAt: isClaimed ? row[9] : "",
        description: row[10]
      };
    });
    
    return createJsonResponse({
      success: true,
      timestamp: Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss"),
      stats: {
        total: list.length,
        available: availableCount,
        claimed: claimedCount
      },
      data: list
    });
    
  } catch (err) {
    return createJsonResponse({ success: false, message: "Error fetching data: " + err.toString() });
  }
}

/**
 * Audit Log Helper
 */
function logAudit(action, result, psCode, teamName, remarks) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(SHEET_NAME_LOG);
    if (!logSheet) return;
    
    const ts = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
    logSheet.appendRow([ts, action, result, psCode, teamName, remarks]);
  } catch (e) {
    Logger.log("Failed to write audit log: " + e.toString());
  }
}

/**
 * Admin: Reset all problem statements back to Available
 */
function resetAllSelections() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_PS);
  if (!sheet) return;
  
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    for (let r = 2; r <= lastRow; r++) {
      sheet.getRange(r, 8).setValue("Available");
      sheet.getRange(r, 9).setValue("");
      sheet.getRange(r, 10).setValue("");
    }
  }
  logAudit("ADMIN_RESET", "SUCCESS", "ALL", "ADMIN", "Reset all claims to Available");
}

/**
 * JSON Response builder
 */
function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Problem Statements Database from SIH 2025 Booklet
 */
function getProblemStatementsData() {
  return """ + js_array + """;
}
"""

with open("Code.gs", "w", encoding="utf-8") as f:
    f.write(template)

print("Generated Code.gs successfully!")
