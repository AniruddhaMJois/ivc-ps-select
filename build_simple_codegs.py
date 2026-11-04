# build_simple_codegs.py
import json

with open("problem_statements.json") as f:
    ps_data = json.load(f)

# Keep code, domain, title, category, description
clean_list = [
    {
        "code": p["code"],
        "domain": p["domain"],
        "title": p["title"],
        "category": p["category"]
    }
    for p in ps_data
]

js_array = json.dumps(clean_list, indent=2, ensure_ascii=False)

code = f"""/**
 * =========================================================================
 * IDEATHON 2025 - VVCE MYSURU (SIMPLE PS SELECTOR)
 * Google Apps Script Backend (Code.gs)
 * =========================================================================
 */

const SHEET_NAME = "Problem_Statements";

/**
 * 1. 1-CLICK SETUP: Creates the table with all 30 Problem Statements
 * Run this function once from toolbar (Run -> setupSheet).
 */
function setupSheet() {{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {{
    sheet = ss.insertSheet(SHEET_NAME);
  }} else {{
    sheet.clear();
  }}

  // Table Headers
  const headers = ["PS Code", "Domain", "Title", "Category", "Status", "Team Name", "Selected At"];
  sheet.appendRow(headers);

  // Style Header Row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#0F172A"); // Dark Slate
  headerRange.setFontColor("#FFFFFF");
  headerRange.setFontWeight("bold");
  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);

  // Insert all 30 Problem Statements
  const psList = getProblemStatements();
  const rows = psList.map(p => [
    p.code,
    p.domain,
    p.title,
    p.category,
    "Available", // Status
    "",          // Team Name
    ""           // Timestamp
  ]);

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // Color formatting: Green for Available, Red for Taken
  const statusRange = sheet.getRange(2, 5, rows.length, 1);
  const ruleAvail = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Available")
    .setBackground("#DCFCE7")
    .setFontColor("#15803D")
    .setRanges([statusRange])
    .build();
  const ruleTaken = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Taken")
    .setBackground("#FEE2E2")
    .setFontColor("#B91C1C")
    .setRanges([statusRange])
    .build();
  sheet.setConditionalFormatRules([ruleAvail, ruleTaken]);

  sheet.autoResizeColumns(1, headers.length);
  sheet.setColumnWidth(3, 400); // Title column
}}

/**
 * 2. GET API: Returns the table of PS with real-time status
 */
function doGet(e) {{
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getPS";

  if (action === "ping") {{
    return sendJson({{ status: "ok", message: "API is Online" }});
  }}

  if (action === "claim") {{
    return handleClaim(e.parameter.teamName, e.parameter.psCode);
  }}

  // Default: Return all rows from the sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {{
    setupSheet();
    sheet = ss.getSheetByName(SHEET_NAME);
  }}

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const data = rows.map(r => ({{
    code: r[0],
    domain: r[1],
    title: r[2],
    category: r[3],
    status: r[4] || "Available",
    teamName: r[5] || "",
    selectedAt: r[6] || ""
  }}));

  return sendJson({{ success: true, data: data }});
}}

/**
 * 3. POST API: Receives teamName and psCode and updates Google Sheet
 */
function doPost(e) {{
  try {{
    let payload = {{}};
    if (e && e.postData && e.postData.contents) {{
      try {{
        payload = JSON.parse(e.postData.contents);
      }} catch (err) {{
        payload = e.parameter;
      }}
    }} else if (e && e.parameter) {{
      payload = e.parameter;
    }}

    return handleClaim(payload.teamName, payload.psCode);
  }} catch (err) {{
    return sendJson({{ success: false, message: "Error: " + err.toString() }});
  }}
}}

/**
 * 4. ATOMIC CLAIM: Updates Google Sheet with Lock (Only 1 Team per PS)
 */
function handleClaim(teamName, psCode) {{
  teamName = (teamName || "").trim();
  psCode = (psCode || "").trim().toUpperCase();

  if (!teamName) return sendJson({{ success: false, message: "Please enter your Team Name!" }});
  if (!psCode) return sendJson({{ success: false, message: "Please select a Problem Statement!" }});

  // Concurrency Lock
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(10000);
  if (!hasLock) {{
    return sendJson({{ success: false, message: "Server busy. Please try again in a few seconds." }});
  }}

  try {{
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {{
      setupSheet();
      sheet = ss.getSheetByName(SHEET_NAME);
    }}

    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();

    // Check if this Team Name already selected a PS
    for (let i = 0; i < rows.length; i++) {{
      const existingTeam = (rows[i][5] || "").toString().trim().toLowerCase();
      if (existingTeam && existingTeam === teamName.toLowerCase()) {{
        return sendJson({{
          success: false,
          message: 'Team "' + teamName + '" has already selected PS ' + rows[i][0] + '! Only 1 PS per team is allowed.'
        }});
      }}
    }}

    // Find requested PS Code
    let targetRow = -1;
    for (let i = 0; i < rows.length; i++) {{
      if (rows[i][0].toString().trim().toUpperCase() === psCode) {{
        // Check if already taken
        if (rows[i][4] === "Taken" || rows[i][5] !== "") {{
          return sendJson({{
            success: false,
            message: 'Problem Statement ' + psCode + ' is ALREADY TAKEN by Team "' + rows[i][5] + '"!'
          }});
        }}
        targetRow = i + 2;
        break;
      }}
    }}

    if (targetRow === -1) {{
      return sendJson({{ success: false, message: "Problem Statement " + psCode + " not found." }});
    }}

    // Update the row
    const timestamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(targetRow, 5).setValue("Taken");
    sheet.getRange(targetRow, 6).setValue(teamName);
    sheet.getRange(targetRow, 7).setValue(timestamp);

    SpreadsheetApp.flush();

    return sendJson({{
      success: true,
      message: 'Successfully assigned ' + psCode + ' to Team "' + teamName + '"!',
      psCode: psCode,
      teamName: teamName
    }});

  }} catch (err) {{
    return sendJson({{ success: false, message: "Error: " + err.toString() }});
  }} finally {{
    lock.releaseLock();
  }}
}}

function sendJson(obj) {{
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}}

/**
 * 30 Problem Statements from SIH 2025 Booklet
 */
function getProblemStatements() {{
  return {js_array};
}}
"""

with open("Code.gs", "w", encoding="utf-8") as f:
    f.write(code)

print("Generated clean and simple Code.gs!")
