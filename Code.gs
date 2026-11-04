/**
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
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  } else {
    sheet.clear();
  }

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
}

/**
 * 2. GET API: Returns the table of PS with real-time status
 */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getPS";

  if (action === "ping") {
    return sendJson({ status: "ok", message: "API is Online" });
  }

  if (action === "claim") {
    return handleClaim(e.parameter.teamName, e.parameter.psCode);
  }

  // Default: Return all rows from the sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    setupSheet();
    sheet = ss.getSheetByName(SHEET_NAME);
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const data = rows.map(r => ({
    code: r[0],
    domain: r[1],
    title: r[2],
    category: r[3],
    status: r[4] || "Available",
    teamName: r[5] || "",
    selectedAt: r[6] || ""
  }));

  return sendJson({ success: true, data: data });
}

/**
 * 3. POST API: Receives teamName and psCode and updates Google Sheet
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
    return sendJson({ success: false, message: "Error: " + err.toString() });
  }
}

/**
 * 4. ATOMIC CLAIM: Updates Google Sheet with Lock (Only 1 Team per PS)
 */
function handleClaim(teamName, psCode) {
  teamName = (teamName || "").trim();
  psCode = (psCode || "").trim().toUpperCase();

  if (!teamName) return sendJson({ success: false, message: "Please enter your Team Name!" });
  if (!psCode) return sendJson({ success: false, message: "Please select a Problem Statement!" });

  // Concurrency Lock
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(10000);
  if (!hasLock) {
    return sendJson({ success: false, message: "Server busy. Please try again in a few seconds." });
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {
      setupSheet();
      sheet = ss.getSheetByName(SHEET_NAME);
    }

    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();

    // Check if this Team Name already selected a PS
    for (let i = 0; i < rows.length; i++) {
      const existingTeam = (rows[i][5] || "").toString().trim().toLowerCase();
      if (existingTeam && existingTeam === teamName.toLowerCase()) {
        return sendJson({
          success: false,
          message: 'Team "' + teamName + '" has already selected PS ' + rows[i][0] + '! Only 1 PS per team is allowed.'
        });
      }
    }

    // Find requested PS Code
    let targetRow = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0].toString().trim().toUpperCase() === psCode) {
        // Check if already taken
        if (rows[i][4] === "Taken" || rows[i][5] !== "") {
          return sendJson({
            success: false,
            message: 'Problem Statement ' + psCode + ' is ALREADY TAKEN by Team "' + rows[i][5] + '"!'
          });
        }
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow === -1) {
      return sendJson({ success: false, message: "Problem Statement " + psCode + " not found." });
    }

    // Update the row
    const timestamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(targetRow, 5).setValue("Taken");
    sheet.getRange(targetRow, 6).setValue(teamName);
    sheet.getRange(targetRow, 7).setValue(timestamp);

    SpreadsheetApp.flush();

    return sendJson({
      success: true,
      message: 'Successfully assigned ' + psCode + ' to Team "' + teamName + '"!',
      psCode: psCode,
      teamName: teamName
    });

  } catch (err) {
    return sendJson({ success: false, message: "Error: " + err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function sendJson(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 30 Problem Statements from SIH 2025 Booklet
 */
function getProblemStatements() {
  return [
  {
    "code": "CS-01",
    "domain": "CSE / AIML",
    "title": "Smart Community Health Monitoring and Early Warning System for Water-Borne Diseases in Rural Northeast India",
    "category": "Software"
  },
  {
    "code": "CS-02",
    "domain": "CSE / AIML",
    "title": "Smart Tourist Safety Monitoring and Incident Response System using AI, Geo-Fencing and Blockchain-Based Digital ID",
    "category": "Software"
  },
  {
    "code": "CS-03",
    "domain": "CSE / AIML",
    "title": "Image-Based Breed Recognition for Cattle and Buffaloes of India",
    "category": "Software"
  },
  {
    "code": "CS-04",
    "domain": "CSE / AIML",
    "title": "Real-Time Public Transport Tracking for Small Cities",
    "category": "Software"
  },
  {
    "code": "CS-05",
    "domain": "CSE / AIML",
    "title": "Maximizing Section Throughput Using AI-Powered Precise Train Traffic Control",
    "category": "Software"
  },
  {
    "code": "CS-06",
    "domain": "CSE / AIML",
    "title": "Authenticity Validator for Academia",
    "category": "Software"
  },
  {
    "code": "CS-07",
    "domain": "CSE / AIML",
    "title": "Blockchain-Based Blue Carbon Registry and MRV System",
    "category": "Software"
  },
  {
    "code": "CS-08",
    "domain": "CSE / AIML",
    "title": "Integrated Platform for Crowdsourced Ocean Hazard Reporting and Social Media Analytics",
    "category": "Software"
  },
  {
    "code": "CS-09",
    "domain": "CSE / AIML",
    "title": "FloatChat: AI-Powered Conversational Interface for ARGO Ocean Data Discovery and Visualization",
    "category": "Software"
  },
  {
    "code": "CS-10",
    "domain": "CSE / AIML",
    "title": "Smart Traffic Management System for Urban Congestion",
    "category": "Software"
  },
  {
    "code": "CS-11",
    "domain": "CSE / AIML",
    "title": "AI-Based Rockfall Prediction and Alert System for Open-Pit Mines",
    "category": "Software"
  },
  {
    "code": "CS-12",
    "domain": "CSE / AIML",
    "title": "AI-Powered Monitoring of Crop Health, Soil Condition and Pest Risks using Multispectral/Hyperspectral Imaging and Sensor Data",
    "category": "Software"
  },
  {
    "code": "CS-13",
    "domain": "CSE / AIML",
    "title": "AI-Based Drop-Out Prediction and Counseling System",
    "category": "Software"
  },
  {
    "code": "CS-14",
    "domain": "CSE / AIML",
    "title": "Language Agnostic Chatbot",
    "category": "Software"
  },
  {
    "code": "CS-15",
    "domain": "CSE / AIML",
    "title": "Digital Mental Health and Psychological Support System for Students in Higher Education",
    "category": "Software"
  },
  {
    "code": "EC-01",
    "domain": "Electronics and Communication",
    "title": "Renewable Energy Monitoring System for Microgrids",
    "category": "Hardware"
  },
  {
    "code": "EC-02",
    "domain": "Electronics and Communication",
    "title": "AI-Based Development of Laser-Based QR Code Marking on Track Fittings on Indian Railways",
    "category": "Hardware"
  },
  {
    "code": "EC-03",
    "domain": "Electronics and Communication",
    "title": "Development of Indigenous Contactless Integrated Track Monitoring Systems (ITMS) for Track Recording on Indian Railways",
    "category": "Hardware"
  },
  {
    "code": "EC-04",
    "domain": "Electronics and Communication",
    "title": "Development of Sensor for Detection of Microplastics",
    "category": "Hardware"
  },
  {
    "code": "EC-05",
    "domain": "Electronics and Communication",
    "title": "Cost-Effective Solution for Detecting Breakage of Low Voltage AC Distribution Overhead Conductors",
    "category": "Hardware"
  },
  {
    "code": "EC-06",
    "domain": "Electronics and Communication",
    "title": "Implementation of Smart Agriculture for Efficient Cultivation in Hilly Regions",
    "category": "Hardware"
  },
  {
    "code": "EC-07",
    "domain": "Electronics and Communication",
    "title": "Improving Renewable Energy Hosting Capacity in Distribution Feeders and Power Quality during High RE Injection",
    "category": "Hardware"
  },
  {
    "code": "EC-08",
    "domain": "Electronics and Communication",
    "title": "Disaster Response Drone for Remote Areas",
    "category": "Hardware"
  },
  {
    "code": "EC-09",
    "domain": "Electronics and Communication",
    "title": "Hardware to Detect and Prevent Unauthorized Use of Electric Fences",
    "category": "Hardware"
  },
  {
    "code": "EC-10",
    "domain": "Electronics and Communication",
    "title": "Embedded Intelligent Microscopy System for Identification and Counting of Microscopic Marine Organisms",
    "category": "Hardware"
  },
  {
    "code": "ME-01",
    "domain": "Mechanical and Civil",
    "title": "Improved Onion Storage Technology for Enhancing Shelf Life of Onions",
    "category": "Hardware"
  },
  {
    "code": "ME-02",
    "domain": "Mechanical and Civil",
    "title": "Low-Cost Smart Transportation Solution for Agri Produce from Remote Farms to the Nearest Motorable Road in the NER Region",
    "category": "Hardware"
  },
  {
    "code": "ME-03",
    "domain": "Mechanical and Civil",
    "title": "Design and Implementation of Solar-Powered Dewatering in Mining Operations",
    "category": "Hardware"
  },
  {
    "code": "CV-01",
    "domain": "Mechanical and Civil",
    "title": "Application for On-Spot Assessment of Roof Top Rain Water Harvesting and Artificial Recharge Potential and Size of RTRWH and AR Structures",
    "category": "Software"
  },
  {
    "code": "CV-02",
    "domain": "Mechanical and Civil",
    "title": "Grey Water Management and Reuse, Wetland Management and Water Conservation",
    "category": "Hardware"
  },
  {
    "code": "CV-03",
    "domain": "Mechanical and Civil",
    "title": "Design of a 4-Storey Commercial Office Building using Revit Architecture and Revit Structure (BIM)",
    "category": "Software"
  },
  {
    "code": "CV-04",
    "domain": "Mechanical and Civil",
    "title": "Solution for Non-Revenue Loss in Water Supply, Improving Awareness of Water Conservation, and Treatment and Reuse of Waste Water for Domestic Purposes",
    "category": "Hardware"
  }
];
}
