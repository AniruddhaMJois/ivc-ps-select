/**
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
      "Loaded 30 Problem Statements across CSE/AIML (15), ECE (10), and Mech/Civil (5).\n\n" +
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
  return [
  {
    "id": "CS-01",
    "code": "CS-01",
    "sihNo": "SIH25001",
    "domain": "CSE / AIML",
    "category": "Software",
    "title": "Smart Community Health Monitoring and Early Warning System for Water-Borne Diseases in Rural Northeast India",
    "org": "Ministry of Development of North Eastern Region",
    "theme": "MedTech / BioTech / HealthTech",
    "description": "Water-borne diseases such as diarrhoea, typhoid and hepatitis A repeatedly affect remote villages of the North East, where hilly terrain, heavy monsoons and scattered settlements make it difficult for health departments to notice an outbreak until many people are already sick. Reports from health workers arrive late, water quality is tested rarely and the two sets of information are never combined. This case study asks teams to design a smart monitoring and early warning system that brings together community health reports, symptom trends, water quality readings and rainfall data to detect an emerging outbreak and alert health officials and villagers in time. Teams must consider low connectivity, local languages, data privacy and the limited training of frontline workers. A strong solution explains the data sources, the prediction or anomaly detection approach, the alert workflow, the village-level interface and how success would be measured."
  },
  {
    "id": "CS-02",
    "code": "CS-02",
    "sihNo": "SIH25002",
    "domain": "CSE / AIML",
    "category": "Software",
    "title": "Smart Tourist Safety Monitoring and Incident Response System using AI, Geo-Fencing and Blockchain-Based Digital ID",
    "org": "Ministry of Development of North Eastern Region",
    "theme": "Travel & Tourism",
    "description": "Tourism is a major opportunity for the North East, but tourists visiting remote, forested and border areas face risks such as getting lost, sudden weather changes, accidents and, at times, crime, while emergency response is slow because the exact location and identity of the visitor are unknown. This case study asks teams to build a tourist safety platform that issues a secure digital identity to visitors, uses geo-fencing to warn them when they enter risky zones, detects unusual behaviour such as prolonged inactivity or route deviation, and triggers a fast incident response to the nearest police or rescue unit. Teams must balance safety with privacy, work in areas of weak network coverage and keep the experience simple for foreign and domestic tourists alike. The proposal should cover the identity design, the AI-based alerting, the response workflow and the control-room dashboard for authorities."
  },
  {
    "id": "CS-03",
    "code": "CS-03",
    "sihNo": "SIH25004",
    "domain": "CSE / AIML",
    "category": "Software",
    "title": "Image-Based Breed Recognition for Cattle and Buffaloes of India",
    "org": "Ministry of Fisheries, Animal Husbandry & Dairying",
    "theme": "Agriculture, FoodTech & Rural Development",
    "description": "India has dozens of recognised cattle and buffalo breeds, and correct breed identification is essential for registration, breeding programmes, milk-yield planning and government schemes, yet it is still done by trained experts through visual inspection, which is slow and inconsistent in the field. This case study asks teams to develop an image-based system that recognises the breed of a cattle or buffalo from photographs taken on an ordinary smartphone. Teams must deal with very similar-looking breeds, variation in age and body posture, cluttered backgrounds, different lighting and a limited availability of well-labelled images. The system should be usable by field workers with little training and in areas with poor connectivity. A good solution describes data collection, the model architecture, handling of uncertain predictions, the mobile workflow and how accuracy will be validated with veterinary experts."
  },
  {
    "id": "CS-04",
    "code": "CS-04",
    "sihNo": "SIH25013",
    "domain": "CSE / AIML",
    "category": "Software",
    "title": "Real-Time Public Transport Tracking for Small Cities",
    "org": "Government of Punjab",
    "theme": "Transportation & Logistics",
    "description": "In small and medium cities, passengers wait at stops without knowing when the next bus will arrive, which discourages public transport use and pushes people toward private vehicles, while operators have little visibility into delays, bunching and route performance. Large cities use expensive intelligent transport systems that small city bus services cannot afford. This case study asks teams to design a low-cost real-time tracking solution for city buses that shows live location and expected arrival time to passengers and gives operators useful operational data. Teams must decide how to collect location data affordably, for example through low-cost GPS units or driver smartphones, handle poor network coverage and unreliable schedules, and make the passenger interface simple for people with basic phones. A strong proposal covers the architecture, arrival time estimation, the passenger and operator interfaces, privacy of drivers and a rollout plan for a pilot city."
  },
  {
    "id": "CS-05",
    "code": "CS-05",
    "sihNo": "SIH25022",
    "domain": "CSE / AIML",
    "category": "Software",
    "title": "Maximizing Section Throughput Using AI-Powered Precise Train Traffic Control",
    "org": "Ministry of Railways",
    "theme": "Transportation & Logistics",
    "description": "A railway section can carry only a limited number of trains, and every delay, unplanned halt or poor crossing decision reduces the number of trains that actually pass through it. Section controllers currently make many of these decisions using experience and fixed rules, which becomes difficult as traffic grows and as passenger, freight and priority trains compete for the same tracks. This case study asks teams to build an AI-powered decision support system that recommends precise train movements, such as precedence, crossing and holding decisions, to maximise throughput while respecting safety rules and punctuality of priority trains. Teams must model block sections, signalling constraints and train characteristics, and handle disruptions in real time. A good solution explains the optimisation or reinforcement learning approach, a simulation to compare against current practice, the controller interface and how recommendations are explained and validated."
  },
  {
    "id": "CS-06",
    "code": "CS-06",
    "sihNo": "SIH25029",
    "domain": "CSE / AIML",
    "category": "Software",
    "title": "Authenticity Validator for Academia",
    "org": "Government of Jharkhand",
    "theme": "Smart Education",
    "description": "Fake degree certificates, forged mark sheets and manipulated transcripts are a growing problem for universities, employers and government recruiters, and verifying a document today usually means writing to the issuing institution and waiting for weeks. Honest students suffer delays while forged documents sometimes pass unnoticed. This case study asks teams to design an authenticity validator that lets institutions, employers and agencies check the genuineness of academic documents quickly and reliably. Teams may combine document forensics, secure QR codes or digital signatures, institution registries and tamper-evident storage. Challenges include supporting scanned and legacy documents with very different formats, keeping the process simple for small colleges, protecting student data and handling disputes. A strong proposal includes the validation workflow, the technology used to detect tampering, the onboarding approach for institutions and the reporting available to verifiers."
  },
  {
    "id": "CS-07",
    "code": "CS-07",
    "sihNo": "SIH25038",
    "domain": "CSE / AIML",
    "category": "Software",
    "title": "Blockchain-Based Blue Carbon Registry and MRV System",
    "org": "Ministry of Earth Sciences (MoES)",
    "theme": "Clean & Green Technology",
    "description": "Coastal ecosystems such as mangroves, seagrass meadows and salt marshes store large amounts of carbon, and restoring them can generate carbon credits, but projects need trustworthy ways to measure, report and verify the carbon actually captured. Without transparent records, credits can be double counted or overstated, and small community projects are left out. This case study asks teams to design a blockchain-based registry and monitoring, reporting and verification (MRV) system for blue carbon projects. Teams must decide how to combine satellite imagery, field surveys and sensor data to estimate carbon stock, how to record project ownership and credit issuance immutably, and how to let communities and auditors take part. A strong solution describes the data pipeline, the estimation approach, the registry design, safeguards against fraud and a simple interface for project owners."
  },
  {
    "id": "CS-08",
    "code": "CS-08",
    "sihNo": "SIH25039",
    "domain": "CSE / AIML",
    "category": "Software",
    "title": "Integrated Platform for Crowdsourced Ocean Hazard Reporting and Social Media Analytics",
    "org": "Ministry of Earth Sciences (MoES)",
    "theme": "Disaster Management",
    "description": "Coastal communities are often the first to see signs of ocean hazards such as unusually high waves, flooding, swell surges, oil spills or abnormal sea behaviour, but their observations rarely reach forecasting agencies quickly or in a usable form. At the same time, social media carries many relevant posts that are buried in noise. This case study asks teams to build a platform where citizens, fishers and volunteers can report ocean hazards with photos, location and short descriptions, and where social media and news streams are analysed to detect emerging hazard signals. Teams must handle false reports, multilingual and informal text, geo-tagging and the need to present verified information to authorities in time. A good proposal covers the reporting app, the text and image analysis pipeline, validation methods, the situational dashboard and how alerts flow back to coastal communities."
  },
  {
    "id": "CS-09",
    "code": "CS-09",
    "sihNo": "SIH25040",
    "domain": "CSE / AIML",
    "category": "Software",
    "title": "FloatChat: AI-Powered Conversational Interface for ARGO Ocean Data Discovery and Visualization",
    "org": "Ministry of Earth Sciences (MoES)",
    "theme": "Miscellaneous",
    "description": "The ARGO programme deploys autonomous floats that measure temperature, salinity and other properties of the ocean, producing a vast and valuable dataset, but the data is stored in specialised scientific formats that are difficult for students, policymakers and even many researchers to explore. This case study asks teams to build a conversational interface that lets users ask questions in plain language, such as how surface temperature has changed in a region, and receive the right data, charts and explanations. Teams must design a pipeline that converts questions into reliable data queries, handles large scientific datasets efficiently, produces correct visualisations and avoids confidently giving wrong answers. A strong solution describes the data ingestion and indexing, the language model and retrieval approach, safeguards for accuracy, the visualisation layer and the way answers cite the underlying data."
  },
  {
    "id": "CS-10",
    "code": "CS-10",
    "sihNo": "SIH25050",
    "domain": "CSE / AIML",
    "category": "Software",
    "title": "Smart Traffic Management System for Urban Congestion",
    "org": "Government of Odisha",
    "theme": "Transportation & Logistics",
    "description": "Rapid urban growth has left many Indian cities with congested junctions, long waiting times and rising pollution, while traffic signals still run on fixed timers that cannot respond to the real flow of vehicles. Indian traffic adds further complexity through mixed vehicle types, informal parking and unpredictable behaviour. This case study asks teams to design a smart traffic management system that senses traffic conditions, adapts signals and informs drivers and authorities in real time. Teams may use cameras, existing sensors or mobile data, and must handle poor weather and lighting, coordinate multiple junctions and give priority to emergency vehicles. A strong proposal explains the sensing approach, the control algorithm, a simulation demonstrating reduced delay, the way emergency vehicles are detected and the dashboard used by traffic police for monitoring and planning."
  },
  {
    "id": "CS-11",
    "code": "CS-11",
    "sihNo": "SIH25071",
    "domain": "CSE / AIML",
    "category": "Software",
    "title": "AI-Based Rockfall Prediction and Alert System for Open-Pit Mines",
    "org": "Ministry of Mines",
    "theme": "Disaster Management",
    "description": "Rockfalls and slope failures in open-pit mines are a leading cause of accidents and equipment damage, and they are difficult to predict because they depend on geology, rainfall, blasting, slope geometry and the gradual movement of the rock mass. Current monitoring relies on periodic inspection and a limited number of instruments. This case study asks teams to develop an AI-based system that predicts the likelihood of rockfall in different parts of a mine and sends timely alerts to supervisors and workers. Teams may combine slope stability data, sensor readings, drone or radar imagery, weather and blasting records. Challenges include limited failure data, the high cost of false alarms versus missed alarms, and the need to deliver warnings in noisy mine environments. A strong solution describes the data sources, the prediction approach, the risk map, the alert system and validation using past incidents."
  },
  {
    "id": "CS-12",
    "code": "CS-12",
    "sihNo": "SIH25099",
    "domain": "CSE / AIML",
    "category": "Software",
    "title": "AI-Powered Monitoring of Crop Health, Soil Condition and Pest Risks using Multispectral/Hyperspectral Imaging and Sensor Data",
    "org": "MathWorks India Pvt. Ltd.",
    "theme": "Agriculture, FoodTech & Rural Development",
    "description": "Precision agriculture depends on knowing, field by field and week by week, how healthy the crop is, what the soil needs and where pests are likely to appear, but conventional scouting is slow, subjective and impossible at the scale of large farms. Multispectral and hyperspectral imaging from drones or satellites, together with field sensors, can reveal stress long before it is visible to the eye. This case study asks teams to build an AI-powered system that processes such imagery and sensor data to monitor crop health, soil condition and pest risk and turns the results into clear recommendations for farmers. Challenges include large data volumes, calibration between sensors, limited labelled data and presenting results in a form that non-experts can act upon. A good proposal explains the data pipeline, the models, the validation method and the farmer-facing interface."
  },
  {
    "id": "CS-13",
    "code": "CS-13",
    "sihNo": "SIH25102",
    "domain": "CSE / AIML",
    "category": "Software",
    "title": "AI-Based Drop-Out Prediction and Counseling System",
    "org": "Government of Rajasthan",
    "theme": "Smart Automation",
    "description": "A significant number of students leave school or college before completing their course, usually after a period of falling attendance, weak performance, financial stress or personal difficulty that nobody connected in time. Teachers and counsellors are stretched thin and lack tools to spot which students need help first. This case study asks teams to build an AI-based system that uses attendance, assessment results and other available indicators to predict which students are at risk of dropping out and to support timely counselling. The system must be fair and explainable, avoid labelling students unfairly, respect privacy and fit into the daily workflow of teachers. A strong solution describes the risk model, the explanation shown to counsellors, the intervention and follow-up workflow, safeguards against bias and a plan to measure whether interventions actually reduce dropouts."
  },
  {
    "id": "CS-14",
    "code": "CS-14",
    "sihNo": "SIH25104",
    "domain": "CSE / AIML",
    "category": "Software",
    "title": "Language Agnostic Chatbot",
    "org": "Government of Rajasthan",
    "theme": "Smart Education",
    "description": "Students, parents and citizens interact with institutions and departments in many languages, but most chatbots work well in only one or two, leaving large groups without help for admissions, schemes, schedules and common queries. This case study asks teams to build a chatbot that understands and answers in whatever language the user writes or speaks, including mixed-language text commonly used on phones. Teams must address language detection, translation or multilingual understanding, accurate retrieval of institutional information and graceful handling of questions the bot cannot answer. Voice support, low-bandwidth operation and easy updating of the knowledge base by non-technical staff are important practical requirements. A strong proposal describes the architecture, the language handling approach, how answers are grounded in official content, evaluation across several languages and the escalation path to a human."
  },
  {
    "id": "CS-15",
    "code": "CS-15",
    "sihNo": "SIH25092",
    "domain": "CSE / AIML",
    "category": "Software",
    "title": "Digital Mental Health and Psychological Support System for Students in Higher Education",
    "org": "Government of Jammu and Kashmir",
    "theme": "MedTech / BioTech / HealthTech",
    "description": "Students in higher education face academic pressure, isolation, financial worries and uncertainty about careers, yet many hesitate to seek help because of stigma, limited counselling staff and lack of privacy. Problems often become serious before anyone notices. This case study asks teams to design a digital mental health and psychological support system for college students that offers confidential self-assessment, guided support, access to trained counsellors and, where needed, escalation to professional help. The solution must be sensitive, safe and clearly not a replacement for clinical care, with strong privacy protection and careful handling of crisis situations. Teams should consider anonymity, multilingual access, peer support, awareness content and ways to help institutions understand overall wellbeing trends without exposing individuals. A strong proposal covers the user journey, safety design, counsellor tools and the plan for responsible deployment."
  },
  {
    "id": "EC-01",
    "code": "EC-01",
    "sihNo": "SIH25051",
    "domain": "Electronics and Communication",
    "category": "Hardware",
    "title": "Renewable Energy Monitoring System for Microgrids",
    "org": "Government of Odisha",
    "theme": "Renewable / Sustainable Energy",
    "description": "Solar-based microgrids are increasingly used in villages, campuses and remote facilities, but operators often lack real-time visibility into how much energy is generated, stored and consumed, so faults, battery degradation and wasteful usage go unnoticed until supply fails. Commercial monitoring systems are expensive and difficult to maintain locally. This case study asks teams to design a low-cost hardware monitoring solution that tracks generation, storage and load in real time and alerts operators to abnormal behaviour. Teams must choose accurate and affordable sensors, a reliable data link such as LoRa, GSM or Wi-Fi, and a dashboard that a local technician can understand. A strong answer also includes basic fault prediction, for example detecting declining panel output or battery health, along with a realistic cost estimate and a power-efficient design."
  },
  {
    "id": "EC-02",
    "code": "EC-02",
    "sihNo": "SIH25021",
    "domain": "Electronics and Communication",
    "category": "Hardware",
    "title": "AI-Based Development of Laser-Based QR Code Marking on Track Fittings on Indian Railways",
    "org": "Ministry of Railways",
    "theme": "Transportation & Logistics",
    "description": "Indian Railways uses millions of track fittings such as elastic rail clips, liners and pads, and the lack of permanent identification makes it hard to trace a component's manufacturer, batch, installation date and inspection history. When a failure occurs, identifying the affected batch is slow and expensive. This case study asks teams to develop a system for marking QR or data-matrix codes on track fittings using laser technology, supported by AI for quality checking and reliable reading in the field. The marking must be durable on metal and rubber surfaces exposed to dust, rust, rain and vibration, and the reader must work in poor light and on dirty surfaces. A good proposal covers the marking method, code design, vision-based verification, mobile reading, database integration and reliability testing."
  },
  {
    "id": "EC-03",
    "code": "EC-03",
    "sihNo": "SIH25020",
    "domain": "Electronics and Communication",
    "category": "Hardware",
    "title": "Development of Indigenous Contactless Integrated Track Monitoring Systems (ITMS) for Track Recording on Indian Railways",
    "org": "Ministry of Railways",
    "theme": "Smart Automation",
    "description": "Regular measurement of track geometry and condition is essential for safe and comfortable train operation, but traditional track recording relies on contact-based equipment and imported systems that are costly and need specialised vehicles and maintenance. This case study asks teams to develop an indigenous contactless integrated track monitoring system that records track parameters while moving, for example using lasers, cameras and inertial sensors mounted on a rail vehicle. Teams must address measurement accuracy at speed, vibration, varying light and weather, synchronisation of multiple sensors and processing of large data streams. A strong solution describes the sensor suite, the mounting and data acquisition design, the algorithms that convert raw data into track parameters, the accuracy validation approach and a cost comparison with imported systems."
  },
  {
    "id": "EC-04",
    "code": "EC-04",
    "sihNo": "SIH25036",
    "domain": "Electronics and Communication",
    "category": "Hardware",
    "title": "Development of Sensor for Detection of Microplastics",
    "org": "Ministry of Earth Sciences (MoES)",
    "theme": "Miscellaneous",
    "description": "Microplastics are now found in oceans, rivers, drinking water and even food, yet detecting and counting them still depends on laboratory methods that are slow, expensive and unsuitable for continuous monitoring. This case study asks teams to develop a sensor that can detect microplastics in water samples quickly and at low cost. Teams may explore optical methods such as fluorescence, scattering or spectroscopy, microfluidic sample handling or electrical sensing. Major challenges include distinguishing plastic particles from organic matter and sand, handling very small particle sizes, calibration and keeping the device portable and affordable for field use. A strong proposal explains the sensing principle, the sample preparation approach, the electronics and signal processing, expected detection limits, validation against laboratory methods and a plan for field deployment."
  },
  {
    "id": "EC-05",
    "code": "EC-05",
    "sihNo": "SIH25063",
    "domain": "Electronics and Communication",
    "category": "Hardware",
    "title": "Cost-Effective Solution for Detecting Breakage of Low Voltage AC Distribution Overhead Conductors",
    "org": "Government of Kerala",
    "theme": "Disaster Management",
    "description": "When a low voltage overhead conductor snaps and falls, it can remain live on the ground and cause fatal electric shocks to people and animals, especially during storms and heavy rain, because protection devices often do not trip for such faults. Utilities need a way to detect conductor breakage quickly and isolate the line before harm occurs. This case study asks teams to develop a cost-effective hardware solution that detects the breakage of low voltage AC overhead conductors and signals or disconnects the affected line. Teams must work with the electrical characteristics of distribution networks, avoid false trips from normal load changes and keep the cost low enough to be installed widely. A good solution describes the detection principle, the electronics, the communication to the utility, field installation, testing under realistic fault conditions and a cost estimate."
  },
  {
    "id": "EC-06",
    "code": "EC-06",
    "sihNo": "SIH25062",
    "domain": "Electronics and Communication",
    "category": "Hardware",
    "title": "Implementation of Smart Agriculture for Efficient Cultivation in Hilly Regions",
    "org": "Government of Sikkim",
    "theme": "Agriculture, FoodTech & Rural Development",
    "description": "Farming in hilly regions such as Sikkim faces steep terrain, small and scattered terraces, variable microclimates, soil erosion and limited access to irrigation and markets, so techniques designed for plains do not work well. This case study asks teams to design a smart agriculture system suited to hill farming that uses sensors and automation to improve water use, monitor soil and weather and guide decisions on planting, irrigation and crop protection. The hardware must be low-power, rugged and able to operate where network coverage is weak, and the interface must be simple for farmers who may not be technically trained. A strong proposal describes the sensing and control hardware, the communication approach, the decision logic, how the system copes with terrain and weather, the cost per farm and the expected benefits for yield and income."
  },
  {
    "id": "EC-07",
    "code": "EC-07",
    "sihNo": "SIH25064",
    "domain": "Electronics and Communication",
    "category": "Hardware",
    "title": "Improving Renewable Energy Hosting Capacity in Distribution Feeders and Power Quality during High RE Injection",
    "org": "Government of Kerala",
    "theme": "Renewable / Sustainable Energy",
    "description": "As rooftop solar and other renewable sources connect to distribution feeders, the network can face voltage rise, reverse power flow, harmonics and instability, which limits how much renewable energy a feeder can safely accept. This case study asks teams to develop a hardware-based solution that increases the renewable energy hosting capacity of distribution feeders while maintaining power quality when renewable injection is high. Teams may explore power electronic devices, smart inverters with voltage support, energy storage, monitoring and control schemes or reactive power compensation. Challenges include keeping the solution affordable, coordinating many small sources and verifying performance under realistic feeder conditions. A strong proposal explains the problem analysis, the proposed hardware and control approach, simulation results, the communication and monitoring design and the expected improvement in hosting capacity."
  },
  {
    "id": "EC-08",
    "code": "EC-08",
    "sihNo": "SIH25047",
    "domain": "Electronics and Communication",
    "category": "Hardware",
    "title": "Disaster Response Drone for Remote Areas",
    "org": "Government of Odisha",
    "theme": "Robotics and Drones",
    "description": "After cyclones, floods and landslides, roads and bridges are often cut off, and relief teams cannot quickly learn where people are stranded or deliver urgent supplies such as medicines, food and communication devices. This case study asks teams to design a drone system for disaster response in remote areas that can survey affected zones, locate survivors and carry small payloads. The drone must operate in wind and rain, cover meaningful distances on limited battery power and be simple enough for relief workers to deploy. Teams must consider the airframe and payload design, sensors such as thermal and visual cameras, navigation without reliable GPS or network, and safe operation. A strong proposal covers the drone architecture, power and range calculations, sensing and communication, a field operation plan and the cost of a deployable unit."
  },
  {
    "id": "EC-09",
    "code": "EC-09",
    "sihNo": "SIH25077",
    "domain": "Electronics and Communication",
    "category": "Hardware",
    "title": "Hardware to Detect and Prevent Unauthorized Use of Electric Fences",
    "org": "Government of Kerala",
    "theme": "Smart Automation",
    "description": "Electric fences are widely used to protect farms from wild animals, but when installed or modified illegally, for example by connecting them directly to the mains supply, they can electrocute animals and people. Authorities need an effective way to detect unauthorised or unsafe use and prevent accidents. This case study asks teams to develop a hardware solution that monitors electric fence installations, detects unauthorised connections or unsafe voltage and current levels, and cuts the supply or alerts authorities when a violation occurs. The device must be rugged, weather resistant, low cost and tamper-proof, and must not interfere with legitimate low-energy fence energisers. A strong proposal describes the sensing and detection logic, the protection and switching hardware, the communication with the authority, field testing and how installations can be certified and tracked."
  },
  {
    "id": "EC-10",
    "code": "EC-10",
    "sihNo": "SIH25043",
    "domain": "Electronics and Communication",
    "category": "Hardware",
    "title": "Embedded Intelligent Microscopy System for Identification and Counting of Microscopic Marine Organisms",
    "org": "Ministry of Earth Sciences (MoES)",
    "theme": "Smart Automation",
    "description": "Plankton and other microscopic marine organisms form the base of the ocean food chain, and monitoring them is important for fisheries, water quality and climate studies, but identifying and counting them under a microscope is slow, tiring and depends on scarce experts. This case study asks teams to develop an embedded, intelligent microscopy system that automatically captures images of water samples and identifies and counts microscopic marine organisms. Teams must design the optics and illumination, sample handling, an embedded processor capable of running image analysis and an AI model trained to distinguish many species from debris. Portability, low power and ease of use on a research vessel or coastal station are important requirements. A strong proposal describes the optical and mechanical design, the imaging pipeline, the model and its training data, accuracy targets and the output format for scientists."
  },
  {
    "id": "ME-01",
    "code": "ME-01",
    "sihNo": "SIH25053",
    "domain": "Mechanical and Civil",
    "category": "Hardware",
    "title": "Improved Onion Storage Technology for Enhancing Shelf Life of Onions",
    "org": "Ministry of Consumer Affairs, Food & Public Distribution",
    "theme": "Agriculture, FoodTech & Rural Development",
    "description": "Onion prices swing sharply from season to season because a large part of the crop is lost to rotting, sprouting and weight loss in storage, and traditional storage structures give farmers and traders little control over temperature, humidity and ventilation. This case study asks teams to design an improved onion storage technology that extends shelf life and reduces losses while remaining affordable for farmers and small traders. Teams should study the factors that cause onion spoilage, calculate the airflow, temperature and humidity requirements for a chosen storage capacity, and propose structural design, ventilation or cooling methods and possible use of solar power or low-energy systems. Practical constraints include cost, local materials, ease of loading and unloading and maintenance by users with limited technical skills. A strong proposal includes design calculations, drawings, the control approach, expected reduction in losses and a payback analysis."
  },
  {
    "id": "ME-02",
    "code": "ME-02",
    "sihNo": "SIH25003",
    "domain": "Mechanical and Civil",
    "category": "Hardware",
    "title": "Low-Cost Smart Transportation Solution for Agri Produce from Remote Farms to the Nearest Motorable Road in the NER Region",
    "org": "Ministry of Development of North Eastern Region",
    "theme": "Transportation & Logistics",
    "description": "In the hills of the North East, many farms lie far from the nearest motorable road, and farmers must carry vegetables, fruits and spices on their backs over steep, narrow paths, which causes damage to the produce, loss of time and physical strain, and reduces the price they finally receive. This case study asks teams to design a low-cost smart transportation solution that moves agricultural produce from remote farms to the nearest motorable road. Options could include ropeways, cable carriers, tracked or small powered carriers and electric or human-assisted vehicles. The design must handle steep slopes, rain, uneven terrain, limited power availability and low cost of building and maintenance by village communities. A strong proposal presents the mechanism, load and power calculations, safety features, the control or tracking element that makes it smart and an estimate of cost per farm cluster."
  },
  {
    "id": "ME-03",
    "code": "ME-03",
    "sihNo": "SIH25072",
    "domain": "Mechanical and Civil",
    "category": "Hardware",
    "title": "Design and Implementation of Solar-Powered Dewatering in Mining Operations",
    "org": "Ministry of Mines",
    "theme": "Renewable / Sustainable Energy",
    "description": "Mines must continuously pump out water that collects in pits and underground workings, and this dewatering is one of the largest consumers of diesel and grid electricity in mining, adding to cost and emissions. This case study asks teams to design a solar-powered dewatering system for mining operations. The design must estimate the volume and head of water to be pumped, size the pumps, solar array and any energy storage, and handle variable inflow during rain, dust deposition on panels and the harsh mining environment. Teams should consider hybrid operation with the grid or generators for reliability, safe installation in mining areas and ways to reduce pumping energy. A strong proposal includes hydraulic and electrical calculations, a system layout, the control strategy, an economic comparison with diesel or grid pumping and a plan for maintenance and safety."
  },
  {
    "id": "CE-01",
    "code": "CE-01",
    "sihNo": "SIH25065",
    "domain": "Mechanical and Civil",
    "category": "Software",
    "title": "Application for On-Spot Assessment of Roof Top Rain Water Harvesting and Artificial Recharge Potential and Size of RTRWH and AR Structures",
    "org": "Ministry of Jal Shakti",
    "theme": "Smart Automation",
    "description": "Rooftop rainwater harvesting and artificial recharge can significantly improve groundwater levels in cities and villages, but property owners, engineers and officials often lack a quick, reliable way to judge how much water a particular site can harvest and what size of storage or recharge structure is appropriate. This case study asks teams to design an application that performs an on-spot assessment of rooftop rainwater harvesting and artificial recharge potential and recommends the size and type of structure. The app would need to use rooftop area and surface, local rainfall data, soil and groundwater conditions and space available, and then apply sound hydrological and civil engineering calculations. A strong solution explains the inputs and data sources, the calculation method, the recommended structure designs, the user interface for field officers and citizens and the way estimated benefits are presented."
  },
  {
    "id": "CE-02",
    "code": "CE-02",
    "sihNo": "SIH25090",
    "domain": "Mechanical and Civil",
    "category": "Hardware",
    "title": "Grey Water Management and Reuse, Wetland Management and Water Conservation",
    "org": "Government of Kerala",
    "theme": "Clean & Green Technology",
    "description": "Households, hostels and institutions produce large volumes of grey water from kitchens, bathrooms and laundry, which is usually discharged directly into drains or the soil, wasting a resource that could be reused for gardening, flushing and other non-potable purposes, while also stressing sewage systems and natural wetlands. This case study asks teams to design practical systems for grey water management and reuse, supported by wetland management and water conservation measures. Teams should study the quality and quantity of grey water, propose treatment methods suitable for small and medium scale such as filtration, constructed wetlands and settling units, and design storage and distribution for reuse. Cost, land requirement, odour, maintenance and public acceptance are important constraints. A strong proposal includes design calculations, a layout, expected water savings, monitoring of treated water quality and a plan for operation and maintenance."
  }
];
}
