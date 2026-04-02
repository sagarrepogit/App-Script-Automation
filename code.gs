// ─────────────────────────────────────────
// GLOBAL CONFIGURATION
// ─────────────────────────────────────────
var props = PropertiesService.getUserProperties();

// Project & authentication details
var projectId = props.getProperty('project_id') || '27433';
var cycleId = props.getProperty('CYCLE_ID');
var build_id = props.getProperty('build_id');
var EMAIL = props.getProperty('EMAIL');
var PASSWORD = props.getProperty('PASSWORD');

// Subdomain dynamic banaya gaya (scalable design)
var SUBDOMAIN = props.getProperty('SUBDOMAIN') || 'vfqatar-prod';


// Sheet ke dynamic column indexes (0-based: column A = 0)
var LINK_COL_INDEX = Number(props.getProperty('DRIVE_LINK_COLUMN_INDEX'));
var TC_NAME_COL_INDEX = Number(props.getProperty('TC_NAME_COLUMN_INDEX'));

var KUALITEE_BASE_URL = 'https://apiss3.kualitee.com/api/v2';


// ─────────────────────────────────────────
// HELPER: Sleep / Delay
// ─────────────────────────────────────────
function sleep() {
  Utilities.sleep(200); // 🔥 fixed delay
}
// ─────────────────────────────────────────
// SETTINGS PAGE SUPPORT
// ─────────────────────────────────────────
// Settings form is filled via HtmlTemplate (server-side) so library users do not need
// getSavedUserData in the bound script. Saving still requires saveUserData() there;
// see BOUND_SCRIPT_STUBS.txt
function getSavedUserData() {
  return {
    EMAIL: props.getProperty('EMAIL') || '',
    PASSWORD: props.getProperty('PASSWORD') || '',
    CYCLE_ID: props.getProperty('CYCLE_ID') || '',
    build_id: props.getProperty('build_id') || '',
    TC_NAME_COLUMN_INDEX: props.getProperty('TC_NAME_COLUMN_INDEX') || '',
    DRIVE_LINK_COLUMN_INDEX: props.getProperty('DRIVE_LINK_COLUMN_INDEX') || '',
    project_id: props.getProperty('project_id') || '27433',
    SUBDOMAIN: props.getProperty('SUBDOMAIN') || 'vfqatar-prod'
  };
}

function saveUserData(data) {
  props.setProperties({
    'EMAIL': data.EMAIL,
    'PASSWORD': data.PASSWORD,
    'CYCLE_ID': data.CYCLE_ID,
    'build_id': data.build_id,
    'project_id': data.project_id || '27433',
    'SUBDOMAIN': data.SUBDOMAIN || 'vfqatar-prod',
    'TC_NAME_COLUMN_INDEX': data.TC_NAME_COLUMN_INDEX,
    'DRIVE_LINK_COLUMN_INDEX': data.DRIVE_LINK_COLUMN_INDEX
  });
  return true;
}


// ─────────────────────────────────────────
// MENU + UI
// ─────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu("🚀 Sagar's AI Tool")
    .addItem('⚙️ Open Settings', 'showSettings')
    .addSeparator()
    .addItem('🤖 Run AI Automation', 'Run_AI_KualiteeAutomation')
    .addToUi();
}

function showSettings() {
  var d = getSavedUserData();
  var t = HtmlService.createTemplateFromFile('SettingPage');
  t.email = d.EMAIL;
  t.password = d.PASSWORD;
  t.cycleId = d.CYCLE_ID;
  t.buildId = d.build_id;
  t.tcNameCol = d.TC_NAME_COLUMN_INDEX;
  t.driveLinkCol = d.DRIVE_LINK_COLUMN_INDEX;
  var html = t.evaluate()
    .setWidth(450)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "Sagar's Tool Configuration");
}


// ─────────────────────────────────────────
// COMMON API HANDLER (🔥 FIXED)
// ─────────────────────────────────────────
// ─────────────────────────────────────────
// COMMON API HANDLER (TOKEN + RETRY LOGIC)
// ─────────────────────────────────────────
// Yeh ek centralized function hai jo:
// 1. Token attach karta hai
// 2. API call karta hai
// 3. Token expire hone par auto-refresh karta hai

function callKualiteeAPI(endpoint, payload, isFormData) {
  // Existing token lo
  var token = getKualiteeToken();
  payload.token = token;

  var url = KUALITEE_BASE_URL + endpoint;

  var options = {
    method: 'post',
    muteHttpExceptions: true
  };

  if (isFormData) {
    options.payload = payload;
  } else {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }
  // First API attempt
  var response = UrlFetchApp.fetch(url, options);
  var result = parseFetchResponse_(response);

  // 🔥 Auto token refresh
  // Agar token expire ho gaya hai → retry karo
  if (result.json && result.json.token_expire === true) {
    // Force new login
    var newToken = getKualiteeToken(true);
    payload.token = newToken;

    options.payload = isFormData ? payload : JSON.stringify(payload);
     // Retry request
    response = UrlFetchApp.fetch(url, options);
    result = parseFetchResponse_(response);
  }

  return result.json;
}


// ─────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────

// API response ko safely parse karta hai
function parseFetchResponse_(resp) {
  var text = resp.getContentText();
  var code = resp.getResponseCode();

  var json = null;
  try { json = JSON.parse(text); } catch (e) {} // Ignore parsing error

  return { ok: code >= 200 && code < 300, code, text, json };
}
  // String normalization (matching improve karne ke liye)
// Example: "Login Test " → "login test"
function normalize_(str) {
  return String(str || '').toLowerCase().trim();
}


// ─────────────────────────────────────────
// TOKEN MANAGEMENT
// ─────────────────────────────────────────
// Token ko cache kiya jata hai taaki baar-baar login na karna pade

function getKualiteeToken(forceRefresh = false) {

  var cache = CacheService.getScriptCache();
     
      // Agar force refresh hai → old token hata do

  if (forceRefresh) cache.remove('KUALITEE_TOKEN');

  var token = cache.get('KUALITEE_TOKEN');
    // Agar force refresh hai → old token hata do
    // Agar token already hai → reuse karo
  if (token && !forceRefresh) return token;

  // Naya token generate karo (login API call)
  var url = KUALITEE_BASE_URL + '/auth/signin';

  var payload = {
    email_id: EMAIL,
    password: PASSWORD,
    subdomain: SUBDOMAIN,
    timezone: "Asia/Calcutta"
  };

  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  });

  var result = JSON.parse(resp.getContentText());

  // Token cache karo (6 hours)
  // Token cache karo (6 hours)
if (result.token) {
    cache.put('KUALITEE_TOKEN', result.token, 21600);
    return result.token;
}

throw new Error("Login failed: " + JSON.stringify(result));
}


// ─────────────────────────────────────────
// OPTIMIZED TEST CASE FETCH (🔥 FAST)
// ─────────────────────────────────────────
// ─────────────────────────────────────────
// FETCH ALL TEST CASES (OPTIMIZED)
// ─────────────────────────────────────────
// IMPORTANT: Yeh function sirf 1 baar API call karta hai
// Aur saare test cases ko map mein store karta hai
// Isse performance dramatically improve hoti hai

function getAllTestCasesMap_() {

  // ────────────────────────────────
  // STEP 1: Prepare API payload
  // ────────────────────────────────
  // Yeh payload /test_case_execution/list endpoint ko bhejne ke liye hai
  // start: 0 → first record se start
  // length: 500 → ek hi call me max 500 test cases fetch karenge
  // (Optimized: multiple calls ki zarurat nahi, batch fetch)
  var payload = {
    project_id: String(projectId),
    cycle_id: String(cycleId),
    build_id: String(build_id),
    start: 0,
    length: 500 // Large batch fetch
  };

  // ────────────────────────────────
  // STEP 2: Fetch test case list from Kualitee
  // ────────────────────────────────
  var res = callKualiteeAPI('/test_case_execution/list', payload);

  // Agar API call fail ho gayi ya data nahi mila → error throw karo
  if (!res || !res.data) {
    throw new Error("Failed to fetch test cases");
  }

  // ────────────────────────────────
  // STEP 3: Create a Map for quick lookup
  // ────────────────────────────────
  // map object banaya: key = normalized test case name, value = tc details
  // Normalization: lowercase + trim → "Login Test " aur "login test" match ho jaye
  var map = {};

  res.data.forEach(tc => {
    map[normalize_(tc.tc_name)] = {
      tcId: tc.testcase_id,
      testscenario_id: tc.testscenario_id
    };
  });

  // Map me ab saare test cases stored hain:
  // Example:
  // map["login test"] = { tcId: 1234, testscenario_id: 5678 }
  // Ye lookup O(1) time me possible hai → optimize performance
  // Pehle: har row ke liye list API call (O(n)) → Ab: 1 API call aur O(1) lookup

  // ────────────────────────────────
  // STEP 4: Return the Map
  // ────────────────────────────────
  return map;
}

// ─────────────────────────────────────────
// EXECUTE TEST CASE
// ─────────────────────────────────────────
// Test case ko "Passed" ya desired status mein mark karta hai

function executeTestCase_(tc) {

  // Delay before execute test case (rate limit safe)
    sleep(); // 200ms wait


  var payload = {
    project_id: String(projectId),
    cycle_id: String(cycleId),
    build_id: String(build_id),
    tc_id: String(tc.tcId),
    testscenario_id: String(tc.testscenario_id),
    status: "Passed",
    execute: "yes",
    time: 0
  };

  var res = callKualiteeAPI('/test_case_execution/execute', payload);

  if (res.executed_results && res.executed_results.length > 0) {
    return res.executed_results[0].id;
  }

  throw new Error("Execution failed");
}


// ─────────────────────────────────────────
// UPLOAD EVIDENCE
// ─────────────────────────────────────────
// ─────────────────────────────────────────
// UPLOAD EVIDENCE FROM GOOGLE DRIVE
// ─────────────────────────────────────────
// Google Drive link se file fetch karke Kualitee par upload karta hai

function uploadEvidence_(tcId, executionId, driveLink, statusCell) {
  try {
    // 1️⃣ File ID extract
    var fileIdMatch = String(driveLink).match(/[-\w]{25,}/);
    if (!fileIdMatch) throw new Error("Invalid Drive link");
    var fileId = fileIdMatch[0];

    // 2️⃣ File object
    var file = DriveApp.getFileById(fileId);

    // 3️⃣ Allowed file types
    var allowedTypes = ['gif','jpg','png','jpeg','pdf','docx','csv','xls','ppt','mp4','webm','msg','eml','zip','xml','pcap'];
    var fileName = file.getName();
    var ext = fileName.split('.').pop().toLowerCase();

    if (allowedTypes.indexOf(ext) === -1) {
      throw new Error("File type not allowed: " + ext);
    }

    // 4️⃣ File size check: 50MB
    if (file.getSize() > 50 * 1024 * 1024) {
      throw new Error("File exceeds 50MB limit");
    }

    // 5️⃣ Prepare blob
    var blob = file.getBlob();

    // 6️⃣ Upload to Kualitee
    var formData = {
      project_id: String(projectId),
      cycle_id: String(cycleId),
      testcase_id: String(tcId),
      execution_id: String(executionId),
      type: 'tc',
      'attachment[]': blob
    };

    callKualiteeAPI('/test_case_execution/execution_attachments', formData, true);

    // ✅ Success status
    statusCell.setValue("Uploaded ✅") // sirf capital letters
      .setBackground("#b6d7a8")
      .setFontWeight("bold");

  } catch (e) {
    statusCell.setValue(e.message).setBackground("#ea9999");
  }
}

// ─────────────────────────────────────────
// MAIN FUNCTION
// ─────────────────────────────────────────
// ─────────────────────────────────────────
// MAIN AUTOMATION FUNCTION (ENTRY POINT)
// ─────────────────────────────────────────
// Yeh function poora workflow execute karta hai:
// 1. Sheet read karta hai
// 2. Test case match karta hai
// 3. Execute karta hai
// 4. Evidence upload karta hai

function Run_AI_KualiteeAutomation() {

  var rawLinkCol = props.getProperty('DRIVE_LINK_COLUMN_INDEX');
  var rawTcCol = props.getProperty('TC_NAME_COLUMN_INDEX');
  if (rawLinkCol == null || String(rawLinkCol).trim() === '' ||
      rawTcCol == null || String(rawTcCol).trim() === '') {
    throw new Error("Column index not set");
  }
  if (isNaN(LINK_COL_INDEX) || isNaN(TC_NAME_COL_INDEX) ||
      LINK_COL_INDEX < 0 || TC_NAME_COL_INDEX < 0) {
    throw new Error("Column index invalid");
  }

  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var testCaseMap = getAllTestCasesMap_(); // 🔥 only once

  for (var i = 1; i < data.length; i++) {

    var tcName = data[i][TC_NAME_COL_INDEX];
    var driveLink = data[i][LINK_COL_INDEX];
    var statusCell = sheet.getRange(i + 1, LINK_COL_INDEX + 2);
    var currentStatus = data[i][LINK_COL_INDEX + 1];

    if (String(currentStatus).toLowerCase().includes("uploaded")) continue; // skip isme uploaded chahe jaise lika ho fark nahi padta hai emoji ka bhi koi issue nahi hoga
    if (!tcName || !driveLink) continue; // skip if either is missing;

    try {
      var tc = testCaseMap[normalize_(tcName)];

      if (!tc) throw new Error("TC not found");

      var executionId = executeTestCase_(tc);

      // Delay before upload (rate limit safe)
    sleep(); // 200ms wait

      uploadEvidence_(tc.tcId, executionId, driveLink, statusCell);

    } catch (err) {
      statusCell.setValue(err.message).setBackground("#ea9999");
    }
  }
}
