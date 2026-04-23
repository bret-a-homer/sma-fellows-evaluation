// =============================================================================
// SMA Harvard Fellows Program — Evaluation Tool
// Google Apps Script Backend
// =============================================================================
// SETUP INSTRUCTIONS
// 1. Open your Google Sheet, then Extensions → Apps Script
// 2. Paste this entire file, replacing any existing content
// 3. Run the setup() function once (Run → Run function → setup)
//    - Authorize the script when prompted
//    - This creates headers, the data dictionary tab, and schedules triggers
// 4. Deploy as web app: Deploy → New deployment
//    - Execute as: Me
//    - Who has access: Anyone
//    - Copy the web app URL and paste it into index.html as APPS_SCRIPT_URL
// 5. To update the script later: Deploy → Manage deployments → edit ✏️ →
//    increment version → Deploy  (NEVER create a new deployment — URL will change)
// =============================================================================

// ── Config ────────────────────────────────────────────────────────────────────
const SHEET_NAME      = 'Responses';
const DICT_SHEET_NAME = 'Data Dictionary';
const BACKUP_PREFIX   = 'Backup_';   // prefix for dated backup tabs
const BACKUP_DAYS     = 90;          // how many days of backups to keep

// ── Reference data (mirrors frontend constants) ───────────────────────────────
const MOTIVATION_LABELS = [
  'External',
  'Introjected',
  'Identified',
  'Integrated',
  'Intrinsic',
  'Fully Intrinsic',
];

const CAUSE_AREAS = [
  { key: 'Global_Health',    full: 'Global health and development' },
  { key: 'Animal_Welfare',   full: 'Animal welfare and alternative proteins' },
  { key: 'AI_Safety',        full: 'AI safety and governance' },
  { key: 'Biosecurity',      full: 'Biosecurity and pandemic preparedness' },
  { key: 'Nuclear_Risk',     full: 'Nuclear risk reduction' },
  { key: 'Climate',          full: 'Climate change and energy transition' },
];

const CAUSE_NUM = { 'Not familiar': 0, 'Somewhat familiar': 1, 'Actively engaged': 2 };

const CAREER_VALUES = [
  { key: 'Altruism',                  name: 'Altruism' },
  { key: 'Prestige',                  name: 'Prestige' },
  { key: 'Economic_Return',           name: 'Economic Return' },
  { key: 'Intellectual_Stimulation',  name: 'Intellectual Stimulation' },
  { key: 'Independence',              name: 'Independence' },
  { key: 'Achievement',               name: 'Achievement' },
];

const SE_ITEMS = [
  { key: 'Theory_Change_Dev',  full: 'Developing a theory of change' },
  { key: 'Theory_Change_Eval', full: 'Evaluating a theory of change' },
  { key: 'Counterfactual',     full: 'Considering counterfactual tradeoffs for different altruistic pursuits' },
  { key: 'Finding_Opps',       full: 'Finding high-impact, altruistic career opportunities' },
  { key: 'Networking',         full: 'Establishing professional contacts to advance a career in high-impact altruism' },
];

const BEHAVIORAL_ITEMS = [
  { key: 'EA_Event',      full: 'Attended an EA or SMA event in the past 12 months' },
  { key: 'EA_Content',    full: 'Read or listened to EA-adjacent content (books, podcasts, articles)' },
  { key: 'Career_Convo',  full: 'Had a serious career conversation with someone from an EA-aligned org' },
  { key: 'Applied',       full: 'Applied to, or seriously researched, a high-impact job or fellowship' },
];

const BARRIERS = [
  { key: 'Financial_Security',    full: 'Concern about financial security or income' },
  { key: 'Family_Peer_Pressure',  full: 'Pressure from family or peers to take a conventional path' },
  { key: 'Cause_Area_Uncertainty',full: 'Uncertainty about which cause area to focus on' },
  { key: 'Skills_Credentials',    full: 'Lack of relevant skills or credentials' },
  { key: 'Org_Culture',           full: 'Navigating a new organizational culture' },
  { key: 'Managing_Up',           full: 'Managing up to my supervisor' },
  { key: 'Quality_Work',          full: 'Producing high-quality work quickly' },
  { key: 'Policy_Landscape',      full: 'Understanding the policy or political landscape' },
  { key: 'Senior_Stakeholders',   full: 'Building relationships with senior stakeholders' },
  { key: 'Nonprofit_Funding',     full: 'Understanding nonprofit funding and sustainability' },
  { key: 'Short_Term_Impact',     full: 'Having an impact despite being there short-term' },
  { key: 'Ambiguity',             full: 'Dealing with ambiguity in my role' },
  { key: 'Presenting_Leadership', full: 'Presenting my work to leadership' },
  { key: 'Work_Life_Balance',     full: 'Maintaining work-life balance' },
  { key: 'Other',                 full: 'Other' },
];

// =============================================================================
// COLUMN SCHEMA
// =============================================================================
function getHeaders() {
  const h = [];

  // ── Metadata ────────────────────────────────────────────────────────────────
  h.push('Email', 'Name', 'Link_Status');
  h.push('T1_Submitted', 'T2_Submitted', 'T3_Submitted', 'T4_Submitted');

  // ── T1 ──────────────────────────────────────────────────────────────────────
  h.push('T1_Career_Vision');
  CAREER_VALUES.forEach(v => h.push('T1_Rank_' + v.key));
  h.push('T1_Motivation_Index', 'T1_Motivation_Label');
  h.push('T1_Commitment');
  h.push('T1_Orgs_Considering');
  BEHAVIORAL_ITEMS.forEach(b => h.push('T1_Behavior_' + b.key));
  CAUSE_AREAS.forEach(c => h.push('T1_Cause_' + c.key, 'T1_Cause_' + c.key + '_Num'));
  h.push('T1_Peer_Influence', 'T1_Peer_Conv_YN', 'T1_Peer_Conv_Count');
  SE_ITEMS.forEach(s => h.push('T1_SE_' + s.key));
  h.push('T1_Career_Direction');

  // ── T2 ──────────────────────────────────────────────────────────────────────
  h.push('T2_Career_Vision', 'T2_Thinking_Shifted');
  CAREER_VALUES.forEach(v => h.push('T2_Rank_' + v.key));
  h.push('T2_Commitment_Now', 'T2_Commitment_Then');
  h.push('T2_Motivation_Now_Index', 'T2_Motivation_Now_Label');
  h.push('T2_Motivation_Then_Index', 'T2_Motivation_Then_Label');
  h.push('T2_Impactful_Sessions');
  CAUSE_AREAS.forEach(c => h.push('T2_Cause_' + c.key, 'T2_Cause_' + c.key + '_Num'));
  SE_ITEMS.forEach(s => h.push('T2_SE_' + s.key));
  BARRIERS.forEach(b => h.push('T2_Barrier_' + b.key));
  h.push('T2_Barrier_Other_Text');
  h.push('T2_Placement_Readiness', 'T2_Career_Capital_Goals', 'T2_Peer_Influence');

  // ── T3 & T4 ─────────────────────────────────────────────────────────────────
  ['T3', 'T4'].forEach(function(T) {
    h.push(T + '_Org_Name', T + '_Role');
    h.push(T + '_BARS_Intellectual_Rigor', T + '_BARS_Prof_Development', T + '_BARS_Impact_Meaningfulness');
    h.push(T + '_Readiness_Retrospective');
    BARRIERS.forEach(b => h.push(T + '_Barrier_' + b.key));
    h.push(T + '_Barrier_Other_Text', T + '_Barrier_Detail');
    h.push(T + '_Career_Capital_Delivery');
    h.push(T + '_NPS');
    h.push(T + '_Most_Valuable', T + '_Suggested_Improvement');
    CAREER_VALUES.forEach(v => h.push(T + '_Rank_' + v.key));
    h.push(T + '_Career_Vision');
    h.push(T + '_Peer_Conv_Count', T + '_Most_Significant_Conv');
    for (var n = 1; n <= 5; n++) {
      h.push(T + '_Role' + n + '_Name', T + '_Role' + n + '_Status', T + '_Role' + n + '_Sector');
    }
    h.push(T + '_Roles_Count');
    h.push(T + '_Career_Direction', T + '_Career_Dir_Factors', T + '_Career_Dir_Influences');
  });

  return h;
}

// =============================================================================
// COLUMN DATA BUILDER
// Maps the frontend state.answers[tp] payload to { columnName: value } pairs
// =============================================================================
function buildColumnData(tp, data, name) {
  var T   = tp.toUpperCase();   // 'T1' | 'T2' | 'T3' | 'T4'
  var cols = {};

  cols[T + '_Submitted'] = new Date().toISOString();

  // Helper: get career value rank from the saved rank array
  function careerRanks(rankKey) {
    var arr = data[rankKey];
    if (!Array.isArray(arr)) return;
    CAREER_VALUES.forEach(function(v) {
      var idx = arr.indexOf(v.name);
      cols[T + '_Rank_' + v.key] = idx >= 0 ? idx + 1 : '';
    });
  }

  // Helper: cause area grid (stored as object with numeric string keys)
  function causeAreas(dataKey, prefix) {
    var obj = data[dataKey] || {};
    CAUSE_AREAS.forEach(function(c, i) {
      var val = obj[i] !== undefined ? obj[i] : (obj[String(i)] || '');
      cols[prefix + '_Cause_' + c.key]         = val;
      cols[prefix + '_Cause_' + c.key + '_Num'] = val !== '' ? (CAUSE_NUM[val] !== undefined ? CAUSE_NUM[val] : '') : '';
    });
  }

  // Helper: self-efficacy grid (stored as object with numeric string keys)
  function selfEfficacy(dataKey, prefix) {
    var obj = data[dataKey] || {};
    SE_ITEMS.forEach(function(s, i) {
      cols[prefix + '_SE_' + s.key] = obj[i] !== undefined ? obj[i] : (obj[String(i)] || '');
    });
  }

  // Helper: barriers checklist (stored as array of full text strings)
  function barriers(dataKey, otherKey, detailKey, prefix) {
    var arr = data[dataKey] || [];
    BARRIERS.forEach(function(b) {
      cols[prefix + '_Barrier_' + b.key] = arr.indexOf(b.full) >= 0 ? 1 : 0;
    });
    cols[prefix + '_Barrier_Other_Text'] = data[otherKey] || '';
    if (detailKey) cols[prefix + '_Barrier_Detail'] = data[detailKey] || '';
  }

  // Helper: motivation index + label
  function motivation(valKey, indexCol, labelCol) {
    var val = data[valKey];
    if (val !== undefined && val !== '') {
      var idx = parseInt(val);
      cols[indexCol] = idx;
      cols[labelCol] = MOTIVATION_LABELS[idx] || '';
    }
  }

  // ── T1 ────────────────────────────────────────────────────────────────────
  if (tp === 't1') {
    if (name) cols['Name'] = name;
    cols['T1_Career_Vision']    = data['t1-q2']    || '';
    careerRanks('t1-rank');
    motivation('t1-motiv-val', 'T1_Motivation_Index', 'T1_Motivation_Label');
    cols['T1_Commitment']       = data['t1-commit'] || '';
    cols['T1_Orgs_Considering'] = data['t1-q3']    || '';
    var behaviors = data['t1-q4'] || [];
    BEHAVIORAL_ITEMS.forEach(function(b) {
      cols['T1_Behavior_' + b.key] = behaviors.indexOf(b.full) >= 0 ? 1 : 0;
    });
    causeAreas('t1-q6', 'T1');
    cols['T1_Peer_Influence']   = data['t1-q7']        || '';
    cols['T1_Peer_Conv_YN']     = data['t1-q8-yn']     || '';
    cols['T1_Peer_Conv_Count']  = data['t1-q8-num']    || '';
    selfEfficacy('t1-se', 'T1');
    cols['T1_Career_Direction'] = data['t1-q11']       || '';
  }

  // ── T2 ────────────────────────────────────────────────────────────────────
  if (tp === 't2') {
    if (name) cols['Name'] = name;
    cols['T2_Career_Vision']     = data['t2-vision'] || '';
    cols['T2_Thinking_Shifted']  = data['t2-q3']     || '';
    careerRanks('t2-rank');
    cols['T2_Commitment_Now']    = data['t2-q1']     || '';
    cols['T2_Commitment_Then']   = data['t2-q2']     || '';
    motivation('t2-motiv-now-val',  'T2_Motivation_Now_Index',  'T2_Motivation_Now_Label');
    motivation('t2-motiv-then-val', 'T2_Motivation_Then_Index', 'T2_Motivation_Then_Label');
    cols['T2_Impactful_Sessions']  = data['t2-q4'] || '';
    causeAreas('t2-q5', 'T2');
    selfEfficacy('t2-se', 'T2');
    barriers('t2-barriers', 't2-barriers-other', null, 'T2');
    cols['T2_Placement_Readiness']  = data['t2-q6'] || '';
    cols['T2_Career_Capital_Goals'] = data['t2-q7'] || '';
    cols['T2_Peer_Influence']       = data['t2-q8'] || '';
  }

  // ── T3 / T4 ───────────────────────────────────────────────────────────────
  if (tp === 't3' || tp === 't4') {
    if (name) cols['Name'] = name;
    cols[T + '_Org_Name']  = data[tp + '-org']  || '';
    cols[T + '_Role']      = data[tp + '-role'] || '';
    cols[T + '_BARS_Intellectual_Rigor']   = data[tp + '-q1a'] || '';
    cols[T + '_BARS_Prof_Development']     = data[tp + '-q1b'] || '';
    cols[T + '_BARS_Impact_Meaningfulness']= data[tp + '-q1c'] || '';
    cols[T + '_Readiness_Retrospective']   = data[tp + '-ready'] || '';
    barriers(tp + '-barriers', tp + '-barriers-other', tp + '-barriers-detail', T);
    cols[T + '_Career_Capital_Delivery'] = data[tp + '-q2'] || '';
    cols[T + '_NPS']                     = data[tp + '-q3'] || '';
    cols[T + '_Most_Valuable']           = data[tp + '-q4'] || '';
    cols[T + '_Suggested_Improvement']   = data[tp + '-q5'] || '';
    careerRanks(tp + '-rank');
    cols[T + '_Career_Vision']           = data[tp + '-q7'] || '';
    cols[T + '_Peer_Conv_Count']         = data[tp + '-q8'] || '';
    cols[T + '_Most_Significant_Conv']   = data[tp + '-q9'] || '';
    // Roles (dynamic list — up to 5)
    var roles = data[tp + '-q10'] || [];
    var rolesCount = 0;
    for (var n = 1; n <= 5; n++) {
      var role = roles[n - 1] || {};
      cols[T + '_Role' + n + '_Name']   = role.role   || '';
      cols[T + '_Role' + n + '_Status'] = role.status || '';
      cols[T + '_Role' + n + '_Sector'] = role.sector || '';
      if (role.role) rolesCount++;
    }
    cols[T + '_Roles_Count']           = rolesCount;
    cols[T + '_Career_Direction']      = data[tp + '-q11']  || '';
    cols[T + '_Career_Dir_Factors']    = data[tp + '-q12a'] || '';
    cols[T + '_Career_Dir_Influences'] = data[tp + '-q12b'] || '';
  }

  return cols;
}

// =============================================================================
// CORE HELPERS
// =============================================================================
function normalizeEmail(email) {
  if (!email) return '';
  return String(email).trim().toLowerCase();
}

function findRowByEmail(sheet, email) {
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var emailCol = headers.indexOf('Email');
  if (emailCol === -1) return -1;
  for (var i = 1; i < data.length; i++) {
    if (normalizeEmail(data[i][emailCol]) === email) return i + 1; // 1-indexed
  }
  return -1;
}

function getLinkStatus(existingRow, headers, tp) {
  var tpNum = parseInt(tp.replace('t', ''));
  // Re-submission?
  var tsIdx = headers.indexOf('T' + tpNum + '_Submitted');
  if (tsIdx !== -1 && existingRow[tsIdx]) return 'Duplicate_overwritten';
  // Missing earlier touchpoints?
  for (var i = tpNum - 1; i >= 1; i--) {
    var prevIdx = headers.indexOf('T' + i + '_Submitted');
    if (prevIdx !== -1 && !existingRow[prevIdx]) return 'T' + tpNum + '_no_T' + i;
  }
  return '';
}

function ensureHeaders(sheet) {
  var headers = getHeaders();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    return;
  }
  var existing = sheet.getRange(1, 1, 1, 1).getValue();
  if (existing === 'Email') return; // Already set up
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
}

// =============================================================================
// WEB APP ENTRY POINTS
// =============================================================================
function doPost(e) {
  try {
    var payload = (e.parameter && e.parameter.payload)
      ? JSON.parse(e.parameter.payload)
      : JSON.parse(e.postData.contents);
    var ss      = SpreadsheetApp.getActiveSpreadsheet();

    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME, 0);

    ensureHeaders(sheet);

    var email = normalizeEmail(payload.email);
    if (!email) {
      return jsonResponse({ success: false, error: 'Email is required.' });
    }

    var tp      = payload.touchpoint;
    var colData = buildColumnData(tp, payload.data || {}, payload.name || '');
    var headers = getHeaders();

    var rowIndex = findRowByEmail(sheet, email);

    if (rowIndex === -1) {
      // ── New row ────────────────────────────────────────────────────────────
      var newRow = new Array(headers.length).fill('');
      newRow[headers.indexOf('Email')] = email;
      if (payload.name) newRow[headers.indexOf('Name')] = payload.name;

      var tpNum = parseInt(tp.replace('t', ''));
      if (tpNum > 1) {
        newRow[headers.indexOf('Link_Status')] = 'T' + tpNum + '_no_T' + (tpNum - 1);
      }

      Object.keys(colData).forEach(function(col) {
        var idx = headers.indexOf(col);
        if (idx !== -1) newRow[idx] = colData[col];
      });

      sheet.appendRow(newRow);

    } else {
      // ── Update existing row ────────────────────────────────────────────────
      var range       = sheet.getRange(rowIndex, 1, 1, headers.length);
      var existingRow = range.getValues()[0];

      existingRow[headers.indexOf('Link_Status')] = getLinkStatus(existingRow, headers, tp);
      if (payload.name) existingRow[headers.indexOf('Name')] = payload.name;

      Object.keys(colData).forEach(function(col) {
        var idx = headers.indexOf(col);
        if (idx !== -1) existingRow[idx] = colData[col];
      });

      range.setValues([existingRow]);
    }

    return jsonResponse({ success: true });

  } catch (err) {
    console.error('doPost error: ' + err.toString());
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// Health-check endpoint — returns {status:'ok'} for manual verification
function doGet(e) {
  return jsonResponse({ status: 'ok', timestamp: new Date().toISOString() });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================================================
// AUTOMATED TRIGGERS
// =============================================================================

// Runs weekly — keeps authorization fresh, produces an execution log entry
function keepAlive() {
  console.log('SMA Eval keep-alive: ' + new Date().toISOString());
}

// Runs weekly — copies the Responses tab to a dated backup tab in the same
// spreadsheet. Requires only Sheets access (no Drive permissions needed).
function weeklyBackup() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { console.log('Backup skipped — Responses sheet not found.'); return; }

  var date       = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var backupName = BACKUP_PREFIX + date;

  // Skip if today's backup already exists
  if (ss.getSheetByName(backupName)) {
    console.log('Backup already exists for today: ' + backupName);
    return;
  }

  // Copy the Responses tab and rename it
  var backup = sheet.copyTo(ss);
  backup.setName(backupName);

  // Move backup tab to the end so it stays out of the way
  ss.moveActiveSheet(ss.getSheets().length);

  // Remove backup tabs older than BACKUP_DAYS
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - BACKUP_DAYS);
  ss.getSheets().forEach(function(s) {
    var name = s.getName();
    if (name.indexOf(BACKUP_PREFIX) === 0) {
      var dateStr   = name.replace(BACKUP_PREFIX, '');
      var sheetDate = new Date(dateStr);
      if (!isNaN(sheetDate.getTime()) && sheetDate < cutoff) {
        ss.deleteSheet(s);
      }
    }
  });

  console.log('Backup tab created: ' + backupName);
}

// =============================================================================
// ONE-TIME SETUP — run this function once after pasting the script
// =============================================================================
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Responses sheet
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME, 0);
  ensureHeaders(sheet);

  // Data dictionary
  setupDataDictionary(ss);

  // Triggers
  setupTriggers();

  Logger.log('✓ Setup complete. Deploy the script as a web app and paste the URL into index.html.');
}

function setupTriggers() {
  // Remove any existing triggers for these functions to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'keepAlive' || fn === 'weeklyBackup') ScriptApp.deleteTrigger(t);
  });

  // Keep-alive: every Sunday at 9 am
  ScriptApp.newTrigger('keepAlive')
    .timeBased().everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(9).create();

  // Weekly backup: every Monday at 8 am
  ScriptApp.newTrigger('weeklyBackup')
    .timeBased().everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();

  Logger.log('✓ Triggers registered: keepAlive (Sun 9am), weeklyBackup (Mon 8am)');
}

// =============================================================================
// DATA DICTIONARY
// =============================================================================
function setupDataDictionary(ss) {
  var dictSheet = ss.getSheetByName(DICT_SHEET_NAME);
  if (dictSheet) { dictSheet.clearContents(); } else { dictSheet = ss.insertSheet(DICT_SHEET_NAME); }

  var rows = [
    ['Column Name', 'Full Question / Description', 'Touchpoint', 'Data Type', 'Valid Values / Scale'],

    // Metadata
    ['Email',       'Respondent email — unique linking key across touchpoints (normalized to lowercase)', 'All', 'Text', 'Email format'],
    ['Name',        'Respondent full name (updated at each touchpoint — most recent wins)',                'All', 'Text', 'Free text'],
    ['Link_Status', 'Data linkage quality flag. Blank = all touchpoints linked cleanly.',                  'System', 'Text',
     'Blank (clean) | T2_no_T1 | T3_no_T2 | T3_no_T1 | T4_no_T3 | T4_no_T2 | T4_no_T1 | Duplicate_overwritten'],
    ['T1_Submitted', 'Timestamp of T1 submission', 'T1', 'DateTime', 'ISO 8601'],
    ['T2_Submitted', 'Timestamp of T2 submission', 'T2', 'DateTime', 'ISO 8601'],
    ['T3_Submitted', 'Timestamp of T3 submission', 'T3', 'DateTime', 'ISO 8601'],
    ['T4_Submitted', 'Timestamp of T4 submission', 'T4', 'DateTime', 'ISO 8601'],

    // T1
    ['T1_Career_Vision',    'In 2–3 sentences, describe what a successful career looks like to you right now. (T1)', 'T1', 'Text', 'Free text'],
  ];

  CAREER_VALUES.forEach(function(v) {
    rows.push(['T1_Rank_' + v.key,
      'Career values ranking — rank given to "' + v.name + '" at T1. Rank 1 = most valued.',
      'T1', 'Integer', '1 (most valued) – 6 (least valued)']);
  });

  rows.push(
    ['T1_Motivation_Index', 'Motivational regulation type selected (T1) — numeric index. Based on Self-Determination Theory (SDT).', 'T1', 'Integer', '0=External, 1=Introjected, 2=Identified, 3=Integrated, 4=Intrinsic, 5=Fully Intrinsic'],
    ['T1_Motivation_Label', 'Motivational regulation type selected (T1) — category label.',                                           'T1', 'Text',    'External | Introjected | Identified | Integrated | Intrinsic | Fully Intrinsic'],
    ['T1_Commitment',       'How committed are you right now to pursuing a high-impact career? (T1)',                                  'T1', 'Integer', '1 (not at all committed) – 7 (fully committed)'],
    ['T1_Orgs_Considering', 'Name the organizations or roles you are most actively considering at this moment. (T1)',                  'T1', 'Text',    'Free text']
  );

  BEHAVIORAL_ITEMS.forEach(function(b) {
    rows.push(['T1_Behavior_' + b.key,
      'Done in past 12 months: "' + b.full + '" (T1)',
      'T1', 'Binary', '0 = No, 1 = Yes']);
  });

  CAUSE_AREAS.forEach(function(c) {
    rows.push(
      ['T1_Cause_' + c.key,           'Cause area familiarity — "' + c.full + '" (T1)',                 'T1', 'Text',    'Not familiar | Somewhat familiar | Actively engaged'],
      ['T1_Cause_' + c.key + '_Num',  'Cause area familiarity — "' + c.full + '" (T1) — numeric code', 'T1', 'Integer', '0=Not familiar, 1=Somewhat familiar, 2=Actively engaged']
    );
  });

  rows.push(
    ['T1_Peer_Influence',   'To what extent do you actively shape how peers think about careers and impact? (T1)', 'T1', 'Integer', '1 (never) – 7 (constantly and deliberately)'],
    ['T1_Peer_Conv_YN',     'Had any substantive peer career conversations in the past month? (T1)',               'T1', 'Text',    'Yes | No'],
    ['T1_Peer_Conv_Count',  'How many substantive peer career conversations in the past month? (T1)',              'T1', 'Integer', '0+']
  );

  SE_ITEMS.forEach(function(s) {
    rows.push(['T1_SE_' + s.key,
      'Skill confidence: "' + s.full + '" (T1)',
      'T1', 'Integer', '1 (would not know where to begin) – 7 (could do this expertly and help others)']);
  });

  rows.push(['T1_Career_Direction', 'Are you currently pursuing a high-impact career as your primary path? (T1)', 'T1', 'Text', 'Yes | No | Still deciding']);

  // T2
  rows.push(
    ['T2_Career_Vision',    'In 2–3 sentences, describe what a successful career looks like to you right now. (T2 — repeated from T1)', 'T2', 'Text', 'Free text'],
    ['T2_Thinking_Shifted', 'Has your answer to "what does a successful career look like" changed since before the orientation? (T2)',   'T2', 'Text', 'Free text']
  );

  CAREER_VALUES.forEach(function(v) {
    rows.push(['T2_Rank_' + v.key,
      'Career values ranking — rank given to "' + v.name + '" at T2 (post-orientation). Rank 1 = most valued.',
      'T2', 'Integer', '1–6']);
  });

  rows.push(
    ['T2_Commitment_Now',          'How committed are you right now to pursuing a high-impact career? (T2 — post-orientation)',                                          'T2', 'Integer', '1–7'],
    ['T2_Commitment_Then',         'Thinking back to before orientation — how committed were you then? (T2 — retrospective baseline for reference-shift correction)',    'T2', 'Integer', '1–7'],
    ['T2_Motivation_Now_Index',    'Motivational regulation — post-orientation numeric index (T2-Now). Pair with T2_Motivation_Then_Index for reference-shift analysis.','T2', 'Integer', '0–5 (see T1_Motivation_Index for scale)'],
    ['T2_Motivation_Now_Label',    'Motivational regulation — post-orientation label (T2-Now)',                                                                          'T2', 'Text',    'External | Introjected | Identified | Integrated | Intrinsic | Fully Intrinsic'],
    ['T2_Motivation_Then_Index',   'Motivational regulation — retrospective pre-orientation numeric index (T2-Then)',                                                    'T2', 'Integer', '0–5'],
    ['T2_Motivation_Then_Label',   'Motivational regulation — retrospective pre-orientation label (T2-Then)',                                                            'T2', 'Text',    'External | Introjected | Identified | Integrated | Intrinsic | Fully Intrinsic'],
    ['T2_Impactful_Sessions',      'Which one or two sessions most shifted how you think about your career, and why? (T2)',                                              'T2', 'Text',    'Free text']
  );

  CAUSE_AREAS.forEach(function(c) {
    rows.push(
      ['T2_Cause_' + c.key,           'Cause area familiarity — "' + c.full + '" (T2 — post-orientation)', 'T2', 'Text',    'Not familiar | Somewhat familiar | Actively engaged'],
      ['T2_Cause_' + c.key + '_Num',  'Cause area familiarity — "' + c.full + '" (T2) — numeric code',    'T2', 'Integer', '0=Not familiar, 1=Somewhat familiar, 2=Actively engaged']
    );
  });

  SE_ITEMS.forEach(function(s) {
    rows.push(['T2_SE_' + s.key, 'Skill confidence: "' + s.full + '" (T2 — post-orientation)', 'T2', 'Integer', '1–7']);
  });

  BARRIERS.forEach(function(b) {
    rows.push(['T2_Barrier_' + b.key,
      'Anticipated barrier before placement: "' + b.full + '" (T2)',
      'T2', 'Binary', '0 = Not selected, 1 = Selected']);
  });

  rows.push(
    ['T2_Barrier_Other_Text',    'Anticipated barrier — "Other" free-text description (T2)',                          'T2', 'Text',    'Free text'],
    ['T2_Placement_Readiness',   'How prepared do you feel to add genuine value in your fellowship placement? (T2)',   'T2', 'Integer', '1 (not at all prepared) – 7 (fully prepared)'],
    ['T2_Career_Capital_Goals',  'What do you most hope to gain from your placement in terms of career capital? (T2)', 'T2', 'Text',    'Free text'],
    ['T2_Peer_Influence',        'To what extent do you actively shape how peers think about careers and impact? (T2 — post-orientation)', 'T2', 'Integer', '1–7']
  );

  // T3 and T4 (identical structure)
  ['T3', 'T4'].forEach(function(T) {
    var tp   = T.toLowerCase();
    var wave = T === 'T3' ? 'Post-Internship (T3)' : 'Six-month follow-up (T4)';
    rows.push(
      [T + '_Org_Name',    'Internship placement — organization name (' + wave + ')',  T, 'Text', 'Free text'],
      [T + '_Role',        'Internship placement — role or title (' + wave + ')',       T, 'Text', 'Free text'],
      [T + '_BARS_Intellectual_Rigor',    'Placement quality — intellectual rigor (BARS 1–7) (' + wave + ')',       T, 'Integer', '1 (entirely routine) – 7 (most demanding experience I\'ve had)'],
      [T + '_BARS_Prof_Development',      'Placement quality — professional development value (BARS 1–7) (' + wave + ')', T, 'Integer', '1 (no meaningful gain) – 7 (transformative for career trajectory)'],
      [T + '_BARS_Impact_Meaningfulness', 'Placement quality — impact meaningfulness (BARS 1–7) (' + wave + ')',   T, 'Integer', '1 (no meaningful impact) – 7 (lasting positive impact)'],
      [T + '_Readiness_Retrospective',    'Looking back, how ready were you for this placement? (' + wave + ')',   T, 'Text',    'Yes | Somewhat | No']
    );

    BARRIERS.forEach(function(b) {
      rows.push([T + '_Barrier_' + b.key,
        'Experienced barrier during placement: "' + b.full + '" (' + wave + ')',
        T, 'Binary', '0 = Not selected, 1 = Selected']);
    });

    rows.push(
      [T + '_Barrier_Other_Text',        'Experienced barrier — "Other" free-text (' + wave + ')',                                           T, 'Text',    'Free text'],
      [T + '_Barrier_Detail',            'Additional context on barriers experienced (' + wave + ')',                                         T, 'Text',    'Free text'],
      [T + '_Career_Capital_Delivery',   'Did your placement deliver on what you hoped for in terms of career capital? (' + wave + ')',       T, 'Text',    'Free text'],
      [T + '_NPS',                       'How likely are you to recommend SMA to a peer? (' + wave + ')',                                     T, 'Integer', '0–10'],
      [T + '_Most_Valuable',             'What was the single most valuable thing about the SMA program for you? (' + wave + ')',             T, 'Text',    'Free text'],
      [T + '_Suggested_Improvement',     'What is the one thing you would change about the program? (' + wave + ')',                          T, 'Text',    'Free text']
    );

    CAREER_VALUES.forEach(function(v) {
      rows.push([T + '_Rank_' + v.key,
        'Career values ranking — rank given to "' + v.name + '" at ' + wave + '. Rank 1 = most valued.',
        T, 'Integer', '1–6']);
    });

    rows.push(
      [T + '_Career_Vision',         'In 2–3 sentences, describe what a successful career looks like to you right now. (' + wave + ')', T, 'Text',    'Free text'],
      [T + '_Peer_Conv_Count',       'Since the program ended, how many substantive peer career conversations have you had? (' + wave + ')', T, 'Integer', '0+'],
      [T + '_Most_Significant_Conv', 'Describe the most significant career conversation you had with a peer since the program ended. (' + wave + ')', T, 'Text', 'Free text']
    );

    for (var n = 1; n <= 5; n++) {
      rows.push(
        [T + '_Role' + n + '_Name',   'Roles / applications — entry ' + n + ': organization and role title (' + wave + ')',  T, 'Text', 'Free text'],
        [T + '_Role' + n + '_Status', 'Roles / applications — entry ' + n + ': application status (' + wave + ')',           T, 'Text', 'Applied | Offer received | Accepted | Declined'],
        [T + '_Role' + n + '_Sector', 'Roles / applications — entry ' + n + ': sector (' + wave + ')',                       T, 'Text', 'High-impact nonprofit/policy | Conventional consulting/finance/tech | Undecided | Other']
      );
    }

    rows.push(
      [T + '_Roles_Count',             'Number of role/application entries filled in (' + wave + ')',                                        T, 'Integer', '0–5'],
      [T + '_Career_Direction',        'Are you currently pursuing a high-impact career as your primary path? (' + wave + ')',               T, 'Text',    'Yes | No | Still deciding'],
      [T + '_Career_Dir_Factors',      'What factors most influenced your career direction? (shown when Career Direction = No) (' + wave + ')', T, 'Text', 'Free text'],
      [T + '_Career_Dir_Influences',   'What most influenced your decision to pursue this path, and what role did SMA play? (shown when Career Direction = Yes) (' + wave + ')', T, 'Text', 'Free text']
    );
  });

  dictSheet.getRange(1, 1, rows.length, 5).setValues(rows);
  dictSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1e1e1e').setFontColor('#ffffff');
  dictSheet.setFrozenRows(1);
  dictSheet.autoResizeColumns(1, 5);

  Logger.log('✓ Data dictionary created: ' + (rows.length - 1) + ' column entries.');
}
