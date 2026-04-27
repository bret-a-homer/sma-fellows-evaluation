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
  { key: 'Food_Transition',  full: 'Food transition and sustainable food systems' },
  { key: 'Democracy',        full: 'Democracy and political reform' },
  { key: 'Tax_Fairness',     full: 'Tax fairness and economic redistribution' },
  { key: 'Abundance',        full: 'Abundance and progress studies' },
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
  { key: 'ITN_Framework',      full: 'Applying the importance, tractability, and neglectedness (ITN) framework to prioritize causes and social impact work' },
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
  h.push('Email', 'Name', 'Cohort', 'Link_Status');
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
  h.push('T1_Prior_Internships', 'T1_Counterfactual');
  for (var i = 1; i <= 3; i++) h.push('T1_Path'+i+'_Desc', 'T1_Path'+i+'_Pct', 'T1_Path'+i+'_Sector');

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
    if (T === 'T4') {
      for (var j = 1; j <= 3; j++) h.push('T4_Path'+j+'_Desc', 'T4_Path'+j+'_Pct', 'T4_Path'+j+'_Sector');
      h.push('T4_Motivation_Index', 'T4_Motivation_Label');
      h.push('T4_Theory_Of_Change');
      h.push('T4_Career_Direction_Then');
    }
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
    for (var pn = 1; pn <= 3; pn++) {
      cols['T1_Path'+pn+'_Desc']   = data['t1-path'+pn+'-desc']   || '';
      cols['T1_Path'+pn+'_Pct']    = data['t1-path'+pn+'-pct']    || '';
      cols['T1_Path'+pn+'_Sector'] = data['t1-path'+pn+'-sector'] || '';
    }
    var internships = (data['t1-internships'] || []).filter(function(r) { return r.role || r.org || r.sector; });
    cols['T1_Prior_Internships'] = JSON.stringify(internships);
    var hcSectors = ['Finance', 'Technology', 'Consulting'];
    var lcSectors = ['Nonprofit / Public Service'];
    var hcCount = 0, lcCount = 0;
    internships.forEach(function(r) {
      if (hcSectors.indexOf(r.sector) >= 0) hcCount++;
      else if (lcSectors.indexOf(r.sector) >= 0) lcCount++;
    });
    if (internships.length === 0) {
      cols['T1_Counterfactual'] = 'Mid';
    } else if (hcCount === internships.length) {
      cols['T1_Counterfactual'] = 'High';
    } else if (lcCount > hcCount) {
      cols['T1_Counterfactual'] = 'Low';
    } else {
      cols['T1_Counterfactual'] = 'Mid';
    }
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
    if (tp === 't4') {
      for (var pn = 1; pn <= 3; pn++) {
        cols['T4_Path'+pn+'_Desc']   = data['t4-path'+pn+'-desc']   || '';
        cols['T4_Path'+pn+'_Pct']    = data['t4-path'+pn+'-pct']    || '';
        cols['T4_Path'+pn+'_Sector'] = data['t4-path'+pn+'-sector'] || '';
      }
      motivation('t4-motiv-val', 'T4_Motivation_Index', 'T4_Motivation_Label');
      cols['T4_Theory_Of_Change']       = data['t4-toc']        || '';
      cols['T4_Career_Direction_Then']  = data['t4-q11-then']   || '';
    }
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
  var lastCol = sheet.getLastColumn();
  var existing = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (existing[0] !== 'Email') {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    return;
  }
  // Append any columns added in later versions (e.g. Cohort)
  var missing = headers.filter(function(h) { return existing.indexOf(h) === -1; });
  if (missing.length > 0) {
    var range = sheet.getRange(1, lastCol + 1, 1, missing.length);
    range.setValues([missing]);
    range.setFontWeight('bold');
  }
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

      if (payload.cohort) newRow[headers.indexOf('Cohort')] = payload.cohort;

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
      var cohortIdx = headers.indexOf('Cohort');
      if (payload.cohort && cohortIdx !== -1 && !existingRow[cohortIdx]) {
        existingRow[cohortIdx] = payload.cohort;
      }

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

// Dashboard — served at the web app URL (no parameters = all cohorts)
function doGet(e) {
  var cohort = (e && e.parameter && e.parameter.cohort) ? e.parameter.cohort : 'all';
  var data   = computeDashboardData(cohort);

  // Always fetch the full cohort list so the dropdown works even when
  // a filtered cohort returns zero rows
  var allData     = (cohort !== 'all') ? computeDashboardData('all') : data;
  var cohortsList = (allData && allData.cohorts) ? allData.cohorts : [];

  if (!allData || allData.n === 0) {
    return HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif;padding:40px;color:#666">' +
      'No data found. Run <strong>setup()</strong> then <strong>generateSampleData()</strong> in the Apps Script editor.</p>'
    );
  }

  // If the selected cohort has no rows, pass empty-state data so the
  // dashboard renders the dropdown and an empty state rather than crashing
  if (!data || data.n === 0) {
    var emptySnap={n:0,recMean:0,careerDirT1:{Yes:0,Deciding:0,No:0},careerDirT4:{Yes:0,Deciding:0,No:0},careerDirT4Then:{Yes:0,Deciding:0,No:0},placementReady:0,convT1:0,convT4:0};
    data = { n: 0, cohorts: cohortsList, snapshot: emptySnap, snapshotHighCF: emptySnap, selfEfficacy:{items:[],t1:[],t2:[]}, commitment:{t1now:0,t2then:0,t2now:0}, motivation:{t1now:0,t2then:0,t2now:0,t4now:0,t1dist:[0,0,0,0,0],t2thenDist:[0,0,0,0,0],t2nowDist:[0,0,0,0,0],t4dist:[0,0,0,0,0],labels:['External','Introjected','Identified','Integrated','Intrinsic']}, careerValues:{names:[],t1:[],t2:[],t4:[]}, barriers:{labels:[],anticipated:[],experienced:[]}, placement:{dims:[],dist:[]} };
  }

  var tmpl = HtmlService.createTemplateFromFile('dashboard');
  tmpl.dataJson     = JSON.stringify(data);
  tmpl.cohortsList  = JSON.stringify(cohortsList);
  tmpl.selected     = cohort;
  return tmpl.evaluate()
    .setTitle('SMA Fellows Evaluation — Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function computeDashboardData(cohortFilter) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return null;

  var raw  = sheet.getDataRange().getValues();
  var hdr  = raw[0];
  var rows = raw.slice(1);
  function ci(name) { return hdr.indexOf(name); }

  rows = rows.filter(function(r) { return String(r[ci('Email')] || '').trim(); });

  var allCohorts = [];
  rows.forEach(function(r) {
    var c = String(r[ci('Cohort')] || '').trim() || 'Unknown';
    if (allCohorts.indexOf(c) < 0) allCohorts.push(c);
  });
  allCohorts.sort();

  if (cohortFilter && cohortFilter !== 'all') {
    rows = rows.filter(function(r) {
      return String(r[ci('Cohort')] || '').trim() === cohortFilter;
    });
  }

  var n = rows.length;
  if (!n) return { n: 0, cohorts: allCohorts };

  function num(r, name) {
    var v = r[ci(name)];
    return (v !== '' && v !== null && v !== undefined) ? Number(v) : null;
  }
  function avg(vals) {
    var v = vals.filter(function(x) { return x !== null && !isNaN(x); });
    return v.length ? v.reduce(function(a,b){return a+b;},0)/v.length : 0;
  }

  var t2r = rows.filter(function(r){ return r[ci('T2_Submitted')]; });
  var t3r = rows.filter(function(r){ return r[ci('T3_Submitted')]; });
  var t4r = rows.filter(function(r){ return r[ci('T4_Submitted')]; });

  // Snapshot
  var npsVals = [];
  rows.forEach(function(r){
    var v = num(r,'T4_NPS') !== null ? num(r,'T4_NPS') : num(r,'T3_NPS');
    if (v !== null) npsVals.push(v);
  });

  function dirPcts(subset, col) {
    var valid = subset.filter(function(r){ return r[ci(col)]; });
    if (!valid.length) return {Yes:0,Deciding:0,No:0};
    var tot = valid.length;
    return {
      Yes:      Math.round(100*valid.filter(function(r){return r[ci(col)]==='Yes';}).length/tot),
      Deciding: Math.round(100*valid.filter(function(r){return r[ci(col)]==='Still deciding';}).length/tot),
      No:       Math.round(100*valid.filter(function(r){return r[ci(col)]==='No';}).length/tot)
    };
  }

  var rdy = t3r.filter(function(r){return num(r,'T3_Readiness_Retrospective')!==null;});
  var t1cv = rows.filter(function(r){return r[ci('T1_Peer_Conv_YN')];});
  var t4cv = t4r.filter(function(r){return num(r,'T4_Peer_Conv_Count')!==null;});

  // SE
  var seT1 = SE_ITEMS.map(function(s){return avg(rows.map(function(r){return num(r,'T1_SE_'+s.key);}));});
  var seT2 = SE_ITEMS.map(function(s){return avg(t2r.map(function(r){return num(r,'T2_SE_'+s.key);}));});

  // Career values
  var cvT1 = CAREER_VALUES.map(function(v){return avg(rows.map(function(r){return num(r,'T1_Rank_'+v.key);}));});
  var cvT2 = CAREER_VALUES.map(function(v){return avg(t2r.map(function(r){return num(r,'T2_Rank_'+v.key);}));});
  var cvT4 = CAREER_VALUES.map(function(v){return avg(t4r.map(function(r){return num(r,'T4_Rank_'+v.key);}));});

  // Barriers
  var nonOther = BARRIERS.filter(function(b){return b.key!=='Other';});
  var bAnt = nonOther.map(function(b){
    return t2r.length ? Math.round(100*t2r.filter(function(r){return Number(r[ci('T2_Barrier_'+b.key)])===1;}).length/t2r.length) : 0;
  });
  var bExp = nonOther.map(function(b){
    return t3r.length ? Math.round(100*t3r.filter(function(r){return Number(r[ci('T3_Barrier_'+b.key)])===1;}).length/t3r.length) : 0;
  });

  // BARS
  var barsCols = ['T3_BARS_Intellectual_Rigor','T3_BARS_Prof_Development','T3_BARS_Impact_Meaningfulness'];
  var barsDist = barsCols.map(function(col){
    return [1,2,3,4,5].map(function(v){return t3r.filter(function(r){return Number(r[ci(col)])===v;}).length;});
  });

  // Motivation distributions — values are 0-indexed (0=External … 5=Fully Intrinsic)
  // Bucket into 5 display bins: 0-4, capping Fully Intrinsic (5) into bin 4
  var motT1dist=[0,0,0,0,0], motT2thenDist=[0,0,0,0,0], motT2nowDist=[0,0,0,0,0], motT4dist=[0,0,0,0,0];
  rows.forEach(function(r){
    var v=num(r,'T1_Motivation_Index');
    if(v!==null){var idx=Math.min(Math.max(Math.round(v),0),4); motT1dist[idx]++;}
  });
  t2r.forEach(function(r){
    var v=num(r,'T2_Motivation_Then_Index'); if(v!==null){var i2=Math.min(Math.max(Math.round(v),0),4); motT2thenDist[i2]++;}
    var w=num(r,'T2_Motivation_Now_Index');  if(w!==null){var i3=Math.min(Math.max(Math.round(w),0),4); motT2nowDist[i3]++;}
  });
  t4r.forEach(function(r){
    var v=num(r,'T4_Motivation_Index');
    if(v!==null){var i4=Math.min(Math.max(Math.round(v),0),4); motT4dist[i4]++;}
  });

  // High-counterfactual subset
  var hcfRows = rows.filter(function(r){ return String(r[ci('T1_Counterfactual')]||'')==='High'; });
  var hcfEmailSet = {};
  hcfRows.forEach(function(r){ hcfEmailSet[String(r[ci('Email')]||'').trim().toLowerCase()]=true; });
  function inHcf(r){ return !!hcfEmailSet[String(r[ci('Email')]||'').trim().toLowerCase()]; }
  var hcfT2r=t2r.filter(inHcf), hcfT3r=t3r.filter(inHcf), hcfT4r=t4r.filter(inHcf);
  var hcfNps=[];
  hcfRows.forEach(function(r){
    var v=num(r,'T4_NPS')!==null?num(r,'T4_NPS'):num(r,'T3_NPS');
    if(v!==null) hcfNps.push(v);
  });
  var hcfRdy=hcfT3r.filter(function(r){return num(r,'T3_Readiness_Retrospective')!==null;});
  var hcfT1cv=hcfRows.filter(function(r){return r[ci('T1_Peer_Conv_YN')];});
  var hcfT4cv=hcfT4r.filter(function(r){return num(r,'T4_Peer_Conv_Count')!==null;});

  function snapObj(sRows, sT2r, sT3r, sT4r, sNps, sRdy, sT1cv, sT4cv) {
    var dc=sT4r.length?dirPcts(sT4r,'T4_Career_Direction'):{Yes:0,Deciding:0,No:0};
    return {
      n: sRows.length,
      recMean: avg(sNps),
      careerDirT1: dirPcts(sRows,'T1_Career_Direction'),
      careerDirT4: dc,
      careerDirT4Then: dirPcts(sT4r,'T4_Career_Direction_Then'),
      placementReady: sRdy.length ? Math.round(100*sRdy.filter(function(r){return num(r,'T3_Readiness_Retrospective')>=4;}).length/sRdy.length) : 0,
      convT1: sT1cv.length ? Math.round(100*sT1cv.filter(function(r){return r[ci('T1_Peer_Conv_YN')]==='Yes';}).length/sT1cv.length) : 0,
      convT4: sT4cv.length ? Math.round(100*sT4cv.filter(function(r){return num(r,'T4_Peer_Conv_Count')>0;}).length/sT4cv.length) : 0
    };
  }

  return {
    n: n, cohorts: allCohorts,
    snapshot: snapObj(rows, t2r, t3r, t4r, npsVals, rdy, t1cv, t4cv),
    snapshotHighCF: snapObj(hcfRows, hcfT2r, hcfT3r, hcfT4r, hcfNps, hcfRdy, hcfT1cv, hcfT4cv),
    selfEfficacy: { items: SE_ITEMS.map(function(s){return s.full;}), t1: seT1, t2: seT2 },
    commitment: {
      t1now:  avg(rows.map(function(r){return num(r,'T1_Commitment');})),
      t2then: avg(t2r.map(function(r){return num(r,'T2_Commitment_Then');})),
      t2now:  avg(t2r.map(function(r){return num(r,'T2_Commitment_Now');}))
    },
    motivation: {
      t1now:  avg(rows.map(function(r){return num(r,'T1_Motivation_Index');})),
      t2then: avg(t2r.map(function(r){return num(r,'T2_Motivation_Then_Index');})),
      t2now:  avg(t2r.map(function(r){return num(r,'T2_Motivation_Now_Index');})),
      t4now:  avg(t4r.map(function(r){return num(r,'T4_Motivation_Index');})),
      t1dist: motT1dist, t2thenDist: motT2thenDist, t2nowDist: motT2nowDist, t4dist: motT4dist,
      labels: ['External','Introjected','Identified','Integrated','Intrinsic']
    },
    careerValues: { names: CAREER_VALUES.map(function(v){return v.name;}), t1:cvT1, t2:cvT2, t4:cvT4 },
    barriers: { labels: nonOther.map(function(b){return b.full;}), anticipated:bAnt, experienced:bExp },
    placement: { dims: ['Intellectual Rigor','Professional Development','Impact & Meaningfulness'], dist: barsDist }
  };
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

  // Derived analysis tabs
  setupDerivedTabs(ss);

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
// SAMPLE DATA GENERATOR
// Run generateSampleData() once to populate a realistic 2024-25 cohort.
// WARNING: appends rows — clear the Responses sheet first if needed.
// =============================================================================
function generateSampleData() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('Run setup() first.'); return; }

  var h = getHeaders();
  function set(row, col, val) { var i = h.indexOf(col); if (i >= 0) row[i] = val; }
  function setRanks(row, tp, order) {
    CAREER_VALUES.forEach(function(v) {
      var r = order.indexOf(v.key) + 1;
      set(row, tp + '_Rank_' + v.key, r > 0 ? r : 6);
    });
  }
  function setSE(row, tp, scores) {
    SE_ITEMS.forEach(function(s, i) { set(row, tp + '_SE_' + s.key, scores[i] || 4); });
  }
  function setBarriers(row, tp, keys) {
    BARRIERS.forEach(function(b) { set(row, tp + '_Barrier_' + b.key, keys.indexOf(b.key) >= 0 ? 1 : 0); });
  }

  var cohort = '2024-25';
  var placements = [
    {org:'GiveWell',                         role:'Research Analyst'},
    {org:'Open Philanthropy',                role:'Program Associate'},
    {org:'Centre for Effective Altruism',    role:'Operations Fellow'},
    {org:'80,000 Hours',                     role:'Research Fellow'},
    {org:'Animal Charity Evaluators',        role:'Research Analyst'},
    {org:'Founders Pledge',                  role:'Research Associate'},
    {org:'Rethink Priorities',               role:'Policy Research Fellow'},
    {org:'Open Philanthropy',                role:'Biosecurity Fellow'},
    {org:'GiveWell',                         role:'Analyst'},
    {org:'Future of Humanity Institute',     role:'Research Intern'},
    {org:'Longview Philanthropy',            role:'Research Associate'},
    {org:'Open Philanthropy',                role:'AI Safety Fellow'},
    {org:'Centre for Effective Altruism',    role:'Community Building Fellow'},
  ];

  // Columns: name, email, t1Commit, t1Motiv, t1Dir, t2cNow, t2cThen, t2mNow, t2mThen, t4Dir, nps, barsIR, barsPD, barsIM, placIdx, t4Motiv
  var fellows = [
    ['Emma Chen',       'e.chen@harvard.edu',       6,3,'Yes',           6,5,4,3,'Yes',           9, 4,4,5, 0, 4],
    ['James Osei',      'j.osei@mit.edu',            4,2,'Still deciding',5,3,3,2,'Yes',           8, 4,3,4, 1, 4],
    ['Priya Sharma',    'p.sharma@yale.edu',         7,4,'Yes',           7,6,5,4,'Yes',           10,5,5,5, 2, 5],
    ['Marcus Williams', 'm.williams@columbia.edu',   5,3,'Yes',           6,4,4,3,'Yes',           8, 4,3,4, 3, 4],
    ['Sofia Rodriguez', 's.rodriguez@brown.edu',     3,1,'Still deciding',5,3,3,2,'Yes',           7, 3,3,4, 4, 3],
    ['Liam Nakamura',   'l.nakamura@stanford.edu',   6,3,'Yes',           6,5,4,3,'Yes',           9, 5,4,5, 5, 4],
    ['Aisha Patel',     'a.patel@princeton.edu',     7,4,'Yes',           7,6,5,4,'Yes',           9, 5,5,5, 6, 5],
    ['Tyler Brooks',    't.brooks@upenn.edu',        3,2,'No',            5,3,3,2,'Still deciding',6, 3,3,3, 7, 3],
    ['Maya Johnson',    'm.johnson@dartmouth.edu',   6,3,'Yes',           6,5,4,3,'Yes',           8, 4,4,5, 8, 4],
    ['Noah Schmidt',    'n.schmidt@cornell.edu',     5,2,'Still deciding',5,4,3,2,'Yes',           7, 4,3,4, 9, 3],
    ['Isabella Torres', 'i.torres@uchicago.edu',     6,4,'Yes',           7,5,5,4,'Yes',           10,5,5,5,10, 5],
    ['Ethan Kim',       'e.kim@duke.edu',            4,2,'Still deciding',5,3,3,2,'Still deciding',7, 3,3,4,11, 3],
    ['Ava Thompson',    'a.thompson@vanderbilt.edu', 7,4,'Yes',           7,6,5,4,'Yes',           9, 5,5,5,12, 5],
    ['Oliver Davis',    'o.davis@northwestern.edu',  5,3,'Still deciding',5,4,3,3,'Still deciding',6, 4,3,4, 0, 3],
    ['Mia Anderson',    'm.anderson@georgetown.edu', 4,2,'No',            6,3,4,2,'Yes',           8, 4,4,4, 1, 4],
  ];

  var t1SE = [
    [5,4,4,4,5],[3,3,3,3,4],[6,6,5,5,6],[4,4,4,4,5],[2,3,3,3,3],
    [5,5,4,4,5],[6,5,5,5,6],[3,3,3,3,3],[5,4,4,5,5],[4,3,3,4,4],
    [5,5,5,5,6],[3,4,3,4,4],[6,6,5,5,6],[4,4,4,4,4],[3,3,4,3,4]
  ];

  var valueOrdersT1 = [
    ['Altruism','Intellectual_Stimulation','Achievement','Independence','Economic_Return','Prestige'],
    ['Altruism','Achievement','Intellectual_Stimulation','Independence','Economic_Return','Prestige'],
    ['Altruism','Intellectual_Stimulation','Independence','Achievement','Prestige','Economic_Return'],
    ['Altruism','Achievement','Independence','Intellectual_Stimulation','Economic_Return','Prestige'],
    ['Economic_Return','Altruism','Achievement','Intellectual_Stimulation','Independence','Prestige'],
    ['Altruism','Intellectual_Stimulation','Achievement','Independence','Prestige','Economic_Return'],
    ['Altruism','Independence','Intellectual_Stimulation','Achievement','Economic_Return','Prestige'],
    ['Economic_Return','Prestige','Achievement','Independence','Intellectual_Stimulation','Altruism'],
    ['Altruism','Achievement','Intellectual_Stimulation','Independence','Prestige','Economic_Return'],
    ['Intellectual_Stimulation','Altruism','Achievement','Independence','Economic_Return','Prestige'],
    ['Altruism','Intellectual_Stimulation','Independence','Achievement','Economic_Return','Prestige'],
    ['Economic_Return','Achievement','Independence','Intellectual_Stimulation','Altruism','Prestige'],
    ['Altruism','Intellectual_Stimulation','Achievement','Independence','Economic_Return','Prestige'],
    ['Achievement','Intellectual_Stimulation','Independence','Altruism','Economic_Return','Prestige'],
    ['Altruism','Economic_Return','Achievement','Intellectual_Stimulation','Independence','Prestige'],
  ];
  // T4: Altruism moves to #1, Economic_Return moves toward bottom
  var valueOrdersT4 = valueOrdersT1.map(function(r) {
    var c = r.slice();
    var ai = c.indexOf('Altruism'); if (ai > 0) { c.splice(ai,1); c.unshift('Altruism'); }
    var ei = c.indexOf('Economic_Return'); if (ei >= 0 && ei < 4) { c.splice(ei,1); c.push('Economic_Return'); }
    return c;
  });

  var anticipatedB = [
    ['Cause_Area_Uncertainty','Skills_Credentials'],
    ['Financial_Security','Cause_Area_Uncertainty','Skills_Credentials'],
    ['Skills_Credentials'],
    ['Cause_Area_Uncertainty','Skills_Credentials','Policy_Landscape'],
    ['Financial_Security','Family_Peer_Pressure','Cause_Area_Uncertainty'],
    ['Cause_Area_Uncertainty','Skills_Credentials'],
    ['Skills_Credentials','Policy_Landscape'],
    ['Financial_Security','Family_Peer_Pressure','Cause_Area_Uncertainty'],
    ['Cause_Area_Uncertainty'],
    ['Financial_Security','Cause_Area_Uncertainty','Skills_Credentials'],
    ['Skills_Credentials'],
    ['Financial_Security','Family_Peer_Pressure'],
    ['Cause_Area_Uncertainty','Skills_Credentials'],
    ['Financial_Security','Cause_Area_Uncertainty'],
    ['Financial_Security','Family_Peer_Pressure','Cause_Area_Uncertainty'],
  ];
  var experiencedB = [
    ['Ambiguity','Work_Life_Balance','Managing_Up'],
    ['Ambiguity','Work_Life_Balance','Org_Culture'],
    ['Work_Life_Balance','Quality_Work'],
    ['Ambiguity','Policy_Landscape','Managing_Up'],
    ['Financial_Security','Ambiguity','Work_Life_Balance'],
    ['Ambiguity','Senior_Stakeholders'],
    ['Policy_Landscape','Ambiguity','Work_Life_Balance'],
    ['Work_Life_Balance','Ambiguity','Managing_Up'],
    ['Ambiguity','Work_Life_Balance'],
    ['Ambiguity','Work_Life_Balance','Org_Culture'],
    ['Quality_Work','Work_Life_Balance'],
    ['Ambiguity','Work_Life_Balance','Financial_Security'],
    ['Work_Life_Balance','Ambiguity'],
    ['Ambiguity','Work_Life_Balance','Managing_Up'],
    ['Ambiguity','Financial_Security','Work_Life_Balance'],
  ];

  var careerVisions = [
    'I want to work at the intersection of policy and research to address global health challenges at scale.',
    'My goal is to build a career in AI safety research, helping ensure advanced AI systems remain beneficial.',
    'I see myself leading programs at an EA organization focused on global health and development.',
    'I want to work on biosecurity policy, helping governments prepare for pandemic risks.',
    'I am still exploring, but I am drawn to paths that combine analytical rigor with meaningful impact.',
    'I want to contribute to AI governance and help shape policy for safe and beneficial AI development.',
    'My vision is to work in grantmaking, directing significant resources toward the most effective interventions.',
    'I want to pursue a career in finance that allows me to earn and give effectively.',
    'I see myself as a community builder, helping grow the EA ecosystem and supporting others.',
    'I want to do research on animal welfare, particularly around improving conditions in factory farming.',
    'My goal is to work in climate policy, bridging the gap between scientific research and policy action.',
    'I want to be an entrepreneur working on a solution to a neglected global problem.',
    'I see myself working in international development, using rigorous evidence to improve program effectiveness.',
    'I am interested in earning to give while building skills, with a long-term goal in AI safety.',
    'My vision is to work in public health, particularly on malaria prevention in Sub-Saharan Africa.',
  ];
  var orgsConsidering = [
    'GiveWell, Open Philanthropy, WHO',
    'OpenAI, DeepMind, Center for Human-Compatible AI',
    'GiveWell, Against Malaria Foundation, Open Philanthropy',
    'Johns Hopkins Center for Health Security, Georgetown GHSS',
    'Still exploring — interested in 80,000 Hours, EA Funds',
    'Partnership on AI, Centre for the Governance of AI, Open Philanthropy',
    'Open Philanthropy, Founders Pledge, Longview Philanthropy',
    'Jane Street, Citadel — planning to earn to give',
    'Centre for Effective Altruism, EA Infrastructure Fund',
    'Animal Charity Evaluators, Humane Society, Good Food Institute',
    'ClimateWorks, Founders Pledge Climate Fund',
    'Wave, WorldCover, Kenya Red Cross',
    'GiveDirectly, Innovations for Poverty Action, J-PAL',
    'Currently in finance — considering transition to direct work',
    'Partners in Health, MSF, Against Malaria Foundation',
  ];
  var mostValuable = [
    'The exposure to different cause areas and time to think carefully about where I could have the most impact.',
    'Building relationships with peers who share similar values and career goals.',
    'The structured curriculum on EA principles and how to evaluate impact rigorously.',
    'Learning frameworks for thinking about counterfactual impact and career capital.',
    'The placement itself — seeing how EA principles apply in a real organizational context.',
    'Guest speakers from leading EA organizations who shared candid insights about their work.',
    'The orientation sessions on motivation and career values — helped me clarify what I actually care about.',
    'Networking opportunities and meeting people already working on high-impact problems.',
    'The self-reflection exercises, especially around motivation type and what drives my career choices.',
    'Having dedicated time to think seriously about career strategy outside of academic pressures.',
    'The combination of intellectual rigor in the curriculum and practical placement experience.',
    'Peer discussions about cause area prioritization — challenged my assumptions in useful ways.',
    'The mentorship component and access to senior practitioners in the EA space.',
    'Learning how to evaluate career opportunities from an impact perspective rather than prestige.',
    'The placement experience — seeing the gap between theory and practice in a real organization.',
  ];
  var improvements = [
    'More time dedicated to exploring specific cause areas before the placement.',
    'A structured mentorship program would add a lot of value.',
    'More preparation for the practical challenges of workplace culture.',
    'Stronger connections between orientation content and the placement experience.',
    'More diverse placement options across different cause areas.',
    'Earlier placement matching to allow more time to prepare.',
    'More alumni engagement — hearing from people 2-3 years out would be valuable.',
    'Better support for fellows uncertain about cause area prioritization.',
    'Additional sessions on managing up and navigating organizational dynamics.',
    'More structured reflection time during the placement phase.',
    'A peer buddy system to support fellows during placement.',
    'More concrete resources for career planning after the fellowship ends.',
    'Clearer guidance on what to do if the placement is not a good fit.',
    'More opportunities to connect with fellows from different cohorts.',
    'A session on the transition from academic to professional environments.',
  ];

  fellows.forEach(function(f, i) {
    var row = new Array(h.length).fill('');
    var p = placements[f[14] % placements.length];
    var t2SE = t1SE[i].map(function(v){ return Math.min(7, v+1); });

    set(row,'Email',   f[1]);
    set(row,'Name',    f[0]);
    set(row,'Cohort',  cohort);
    set(row,'T1_Submitted', new Date('2024-09-15').toISOString());
    if (i < 13) set(row,'T2_Submitted', new Date('2024-10-20').toISOString());
    if (i < 12) set(row,'T3_Submitted', new Date('2025-01-15').toISOString());
    if (i < 11) set(row,'T4_Submitted', new Date('2025-03-20').toISOString());

    // T1
    set(row,'T1_Career_Vision',    careerVisions[i]);
    set(row,'T1_Orgs_Considering', orgsConsidering[i]);
    setRanks(row,'T1', valueOrdersT1[i]);
    set(row,'T1_Motivation_Index', f[3]);
    set(row,'T1_Motivation_Label', MOTIVATION_LABELS[f[3]]);
    set(row,'T1_Commitment',       f[2]);
    setSE(row,'T1', t1SE[i]);
    set(row,'T1_Career_Direction', f[4]);
    var ctfSample = ['High','Low','Mid','Mid','High','Low','High','Mid','Low','Mid'];
    set(row,'T1_Counterfactual',   ctfSample[i % ctfSample.length]);
    set(row,'T1_Prior_Internships','[]');
    set(row,'T1_Peer_Influence',   3 + (i % 4));
    set(row,'T1_Peer_Conv_YN',     i % 3 === 0 ? 'No' : 'Yes');
    set(row,'T1_Peer_Conv_Count',  i % 3 === 0 ? 0 : 2 + (i % 4));
    ['EA_Event','EA_Content','Career_Convo','Applied'].forEach(function(b,bi){
      set(row,'T1_Behavior_'+b, (i+bi)%3===0 ? 0 : 1);
    });
    CAUSE_AREAS.forEach(function(c,ci){
      var vals = ['Actively engaged','Somewhat familiar','Not familiar'];
      var v = vals[(ci+i)%3];
      set(row,'T1_Cause_'+c.key,     v);
      set(row,'T1_Cause_'+c.key+'_Num', CAUSE_NUM[v]);
    });

    // T2
    if (i < 13) {
      set(row,'T2_Career_Vision',      careerVisions[i]+' — updated after orientation.');
      set(row,'T2_Thinking_Shifted',   'Orientation sharpened my thinking on cause prioritization and scale. I feel more confident about where to focus.');
      set(row,'T2_Impactful_Sessions', 'The sessions on cause prioritization and the guest speaker from Open Philanthropy were the highlights.');
      setRanks(row,'T2', valueOrdersT1[i]);
      set(row,'T2_Commitment_Now',  f[5]);
      set(row,'T2_Commitment_Then', f[6]);
      set(row,'T2_Motivation_Now_Index',  f[7]);
      set(row,'T2_Motivation_Now_Label',  MOTIVATION_LABELS[f[7]]);
      set(row,'T2_Motivation_Then_Index', f[8]);
      set(row,'T2_Motivation_Then_Label', MOTIVATION_LABELS[f[8]]);
      setSE(row,'T2', t2SE);
      CAUSE_AREAS.forEach(function(c,ci){
        var vals = ['Actively engaged','Somewhat familiar','Not familiar'];
        var v = vals[Math.max(0,(ci+i)%3-1)] || 'Somewhat familiar';
        set(row,'T2_Cause_'+c.key,     v);
        set(row,'T2_Cause_'+c.key+'_Num', CAUSE_NUM[v]);
      });
      setBarriers(row,'T2', anticipatedB[i]);
      set(row,'T2_Placement_Readiness',  3+(i%3));
      set(row,'T2_Career_Capital_Goals', 'Build relationships with practitioners, develop research skills, and clarify which organizations are doing the most effective work.');
      set(row,'T2_Peer_Influence',       4+(i%3));
    }

    // T3
    if (i < 12) {
      set(row,'T3_Org_Name',  p.org);
      set(row,'T3_Role',      p.role);
      set(row,'T3_BARS_Intellectual_Rigor',    f[11]);
      set(row,'T3_BARS_Prof_Development',      f[12]);
      set(row,'T3_BARS_Impact_Meaningfulness', f[13]);
      set(row,'T3_Readiness_Retrospective',    [4,3,5,4,3,4,5,2,4,3,5,3][i] || 4);
      setBarriers(row,'T3', experiencedB[i]);
      set(row,'T3_Career_Capital_Delivery', 4+(i%2));
      set(row,'T3_NPS',                     f[10]);
      set(row,'T3_Most_Valuable',           mostValuable[i]);
      set(row,'T3_Suggested_Improvement',   improvements[i]);
      setRanks(row,'T3', valueOrdersT4[i]);
      set(row,'T3_Career_Vision',     careerVisions[i]+' — strengthened by the placement.');
      set(row,'T3_Peer_Conv_Count',   3+(i%4));
      set(row,'T3_Most_Significant_Conv', 'A conversation with my supervisor about long-term career strategy and the transition from fellowship to full-time work.');
      set(row,'T3_Role1_Name',   p.role);
      set(row,'T3_Role1_Status', 'Actively applying');
      set(row,'T3_Role1_Sector', 'Nonprofit / EA');
      set(row,'T3_Roles_Count',  1+(i%3));
      set(row,'T3_Career_Direction',      f[9]);
      set(row,'T3_Career_Dir_Factors',    'The quality of the people at my placement org and the intellectual rigor of the work.');
      set(row,'T3_Career_Dir_Influences', 'Conversations with my supervisor and seeing the direct impact of the organization\'s work.');
    }

    // T4
    if (i < 11) {
      set(row,'T4_Org_Name',  p.org);
      set(row,'T4_Role',      p.role);
      set(row,'T4_BARS_Intellectual_Rigor',    Math.min(5,f[11]+(i%2)));
      set(row,'T4_BARS_Prof_Development',      Math.min(5,f[12]+(i%2)));
      set(row,'T4_BARS_Impact_Meaningfulness', f[13]);
      set(row,'T4_Readiness_Retrospective',    Math.min(5,([4,3,5,4,3,4,5,2,4,3,5][i]||4)+1));
      setBarriers(row,'T4', experiencedB[i].slice(0,1));
      set(row,'T4_Career_Capital_Delivery', Math.min(5,4+(i%2)));
      set(row,'T4_NPS',                     f[10]);
      set(row,'T4_Most_Valuable',           mostValuable[i]);
      set(row,'T4_Suggested_Improvement',   improvements[i]);
      setRanks(row,'T4', valueOrdersT4[i]);
      set(row,'T4_Career_Vision',     careerVisions[i]+' — now focused on near-term high-impact work at an EA-aligned organization.');
      set(row,'T4_Peer_Conv_Count',   4+(i%4));
      set(row,'T4_Most_Significant_Conv', 'A career strategy session with an 80,000 Hours advisor that helped clarify my comparative advantage and next steps.');
      set(row,'T4_Role1_Name',   p.role+' (continued)');
      set(row,'T4_Role1_Status', i%3===0 ? 'Offer accepted' : 'Actively applying');
      set(row,'T4_Role1_Sector', 'Nonprofit / EA');
      set(row,'T4_Roles_Count',  2+(i%2));
      set(row,'T4_Career_Direction',      f[9]);
      set(row,'T4_Career_Dir_Factors',    'The quality of the organization\'s work, the people I\'d be working with, and the clarity of their theory of change.');
      set(row,'T4_Career_Dir_Influences', 'My placement experience, conversations with mentors, and clearer thinking about where I can have the most impact.');
      set(row,'T4_Motivation_Index',  f[15]);
      set(row,'T4_Motivation_Label',  MOTIVATION_LABELS[f[15]]);
    }

    sheet.appendRow(row);
  });

  Logger.log('✓ Generated 15 sample fellows for cohort ' + cohort + '. Run setupDerivedTabs() to refresh the analysis tabs.');
}

// =============================================================================
// DERIVED ANALYSIS TABS
// Run setupDerivedTabs() any time to create or rebuild these tabs.
// Safe to re-run — clears and rebuilds without touching Responses data.
// =============================================================================

function colLetter(n) {
  var s = '';
  while (n > 0) {
    var rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function setupDerivedTabs(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  setupCompletionTab(ss);
  setupQualitativeTab(ss);
  setupFellowLookupTab(ss);
  buildAnalyticsTab();
  Logger.log('✓ Derived tabs created: Completion, Qualitative Responses, Fellow Lookup, Analytics, Explore');
}

function setupCompletionTab(ss) {
  var tabName = 'Completion';
  var tab = ss.getSheetByName(tabName);
  if (tab) { tab.clearContents(); tab.clearFormats(); } else { tab = ss.insertSheet(tabName); }

  var h  = getHeaders();
  var eA  = colLetter(h.indexOf('Email')        + 1);
  var eB  = colLetter(h.indexOf('Name')         + 1);
  var eC  = colLetter(h.indexOf('Cohort')       + 1);
  var eD  = colLetter(h.indexOf('Link_Status')  + 1);
  var eT1 = colLetter(h.indexOf('T1_Submitted') + 1);
  var eT2 = colLetter(h.indexOf('T2_Submitted') + 1);
  var eT3 = colLetter(h.indexOf('T3_Submitted') + 1);
  var eT4 = colLetter(h.indexOf('T4_Submitted') + 1);

  tab.getRange('A1').setValue('Completion Tracker').setFontSize(14).setFontWeight('bold');
  tab.getRange('A2').setValue('Auto-refreshes as responses arrive.  Blank submitted date = not yet submitted.');

  tab.getRange('A4:A8').setFontWeight('bold');
  tab.getRange('A4').setValue('Total fellows');
  tab.getRange('B4').setFormula('=COUNTA(Responses!' + eA + '2:' + eA + ')');

  tab.getRange('A5').setValue('T1 complete');
  tab.getRange('B5').setFormula('=COUNTIF(Responses!' + eT1 + '2:' + eT1 + ',"<>")');
  tab.getRange('C5').setFormula('=IFERROR(TEXT(B5/B4,"0%"),"—")');

  tab.getRange('A6').setValue('T2 complete');
  tab.getRange('B6').setFormula('=COUNTIF(Responses!' + eT2 + '2:' + eT2 + ',"<>")');
  tab.getRange('C6').setFormula('=IFERROR(TEXT(B6/B4,"0%"),"—")');

  tab.getRange('A7').setValue('T3 complete');
  tab.getRange('B7').setFormula('=COUNTIF(Responses!' + eT3 + '2:' + eT3 + ',"<>")');
  tab.getRange('C7').setFormula('=IFERROR(TEXT(B7/B4,"0%"),"—")');

  tab.getRange('A8').setValue('T4 complete');
  tab.getRange('B8').setFormula('=COUNTIF(Responses!' + eT4 + '2:' + eT4 + ',"<>")');
  tab.getRange('C8').setFormula('=IFERROR(TEXT(B8/B4,"0%"),"—")');

  tab.getRange('A10').setValue('Fellow-by-Fellow Status').setFontSize(12).setFontWeight('bold');
  var hdrRow = tab.getRange('A11:H11');
  hdrRow.setValues([['Email','Name','Cohort','T1 Submitted','T2 Submitted','T3 Submitted','T4 Submitted','Link Status']]);
  hdrRow.setFontWeight('bold').setBackground('#f0f0f0');
  tab.setFrozenRows(11);

  var maxCol = Math.max(h.indexOf('T4_Submitted'), h.indexOf('Link_Status')) + 1;
  var rangeRef = 'Responses!' + eA + ':' + colLetter(maxCol);
  var sel = [eA,eB,eC,eT1,eT2,eT3,eT4,eD].join(', ');
  var lbl = [eA,eB,eC,eT1,eT2,eT3,eT4,eD].map(function(c){return c+" ''";}).join(', ');
  tab.getRange('A12').setFormula(
    '=IFERROR(QUERY(' + rangeRef + ',"SELECT ' + sel +
    ' WHERE ' + eA + " <> '' ORDER BY " + eC + ', ' + eA +
    ' LABEL ' + lbl + '",0),"No data yet")'
  );

  tab.setColumnWidth(1, 200);
  tab.setColumnWidth(2, 140);
  tab.setColumnWidth(3, 80);
  tab.setColumnWidths(4, 4, 165);
  tab.setColumnWidth(8, 160);
}

function setupQualitativeTab(ss) {
  var tabName = 'Qualitative Responses';
  var tab = ss.getSheetByName(tabName);
  if (tab) { tab.clearContents(); tab.clearFormats(); } else { tab = ss.insertSheet(tabName); }

  var h = getHeaders();
  var qualFields = [
    { col: 'T1_Career_Vision',          label: 'T1: Career vision' },
    { col: 'T1_Orgs_Considering',       label: 'T1: Organizations / roles considering' },
    { col: 'T2_Career_Vision',          label: 'T2: Career vision (updated)' },
    { col: 'T2_Thinking_Shifted',       label: 'T2: How thinking has shifted' },
    { col: 'T2_Impactful_Sessions',     label: 'T2: Most impactful sessions' },
    { col: 'T2_Career_Capital_Goals',   label: 'T2: Career capital goals' },
    { col: 'T2_Barrier_Other_Text',     label: 'T2: Barriers — other (free text)' },
    { col: 'T3_Org_Name',              label: 'T3: Organization name' },
    { col: 'T3_Role',                  label: 'T3: Role / position' },
    { col: 'T3_Most_Valuable',         label: 'T3: Most valuable aspect' },
    { col: 'T3_Suggested_Improvement', label: 'T3: Suggested improvement' },
    { col: 'T3_Career_Vision',         label: 'T3: Career vision (updated)' },
    { col: 'T3_Barrier_Detail',        label: 'T3: Barrier — detail' },
    { col: 'T3_Most_Significant_Conv', label: 'T3: Most significant conversation' },
    { col: 'T3_Career_Dir_Factors',    label: 'T3: Career direction — factors' },
    { col: 'T3_Career_Dir_Influences', label: 'T3: Career direction — influences' },
    { col: 'T4_Org_Name',              label: 'T4: Organization name' },
    { col: 'T4_Role',                  label: 'T4: Role / position' },
    { col: 'T4_Most_Valuable',         label: 'T4: Most valuable aspect' },
    { col: 'T4_Suggested_Improvement', label: 'T4: Suggested improvement' },
    { col: 'T4_Career_Vision',         label: 'T4: Career vision (updated)' },
    { col: 'T4_Barrier_Detail',        label: 'T4: Barrier — detail' },
    { col: 'T4_Most_Significant_Conv', label: 'T4: Most significant conversation' },
    { col: 'T4_Career_Dir_Factors',    label: 'T4: Career direction — factors' },
    { col: 'T4_Career_Dir_Influences', label: 'T4: Career direction — influences' },
  ];

  var eA = colLetter(h.indexOf('Email')  + 1);
  var eB = colLetter(h.indexOf('Name')   + 1);
  var eC = colLetter(h.indexOf('Cohort') + 1);

  var fieldCols = qualFields.map(function(f) {
    var idx = h.indexOf(f.col);
    return idx >= 0 ? colLetter(idx + 1) : null;
  });

  var headerVals = ['Email','Name','Cohort'].concat(qualFields.map(function(f){return f.label;}));
  var hdrRange = tab.getRange(1, 1, 1, headerVals.length);
  hdrRange.setValues([headerVals]).setFontWeight('bold').setBackground('#f0f0f0').setWrap(true);
  tab.setRowHeight(1, 60);
  tab.setFrozenRows(1);
  tab.setFrozenColumns(3);

  var maxIdx = Math.max.apply(null, qualFields.map(function(f){ return h.indexOf(f.col) + 1; }));
  var rangeRef = 'Responses!' + eA + ':' + colLetter(maxIdx);
  var allCols = [eA, eB, eC].concat(fieldCols.filter(function(c){return c;}));
  var sel = allCols.join(', ');
  var lbl = allCols.map(function(c){return c+" ''";}).join(', ');

  tab.getRange('A2').setFormula(
    '=IFERROR(QUERY(' + rangeRef + ',"SELECT ' + sel +
    ' WHERE ' + eA + " <> '' ORDER BY " + eC + ', ' + eA +
    ' LABEL ' + lbl + '",0),"No responses yet")'
  );

  tab.setColumnWidth(1, 180);
  tab.setColumnWidth(2, 130);
  tab.setColumnWidth(3, 80);
  for (var i = 4; i <= headerVals.length; i++) tab.setColumnWidth(i, 220);
}

function setupFellowLookupTab(ss) {
  var tabName = 'Fellow Lookup';
  var tab = ss.getSheetByName(tabName);
  if (tab) { tab.clearContents(); tab.clearFormats(); } else { tab = ss.insertSheet(tabName); }

  var h = getHeaders();
  function fml(colName) {
    var idx = h.indexOf(colName);
    if (idx < 0) return '';
    var c = colLetter(idx + 1);
    return '=IFERROR(INDEX(Responses!' + c + ':' + c + ',MATCH($B$1,Responses!A:A,0)),"")';
  }

  tab.getRange('A1').setValue('Enter email to look up:').setFontWeight('bold');
  tab.getRange('B1').setBackground('#fff9c4')
    .setBorder(true,true,true,true,false,false);

  var row = 3;
  function section(title) {
    if (row > 3) row++;
    var r = tab.getRange(row, 1, 1, 2);
    r.merge().setValue(title).setFontWeight('bold').setFontSize(11).setBackground('#e8e8e8');
    row++;
  }
  function field(label, colName) {
    tab.getRange(row, 1).setValue(label);
    if (colName) tab.getRange(row, 2).setFormula(fml(colName));
    row++;
  }

  section('Identity & Status');
  field('Name',         'Name');
  field('Cohort',       'Cohort');
  field('Link Status',  'Link_Status');
  field('T1 Submitted', 'T1_Submitted');
  field('T2 Submitted', 'T2_Submitted');
  field('T3 Submitted', 'T3_Submitted');
  field('T4 Submitted', 'T4_Submitted');

  section('T1 — Pre-Orientation');
  field('Career vision',     'T1_Career_Vision');
  field('Orgs considering',  'T1_Orgs_Considering');
  field('Commitment (1–7)',  'T1_Commitment');
  field('Motivation type',   'T1_Motivation_Label');
  field('Career direction',  'T1_Career_Direction');
  SE_ITEMS.forEach(function(s) { field('SE — ' + s.full, 'T1_SE_' + s.key); });

  section('T2 — Post-Orientation');
  field('Career vision',            'T2_Career_Vision');
  field('How thinking shifted',     'T2_Thinking_Shifted');
  field('Commitment now (1–7)',     'T2_Commitment_Now');
  field('Commitment then (1–7)',    'T2_Commitment_Then');
  field('Motivation now',           'T2_Motivation_Now_Label');
  field('Motivation then',          'T2_Motivation_Then_Label');
  field('Most impactful sessions',  'T2_Impactful_Sessions');
  field('Career capital goals',     'T2_Career_Capital_Goals');

  section('T3 — Mid-Placement');
  field('Organization',              'T3_Org_Name');
  field('Role',                      'T3_Role');
  field('BARS — Intellectual rigor (1–5)',     'T3_BARS_Intellectual_Rigor');
  field('BARS — Prof. development (1–5)',      'T3_BARS_Prof_Development');
  field('BARS — Impact & meaning (1–5)',       'T3_BARS_Impact_Meaningfulness');
  field('Felt ready for placement',  'T3_Readiness_Retrospective');
  field('NPS (0–10)',                'T3_NPS');
  field('Most valuable aspect',      'T3_Most_Valuable');
  field('Suggested improvement',     'T3_Suggested_Improvement');
  field('Career vision',             'T3_Career_Vision');
  field('Career direction',          'T3_Career_Direction');
  field('Most significant conversation', 'T3_Most_Significant_Conv');

  section('T4 — Six-Month Follow-Up');
  field('Organization',              'T4_Org_Name');
  field('Role',                      'T4_Role');
  field('BARS — Intellectual rigor (1–5)',     'T4_BARS_Intellectual_Rigor');
  field('BARS — Prof. development (1–5)',      'T4_BARS_Prof_Development');
  field('BARS — Impact & meaning (1–5)',       'T4_BARS_Impact_Meaningfulness');
  field('Felt ready for placement',  'T4_Readiness_Retrospective');
  field('NPS (0–10)',                'T4_NPS');
  field('Most valuable aspect',      'T4_Most_Valuable');
  field('Suggested improvement',     'T4_Suggested_Improvement');
  field('Career vision',             'T4_Career_Vision');
  field('Career direction',          'T4_Career_Direction');
  field('Career direction factors',  'T4_Career_Dir_Factors');
  field('Career direction influences','T4_Career_Dir_Influences');
  field('Most significant conversation', 'T4_Most_Significant_Conv');

  tab.setColumnWidth(1, 230);
  tab.setColumnWidth(2, 420);
  tab.setFrozenRows(1);
}

// =============================================================================
// MENU
// =============================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SMA Fellows')
    .addItem('Rebuild Analytics & Explore tabs', 'buildAnalyticsTab')
    .addSeparator()
    .addItem('Rebuild all derived tabs', 'setupDerivedTabs')
    .addItem('Generate sample data',     'generateSampleData')
    .addToUi();
}

// =============================================================================
// ANALYTICS TAB  (live Sheets formulas mirroring the dashboard)
// Run buildAnalyticsTab() from the SMA Fellows menu, or it runs during setup().
// Safe to re-run — clears and rebuilds; never touches Responses data.
// =============================================================================
function buildAnalyticsTab() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var respSheet = ss.getSheetByName(SHEET_NAME);
  if (!respSheet) { Logger.log('Run setup() first.'); return; }

  // ── Read Responses headers once ──────────────────────────────────────────
  var lastCol    = respSheet.getLastColumn();
  var headers    = respSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var lastColLtr = colLetter(lastCol);

  function cIdx(name) { return headers.indexOf(name); }
  function cLtr(name) { var i=cIdx(name); return i<0?null:colLetter(i+1); }
  function cRng(name) { var l=cLtr(name); return l?'Responses!'+l+':'+l:null; }

  var cohortNum = cIdx('Cohort') + 1;                           // 1-based
  var chRng     = 'Responses!'+colLetter(cohortNum)+':'+colLetter(cohortNum);

  // ── Formula builders (all respect the cohort dropdown in B2) ─────────────
  // Average non-blank values
  function avgF(col) {
    var r=cRng(col); if(!r) return '"[missing: '+col+']"';
    return '=IFERROR(IF(B$4="All",AVERAGEIF('+r+',"<>"),AVERAGEIFS('+r+','+chRng+',B$4,'+r+',"<>")),0)';
  }
  // Count non-blank rows (subtracts header when "All")
  function cntF(col) {
    var r=cRng(col); if(!r) return '"[missing: '+col+']"';
    return '=IF(B$4="All",MAX(COUNTIF('+r+',"<>")-1,0),COUNTIFS('+chRng+',B$4,'+r+',"<>"))';
  }
  // Count rows where col = val
  function cntValF(col, val) {
    var r=cRng(col); if(!r) return '"[missing: '+col+']"';
    return '=IF(B$4="All",COUNTIF('+r+',"'+val+'"),COUNTIFS('+chRng+',B$4,'+r+',"'+val+'"))';
  }
  // % of non-blank rows where col = val (×100 for display)
  function pctF(col, val) {
    var r=cRng(col); if(!r) return '"[missing: '+col+']"';
    return '=IFERROR(IF(B$4="All",'+
      'COUNTIF('+r+',"'+val+'")/MAX(COUNTIF('+r+',"<>"),1),'+
      'COUNTIFS('+chRng+',B$4,'+r+',"'+val+'")/MAX(COUNTIFS('+chRng+',B$4,'+r+',"<>"),1))*100,0)';
  }

  // ── Sheet setup ───────────────────────────────────────────────────────────
  var sh = ss.getSheetByName('Analytics');
  if (!sh) sh = ss.insertSheet('Analytics');
  else { sh.clearContents(); sh.clearFormats(); sh.clearConditionalFormatRules(); }
  sh.setTabColor('#c9622f');

  var rw = 1;
  var BG='#161616', BG2='#1e1e1e', BG3='#111111', TXT='#f2ede4', SEC='#9e9a91', ACC='#c9622f';

  // ── Writer helpers ────────────────────────────────────────────────────────
  function sectionHead(title) {
    sh.getRange(rw,1,1,6).merge()
      .setValue(title).setFontWeight('bold').setFontSize(11)
      .setFontColor(TXT).setBackground('#2c1a10');
    rw++;
  }
  function subHead(labels) {
    labels.forEach(function(l,i){ sh.getRange(rw,i+1).setValue(l).setFontWeight('bold').setFontColor(SEC).setFontSize(10); });
    sh.getRange(rw,1,1,labels.length).setBackground(BG2);
    rw++;
  }
  // Primary metric row: label | formula | note
  function mRow(label, formula, fmt, note) {
    sh.getRange(rw,1).setValue(label).setFontColor(SEC);
    if (formula) {
      var c=sh.getRange(rw,2);
      if (typeof formula==='string'&&formula.startsWith('=')) c.setFormula(formula); else c.setValue(formula);
      if (fmt) c.setNumberFormat(fmt);
      c.setFontWeight('bold').setFontColor(TXT);
    }
    if (note) sh.getRange(rw,3).setValue(note).setFontColor('#555').setFontSize(10).setFontStyle('italic');
    rw++;
  }
  // Indented delta row (computed from other cells)
  function dRow(label, formula, fmt, note) {
    sh.getRange(rw,1).setValue(label).setFontColor('#666').setFontStyle('italic');
    if (formula) {
      var c=sh.getRange(rw,2);
      if (formula.startsWith('=')) c.setFormula(formula); else c.setValue(formula);
      if (fmt) c.setNumberFormat(fmt);
      c.setFontColor('#aaa');
    }
    if (note) sh.getRange(rw,3).setValue(note).setFontColor('#444').setFontSize(10).setFontStyle('italic');
    rw++;
  }
  // Table row: label in A, formulas in B C D E
  function tRow(label, formulas, fmt) {
    sh.getRange(rw,1).setValue(label).setFontColor(SEC);
    formulas.forEach(function(f,i){
      if (f===null||f===undefined) return;
      var gc=sh.getRange(rw,i+2);
      if (typeof f==='string'&&f.startsWith('=')) gc.setFormula(f); else gc.setValue(f);
      if (fmt) gc.setNumberFormat(fmt);
      gc.setFontColor(TXT);
    });
    rw++;
  }
  function blank() { rw++; }

  // ════════════════════════════════════════════════════════════════════════
  // TITLE + COHORT FILTER
  // ════════════════════════════════════════════════════════════════════════
  sh.getRange(rw,1,1,6).merge().setValue('SMA Fellows Program — Analytics Validation')
    .setFontSize(16).setFontWeight('bold').setFontColor(TXT).setBackground(BG3);
  rw++;
  sh.getRange(rw,1,1,6).merge()
    .setValue('Live Sheets formulas that mirror the dashboard calculations. Change the cohort dropdown to filter every section below.')
    .setFontColor(SEC).setFontSize(11).setFontStyle('italic').setWrap(true).setBackground(BG3);
  rw++;
  blank();

  // Cohort filter
  sh.getRange(rw,1).setValue('Cohort filter').setFontWeight('bold').setFontColor(TXT);
  var filterCell = sh.getRange(rw,2);
  filterCell.setValue('All').setFontWeight('bold').setFontColor(ACC).setBackground(BG2)
    .setBorder(true,true,true,true,false,false, ACC, SpreadsheetApp.BorderStyle.SOLID);
  var cohortVals = ['All'];
  if (respSheet.getLastRow() > 1) {
    var rawC = respSheet.getRange(2, cohortNum, respSheet.getLastRow()-1, 1).getValues();
    rawC.forEach(function(row){ if(row[0]&&cohortVals.indexOf(String(row[0]))<0) cohortVals.push(String(row[0])); });
  }
  filterCell.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(cohortVals,true).setAllowInvalid(false).build()
  );
  sh.getRange(rw,3).setValue('← change to filter all sections below').setFontColor('#555').setFontSize(10).setFontStyle('italic');
  rw++;
  sh.getRange(rw,1).setValue('Last rebuilt').setFontColor(SEC).setFontSize(10);
  sh.getRange(rw,2).setFormula('=TEXT(NOW(),"mmm d, yyyy h:mm am/pm")').setFontColor(SEC).setFontSize(10);
  rw++;
  sh.setFrozenRows(rw-1);
  blank();

  // ════════════════════════════════════════════════════════════════════════
  // ① SNAPSHOT
  // ════════════════════════════════════════════════════════════════════════
  sectionHead('① SNAPSHOT');
  mRow('Fellows (N)',
    '=IF(B$4="All",MAX(COUNTIF('+chRng+',"<>")-1,0),COUNTIF('+chRng+',B$4))',
    '0', 'Total rows in Responses for this cohort');
  mRow('Avg NPS — T3', avgF('T3_NPS'), '0.00', 'Program recommendation, 0–10');

  var snT1row = rw;
  mRow('Career direction "Yes" — T1 (%)', pctF('T1_Career_Direction','Yes'), '0.0"%"');
  var snT4row = rw;
  mRow('Career direction "Yes" — T4 (%)', pctF('T4_Career_Direction','Yes'), '0.0"%"');
  dRow('  ↳ Naïve shift (pp)', '=B'+snT4row+'-B'+snT1row, '+0.0;-0.0;0.0"%"', 'T4-Now % minus T1 %');

  // Placement ready: "Yes" (form text) or ≥4 (old numeric sample data)
  var rdyR = cRng('T3_Readiness_Retrospective');
  if (rdyR) {
    var rdyA='COUNTIF('+rdyR+',"Yes")+COUNTIF('+rdyR+',">=4")';
    var rdyC='COUNTIFS('+chRng+',B$4,'+rdyR+',"Yes")+COUNTIFS('+chRng+',B$4,'+rdyR+',">=4")';
    var rdyDA='MAX(COUNTIF('+rdyR+',"<>")-1,1)', rdyDC='MAX(COUNTIFS('+chRng+',B$4,'+rdyR+',"<>"),1)';
    mRow('Placement ready — T3 (%)',
      '=IFERROR(IF(B$4="All",('+rdyA+')/'+rdyDA+',('+rdyC+')/'+rdyDC+')*100,0)',
      '0.0"%"', '"Yes" or numeric ≥ 4 on retrospective readiness');
  }

  var snCV1 = rw;
  mRow('Having career conversations — T1 (%)', pctF('T1_Peer_Conv_YN','Yes'), '0.0"%"');
  var t4cvR = cRng('T4_Peer_Conv_Count');
  if (t4cvR) {
    var snCV4 = rw;
    mRow('Having career conversations — T4 (%)',
      '=IFERROR(IF(B$4="All",COUNTIF('+t4cvR+',">=1")/MAX(COUNTIF('+t4cvR+',"<>")-1,1),'+
      'COUNTIFS('+chRng+',B$4,'+t4cvR+',">=1")/MAX(COUNTIFS('+chRng+',B$4,'+t4cvR+',"<>"),1))*100,0)',
      '0.0"%"', 'T4 peer conversation count ≥ 1');
    dRow('  ↳ Change (pp)', '=B'+snCV4+'-B'+snCV1, '+0.0;-0.0;0.0"%"');
  }
  blank();

  // ════════════════════════════════════════════════════════════════════════
  // ② COMMITMENT
  // ════════════════════════════════════════════════════════════════════════
  sectionHead('② COMMITMENT TO HIGH-IMPACT WORK  (1 – 7 scale)');
  var cT1=rw;  mRow('T1 — Now  (naïve baseline)',       avgF('T1_Commitment'),        '0.00');
  var cT2t=rw; mRow('T2 — Then  (corrected baseline)', avgF('T2_Commitment_Then'),   '0.00');
  var cT2n=rw; mRow('T2 — Now  (post-orientation)',    avgF('T2_Commitment_Now'),    '0.00');
  dRow('  ↳ Naïve gain  (T1 → T2-Now)',          '=B'+cT2n+'-B'+cT1,  '+0.00;-0.00;0.00', 'Dashboard: naïve gain');
  dRow('  ↳ Corrected gain  (T2-Then → T2-Now)', '=B'+cT2n+'-B'+cT2t, '+0.00;-0.00;0.00', 'Dashboard: corrected gain — key metric');
  dRow('  ↳ Reference shift  (T2-Then − T1-Now)','=B'+cT2t+'-B'+cT1,  '+0.00;-0.00;0.00', 'Scale recalibration during orientation');
  blank();

  // ════════════════════════════════════════════════════════════════════════
  // ③ MOTIVATION QUALITY
  // ════════════════════════════════════════════════════════════════════════
  sectionHead('③ MOTIVATION QUALITY  (0 = External … 5 = Intrinsic)');
  var mT1=rw;  mRow('T1 — Now',                           avgF('T1_Motivation_Index'),         '0.00');
  var mT2t=rw; mRow('T2 — Then  (corrected baseline)',    avgF('T2_Motivation_Then_Index'),    '0.00');
  var mT2n=rw; mRow('T2 — Now  (post-orientation)',       avgF('T2_Motivation_Now_Index'),     '0.00');
  var mT4=rw;  mRow('T4 — Now  (6-month follow-up)',      avgF('T4_Motivation_Index'),         '0.00');
  dRow('  ↳ Naïve shift  (T1 → T2-Now)',          '=B'+mT2n+'-B'+mT1,  '+0.00;-0.00;0.00');
  dRow('  ↳ Corrected shift  (T2-Then → T2-Now)', '=B'+mT2n+'-B'+mT2t, '+0.00;-0.00;0.00', 'Key metric');
  dRow('  ↳ Long-term naïve  (T1 → T4)',          '=B'+mT4 +'-B'+mT1,  '+0.00;-0.00;0.00');
  dRow('  ↳ Long-term corrected  (T2-Then → T4)', '=B'+mT4 +'-B'+mT2t, '+0.00;-0.00;0.00');
  blank();
  subHead(['Motivation Level','T1 — Now','T2 — Then','T2 — Now','T4 — Now']);
  ['External','Introjected','Identified','Integrated','Intrinsic'].forEach(function(lab,idx){
    tRow(lab,[cntValF('T1_Motivation_Index',idx),cntValF('T2_Motivation_Then_Index',idx),
              cntValF('T2_Motivation_Now_Index',idx),cntValF('T4_Motivation_Index',idx)],'0');
  });
  blank();

  // ════════════════════════════════════════════════════════════════════════
  // ④ SELF-EFFICACY
  // ════════════════════════════════════════════════════════════════════════
  sectionHead('④ SELF-EFFICACY  (avg across fellows, 1 – 7 scale)');
  subHead(['Skill Area','T1 Avg','T2 Avg','Δ T1 → T2']);
  SE_ITEMS.forEach(function(item){
    var t1f=avgF('T1_SE_'+item.key), t2f=avgF('T2_SE_'+item.key);
    sh.getRange(rw,1).setValue(item.full).setFontColor(SEC);
    sh.getRange(rw,2).setFormula(t1f).setNumberFormat('0.00').setFontColor(TXT);
    sh.getRange(rw,3).setFormula(t2f).setNumberFormat('0.00').setFontColor(TXT);
    sh.getRange(rw,4).setFormula('=C'+rw+'-B'+rw).setNumberFormat('+0.00;-0.00;0.00').setFontColor(TXT);
    rw++;
  });
  blank();

  // ════════════════════════════════════════════════════════════════════════
  // ⑤ CAREER VALUES
  // ════════════════════════════════════════════════════════════════════════
  sectionHead('⑤ CAREER VALUES  (avg rank: 1 = highest priority)');
  subHead(['Value','T1 Avg Rank','T2 Avg Rank','T4 Avg Rank','Δ T1 → T4']);
  CAREER_VALUES.forEach(function(v){
    var t1f=avgF('T1_Rank_'+v.key),t2f=avgF('T2_Rank_'+v.key),t4f=avgF('T4_Rank_'+v.key);
    sh.getRange(rw,1).setValue(v.name).setFontColor(SEC);
    sh.getRange(rw,2).setFormula(t1f).setNumberFormat('0.00').setFontColor(TXT);
    sh.getRange(rw,3).setFormula(t2f).setNumberFormat('0.00').setFontColor(TXT);
    sh.getRange(rw,4).setFormula(t4f).setNumberFormat('0.00').setFontColor(TXT);
    sh.getRange(rw,5).setFormula('=D'+rw+'-B'+rw).setNumberFormat('+0.00;-0.00;0.00').setFontColor(TXT);
    rw++;
  });
  blank();

  // ════════════════════════════════════════════════════════════════════════
  // ⑥ CAREER DIRECTION
  // ════════════════════════════════════════════════════════════════════════
  sectionHead('⑥ CAREER DIRECTION  (% of fellows with data at each timepoint)');
  subHead(['Timepoint','Yes %','Still Deciding %','No %','N with data']);
  var cdDefs=[
    {label:'T1 — Now  (naïve baseline)',          col:'T1_Career_Direction'},
    {label:'T4 — Then  (retrospective baseline)', col:'T4_Career_Direction_Then'},
    {label:'T4 — Now',                            col:'T4_Career_Direction'},
  ];
  var cdYesRow={};
  cdDefs.forEach(function(cd){
    sh.getRange(rw,1).setValue(cd.label).setFontColor(SEC);
    sh.getRange(rw,2).setFormula(pctF(cd.col,'Yes')).setNumberFormat('0.0"%"').setFontColor(TXT);
    sh.getRange(rw,3).setFormula(pctF(cd.col,'Still deciding')).setNumberFormat('0.0"%"').setFontColor(TXT);
    sh.getRange(rw,4).setFormula(pctF(cd.col,'No')).setNumberFormat('0.0"%"').setFontColor(TXT);
    sh.getRange(rw,5).setFormula(cntF(cd.col)).setNumberFormat('0').setFontColor(TXT);
    cdYesRow[cd.col]=rw; rw++;
  });
  dRow('  ↳ Naïve shift  (T1 → T4-Now, pp)',
    '=B'+cdYesRow['T4_Career_Direction']+'-B'+cdYesRow['T1_Career_Direction'],
    '+0.0;-0.0;0.0"%"','Dashboard: naïve shift');
  dRow('  ↳ Corrected shift  (T4-Then → T4-Now, pp)',
    '=B'+cdYesRow['T4_Career_Direction']+'-B'+cdYesRow['T4_Career_Direction_Then'],
    '+0.0;-0.0;0.0"%"','Dashboard: corrected shift');
  blank();

  // ════════════════════════════════════════════════════════════════════════
  // ⑦ BARRIERS
  // ════════════════════════════════════════════════════════════════════════
  sectionHead('⑦ BARRIERS  (% of fellows selecting each — stored as 0 / 1)');
  subHead(['Barrier','T2 Anticipated %','T3 Experienced %','Diff (pp)']);
  BARRIERS.forEach(function(b){
    if (b.key==='Other') return;
    var t2r=cRng('T2_Barrier_'+b.key), t3r=cRng('T3_Barrier_'+b.key);
    if (!t2r||!t3r) return;
    var t2f='=IFERROR(IF(B$4="All",AVERAGEIF('+t2r+',"<>"),AVERAGEIFS('+t2r+','+chRng+',B$4,'+t2r+',"<>"))*100,0)';
    var t3f='=IFERROR(IF(B$4="All",AVERAGEIF('+t3r+',"<>"),AVERAGEIFS('+t3r+','+chRng+',B$4,'+t3r+',"<>"))*100,0)';
    sh.getRange(rw,1).setValue(b.full).setFontColor(SEC);
    sh.getRange(rw,2).setFormula(t2f).setNumberFormat('0.0"%"').setFontColor(TXT);
    sh.getRange(rw,3).setFormula(t3f).setNumberFormat('0.0"%"').setFontColor(TXT);
    sh.getRange(rw,4).setFormula('=C'+rw+'-B'+rw).setNumberFormat('+0.0;-0.0;0.0"%"').setFontColor(TXT);
    rw++;
  });
  blank();

  // ════════════════════════════════════════════════════════════════════════
  // ⑧ PLACEMENT QUALITY
  // ════════════════════════════════════════════════════════════════════════
  sectionHead('⑧ PLACEMENT QUALITY — BARS  (T3, 1 – 5 scale)');
  subHead(['Dimension','T3 Avg Rating','N rated']);
  [{label:'Intellectual Rigor',       col:'T3_BARS_Intellectual_Rigor'},
   {label:'Professional Development', col:'T3_BARS_Prof_Development'},
   {label:'Impact Meaningfulness',    col:'T3_BARS_Impact_Meaningfulness'}
  ].forEach(function(d){
    sh.getRange(rw,1).setValue(d.label).setFontColor(SEC);
    sh.getRange(rw,2).setFormula(avgF(d.col)).setNumberFormat('0.00').setFontColor(TXT);
    sh.getRange(rw,3).setFormula(cntF(d.col)).setNumberFormat('0').setFontColor(TXT);
    rw++;
  });

  // ── Column widths + background ────────────────────────────────────────────
  sh.getRange(1,1,rw,6).setBackground(BG).setFontFamily('Arial');
  sh.setColumnWidth(1,310); sh.setColumnWidth(2,120); sh.setColumnWidth(3,120);
  sh.setColumnWidth(4,120); sh.setColumnWidth(5,120); sh.setColumnWidth(6,220);

  Logger.log('✓ Analytics tab built: '+(rw-1)+' rows');

  // ── Build Explore tab ─────────────────────────────────────────────────────
  buildExploreTab(ss, lastColLtr, cohortNum);

  try { SpreadsheetApp.getUi().alert('✓ Analytics and Explore tabs rebuilt successfully.'); } catch(e) {}
}

// =============================================================================
// EXPLORE TAB  (filtered data view + pivot table instructions)
// =============================================================================
function buildExploreTab(ss, lastColLtr, cohortNum) {
  var sh = ss.getSheetByName('Explore');
  if (!sh) sh = ss.insertSheet('Explore');
  else { sh.clearContents(); sh.clearFormats(); }
  sh.setTabColor('#4a7c59');

  var BG='#161616', BG3='#111111', TXT='#f2ede4', SEC='#9e9a91';

  // Title block
  sh.getRange(1,1,1,6).merge().setValue('SMA Fellows — Explore')
    .setFontSize(16).setFontWeight('bold').setFontColor(TXT).setBackground(BG3);
  sh.getRange(2,1,1,6).merge()
    .setValue('Shows all Responses rows for the cohort selected in Analytics!B2. ' +
              'To build custom breakdowns: click any cell in the data below → Insert → Pivot table.')
    .setFontColor(SEC).setFontSize(11).setFontStyle('italic').setWrap(true).setBackground(BG3);
  sh.getRange(3,1,1,6).merge()
    .setValue('Tip: sort or filter this view freely — it re-queries live data every time Analytics!B2 changes.')
    .setFontColor('#555').setFontSize(10).setFontStyle('italic').setBackground(BG3);

  sh.setFrozenRows(4); // rows 1-3 = title; row 4 onward = QUERY output with its own header

  // Filtered data QUERY — cohort col is hardcoded from build-time index
  var emailColNum = 1; // Email is always col 1
  var qAll  = '"SELECT * WHERE Col'+emailColNum+'<>\'\'"';
  var qFilt = '"SELECT * WHERE Col'+cohortNum+'=\'"&Analytics!B2&"\'"';
  sh.getRange(4,1).setFormula(
    '=IFERROR(IF(Analytics!B2="All",'+
      'QUERY(Responses!A:'+lastColLtr+','+qAll+',1),'+
      'QUERY(Responses!A:'+lastColLtr+','+qFilt+',1)'+
    '),"No matching data — check the cohort filter in the Analytics tab")'
  ).setFontColor(TXT).setFontFamily('Arial');

  sh.getRange(1,1,4,6).setBackground(BG3);
  sh.setColumnWidth(1,180); sh.setColumnWidth(2,140);

  Logger.log('✓ Explore tab built');
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
    ['Cohort',      'Program cohort identifier entered at T1 (e.g., 2025-26). Groups responses for cross-year analysis.', 'T1', 'Text', 'Free text — e.g. 2025-26'],
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
