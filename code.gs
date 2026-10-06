/**
 * Cap Plan — Google Apps Script backend (Sheets databank)
 * -------------------------------------------------------------
 * Databank:
 *   https://docs.google.com/spreadsheets/d/1aeaPWGQghoXC__bmEwsgU7203rxzg5jRQz5CN_IInTI/edit
 *
 * SETUP
 * 1. In that Sheet: Extensions → Apps Script (or bind this project to the sheet).
 * 2. Paste this file as Code.gs. Create HTML file "Index" from Index.html.
 * 3. Run setupSpreadsheet() once (authorize), then Deploy → New deployment
 *    → Web app → Execute as: Me → Who has access: your org / anyone with link.
 * 4. Open the web app URL. Re-deploy after replacing Index.html.
 *
 * Live tabs:
 *   Projects, Employees, Current data, Current data_years, Access
 *   cost center information, budget, booking
 *   (+ meta, settings created if missing)
 * Snapshots: "YYYYMMDD data", "YYYYMMDD data_years"
 *
 * Year-series Source codes: Dm (demand/capacity/timeline), Alc (allocation),
 *   Emp (employee). Legacy Bud month series is superseded by the budget tab.
 *   Legacy Sc / M1–M12 headers still read.
 */

/** Production Cap Plan spreadsheet (edit ID only if the databank moves). */
var DATABANK_ID = '1aeaPWGQghoXC__bmEwsgU7203rxzg5jRQz5CN_IInTI';

/** External activity log (load / save / notify). Script owner must have edit access. */
var LOGGER_SHEET_ID = '1RnzfCWrxc1adpgXa3M6JX3pL7MfPXC6_YDFD0Rn34S8';
var LOGGER_TAB = 'Logger';
var LOGGER_HEADERS = ['Timestamp', 'Event', 'User Name', 'Role', 'Email', 'Details'];
/** Tabs scanned for name→email (Activity Logger workbook). */
var EMAIL_DIR_TABS = ['Contacts', 'Activity_Logger', 'Logger'];
var CONTACTS_HEADERS = ['User Name', 'Email', 'Updated'];

var SHEETS = {
  PROJECTS: 'Projects',
  EMPLOYEES: 'Employees',
  DATA: 'Current data',
  YEARS: 'Current data_years',
  META: 'meta',
  SETTINGS: 'settings',
  ACCESS: 'Access',
  COST_CENTERS: 'cost center information',
  BUDGET: 'budget',
  BOOKING: 'booking'
};

/** 1 FTE = 133 hours for SAP transfer conversion. */
var FTE_HOURS = 133;
/** Optional SAP row split threshold (hours); UI default off. */
var TRANSFER_SPLIT_HOURS = 99;
var TRANSFER_TYPES = ['direct', 'indirect', 'finance', 'none'];
var COST_CENTER_LEVELS = ['unit', 'department', 'group', 'role', 'name'];

/** Legacy aliases resolved by sheetByName_() when the canonical tab is missing. */
var SHEET_ALIASES = {
  Projects: ['project information'],
  Employees: ['employee information'],
  'Current data': ['data information', 'row information'],
  'Current data_years': ['data_years', 'row_years'],
  Access: ['user access']
};

var SNAP_DATA_SUFFIX = ' data';
var SNAP_YEARS_SUFFIX = ' data_years';
/** Legacy snapshot prefixes */
var SNAP_DATA_PREFIX_LEGACY = 'data information snapshot ';
var SNAP_ROW_PREFIX_LEGACY = 'row information snapshot ';
var SNAP_YEARS_PREFIX_LEGACY = 'data_years snapshot ';
var SNAP_ROW_YEARS_PREFIX_LEGACY = 'row_years snapshot ';

var SOURCES = {
  DA: 'Demand/Allocation',
  AVAIL: 'Availability',
  HEADER: 'Timeline/Header'
};

/** Internal data_type values (also accept workbook Sc: Dm / Alc). */
var DATA_TYPES = {
  DEMAND: 'demand',
  ALLOCATION: 'allocation',
  TIMELINE: 'timeline',
  AVAILABILITY: 'availability',
  EMPLOYEE: 'employee',
  BUDGET: 'budget'
};

var MONTH_MM = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'];
var MONTH_HEADERS = ['m01', 'm02', 'm03', 'm04', 'm05', 'm06', 'm07', 'm08', 'm09', 'm10', 'm11', 'm12'];
var MONTH_HEADERS_SHORT = ['M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10', 'M11', 'M12'];
var MONTH_HEADERS_LONG = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** Current data — matches databank headers. */
var DATA_HEADERS = [
  'Row_ID', 'Sort', 'Source', 'Version', 'ID Project', 'ID Emp',
  'Workpackage (opt.)', 'Notes Project Lead', 'Notes People Lead', 'Notes Employee',
  'Comments', 'Changed by', 'Alternative Financing', 'Updated at', 'Change history'
];

/** Current data_years — Monthrow_ID / Row_ID / Year / Source / Version / Jan–Dec. */
var YEAR_HEADERS = ['Monthrow_ID', 'Row_ID', 'Year', 'Source', 'Version'].concat(MONTH_HEADERS_LONG);

var PROJECT_HEADERS = [
  'Project_ID', 'Project/Topic', 'Program', 'WBS MA', 'WBS Indy', 'Portfolio', 'Category',
  'Status', 'Link Project', 'Lin Resource', 'Comment', 'Project Lead', 'Project Location',
  'Transfer Type'
];

var EMPLOYEE_HEADERS = [
  'Employee_ID', 'Role', 'Name', 'Comments', 'People Location', 'Group', 'FTE Type'
];

/** Hierarchy cost-center attribution (most-specific level wins per field). */
var COST_CENTER_HEADERS = [
  'CC_ID', 'Level', 'Level Value',
  'Internal Cost Center', 'Internal Activity Type',
  'External Cost Center', 'External Activity Type',
  'RBS Skill Name', 'Hourly Cost'
];

/**
 * Finance month store (budget + booking).
 * Kind=row → Row_ID override; Kind=pair → Cost Center + Activity Type + Project Cost Center.
 * Overarching pair values win over Row_ID for the same month.
 */
var FINANCE_MONTH_HEADERS = [
  'Kind', 'Row_ID', 'Cost Center', 'Activity Type', 'Project Cost Center', 'Year', 'Manual'
].concat(MONTH_HEADERS_LONG);

var BOOKING_HEADERS = FINANCE_MONTH_HEADERS.concat([
  'Transfer Type'
]).concat(MONTH_HEADERS_LONG.map(function (m) { return 'Auto_' + m; }));

/** Access tab — User Name / Role / Scope / Scope Value / Source (+ optional User Email). */
var ACCESS_HEADERS = [
  'User Name', 'Role', 'Scope', 'Scope Value', 'Source', 'User Email'
];
var DEFAULT_ADMIN_ACCESS_ID = 'access_admin_the_admin';
var DEFAULT_ADMIN_USER_NAME = 'the admin';

/** Back-compat name used by older call sites in this file. */
var ROW_HEADERS = DATA_HEADERS;

/**
 * Open the Cap Plan databank.
 * Prefer the container-bound / active spreadsheet (no openById scope needed).
 * Fall back to openById only when the script is standalone or the active sheet differs.
 */
function ss_() {
  var active = null;
  try {
    active = SpreadsheetApp.getActive();
  } catch (e0) {
    active = null;
  }

  // Bound script on the Cap Plan sheet — use it directly.
  if (active && (!DATABANK_ID || active.getId() === DATABANK_ID)) {
    return active;
  }

  if (DATABANK_ID) {
    try {
      return SpreadsheetApp.openById(DATABANK_ID);
    } catch (e) {
      // Missing spreadsheets scope, or script not yet re-authorized after adding openById.
      if (active) return active;
      throw new Error(
        'Cannot open databank ' + DATABANK_ID + '. ' +
        'Open the Sheet → Extensions → Apps Script (container-bound), then run setupSpreadsheet() ' +
        'and approve Sheets access. Detail: ' + (e && e.message ? e.message : e)
      );
    }
  }

  if (active) return active;
  throw new Error('Open this script from the Cap Plan Google Spreadsheet (Extensions → Apps Script).');
}

/* ========== Web entry ========== */

function doGet(e) {
  ensureSchema_();
  var out = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Cap Plan — Resource Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  var params = (e && e.parameter) || {};
  var deep = {
    row: String(params.row || params.rowId || '').trim(),
    tab: String(params.tab || params.mode || '').trim().toLowerCase()
  };
  if (deep.tab !== 'allocation' && deep.tab !== 'demand') {
    deep.tab = deep.tab.indexOf('alloc') >= 0 ? 'allocation' : 'demand';
  }
  if (deep.row) {
    var html = out.getContent();
    var inject = '<script>window.__CAPPLAN_DEEP_LINK=' +
      JSON.stringify(deep) + ';</script>';
    if (html.indexOf('</head>') >= 0) {
      html = html.replace('</head>', inject + '</head>');
    } else {
      html = inject + html;
    }
    out.setContent(html);
  }
  return out;
}

/** Deployed web-app base URL (empty if not deployed yet). */
function webAppBaseUrl_() {
  try {
    return String(ScriptApp.getService().getUrl() || '').trim();
  } catch (err) {
    return '';
  }
}

/** Deep-link into Cap Plan for a planning row. */
function rowDeepLink_(rowId, tab) {
  var base = webAppBaseUrl_();
  if (!base || !rowId) return '';
  var t = String(tab || 'demand').toLowerCase() === 'allocation' ? 'allocation' : 'demand';
  return base +
    '?row=' + encodeURIComponent(String(rowId)) +
    '&tab=' + encodeURIComponent(t);
}

/** Run once from the editor to create tabs + sample rows. */
function setupSpreadsheet() {
  ensureSchema_();
  seedIfEmpty_();
  SpreadsheetApp.getUi().alert('Cap Plan sheets are ready.');
}

/* ========== Bootstrap for UI ========== */

/**
 * Cap Plan is deployed as Execute as: Me (script owner) so the spreadsheet
 * stays private. Visitor identity is resolved automatically when possible:
 *   1) Explicit registeredAs when forceRegistered (Switch user / manual pick)
 *   2) Session.getActiveUser() email (Workspace / same-domain sign-in)
 *   3) Remembered mapping for Session.getTemporaryActiveUserKey()
 *   4) Explicit registeredAs without force (first-time / localStorage)
 *
 * Wave APIs (preferred progressive load):
 *   detectVisitor → getMastersBundle → getPortfolioBundle → getFinanceBundle
 * getBootstrap remains a full one-shot for snapshots / legacy callers.
 *
 * @param {string=} registeredAs Display name selected at startup / Switch user
 * @param {boolean=} forceRegistered When true, prefer registeredAs over auto-detect
 */
function getBootstrap(registeredAs, forceRegistered) {
  ensureSchema_();
  var masters = getMastersBundle_(registeredAs, forceRegistered);
  var portfolio = getPortfolioBundle_(masters);
  var finance = getFinanceBundle_();
  return Object.assign({}, masters, portfolio, finance);
}

/** Wave 1 — projects, employees, access, meta, settings (no planning rows). */
function getMastersBundle(registeredAs, forceRegistered) {
  ensureSchema_();
  return getMastersBundle_(registeredAs, forceRegistered);
}

function getMastersBundle_(registeredAs, forceRegistered) {
  var settings = readSettings_();
  var months = monthsFromSettings_(settings);
  var projects = readProjects_();
  var employees = readEmployees_();
  var meta = readMeta_();
  var access = readAccess_();
  var knownUsers = knownRegistrableUsers_(access, projects, employees);
  var force = !!forceRegistered;
  var resolved = resolveVisitorIdentity_(registeredAs, access, projects, employees, knownUsers, force);
  var identity = resolved.identity;
  var needsRegistration = !identity;
  var role = 'Employee';
  var allowed = ['Employee'];
  if (identity) {
    role = resolveRole_(identity, access, settings.defaultRole, projects);
    allowed = allowedRolesFor_(identity, access, projects);
    if (!force) rememberVisitorIdentity_(identity, resolved.visitorEmail);
  }
  employees = redactEmployeeCommentsForRole_(employees, role, identity || '', access);
  return {
    userEmail: resolved.visitorEmail || '',
    userName: identity || '',
    signedInEmail: resolved.visitorEmail || '',
    ownerEmail: resolved.ownerEmail || '',
    visitorKey: resolved.visitorKey || '',
    executeAsOwner: true,
    identitySource: resolved.source || 'none',
    autoDetected: !!resolved.autoDetected,
    forceRegistered: force,
    registeredAs: identity || '',
    needsRegistration: needsRegistration,
    knownUsers: knownUsers,
    role: role,
    allowedRoles: allowed.length ? allowed : ['Employee'],
    settings: settings,
    months: months,
    projects: projects,
    employees: employees,
    meta: meta,
    access: access,
    costCenters: readCostCenters_(),
    wave: 'masters'
  };
}

/** Wave 2a — Current data + snapshot row identity (no _years month series yet). */
function getPortfolioRowsBundle(registeredAs, forceRegistered) {
  ensureSchema_();
  var settings = readSettings_();
  var months = monthsFromSettings_(settings);
  var projects = readProjects_();
  var employees = readEmployees_();
  var rowRecords = readRowRecords_();
  var rows = assembleRows_(rowRecords, [], projects, employees);

  var snapshotDates = listSnapshotDates_();
  var compareDate = defaultCompareSnapshotDate_(snapshotDates, settings.compareSnapshotDate);
  var compare = [];
  var rowSnapshots = {};
  if (compareDate != null) {
    var snapPack = loadSnapshotRowsPack_(compareDate, projects, employees);
    compare = snapPack.rows;
    rowSnapshots = snapPack.rowSnapshots;
  }

  if (registeredAs) {
    logAppEvent_('load', registeredAs, '',
      'wave=portfolioRows; rows=' + rows.length +
      '; compare=' + (compareDate != null ? formatSnapshotLabel_(compareDate) : 'none'));
  }

  return {
    months: months,
    rows: rows,
    compare: compare,
    rowSnapshots: rowSnapshots,
    snapshotDates: snapshotDates,
    compareSnapshotDate: compareDate,
    snapshotLabel: compareDate != null ? formatSnapshotLabel_(compareDate) : '',
    wave: 'portfolioRows',
    yearsPending: true
  };
}

/** Wave 2b — Current data_years + snapshot _years month series. */
function getPortfolioYearsBundle(registeredAs, forceRegistered) {
  ensureSchema_();
  var settings = readSettings_();
  var projects = readProjects_();
  var employees = readEmployees_();
  var rowRecords = readRowRecords_();
  var yearRecords = readYearRecords_();
  var rows = assembleRows_(rowRecords, yearRecords, projects, employees);
  syncAvailabilityTotals_(rows);

  var snapshotDates = listSnapshotDates_();
  var compareDate = defaultCompareSnapshotDate_(snapshotDates, settings.compareSnapshotDate);
  var compare = [];
  var rowSnapshots = {};
  if (compareDate != null) {
    var snapPack = loadSnapshotPack_(compareDate, projects, employees);
    compare = snapPack.rows;
    rowSnapshots = snapPack.rowSnapshots;
  }

  if (registeredAs) {
    logAppEvent_('load', registeredAs, '',
      'wave=portfolioYears; rows=' + rows.length +
      '; compare=' + (compareDate != null ? formatSnapshotLabel_(compareDate) : 'none'));
  }

  return {
    rows: rows,
    compare: compare,
    rowSnapshots: rowSnapshots,
    snapshotDates: snapshotDates,
    compareSnapshotDate: compareDate,
    snapshotLabel: compareDate != null ? formatSnapshotLabel_(compareDate) : '',
    wave: 'portfolioYears',
    yearsPending: false
  };
}

/** Wave 2 — full portfolio (rows + years). Kept for compatibility. */
function getPortfolioBundle(registeredAs, forceRegistered) {
  ensureSchema_();
  var settings = readSettings_();
  var months = monthsFromSettings_(settings);
  var projects = readProjects_();
  var employees = readEmployees_();
  return getPortfolioBundle_({
    settings: settings,
    months: months,
    projects: projects,
    employees: employees,
    registeredAs: registeredAs,
    role: ''
  });
}

function getPortfolioBundle_(mastersHint) {
  var settings = (mastersHint && mastersHint.settings) || readSettings_();
  var months = (mastersHint && mastersHint.months) || monthsFromSettings_(settings);
  var projects = (mastersHint && mastersHint.projects) || readProjects_();
  var employees = (mastersHint && mastersHint.employees) || readEmployees_();
  var rowRecords = readRowRecords_();
  var yearRecords = readYearRecords_();
  var rows = assembleRows_(rowRecords, yearRecords, projects, employees);
  syncAvailabilityTotals_(rows);

  var snapshotDates = listSnapshotDates_();
  var compareDate = defaultCompareSnapshotDate_(snapshotDates, settings.compareSnapshotDate);
  var compare = [];
  var rowSnapshots = {};
  if (compareDate != null) {
    var snapPack = loadSnapshotPack_(compareDate, projects, employees);
    compare = snapPack.rows;
    rowSnapshots = snapPack.rowSnapshots;
  }

  if (mastersHint && mastersHint.registeredAs) {
    logAppEvent_('load', mastersHint.registeredAs, mastersHint.role || 'Employee',
      'wave=portfolio; rows=' + rows.length +
      '; compare=' + (compareDate != null ? formatSnapshotLabel_(compareDate) : 'none'));
  }

  return {
    months: months,
    rows: rows,
    compare: compare,
    rowSnapshots: rowSnapshots,
    snapshotDates: snapshotDates,
    compareSnapshotDate: compareDate,
    snapshotLabel: compareDate != null ? formatSnapshotLabel_(compareDate) : '',
    wave: 'portfolio',
    yearsPending: false
  };
}

/**
 * Client-triggered reload (no identity re-check). Logs to Activity Logger.
 * @param {string} userName
 * @param {string=} role
 * @param {string=} details
 */
function logPortfolioReload(userName, role, details) {
  ensureSchema_();
  var who = String(userName || '').trim() || 'unknown';
  var r = String(role || '').trim();
  logAppEvent_('reload', who, r, String(details || 'user reload (>24h since last load)'));
  return { ok: true, logged: true, at: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss") };
}

/** Wave 3 — budget + booking finance ledgers. */
function getFinanceBundle() {
  ensureSchema_();
  return getFinanceBundle_();
}

function getFinanceBundle_() {
  return {
    budgetRecords: readFinanceRecords_(SHEETS.BUDGET, false),
    bookingRecords: readFinanceRecords_(SHEETS.BOOKING, true),
    wave: 'finance'
  };
}

/**
 * Lightweight identity probe (no portfolio load) for the boot splash sequence.
 * @param {string=} registeredAs
 * @param {boolean=} forceRegistered
 */
function detectVisitor(registeredAs, forceRegistered) {
  ensureSchema_();
  var access = readAccess_();
  var projects = readProjects_();
  var employees = readEmployees_();
  var knownUsers = knownRegistrableUsers_(access, projects, employees);
  var force = !!forceRegistered;
  var resolved = resolveVisitorIdentity_(registeredAs, access, projects, employees, knownUsers, force);
  var identity = resolved.identity;
  var role = 'Employee';
  var allowed = ['Employee'];
  if (identity) {
    var settings = readSettings_();
    role = resolveRole_(identity, access, settings.defaultRole, projects);
    allowed = allowedRolesFor_(identity, access, projects);
  }
  return {
    userEmail: resolved.visitorEmail || '',
    signedInEmail: resolved.visitorEmail || '',
    ownerEmail: resolved.ownerEmail || '',
    visitorKey: resolved.visitorKey || '',
    identitySource: resolved.source || 'none',
    autoDetected: !!resolved.autoDetected,
    forceRegistered: force,
    registeredAs: identity || '',
    needsRegistration: !identity,
    knownUsers: knownUsers,
    role: role,
    allowedRoles: allowed.length ? allowed : ['Employee']
  };
}

/**
 * Lightweight identity + comment redaction after parallel portfolio load.
 * @param {string} registeredAs
 */
function getIdentityOverlay(registeredAs) {
  ensureSchema_();
  var access = readAccess_();
  var projects = readProjects_();
  var employees = readEmployees_();
  var knownUsers = knownRegistrableUsers_(access, projects, employees);
  var resolved = resolveVisitorIdentity_(registeredAs, access, projects, employees, knownUsers, false);
  var identity = resolved.identity;
  if (!identity) {
    return { ok: false, needsRegistration: true, signedInEmail: resolved.visitorEmail || '' };
  }
  // Google-only gate: require session or remembered mapping
  if (resolved.source !== 'session' && resolved.source !== 'remembered' && !resolved.autoDetected) {
    if (resolved.source === 'registered' || resolved.source === 'forced') {
      return {
        ok: false,
        needsRegistration: true,
        signedInEmail: resolved.visitorEmail || '',
        message: 'Manual identity is disabled'
      };
    }
  }
  var settings = readSettings_();
  var role = resolveRole_(identity, access, settings.defaultRole, projects);
  var allowed = allowedRolesFor_(identity, access, projects);
  rememberVisitorIdentity_(identity, resolved.visitorEmail);
  employees = redactEmployeeCommentsForRole_(employees, role, identity, access);
  logAppEvent_('load', identity, role,
    'source=' + (resolved.source || 'none') + '; overlay=1');
  return {
    ok: true,
    registeredAs: identity,
    role: role,
    allowedRoles: allowed.length ? allowed : ['Employee'],
    employees: employees,
    signedInEmail: resolved.visitorEmail || '',
    identitySource: resolved.source || 'none',
    autoDetected: !!resolved.autoDetected
  };
}

/** Append a row to the external Logger spreadsheet (best-effort). */
function logAppEvent_(event, userName, role, details) {
  try {
    var ss = SpreadsheetApp.openById(LOGGER_SHEET_ID);
    var sh = ss.getSheetByName(LOGGER_TAB);
    if (!sh) {
      sh = ss.insertSheet(LOGGER_TAB);
      sh.appendRow(LOGGER_HEADERS);
      sh.setFrozenRows(1);
    } else if (sh.getLastRow() < 1 || !String(sh.getRange(1, 1).getValue() || '').trim()) {
      sh.getRange(1, 1, 1, LOGGER_HEADERS.length).setValues([LOGGER_HEADERS]);
      sh.setFrozenRows(1);
    }
    var email = '';
    try { email = String(Session.getActiveUser().getEmail() || '').trim(); } catch (e0) {}
    var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sh.appendRow([
      ts,
      String(event || ''),
      String(userName || ''),
      String(role || ''),
      email,
      String(details || '')
    ]);
  } catch (err) {
    // Logger must never block load/save.
  }
}

function loggerSpreadsheet_() {
  return SpreadsheetApp.openById(LOGGER_SHEET_ID);
}

function looksLikeEmail_(v) {
  var s = String(v || '').trim();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
}

/**
 * Build name→email directory from Activity Logger workbook
 * (Contacts, Activity_Logger, Logger) + Access User Email when present.
 */
function buildEmailDirectory_() {
  var map = {}; // lowerName -> { name, email }
  function put(name, email) {
    var n = String(name || '').trim();
    var e = String(email || '').trim();
    if (!n || !e || !looksLikeEmail_(e)) return;
    if (looksLikeEmail_(n)) return; // skip pure-email "names"
    var key = n.toLowerCase();
    map[key] = { name: n, email: e };
  }
  try {
    var ss = loggerSpreadsheet_();
    EMAIL_DIR_TABS.forEach(function (tabName) {
      var sh = ss.getSheetByName(tabName);
      if (!sh || sh.getLastRow() < 2) return;
      var values = sh.getDataRange().getValues();
      if (!values.length) return;
      var h = values[0].map(function (x) { return String(x || '').trim().toLowerCase(); });
      var iName = -1;
      var iEmail = -1;
      h.forEach(function (label, i) {
        if (iName < 0 && (label === 'user name' || label === 'name' || label === 'user')) iName = i;
        if (iEmail < 0 && (label === 'email' || label === 'e-mail' || label === 'user email')) iEmail = i;
      });
      if (iName < 0) iName = 0;
      if (iEmail < 0) {
        // Logger layout: Timestamp, Event, User Name, Role, Email, Details
        if (h.indexOf('email') >= 0) iEmail = h.indexOf('email');
        else if (values[0].length >= 5) iEmail = 4;
      }
      if (iEmail < 0) return;
      for (var r = 1; r < values.length; r++) {
        put(values[r][iName], values[r][iEmail]);
      }
    });
  } catch (err) {}
  try {
    var access = readAccess_();
    (access || []).forEach(function (a) {
      put(a.userName || a.user_name, a.userEmail);
    });
  } catch (err2) {}
  return map;
}

/** Lookup emails for display names (on-demand from Activity Logger). */
function lookupNotificationEmails(names) {
  var dir = buildEmailDirectory_();
  var out = {};
  (names || []).forEach(function (raw) {
    var s = String(raw || '').trim();
    if (!s) return;
    if (looksLikeEmail_(s)) {
      out[s] = { name: s, email: s, source: 'direct' };
      return;
    }
    var hit = dir[s.toLowerCase()];
    if (hit) {
      out[s] = { name: hit.name, email: hit.email, source: 'activity_logger' };
    } else {
      out[s] = { name: s, email: '', source: 'missing' };
    }
  });
  return { ok: true, directory: out };
}

/** Persist a name→email mapping into Contacts on the Activity Logger workbook. */
function rememberNotificationEmail(name, email) {
  var n = String(name || '').trim();
  var e = String(email || '').trim();
  if (!n) throw new Error('Name required');
  if (!looksLikeEmail_(e)) throw new Error('Valid email required');
  var ss = loggerSpreadsheet_();
  var sh = ss.getSheetByName('Contacts');
  if (!sh) {
    sh = ss.insertSheet('Contacts');
    sh.getRange(1, 1, 1, CONTACTS_HEADERS.length).setValues([CONTACTS_HEADERS]);
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() < 1 || !String(sh.getRange(1, 1).getValue() || '').trim()) {
    sh.getRange(1, 1, 1, CONTACTS_HEADERS.length).setValues([CONTACTS_HEADERS]);
    sh.setFrozenRows(1);
  }
  var values = sh.getDataRange().getValues();
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var found = false;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim().toLowerCase() === n.toLowerCase()) {
      sh.getRange(i + 1, 1, 1, 3).setValues([[n, e, ts]]);
      found = true;
      break;
    }
  }
  if (!found) sh.appendRow([n, e, ts]);
  return { ok: true, name: n, email: e };
}

/**
 * Send Cap Plan notification email.
 * payload = {
 *   senderName, senderRole, comment,
 *   recipients: [{ name, email, roleLabel }],
 *   columns: [label...],
 *   rows: [{ cells: { label: value }, type, project, role, name }],
 *   registerEmails?: [{ name, email }]
 * }
 */
function sendCapPlanNotification(payload) {
  payload = payload || {};
  var sender = String(payload.senderName || '').trim() || 'Cap Plan user';
  var comment = String(payload.comment || '').trim();
  var recipients = payload.recipients || [];
  var rows = payload.rows || [];
  if (!recipients.length) throw new Error('No recipients selected');
  if (!rows.length) throw new Error('No rows selected');

  (payload.registerEmails || []).forEach(function (pair) {
    try {
      if (pair && pair.name && pair.email) rememberNotificationEmail(pair.name, pair.email);
    } catch (e0) {}
  });

  var missing = recipients.filter(function (r) {
    return !looksLikeEmail_(r && r.email);
  });
  if (missing.length) {
    throw new Error('Missing email for: ' + missing.map(function (r) {
      return r.name || '?';
    }).join(', '));
  }

  var subject = 'Cap Plan notification by ' + sender;
  var toList = [];
  var seen = {};
  recipients.forEach(function (r) {
    var e = String(r.email || '').trim().toLowerCase();
    if (!e || seen[e]) return;
    seen[e] = 1;
    toList.push(String(r.email).trim());
  });
  if (!toList.length) throw new Error('No valid recipient emails');

  var columns = payload.columns || [];
  if (!columns.length && rows[0] && rows[0].columns) columns = rows[0].columns;
  if (!columns.length && rows[0] && rows[0].cells) columns = Object.keys(rows[0].cells);
  if (!columns.length) columns = ['Type', 'Project', 'Role', 'Name'];

  var rowLines = rows.map(function (r, i) {
    var cells = r.cells || {};
    var bits = columns.map(function (c) {
      return c + '=' + String(cells[c] != null ? cells[c] : '');
    });
    return (i + 1) + '. ' + bits.join(' | ');
  }).join('\n');

  var recipLines = recipients.map(function (r) {
    return '- ' + String(r.name || '') +
      (r.roleLabel ? (' (' + r.roleLabel + ')') : '') +
      ' <' + String(r.email || '') + '>';
  }).join('\n');

  var body =
    'Cap Plan notification\n' +
    'From: ' + sender + (payload.senderRole ? (' (' + payload.senderRole + ')') : '') + '\n\n' +
    (comment ? ('Comment:\n' + comment + '\n\n') : '') +
    'Selected rows (' + rows.length + '):\n' + rowLines + '\n\n' +
    'Recipients:\n' + recipLines + '\n';

  var htmlHead = columns.map(function (c) {
    return '<th style="text-align:left;padding:4px 8px;border:1px solid #ddd;background:#f4f4f5;white-space:nowrap">' +
      escHtml_(c) + '</th>';
  }).join('') +
    '<th style="text-align:left;padding:4px 8px;border:1px solid #ddd;background:#f4f4f5;white-space:nowrap">Link</th>';

  var htmlRows = rows.map(function (r) {
    var cells = r.cells || {};
    var colors = r.cellColors || {};
    var tds = columns.map(function (c) {
      var color = String(colors[c] || '').trim();
      var style = 'padding:4px 8px;border:1px solid #ddd';
      if (color) style += ';color:' + color + ';font-weight:700';
      return '<td style="' + style + '">' +
        escHtml_(cells[c] != null ? cells[c] : '') + '</td>';
    }).join('');
    var tab = String(r.tab || '').toLowerCase() === 'allocation' ? 'allocation'
      : (r.is_availability ? 'allocation' : 'demand');
    var link = rowDeepLink_(r.id || r.rowId, tab);
    var linkTd = link
      ? ('<td style="padding:4px 8px;border:1px solid #ddd;white-space:nowrap">' +
        '<a href="' + escHtml_(link) + '" style="color:#2563eb;font-weight:600">Open in Cap Plan</a></td>')
      : '<td style="padding:4px 8px;border:1px solid #ddd"></td>';
    return '<tr>' + tds + linkTd + '</tr>';
  }).join('');

  var htmlRecip = recipients.map(function (r) {
    return '<li><b>' + escHtml_(r.name || '') + '</b>' +
      (r.roleLabel ? (' <span style="color:#666">(' + escHtml_(r.roleLabel) + ')</span>') : '') +
      ' &lt;' + escHtml_(r.email || '') + '&gt;</li>';
  }).join('');

  var linkNote = webAppBaseUrl_()
    ? '<p style="color:#52525b;font-size:12px">Each row includes an <b>Open in Cap Plan</b> link that jumps to that row after sign-in.</p>'
    : '<p style="color:#b91c1c;font-size:12px">Deep links unavailable — redeploy the web app so the deployment URL is available.</p>';

  var html =
    '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#18181b">' +
    '<p><b>Cap Plan notification</b> by ' + escHtml_(sender) +
    (payload.senderRole ? (' (' + escHtml_(payload.senderRole) + ')') : '') + '</p>' +
    (comment
      ? ('<p><b>Comment</b></p><p style="white-space:pre-wrap;background:#f4f4f5;padding:10px;border-radius:6px">' +
        escHtml_(comment) + '</p>')
      : '') +
    '<p><b>Selected rows (' + rows.length + ')</b></p>' +
    linkNote +
    '<div style="overflow:auto;max-width:100%">' +
    '<table style="border-collapse:collapse;width:100%;min-width:480px">' +
    '<thead><tr>' + htmlHead + '</tr></thead><tbody>' + htmlRows + '</tbody></table>' +
    '</div>' +
    '<p><b>Recipients</b></p><ul>' + htmlRecip + '</ul>' +
    '</div>';

  MailApp.sendEmail({
    to: toList.join(','),
    subject: subject,
    body: body,
    htmlBody: html,
    name: 'Cap Plan'
  });

  logAppEvent_('notify', sender, payload.senderRole || '',
    'to=' + toList.length + '; rows=' + rows.length + '; cols=' + columns.length + '; ' +
    toList.join(','));

  return { ok: true, sent: toList.length, subject: subject };
}

function escHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Session + PropertiesService identity for Execute-as-Me web apps. */
function resolveVisitorIdentity_(registeredAs, access, projects, employees, knownUsers, forceRegistered) {
  var ownerEmail = '';
  var visitorEmail = '';
  var visitorKey = '';
  try { ownerEmail = String(Session.getEffectiveUser().getEmail() || '').trim(); } catch (e0) {}
  try { visitorEmail = String(Session.getActiveUser().getEmail() || '').trim(); } catch (e1) {}
  try { visitorKey = String(Session.getTemporaryActiveUserKey() || '').trim(); } catch (e2) {}

  // Never treat the script owner as the visitor when ActiveUser is blank.
  // (Under Execute-as-Me, EffectiveUser is always the owner.)
  var out = {
    identity: '',
    visitorEmail: visitorEmail,
    ownerEmail: ownerEmail,
    visitorKey: visitorKey,
    source: 'none',
    autoDetected: false
  };

  var manual = String(registeredAs || '').trim();
  var manualKnown = manual && isKnownRegistrableUser_(manual, knownUsers)
    ? canonicalKnownUserName_(manual, knownUsers)
    : '';

  // Google-only login: never prefer forced/manual over session/remembered.
  // (forceRegistered is ignored for identity — kept only for API compat.)

  var fromEmail = visitorEmail
    ? identityFromVisitorEmail_(visitorEmail, access, projects, employees, knownUsers)
    : '';
  if (fromEmail) {
    out.identity = fromEmail;
    out.source = 'session';
    out.autoDetected = true;
    return out;
  }

  var fromKey = visitorKey ? loadRememberedIdentity_(visitorKey, visitorEmail) : '';
  if (fromKey && isKnownRegistrableUser_(fromKey, knownUsers)) {
    out.identity = canonicalKnownUserName_(fromKey, knownUsers);
    out.source = 'remembered';
    out.autoDetected = true;
    return out;
  }

  // Manual registeredAs is no longer accepted as a login path.
  if (manualKnown && visitorEmail) {
    // Allow only when the manual name matches the session email mapping attempt
    // (already failed above) — do not grant by typed name alone.
  }

  // Email seen but not yet mapped to Access — still expose it for the UI.
  if (visitorEmail) out.source = 'session_unmapped';
  return out;
}

/**
 * Map a Google account email to an Access / Employee / Project Lead display name.
 * Matches Access User Email, Access User Name (== email or derived display name),
 * Employees.name, Topics.project_lead.
 */
function identityFromVisitorEmail_(email, access, projects, employees, knownUsers) {
  var em = String(email || '').trim().toLowerCase();
  if (!em) return '';
  var derived = displayNameFromEmail_(email);
  var derivedLc = derived.toLowerCase();
  var local = em.indexOf('@') >= 0 ? em.split('@')[0] : em;

  var hit = '';
  (access || []).forEach(function (a) {
    if (hit || a.is_auto || a.isAuto) return;
    var mail = String(a.userEmail || '').trim().toLowerCase();
    var name = String(a.userName || a.user_name || a.user || '').trim();
    var nameLc = name.toLowerCase();
    if (mail && mail === em) hit = name || email;
    else if (nameLc && nameLc === em) hit = name;
    else if (nameLc && derivedLc && nameLc === derivedLc) hit = name;
    else if (nameLc && local && nameLc === local) hit = name;
  });
  if (hit) return canonicalKnownUserName_(hit, knownUsers) || hit;

  (employees || []).forEach(function (e) {
    if (hit) return;
    var name = String(e.name || '').trim();
    var nameLc = name.toLowerCase();
    if (!nameLc) return;
    if (nameLc === em || nameLc === derivedLc || nameLc === local) hit = name;
  });
  if (hit) return canonicalKnownUserName_(hit, knownUsers) || hit;

  (projects || []).forEach(function (p) {
    if (hit) return;
    var name = String(p.project_lead || '').trim();
    var nameLc = name.toLowerCase();
    if (!nameLc) return;
    if (nameLc === em || nameLc === derivedLc || nameLc === local) hit = name;
  });
  if (hit) return canonicalKnownUserName_(hit, knownUsers) || hit;

  if (isKnownRegistrableUser_(email, knownUsers)) {
    return canonicalKnownUserName_(email, knownUsers);
  }
  if (derived && isKnownRegistrableUser_(derived, knownUsers)) {
    return canonicalKnownUserName_(derived, knownUsers);
  }
  return '';
}

function canonicalKnownUserName_(name, knownUsers) {
  var want = String(name || '').trim().toLowerCase();
  for (var i = 0; i < (knownUsers || []).length; i++) {
    if (String(knownUsers[i] || '').trim().toLowerCase() === want) {
      return String(knownUsers[i]).trim();
    }
  }
  return String(name || '').trim();
}

function visitorProps_() {
  return PropertiesService.getScriptProperties();
}

function loadRememberedIdentity_(visitorKey, visitorEmail) {
  var props = visitorProps_();
  var byKey = visitorKey ? props.getProperty('capplan.visitor.' + visitorKey) : '';
  if (byKey) return String(byKey).trim();
  var em = String(visitorEmail || '').trim().toLowerCase();
  if (em) {
    var byEmail = props.getProperty('capplan.email.' + em);
    if (byEmail) return String(byEmail).trim();
  }
  return '';
}

/** Link this browser/Google visitor to a display name for later auto-login. */
function rememberVisitorIdentity_(identity, visitorEmail) {
  var name = String(identity || '').trim();
  if (!name) return { ok: false };
  var visitorKey = '';
  var email = String(visitorEmail || '').trim();
  try { visitorKey = String(Session.getTemporaryActiveUserKey() || '').trim(); } catch (e0) {}
  if (!email) {
    try { email = String(Session.getActiveUser().getEmail() || '').trim(); } catch (e1) {}
  }
  var props = visitorProps_();
  if (visitorKey) props.setProperty('capplan.visitor.' + visitorKey, name);
  if (email) props.setProperty('capplan.email.' + email.toLowerCase(), name);
  return { ok: true, identity: name, visitorKey: visitorKey, visitorEmail: email };
}

/** Client may call after manual registration to persist the mapping. */
function rememberRegistration(registeredAs) {
  var name = String(registeredAs || '').trim();
  if (!name) throw new Error('Name required');
  var access = readAccess_();
  var projects = readProjects_();
  var employees = readEmployees_();
  var known = knownRegistrableUsers_(access, projects, employees);
  if (!isKnownRegistrableUser_(name, known)) {
    throw new Error('Name must match a User Access entry (or Project Lead / Employee).');
  }
  var canonical = canonicalKnownUserName_(name, known);
  return rememberVisitorIdentity_(canonical, '');
}

/** Unique display names eligible for startup registration. */
function knownRegistrableUsers_(access, projects, employees) {
  var names = {};
  function add(v) {
    var s = String(v || '').trim();
    if (!s || s === '—') return;
    names[s] = true;
  }
  (access || []).forEach(function (a) {
    if (a.is_auto || a.isAuto) return;
    add(a.userName || a.user_name || a.user);
    add(a.userEmail);
  });
  (projects || []).forEach(function (p) { add(p.project_lead); });
  (employees || []).forEach(function (e) { add(e.name); });
  return Object.keys(names).sort(function (a, b) {
    return a.toLowerCase().localeCompare(b.toLowerCase());
  });
}

function isKnownRegistrableUser_(name, knownUsers) {
  var want = String(name || '').trim().toLowerCase();
  if (!want) return false;
  for (var i = 0; i < (knownUsers || []).length; i++) {
    if (String(knownUsers[i] || '').trim().toLowerCase() === want) return true;
  }
  return false;
}

/**
 * Delta save.
 * payload = {
 *   rows (dirty only preferred), originalRows (baselines for those rows),
 *   deletedIds, role, userName, settings?, compareSnapshotDate?,
 *   dirtyOnly: true
 * }
 * Returns a light ack (no full bootstrap) so saves stay under ~10s.
 */
function savePortfolio(payload) {
  var t0 = Date.now();
  var settings = readSettings_();
  var months = monthsFromSettings_(settings);
  var role = payload.role || 'Employee';
  // With Execute-as-Me, Session email is the script owner — attribute changes
  // to the registered end-user name from the client.
  var changedBy = String(payload.userName || '').trim();
  if (!changedBy) {
    throw new Error('Registration required: identify yourself before saving.');
  }

  var incoming = payload.rows || [];
  var originalById = indexById_(payload.originalRows || []);
  var access = readAccess_();
  var projects = readProjects_();
  var employees = readEmployees_();
  var actorName = changedBy;
  // Re-resolve role from registered identity (do not trust client role alone)
  var resolvedRole = resolveRole_(actorName, access, role, projects);
  if (resolvedRole) role = resolvedRole;
  // Trust client analysis for ∑ Alloc; only lock fields the role cannot edit.
  incoming = incoming.map(function (r) {
    var prev = originalById[r.id] || originalById[normalizeSheetId_(r.id)];
    if (role === 'Viewer') {
      return prev ? JSON.parse(JSON.stringify(prev)) : r;
    }
    if (role === 'Project Lead' && !rowInProjectLeadScopeServer_(r, actorName, role, access, projects)) {
      return prev ? JSON.parse(JSON.stringify(prev)) : r;
    }
    if (role === 'People Lead' && !rowInPeopleLeadScopeServer_(r, actorName, access, employees)) {
      return prev ? JSON.parse(JSON.stringify(prev)) : r;
    }
    if (role === 'People+Project Lead') {
      var inPL = rowInProjectLeadScopeServer_(r, actorName, role, access, projects);
      var inPeople = rowInPeopleLeadScopeServer_(r, actorName, access, employees);
      if (!inPL && !inPeople) {
        return prev ? JSON.parse(JSON.stringify(prev)) : r;
      }
      var locked = applyRoleLocksOnSave_(r, prev, role, months);
      if (!inPL && prev) {
        locked = applyOutOfProjectLeadLocks_(locked, prev, months);
      }
      if (!inPeople && prev) {
        locked = applyOutOfPeopleLeadLocks_(locked, prev, months);
      }
      return locked;
    }
    return applyRoleLocksOnSave_(r, prev, role, months);
  });

  var deleted = payload.deletedIds || [];
  if (role === 'Viewer' || role === 'Employee') {
    deleted = [];
  }
  var stats = writeDelta_(incoming, originalById, originalById, deleted, changedBy);

  if (payload.settings) {
    var next = Object.assign({}, settings, payload.settings);
    if (payload.compareSnapshotDate != null) {
      next.compareSnapshotDate = String(payload.compareSnapshotDate);
    }
    // Skip sheet write when nothing settings-related changed
    if (!settingsEqualForSave_(settings, next)) writeSettings_(next);
    settings = next;
  }

  var masterStats = { projects: 0, employees: 0, access: 0, costCenters: 0, budget: 0, booking: 0 };
  if (payload.projects && role === 'Admin') {
    writeProjectsDelta_(payload.projects);
    masterStats.projects = (payload.projects || []).length;
  }
  if (payload.employees && (role === 'Admin' || role === 'People Lead' || role === 'People+Project Lead')) {
    var employeesToWrite = sanitizeEmployeesForRole_(payload.employees, employees, role, actorName, access);
    writeEmployeesList_(employeesToWrite);
    masterStats.employees = (employeesToWrite || []).length;
  }
  if (payload.access && role === 'Admin') {
    writeAccessList_(payload.access);
    masterStats.access = (payload.access || []).length;
  }
  if (payload.costCenters && (role === 'Admin' || role === 'People Lead' || role === 'People+Project Lead')) {
    writeCostCenters_(payload.costCenters);
    masterStats.costCenters = (payload.costCenters || []).length;
  }
  // Budget: Project Lead / Admin; Booking: People Lead / Admin
  if (payload.budgetRecords && (role === 'Admin' || role === 'Project Lead' || role === 'People+Project Lead')) {
    writeFinanceRecords_(SHEETS.BUDGET, payload.budgetRecords, false);
    masterStats.budget = (payload.budgetRecords || []).length;
  }
  if (payload.bookingRecords && (role === 'Admin' || role === 'People Lead' || role === 'People+Project Lead')) {
    writeFinanceRecords_(SHEETS.BOOKING, payload.bookingRecords, true);
    masterStats.booking = (payload.bookingRecords || []).length;
  }

  SpreadsheetApp.flush();
  var elapsedMs = Date.now() - t0;
  var saveBits = [];
  saveBits.push((stats.rowsUpserted || 0) + ' rows');
  saveBits.push((stats.yearsUpserted || 0) + ' year-series');
  if (stats.rowsDeleted) saveBits.push(stats.rowsDeleted + ' deleted');
  if (masterStats.projects) saveBits.push(masterStats.projects + ' projects');
  if (masterStats.employees) saveBits.push(masterStats.employees + ' employees');
  if (masterStats.access) saveBits.push(masterStats.access + ' access');
  if (masterStats.costCenters) saveBits.push(masterStats.costCenters + ' cost centers');
  if (masterStats.budget) saveBits.push(masterStats.budget + ' budget');
  if (masterStats.booking) saveBits.push(masterStats.booking + ' booking');
  saveBits.push((elapsedMs / 1000).toFixed(1) + 's');
  logAppEvent_('save', changedBy, role, saveBits.join(' · '));

  return {
    ok: true,
    lightSave: true,
    saveStats: Object.assign({}, stats, { masters: masterStats }),
    stampedAt: stats.stampedAt || null,
    stampedBy: stats.stampedBy || '',
    stampedRowIds: stats.stampedRowIds || [],
    changeHistoryById: stats.changeHistoryById || {},
    savedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"),
    elapsedMs: elapsedMs,
    compareSnapshotDate: settings.compareSnapshotDate || '',
    settings: settings
  };
}

function settingsEqualForSave_(a, b) {
  if (!a || !b) return false;
  var keys = [
    'userName', 'defaultRole', 'colorDelta', 'yearFilters', 'colView', 'showDiffs',
    'compareLabel', 'compareSnapshotDate', 'frozenCols', 'hiddenColumns',
    'thirdMonthColumn', 'allocStatuses', 'showDeletedRows'
  ];
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (JSON.stringify(a[k] == null ? null : a[k]) !== JSON.stringify(b[k] == null ? null : b[k])) {
      return false;
    }
  }
  return true;
}

/** Full copy of live Current data + Current data_years into dated snapshot tabs. */
function createSnapshot(registeredAs, forceRegistered) {
  ensureSchema_();
  var dateNum = todaySnapshotDate_();
  var label = String(dateNum);

  var rowSheet = sheet_(SHEETS.DATA);
  var yearSheet = sheet_(SHEETS.YEARS);
  var snapRowName = label + SNAP_DATA_SUFFIX;
  var snapYearName = label + SNAP_YEARS_SUFFIX;

  copySheetContents_(rowSheet, ensureOrReplaceSheet_(snapRowName, DATA_HEADERS));
  copySheetContents_(yearSheet, ensureOrReplaceSheet_(snapYearName, YEAR_HEADERS));

  var settings = readSettings_();
  settings.compareLabel = formatSnapshotLabel_(dateNum);
  settings.compareSnapshotDate = String(dateNum);
  writeSettings_(settings);

  var boot = getBootstrap(registeredAs || '', !!forceRegistered);
  boot.snapshotCreated = dateNum;
  return boot;
}

function levelResources(payload) {
  ensureSchema_();
  var settings = readSettings_();
  var months = monthsFromSettings_(settings);
  var rows = (payload && payload.rows) || assembleRows_(
    readRowRecords_(), readYearRecords_(), readProjects_(), readEmployees_()
  );
  var delta = Number(settings.colorDelta) || 0;
  var changed = levelAllocations_(rows, months, delta);
  syncAvailabilityTotals_(rows);
  return { rows: rows, changed: changed };
}

function setCompareSnapshot(dateNum) {
  return loadCompareSnapshot(dateNum);
}

/** Load one snapshot pack for compare diffs (not full bootstrap). */
function loadCompareSnapshot(dateNum) {
  ensureSchema_();
  var settings = readSettings_();
  var projects = readProjects_();
  var employees = readEmployees_();
  var snapshotDates = listSnapshotDates_();
  var compareDate = defaultCompareSnapshotDate_(snapshotDates, dateNum);
  if (compareDate == null) {
    settings.compareSnapshotDate = '';
    settings.compareLabel = '';
    writeSettings_(settings);
    return {
      ok: true,
      compare: [],
      rowSnapshots: {},
      compareSnapshotDate: null,
      snapshotLabel: '',
      settings: settings
    };
  }
  var pack = loadSnapshotPack_(compareDate, projects, employees);
  settings.compareSnapshotDate = String(compareDate);
  settings.compareLabel = formatSnapshotLabel_(compareDate);
  writeSettings_(settings);
  return {
    ok: true,
    compare: pack.rows,
    rowSnapshots: pack.rowSnapshots,
    compareSnapshotDate: compareDate,
    snapshotLabel: settings.compareLabel,
    settings: settings
  };
}

/* ========== Schema ========== */

function ensureSchema_() {
  var ss = ss_();
  if (!ss) throw new Error('Open this script from a Google Spreadsheet (Extensions → Apps Script).');

  ensureSheet_(sheetName_(SHEETS.PROJECTS), PROJECT_HEADERS);
  ensureSheet_(sheetName_(SHEETS.EMPLOYEES), EMPLOYEE_HEADERS);
  ensureSheet_(sheetName_(SHEETS.DATA), DATA_HEADERS);
  ensureSheet_(sheetName_(SHEETS.YEARS), YEAR_HEADERS);
  ensureSheet_(SHEETS.META, ['Field', 'Value', 'Sort Order']);
  ensureSheet_(SHEETS.SETTINGS, [
    'userName', 'defaultRole', 'colorDelta', 'yearFilters', 'colView',
    'showDiffs', 'compareLabel', 'compareSnapshotDate', 'frozenCols', 'hiddenColumns',
    'thirdMonthColumn'
  ]);
  ensureSheet_(sheetName_(SHEETS.ACCESS), ACCESS_HEADERS);
  ensureSheet_(SHEETS.COST_CENTERS, COST_CENTER_HEADERS);
  ensureSheet_(SHEETS.BUDGET, FINANCE_MONTH_HEADERS);
  ensureSheet_(SHEETS.BOOKING, BOOKING_HEADERS);
}

/** Resolve canonical sheet name, falling back to legacy aliases. */
function sheetName_(preferred) {
  var ss = ss_();
  if (ss.getSheetByName(preferred)) return preferred;
  var aliases = SHEET_ALIASES[preferred] || [];
  for (var i = 0; i < aliases.length; i++) {
    if (ss.getSheetByName(aliases[i])) return aliases[i];
  }
  return preferred;
}

function sheet_(preferred) {
  return ss_().getSheetByName(sheetName_(preferred));
}

function ensureSheet_(name, headers) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    var existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
    if (!existing[0]) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    } else {
      var have = {};
      existing.forEach(function (h) {
        if (h != null && String(h).trim() !== '') have[String(h).trim()] = true;
      });
      var missing = headers.filter(function (h) { return !have[h]; });
      if (missing.length) {
        sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
      }
    }
  }
  return sh;
}

function ensureOrReplaceSheet_(name, headers) {
  var ss = ss_();
  var existing = ss.getSheetByName(name);
  if (existing) ss.deleteSheet(existing);
  var sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  return sh;
}

function copySheetContents_(fromSh, toSh) {
  var lastRow = fromSh.getLastRow();
  var lastCol = fromSh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;
  var values = fromSh.getRange(1, 1, lastRow, lastCol).getValues();
  toSh.clear();
  toSh.getRange(1, 1, values.length, values[0].length).setValues(values);
  toSh.setFrozenRows(1);
}

function seedIfEmpty_() {
  var shP = sheet_(SHEETS.PROJECTS);
  if (shP.getLastRow() > 1) return;

  shP.getRange(2, 1, 2, PROJECT_HEADERS.length).setValues([
    ['P_DEMO1', 'Demo Topic Alpha', 'Program A', 'WBS-1', 'IND-1', 'Portfolio X', 'Cat 1', 'Active', '', '', '', 'Alex Lead', 'Berlin', 'direct'],
    ['P_DEMO2', 'Demo Topic Beta', 'Program B', 'WBS-2', 'IND-2', 'Portfolio Y', 'Cat 2', 'Active', '', '', '', 'Alex Lead', 'Munich', 'direct']
  ]);

  var shE = sheet_(SHEETS.EMPLOYEES);
  shE.getRange(2, 1, 2, EMPLOYEE_HEADERS.length).setValues([
    ['E_DEMO1', 'Engineer', 'Ada Example', '', 'Berlin', 'Group A', 'project-allocated'],
    ['E_DEMO2', 'PM', 'Ben Example', '', 'Munich', 'Group B', 'project-allocated']
  ]);

  var metaRows = [
    ['Role', 'Engineer', 1], ['Role', 'PM', 2], ['Role', 'Analyst', 3],
    ['Program', 'Program A', 1], ['Program', 'Program B', 2],
    ['Portfolio', 'Portfolio X', 1], ['Portfolio', 'Portfolio Y', 2],
    ['Category', 'Cat 1', 1], ['Category', 'Cat 2', 2],
    ['Status', 'Active', 1], ['Status', 'On Hold', 2],
    ['Group', 'Group A', 1], ['Group', 'Group B', 2],
    ['FTE Type', 'project-allocated', 1], ['FTE Type', 'Internal', 2], ['FTE Type', 'External', 3],
    ['Location', 'Berlin', 1], ['Location', 'Munich', 2], ['Location', 'Indianapolis', 3],
    ['Transfer Type', 'direct', 1], ['Transfer Type', 'indirect', 2],
    ['Transfer Type', 'finance', 3], ['Transfer Type', 'none', 4],
    ['Internal Cost Center', 'CC-100', 1], ['Internal Activity Type', 'ACT-1', 1],
    ['External Cost Center', 'CC-E100', 1], ['External Activity Type', 'ACT-E1', 1]
  ];
  ss_().getSheetByName(SHEETS.META)
    .getRange(2, 1, metaRows.length, 3).setValues(metaRows);

  var shS = ss_().getSheetByName(SHEETS.SETTINGS);
  shS.getRange(2, 1, 1, 10).setValues([[
    Session.getActiveUser().getEmail() || 'Admin',
    'Admin',
    0,
    '2026',
    'Both',
    true,
    '',
    '',
    3,
    '[]'
  ]]);

  var rows = [
    newRow_('Demand/Allocation', 'P_DEMO1', 'Demo Topic Alpha', 'Engineer', 'E_DEMO1', 'Ada Example', 1),
    newRow_('Timeline/Header', 'P_DEMO1', 'Demo Topic Alpha', '', '', '', 0),
    newRow_('Demand/Allocation', 'P_DEMO2', 'Demo Topic Beta', 'PM', 'E_DEMO2', 'Ben Example', 1),
    newRow_('Availability', '', '', 'Engineer', 'E_DEMO1', 'Ada Example', 0),
    newRow_('Availability', '', '', 'PM', 'E_DEMO2', 'Ben Example', 0)
  ];
  rows[0].months['2026-01-01'] = { demand: 0.5, allocation: 0.5 };
  rows[0].months['2026-02-01'] = { demand: 0.5, allocation: 0.3 };
  rows[2].months['2026-01-01'] = { demand: 1, allocation: 0.8 };
  rows[3].months['2026-01-01'] = { demand: 1, allocation: null };
  rows[3].months['2026-02-01'] = { demand: 1, allocation: null };
  rows[4].months['2026-01-01'] = { demand: 1, allocation: null };
  rows[1].phases = { '2026-01-01': 'Design', '2026-02-01': 'Design', '2026-03-01': 'Build' };

  syncAvailabilityTotals_(rows);
  writeDelta_(rows, {}, {}, [], 'Seed');
}

function newRow_(source, projectId, topic, role, empId, name, sort) {
  return {
    id: Utilities.getUuid(),
    sort_order: sort,
    source: source,
    version: 'Cur',
    project_id: projectId || null,
    project_topic: topic || null,
    alternative_financing: null,
    alternative_financing_topic: null,
    workpackage: null,
    role: role || null,
    employee_id: empId || null,
    name: name || null,
    notes_project_lead: null,
    notes_people_lead: null,
    notes_employee: null,
    comments: null,
    changed_by: null,
    updated_at: null,
    is_availability: source === SOURCES.AVAIL,
    is_header: source === SOURCES.HEADER,
    months: {},
    phases: {}
  };
}

/* ========== Settings / Access / Role ========== */

function readSettings_() {
  var sh = ss_().getSheetByName(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  var headers = data[0] || [];
  var row = data[1] || [];
  var o = {};
  headers.forEach(function (h, i) { o[h] = row[i]; });
  var hiddenColumns = [];
  if (o.hiddenColumns) {
    try {
      hiddenColumns = typeof o.hiddenColumns === 'string'
        ? JSON.parse(String(o.hiddenColumns))
        : o.hiddenColumns;
    } catch (e) {
      hiddenColumns = String(o.hiddenColumns).split(',').map(function (s) { return s.trim(); }).filter(Boolean);
    }
  }
  if (!Array.isArray(hiddenColumns)) hiddenColumns = [];
  var colView = String(o.colView || 'Demand/Allocation');
  if (colView === 'Both') colView = 'Demand/Allocation';
  var allowedViews = {
    'Demand': 1, 'Allocation': 1, 'Employee': 1,
    'Demand/Allocation': 1, 'Demand/Allocation/Employee': 1
  };
  var third = String(o.thirdMonthColumn || 'none').toLowerCase();
  if (!allowedViews[colView]) {
    if (String(colView).indexOf('Budget') >= 0) colView = 'Demand/Allocation';
    else colView = 'Demand/Allocation';
  }
  if (third === 'employee_info' && colView === 'Demand/Allocation') {
    colView = 'Demand/Allocation/Employee';
  }
  return {
    userName: String(o.userName || ''),
    defaultRole: String(o.defaultRole || 'Admin'),
    colorDelta: Number(o.colorDelta) || 0,
    yearFilters: String(o.yearFilters || '2026').split(',').map(function (s) { return s.trim(); }).filter(Boolean),
    colView: colView,
    showDiffs: o.showDiffs !== false && o.showDiffs !== 'FALSE' && o.showDiffs !== 0,
    compareLabel: String(o.compareLabel || ''),
    compareSnapshotDate: o.compareSnapshotDate ? String(o.compareSnapshotDate) : '',
    frozenCols: Number(o.frozenCols) || 3,
    hiddenColumns: hiddenColumns,
    thirdMonthColumn: 'none'
  };
}

function writeSettings_(s) {
  var sh = ss_().getSheetByName(SHEETS.SETTINGS);
  // Keep header row in sync when new columns are added.
  sh.getRange(1, 1, 1, 11).setValues([[
    'userName', 'defaultRole', 'colorDelta', 'yearFilters', 'colView',
    'showDiffs', 'compareLabel', 'compareSnapshotDate', 'frozenCols', 'hiddenColumns',
    'thirdMonthColumn'
  ]]);
  var colView = String(s.colView || 'Demand/Allocation');
  if (colView === 'Both') colView = 'Demand/Allocation';
  sh.getRange(2, 1, 1, 11).setValues([[
    s.userName || '',
    s.defaultRole || 'Admin',
    Number(s.colorDelta) || 0,
    (s.yearFilters || ['2026']).join(','),
    colView,
    s.showDiffs !== false,
    s.compareLabel || '',
    s.compareSnapshotDate || '',
    Number(s.frozenCols) || 3,
    JSON.stringify(s.hiddenColumns || []),
    'none'
  ]]);
}

function normalizeAccessRoleType_(t) {
  var s = String(t || '').toLowerCase().replace(/_/g, ' ').trim();
  if (s.indexOf('admin') >= 0) return 'admin';
  if (s.indexOf('viewer') >= 0 || s.indexOf('read only') >= 0 || s.indexOf('readonly') >= 0) return 'viewer';
  if (s.indexOf('people') >= 0) return 'people_lead';
  if (s.indexOf('project') >= 0) return 'project_lead';
  if (s.indexOf('employee') >= 0) return 'employee';
  return s || 'viewer';
}

function normalizeAccessScopeType_(t) {
  var s = String(t || '').toLowerCase().trim();
  if (s === 'global' || s === 'project' || s === 'program' || s === 'department' || s === 'unit') return s;
  return 'global';
}

function readAccess_() {
  var sh = sheet_(SHEETS.ACCESS);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var h = values[0];
  var idx = headerIndex_(h);
  // Databank: User Name, Role, Scope, Scope Value, Source
  // Also accepts extended / legacy column names.
  return values.slice(1).filter(function (r) {
    return r[0] || col_(r, idx, ['User Email', 'user', 'User Name', 'id', 'ID']);
  }).map(function (r, i) {
    var userName = strCol_(r, idx, ['User Name', 'user']) || '';
    var userEmail = strCol_(r, idx, ['User Email']) || '';
    var roleType = normalizeAccessRoleType_(strCol_(r, idx, ['Role Type', 'Role', 'role']) || '');
    var scopeType = strCol_(r, idx, ['Scope Type', 'Scope', 'scope_type']) || '';
    var scopeValue = strCol_(r, idx, ['Scope Value', 'scope_value']) || '';
    if (!scopeType) {
      var proj = strCol_(r, idx, ['project']);
      var prog = strCol_(r, idx, ['programm', 'program']);
      var dept = strCol_(r, idx, ['department']);
      var unit = strCol_(r, idx, ['unit']);
      if (proj) { scopeType = 'project'; scopeValue = proj; }
      else if (prog) { scopeType = 'program'; scopeValue = prog; }
      else if (dept) { scopeType = 'department'; scopeValue = dept; }
      else if (unit) { scopeType = 'unit'; scopeValue = unit; }
      else { scopeType = 'global'; }
    }
    scopeType = normalizeAccessScopeType_(scopeType);
    if (roleType === 'admin' || roleType === 'viewer') {
      scopeType = 'global';
      scopeValue = '';
    }
    var source = strCol_(r, idx, ['Source']) || '';
    var isAutoRaw = col_(r, idx, ['Is Auto', 'is_auto']);
    var isAuto = isAutoRaw === true || isAutoRaw === 1 ||
      String(isAutoRaw).toUpperCase() === 'TRUE' || String(isAutoRaw) === '1' ||
      String(source).toLowerCase() === 'auto';
    var id = strCol_(r, idx, ['id', 'ID']);
    if (!id) {
      id = 'access_' + String(userName + '_' + roleType + '_' + scopeType + '_' + scopeValue)
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '_')
        .replace(/^_|_$/g, '')
        .slice(0, 48);
    }
    if (!id) id = 'A_' + (i + 1);
    var sortOrder = Number(col_(r, idx, ['sort_order', 'Sort Order'])) || (i + 1);
    return {
      id: id || ('A_' + (i + 1)),
      user_name: userName,
      userName: userName,
      userEmail: userEmail,
      role_type: roleType,
      roleType: roleType,
      role: roleType,
      scope_type: scopeType,
      scopeType: scopeType,
      scope_value: scopeValue || null,
      scopeValue: scopeValue || null,
      is_auto: isAuto,
      isAuto: isAuto,
      source: source || (isAuto ? 'Auto' : 'Manual'),
      sort_order: sortOrder,
      project: scopeType === 'project' ? (scopeValue || '') : '',
      program: scopeType === 'program' ? (scopeValue || '') : '',
      programm: scopeType === 'program' ? (scopeValue || '') : '',
      department: scopeType === 'department' ? (scopeValue || '') : '',
      unit: scopeType === 'unit' ? (scopeValue || '') : ''
    };
  });
}

/** Replace Access sheet with structured records (Admin UI). */
function saveAccess(records, registeredAs, forceRegistered) {
  ensureSchema_();
  var list = (records || []).map(function (r, i) {
    var roleType = normalizeAccessRoleType_(r.role_type || r.roleType || r.role || '');
    var scopeType = normalizeAccessScopeType_(r.scope_type || r.scopeType || 'global');
    if (roleType === 'admin' || roleType === 'viewer') scopeType = 'global';
    var scopeValue = scopeType === 'global'
      ? ''
      : String(r.scope_value != null ? r.scope_value : (r.scopeValue || '')).trim();
    var userName = String(r.user_name || r.userName || r.user || '').trim();
    var userEmail = String(r.userEmail || r.user_email || '').trim();
    var isAuto = !!(r.is_auto || r.isAuto);
    return {
      id: String(r.id || ('access_' + Utilities.getUuid().slice(0, 12))),
      user_name: userName,
      userEmail: userEmail,
      role_type: roleType,
      scope_type: scopeType,
      scope_value: scopeValue || null,
      is_auto: false,
      source: 'Manual',
      sort_order: Number(r.sort_order) || (i + 1)
    };
  }).filter(function (r) { return r.user_name; });

  // Ensure built-in admin exists.
  var hasDefault = list.some(function (r) {
    return r.id === DEFAULT_ADMIN_ACCESS_ID ||
      (String(r.user_name).toLowerCase() === DEFAULT_ADMIN_USER_NAME && r.role_type === 'admin');
  });
  if (!hasDefault) {
    list.unshift({
      id: DEFAULT_ADMIN_ACCESS_ID,
      user_name: DEFAULT_ADMIN_USER_NAME,
      role_type: 'admin',
      scope_type: 'global',
      scope_value: null,
      is_auto: false,
      sort_order: 0
    });
  }

  writeAccessList_(list);
  return getBootstrap(registeredAs || '', !!forceRegistered);
}

function writeAccessList_(list) {
  var sh = sheet_(SHEETS.ACCESS);
  if (!sh) return;
  // Keep existing header shape when already present; otherwise write canonical headers.
  if (sh.getLastRow() < 1 || !sh.getRange(1, 1).getValue()) {
    sh.getRange(1, 1, 1, ACCESS_HEADERS.length).setValues([ACCESS_HEADERS]);
    sh.setFrozenRows(1);
  }
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var idx = headerIndex_(headers);
  var width = Math.max(headers.length, ACCESS_HEADERS.length);

  function accessKey_(userName, roleType, scopeType, scopeValue) {
    return [
      String(userName || '').trim().toLowerCase(),
      String(roleType || '').trim().toLowerCase(),
      String(scopeType || '').trim().toLowerCase(),
      String(scopeValue || '').trim().toLowerCase()
    ].join('|');
  }

  var keyToRow = {};
  for (var i = 1; i < values.length; i++) {
    var u = strCol_(values[i], idx, ['User Name', 'user']) || String(values[i][0] || '');
    var role = normalizeAccessRoleType_(strCol_(values[i], idx, ['Role Type', 'Role', 'role']) || '');
    var scope = normalizeAccessScopeType_(strCol_(values[i], idx, ['Scope Type', 'Scope', 'scope_type']) || 'global');
    var sval = strCol_(values[i], idx, ['Scope Value', 'scope_value']) || '';
    keyToRow[accessKey_(u, role, scope, sval)] = i + 1;
  }

  var seen = {};
  var toAppend = [];
  (list || []).forEach(function (r) {
    var scopeType = r.scope_type || 'global';
    var scopeValue = r.scope_value || '';
    var source = 'Manual';
    var line = new Array(width).fill('');
    setAny_(line, idx, ['User Name', 'user'], r.user_name || '');
    setAny_(line, idx, ['User Email'], r.userEmail || r.user_email || '');
    setAny_(line, idx, ['Role Type', 'Role', 'role'], r.role_type || '');
    setAny_(line, idx, ['Scope Type', 'Scope', 'scope_type'], scopeType);
    setAny_(line, idx, ['Scope Value', 'scope_value'], scopeValue);
    setAny_(line, idx, ['Source'], source);
    setAny_(line, idx, ['id', 'ID'], r.id || '');
    setAny_(line, idx, ['is_auto', 'Is Auto'], !!r.is_auto);
    setAny_(line, idx, ['sort_order', 'Sort Order'], Number(r.sort_order) || 0);
    // If sheet still has only positional ACCESS_HEADERS and setAny missed, fall back.
    if (idx['User Name'] == null && idx.user == null && headers.length >= 5 && !line[0]) {
      line = [r.user_name || '', r.role_type || '', scopeType, scopeValue, source];
    }
    var key = accessKey_(r.user_name, r.role_type, scopeType, scopeValue);
    seen[key] = true;
    var sheetRow = keyToRow[key];
    if (sheetRow) {
      if (!sheetRowEquals_(values[sheetRow - 1], line.slice(0, values[sheetRow - 1].length))) {
        sh.getRange(sheetRow, 1, 1, line.length).setValues([line]);
      }
    } else {
      toAppend.push(line);
    }
  });

  var toDelete = [];
  Object.keys(keyToRow).forEach(function (key) {
    if (seen[key]) return;
    // Preserve auto Access rows when the UI only sends manual rows
    var sheetRow = keyToRow[key];
    var isAuto = false;
    if (sheetRow && values[sheetRow - 1]) {
      var autoRaw = col_(values[sheetRow - 1], idx, ['Is Auto', 'is_auto']);
      isAuto = autoRaw === true || autoRaw === 1 || String(autoRaw).toLowerCase() === 'true' ||
        String(strCol_(values[sheetRow - 1], idx, ['Source']) || '').toLowerCase() === 'auto';
    }
    if (!isAuto) toDelete.push(sheetRow);
  });
  deleteSheetRows_(sh, toDelete);
  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
  }
}

/** Auto Project Lead access rows are managed outside Cap Plan — do not regenerate. */
function syncAutoProjectLeads_(projects) {
  return;
}

function allowedRolesFor_(email, access, projects) {
  var roles = {};
  var em = String(email || '').toLowerCase().trim();
  var formatted = displayNameFromEmail_(email).toLowerCase().trim();
  (access || []).forEach(function (a) {
    if (a.is_auto || a.isAuto) return;
    var mail = String(a.userEmail || '').toLowerCase().trim();
    var name = String(a.userName || a.user_name || '').toLowerCase().trim();
    var hit = false;
    if (em) {
      if (mail && mail === em) hit = true;
      if (name && name === em) hit = true;
      if (formatted && name && name === formatted) hit = true;
      if (name && em.indexOf('@') < 0 && name === em) hit = true;
    }
    if (!hit && normalizeRole_(a.roleType || a.role_type || a.role) === 'Admin' && !mail && !em) hit = true;
    if (!hit) return;
    var rt = normalizeRole_(a.roleType || a.role_type || a.role);
    if (rt) roles[rt] = true;
  });
  // Topics.project_lead adds Project Lead unless Access already grants Admin
  if (!roles.Admin) {
    (projects || []).forEach(function (p) {
      var lead = String(p.project_lead || '').trim().toLowerCase();
      if (!lead) return;
      if ((em && lead === em) || (formatted && lead === formatted)) {
        roles['Project Lead'] = true;
      }
    });
  }
  if (roles['Project Lead'] && roles['People Lead']) {
    delete roles['Project Lead'];
    delete roles['People Lead'];
    roles['People+Project Lead'] = true;
  }
  var list = Object.keys(roles);
  if (!list.length) {
    return ['Employee'];
  }
  var order = ['Admin', 'People+Project Lead', 'Project Lead', 'People Lead', 'Viewer', 'Employee'];
  list.sort(function (a, b) {
    return order.indexOf(a) - order.indexOf(b);
  });
  return list;
}

function resolveRole_(email, access, fallback, projects) {
  var allowed = allowedRolesFor_(email, access, projects);
  if (fallback && allowed.indexOf(fallback) >= 0) return fallback;
  return allowed[0] || 'Employee';
}

function normalizeRole_(t) {
  var s = String(t || '').toLowerCase().replace(/_/g, ' ').trim();
  if (s.indexOf('admin') >= 0) return 'Admin';
  if (s.indexOf('viewer') >= 0 || s.indexOf('read only') >= 0 || s.indexOf('readonly') >= 0) return 'Viewer';
  if (s.indexOf('people') >= 0 && s.indexOf('project') >= 0) return 'People+Project Lead';
  if (s.indexOf('project') >= 0) return 'Project Lead';
  if (s.indexOf('people') >= 0) return 'People Lead';
  if (s.indexOf('employee') >= 0) return 'Employee';
  return '';
}

/* ========== Masters ========== */

function readProjects_() {
  var sh = sheet_(SHEETS.PROJECTS);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var h = values[0];
  var idx = headerIndex_(h);
  return values.slice(1).filter(function (r) {
    return isMasterId_(col_(r, idx, ['Project_ID', 'ID']), 'P_');
  }).map(function (r, i) {
    return {
      id: String(col_(r, idx, ['Project_ID', 'ID'])),
      project_topic: strCol_(r, idx, ['Project/Topic']) || '',
      program: strCol_(r, idx, ['Program']) || '',
      wbs_ma: strCol_(r, idx, ['WBS MA']) || '',
      wbs_indy: strCol_(r, idx, ['WBS Indy']) || '',
      portfolio: strCol_(r, idx, ['Portfolio']) || '',
      category: strCol_(r, idx, ['Category']) || '',
      status: strCol_(r, idx, ['Status']) || '',
      project_lead: strCol_(r, idx, ['Project Lead']) || '',
      project_location: strCol_(r, idx, ['Project Location']) || '',
      link_project: strCol_(r, idx, ['Link Project']) || '',
      link_resources: strCol_(r, idx, ['Lin Resource', 'Link Resources', 'Link Resource']) || '',
      comment: strCol_(r, idx, ['Comment']) || '',
      transfer_type: normalizeTransferType_(strCol_(r, idx, ['Transfer Type', 'Standard Transfer'])),
      sort_order: Number(col_(r, idx, ['Sort Order'])) || (i + 1)
    };
  });
}

function writeProjectsDelta_(projects) {
  var sh = sheet_(SHEETS.PROJECTS);
  if (sh.getLastRow() < 1 || !sh.getRange(1, 1).getValue()) {
    sh.getRange(1, 1, 1, PROJECT_HEADERS.length).setValues([PROJECT_HEADERS]);
    sh.setFrozenRows(1);
  }
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var idx = headerIndex_(headers);
  var width = Math.max(headers.length, PROJECT_HEADERS.length);
  var idToRow = {};
  for (var i = 1; i < values.length; i++) {
    var id = normalizeSheetId_(col_(values[i], idx, ['Project_ID', 'ID']));
    if (id) idToRow[id] = i + 1;
  }

  var seen = {};
  var toAppend = [];
  (projects || []).forEach(function (p) {
    var id = normalizeSheetId_(p.id) || ('P_' + Utilities.getUuid().slice(0, 8));
    seen[id] = true;
    var line = new Array(width).fill('');
    setAny_(line, idx, ['Project_ID', 'ID'], id);
    setAny_(line, idx, ['Project/Topic'], p.project_topic || '');
    setAny_(line, idx, ['Program'], p.program || '');
    setAny_(line, idx, ['WBS MA'], p.wbs_ma || '');
    setAny_(line, idx, ['WBS Indy'], p.wbs_indy || '');
    setAny_(line, idx, ['Portfolio'], p.portfolio || '');
    setAny_(line, idx, ['Category'], p.category || '');
    setAny_(line, idx, ['Status'], p.status || '');
    setAny_(line, idx, ['Link Project'], p.link_project || '');
    setAny_(line, idx, ['Lin Resource', 'Link Resources', 'Link Resource'], p.link_resources || '');
    setAny_(line, idx, ['Comment'], p.comment || '');
    setAny_(line, idx, ['Project Lead'], p.project_lead || '');
    setAny_(line, idx, ['Project Location'], p.project_location || '');
    setAny_(line, idx, ['Transfer Type', 'Standard Transfer'], normalizeTransferType_(p.transfer_type || p.transferType));
    if (idx['Project_ID'] == null && idx.ID == null && !line[0]) {
      line = [
        id, p.project_topic || '', p.program || '', p.wbs_ma || '', p.wbs_indy || '',
        p.portfolio || '', p.category || '', p.status || '',
        p.link_project || '', p.link_resources || '', p.comment || '',
        p.project_lead || '', p.project_location || '',
        normalizeTransferType_(p.transfer_type || p.transferType)
      ];
    }
    var sheetRow = idToRow[id];
    if (sheetRow) {
      if (!sheetRowEquals_(values[sheetRow - 1], line.slice(0, values[sheetRow - 1].length))) {
        sh.getRange(sheetRow, 1, 1, line.length).setValues([line]);
      }
    } else {
      toAppend.push(line);
    }
  });

  var toDelete = [];
  Object.keys(idToRow).forEach(function (id) {
    if (!seen[id]) toDelete.push(idToRow[id]);
  });
  deleteSheetRows_(sh, toDelete);
  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
  }
}

function saveProjects(projects, registeredAs, forceRegistered) {
  ensureSchema_();
  writeProjectsDelta_(projects || []);
  return getBootstrap(registeredAs || '', !!forceRegistered);
}

function readEmployees_() {
  var sh = sheet_(SHEETS.EMPLOYEES);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var h = values[0];
  var idx = headerIndex_(h);
  return values.slice(1).filter(function (r) {
    return isMasterId_(col_(r, idx, ['Employee_ID', 'ID']), 'E_');
  }).map(function (r) {
    var group = strCol_(r, idx, ['Group']) || '';
    var role = strCol_(r, idx, ['Role']) || '';
    return {
      id: String(col_(r, idx, ['Employee_ID', 'ID'])),
      name: strCol_(r, idx, ['Name']) || '',
      role: role,
      group_name: group,
      // Department is role-derived — never fall back to Group
      department: roleDepartmentFromRole_(role),
      unit: roleUnitFromRole_(role),
      fte_type: strCol_(r, idx, ['FTE Type']) || '',
      people_location: strCol_(r, idx, ['People Location']) || '',
      comments: strCol_(r, idx, ['Comments']) || ''
    };
  });
}

function saveEmployee(emp) {
  ensureSchema_();
  var list = readEmployees_();
  var found = false;
  for (var i = 0; i < list.length; i++) {
    if (list[i].id === emp.id) {
      list[i] = Object.assign(list[i], emp);
      found = true;
      break;
    }
  }
  if (!found) {
    emp.id = emp.id || ('E_' + Utilities.getUuid().slice(0, 8));
    list.push(emp);
  }
  writeEmployeesList_(list);
  return emp;
}

/** Alias used by Allocation availability people-master edits. */
function upsertEmployee(emp) {
  return saveEmployee(emp);
}

function writeEmployeesList_(list) {
  var sh = sheet_(SHEETS.EMPLOYEES);
  if (sh.getLastRow() < 1 || !sh.getRange(1, 1).getValue()) {
    sh.getRange(1, 1, 1, EMPLOYEE_HEADERS.length).setValues([EMPLOYEE_HEADERS]);
    sh.setFrozenRows(1);
  }
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var idx = headerIndex_(headers);
  var width = Math.max(headers.length, EMPLOYEE_HEADERS.length);
  var idToRow = {};
  for (var i = 1; i < values.length; i++) {
    var id = normalizeSheetId_(col_(values[i], idx, ['Employee_ID', 'ID']));
    if (id) idToRow[id] = i + 1;
  }

  var seen = {};
  var toAppend = [];
  (list || []).forEach(function (e) {
    var id = normalizeSheetId_(e.id) || ('E_' + Utilities.getUuid().slice(0, 8));
    seen[id] = true;
    var line = new Array(width).fill('');
    setAny_(line, idx, ['Employee_ID', 'ID'], id);
    setAny_(line, idx, ['Role'], e.role || '');
    setAny_(line, idx, ['Name'], e.name || '');
    setAny_(line, idx, ['Comments'], e.comments || '');
    setAny_(line, idx, ['People Location'], e.people_location || '');
    setAny_(line, idx, ['Group'], e.group_name || '');
    setAny_(line, idx, ['FTE Type'], e.fte_type || '');
    if (idx['Employee_ID'] == null && idx.ID == null && !line[0]) {
      line = [id, e.role || '', e.name || '', e.comments || '',
        e.people_location || '', e.group_name || '', e.fte_type || ''];
    }
    var sheetRow = idToRow[id];
    if (sheetRow) {
      if (!sheetRowEquals_(values[sheetRow - 1], line.slice(0, values[sheetRow - 1].length))) {
        sh.getRange(sheetRow, 1, 1, line.length).setValues([line]);
      }
    } else {
      toAppend.push(line);
    }
  });

  var toDelete = [];
  Object.keys(idToRow).forEach(function (id) {
    if (!seen[id]) toDelete.push(idToRow[id]);
  });
  deleteSheetRows_(sh, toDelete);
  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
  }
}

function normalizeTransferType_(v) {
  var s = String(v || '').trim().toLowerCase();
  if (TRANSFER_TYPES.indexOf(s) >= 0) return s;
  return 'direct';
}

function normalizeCostCenterLevel_(v) {
  var s = String(v || '').trim().toLowerCase();
  if (COST_CENTER_LEVELS.indexOf(s) >= 0) return s;
  return 'role';
}

function readCostCenters_() {
  var sh = ss_().getSheetByName(SHEETS.COST_CENTERS);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getValues();
  var idx = headerIndex_(values[0]);
  return values.slice(1).filter(function (r) {
    return String(col_(r, idx, ['CC_ID', 'ID']) || '').trim() ||
      String(col_(r, idx, ['Level Value']) || '').trim();
  }).map(function (r, i) {
    return {
      id: String(col_(r, idx, ['CC_ID', 'ID']) || ('CC_' + (i + 1))),
      level: normalizeCostCenterLevel_(strCol_(r, idx, ['Level'])),
      level_value: strCol_(r, idx, ['Level Value']) || '',
      internal_cost_center: strCol_(r, idx, ['Internal Cost Center']) || '',
      internal_activity_type: strCol_(r, idx, ['Internal Activity Type']) || '',
      external_cost_center: strCol_(r, idx, ['External Cost Center']) || '',
      external_activity_type: strCol_(r, idx, ['External Activity Type']) || '',
      rbs_skill_name: strCol_(r, idx, ['RBS Skill Name']) || '',
      hourly_cost: strCol_(r, idx, ['Hourly Cost']) || ''
    };
  });
}

function writeCostCenters_(list) {
  ensureSheet_(SHEETS.COST_CENTERS, COST_CENTER_HEADERS);
  var sh = ss_().getSheetByName(SHEETS.COST_CENTERS);
  var lines = [[].concat(COST_CENTER_HEADERS)];
  (list || []).forEach(function (c, i) {
    if (!c) return;
    var id = String(c.id || ('CC_' + Utilities.getUuid().slice(0, 8)));
    var levelValue = String(c.level_value || c.levelValue || '').trim();
    if (!levelValue && !String(c.internal_cost_center || '').trim()) return;
    lines.push([
      id,
      normalizeCostCenterLevel_(c.level),
      levelValue,
      c.internal_cost_center || c.internalCostCenter || '',
      c.internal_activity_type || c.internalActivityType || '',
      c.external_cost_center || c.externalCostCenter || '',
      c.external_activity_type || c.externalActivityType || '',
      c.rbs_skill_name || c.rbsSkillName || '',
      c.hourly_cost || c.hourlyCost || ''
    ]);
  });
  sh.clear();
  sh.getRange(1, 1, lines.length, COST_CENTER_HEADERS.length).setValues(lines);
  sh.setFrozenRows(1);
}

function emptyMonthMap_() {
  var o = {};
  MONTH_MM.forEach(function (mm) { o[mm] = null; });
  return o;
}

function readMonthCellsFromRow_(r, idx, prefix) {
  var months = emptyMonthMap_();
  MONTH_HEADERS_LONG.forEach(function (label, i) {
    var mm = MONTH_MM[i];
    var names = prefix
      ? [prefix + label, prefix + mm, prefix + 'm' + mm]
      : [label, 'm' + mm, 'M' + String(i + 1), 'M' + mm];
    var v = col_(r, idx, names);
    months[mm] = (v === '' || v == null) ? null : toNum_(v);
  });
  return months;
}

function writeMonthCellsToLine_(line, idx, months, prefix) {
  MONTH_HEADERS_LONG.forEach(function (label, i) {
    var mm = MONTH_MM[i];
    var names = prefix
      ? [prefix + label, prefix + mm]
      : [label, 'm' + mm];
    var v = months && months[mm];
    setAny_(line, idx, names, v == null || v === '' ? '' : Number(v));
  });
}

function financeRecordKey_(rec) {
  var kind = String(rec.kind || 'pair').toLowerCase() === 'row' ? 'row' : 'pair';
  if (kind === 'row') {
    return 'row|' + String(rec.row_id || rec.rowId || '') + '|' + String(rec.year || '');
  }
  return 'pair|' +
    String(rec.cost_center || rec.costCenter || '').trim().toLowerCase() + '|' +
    String(rec.activity_type || rec.activityType || '').trim().toLowerCase() + '|' +
    String(rec.project_cost_center || rec.projectCostCenter || '').trim().toLowerCase() + '|' +
    String(rec.year || '');
}

function readFinanceRecords_(sheetName, isBooking) {
  var sh = ss_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getValues();
  var idx = headerIndex_(values[0]);
  return values.slice(1).filter(function (r) {
    var kind = String(col_(r, idx, ['Kind']) || '').trim();
    var rowId = String(col_(r, idx, ['Row_ID', 'Row ID']) || '').trim();
    var cc = String(col_(r, idx, ['Cost Center']) || '').trim();
    var pcc = String(col_(r, idx, ['Project Cost Center']) || '').trim();
    return kind || rowId || cc || pcc;
  }).map(function (r) {
    var kind = String(strCol_(r, idx, ['Kind']) || 'pair').toLowerCase() === 'row' ? 'row' : 'pair';
    var rec = {
      kind: kind,
      row_id: strCol_(r, idx, ['Row_ID', 'Row ID']) || '',
      cost_center: strCol_(r, idx, ['Cost Center']) || '',
      activity_type: strCol_(r, idx, ['Activity Type']) || '',
      project_cost_center: strCol_(r, idx, ['Project Cost Center']) || '',
      year: Number(col_(r, idx, ['Year'])) || 0,
      manual: /^(1|y|yes|true|manual)$/i.test(String(col_(r, idx, ['Manual']) || '')),
      months: readMonthCellsFromRow_(r, idx, '')
    };
    if (isBooking) {
      rec.transfer_type = normalizeTransferType_(strCol_(r, idx, ['Transfer Type']));
      rec.auto_months = readMonthCellsFromRow_(r, idx, 'Auto_');
    }
    return rec;
  });
}

function writeFinanceRecords_(sheetName, list, isBooking) {
  var headers = isBooking ? BOOKING_HEADERS : FINANCE_MONTH_HEADERS;
  ensureSheet_(sheetName, headers);
  var sh = ss_().getSheetByName(sheetName);
  var lines = [[].concat(headers)];
  var seen = {};
  (list || []).forEach(function (rec) {
    if (!rec) return;
    var key = financeRecordKey_(rec);
    if (seen[key]) return;
    seen[key] = true;
    var kind = String(rec.kind || 'pair').toLowerCase() === 'row' ? 'row' : 'pair';
    var year = Number(rec.year) || 0;
    if (!year) return;
    if (kind === 'row' && !String(rec.row_id || rec.rowId || '').trim()) return;
    if (kind === 'pair' && !String(rec.cost_center || rec.costCenter || '').trim() &&
        !String(rec.project_cost_center || rec.projectCostCenter || '').trim()) return;
    var line = new Array(headers.length).fill('');
    var idx = headerIndex_(headers);
    setAny_(line, idx, ['Kind'], kind);
    setAny_(line, idx, ['Row_ID', 'Row ID'], rec.row_id || rec.rowId || '');
    setAny_(line, idx, ['Cost Center'], rec.cost_center || rec.costCenter || '');
    setAny_(line, idx, ['Activity Type'], rec.activity_type || rec.activityType || '');
    setAny_(line, idx, ['Project Cost Center'], rec.project_cost_center || rec.projectCostCenter || '');
    setAny_(line, idx, ['Year'], year);
    setAny_(line, idx, ['Manual'], rec.manual ? 'Y' : '');
    writeMonthCellsToLine_(line, idx, rec.months || {}, '');
    if (isBooking) {
      setAny_(line, idx, ['Transfer Type'], normalizeTransferType_(rec.transfer_type || rec.transferType));
      writeMonthCellsToLine_(line, idx, rec.auto_months || rec.autoMonths || {}, 'Auto_');
    }
    lines.push(line);
  });
  sh.clear();
  sh.getRange(1, 1, lines.length, headers.length).setValues(lines);
  sh.setFrozenRows(1);
}

/** Replace the full Employees master list (Admin / People Lead UI). */
function saveEmployees(employees, registeredAs, forceRegistered) {
  ensureSchema_();
  writeEmployeesList_(employees || []);
  return getBootstrap(registeredAs || '', !!forceRegistered);
}

/** Admin-only delete of a Projects master row (+ related planning rows). */
function deleteProject(projectId, registeredAs, forceRegistered) {
  ensureSchema_();
  if (!projectId) throw new Error('Missing project id');
  var want = normalizeSheetId_(projectId);
  var sh = sheet_(SHEETS.PROJECTS);
  var values = sh.getDataRange().getValues();
  if (values.length >= 2) {
    var idx = headerIndex_(values[0]);
    var toDelete = [];
    for (var i = 1; i < values.length; i++) {
      if (normalizeSheetId_(col_(values[i], idx, ['Project_ID', 'ID'])) === want) toDelete.push(i + 1);
    }
    deleteSheetRows_(sh, toDelete);
  }

  // Remove demand / timeline / allocation rows tied to this project.
  var related = readRowRecords_().filter(function (r) {
    return normalizeSheetId_(r.project_id) === want;
  }).map(function (r) { return r.id; });
  if (related.length) deleteRowIds_(related);

  return getBootstrap(registeredAs || '', !!forceRegistered);
}

/** Admin-only delete of an Employees master row. */
function deleteEmployee(employeeId, registeredAs, forceRegistered) {
  ensureSchema_();
  if (!employeeId) throw new Error('Missing employee id');
  var list = readEmployees_().filter(function (e) { return e.id !== employeeId; });
  writeEmployeesList_(list);
  return getBootstrap(registeredAs || '', !!forceRegistered);
}

function readMeta_() {
  var sh = ss_().getSheetByName(SHEETS.META);
  var values = sh.getDataRange().getValues();
  var out = {};
  values.slice(1).forEach(function (r) {
    var field = String(r[0] || '');
    var val = String(r[1] || '');
    if (!field || !val) return;
    if (!out[field]) out[field] = [];
    out[field].push(val);
  });
  return out;
}

/* ========== Row information + row_years ========== */

function monthsFromSettings_(settings) {
  var years = settings.yearFilters && settings.yearFilters.length
    ? settings.yearFilters
    : ['2026'];
  var months = [];
  years.forEach(function (y) {
    for (var m = 1; m <= 12; m++) {
      months.push(y + '-' + ('0' + m).slice(-2));
    }
  });
  return months;
}

function readRowRecords_() {
  var sh = sheet_(SHEETS.DATA);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var h = values[0];
  var idx = headerIndex_(h);
  return values.slice(1).filter(function (r) {
    return isPlanningId_(col_(r, idx, ['Row_ID', 'ID']));
  }).map(function (r) {
    var pid = normalizeFk_(col_(r, idx, ['ID Project', 'Project ID']));
    var eid = normalizeFk_(col_(r, idx, ['ID Emp', 'Employee ID']));
    return {
      id: String(col_(r, idx, ['Row_ID', 'ID'])),
      sort_order: Number(col_(r, idx, ['Sort', 'Sort Order'])) || 0,
      source: String(col_(r, idx, ['Source']) || SOURCES.DA),
      version: String(col_(r, idx, ['Version']) || 'Cur'),
      project_id: blankToNull_(pid),
      alternative_financing: blankToNull_(col_(r, idx, ['Alternative Financing'])),
      workpackage: blankToNull_(col_(r, idx, ['Workpackage (opt.)', 'Workpackage'])),
      role: blankToNull_(col_(r, idx, ['Role'])),
      employee_id: blankToNull_(eid),
      notes_project_lead: blankToNull_(col_(r, idx, ['Notes Project Lead'])),
      notes_people_lead: blankToNull_(col_(r, idx, ['Notes People Lead'])),
      notes_employee: blankToNull_(col_(r, idx, ['Notes Employee'])),
      comments: blankToNull_(col_(r, idx, ['Comments'])),
      changed_by: blankToNull_(col_(r, idx, ['Changed by'])),
      updated_at: blankToNull_(col_(r, idx, ['Updated at'])),
      change_history: parseChangeHistory_(col_(r, idx, ['Change history']))
    };
  });
}

function readYearRecords_(sheetName) {
  var name = sheetName || sheetName_(SHEETS.YEARS);
  var sh = ss_().getSheetByName(name);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var h = values[0];
  var idx = headerIndex_(h);
  return values.slice(1).filter(function (r) {
    return isPlanningId_(col_(r, idx, ['R_ID', 'row_id', 'Row_ID']));
  }).map(function (r) {
    var scOrType = col_(r, idx, ['Sc', 'data_type', 'Source']);
    var rec = {
      row_id: String(col_(r, idx, ['R_ID', 'row_id', 'Row_ID'])),
      year: Number(col_(r, idx, ['Y', 'year', 'Year'])) || 0,
      data_type: normalizeDataType_(scOrType),
      version: String(col_(r, idx, ['V', 'Version']) || 'Cur')
    };
    MONTH_HEADERS.forEach(function (mh, i) {
      var shortH = MONTH_HEADERS_SHORT[i];
      var longH = MONTH_HEADERS_LONG[i];
      var val = col_(r, idx, [mh, shortH, longH]);
      rec[mh] = cellValue_(val, rec.data_type === DATA_TYPES.TIMELINE);
    });
    return rec;
  });
}

function assembleRows_(rowRecords, yearRecords, projects, employees) {
  var projectById = {};
  (projects || []).forEach(function (p) { projectById[p.id] = p; });
  var employeeById = {};
  (employees || []).forEach(function (e) { employeeById[e.id] = e; });

  var yearsByRow = {};
  (yearRecords || []).forEach(function (y) {
    if (!yearsByRow[y.row_id]) yearsByRow[y.row_id] = [];
    yearsByRow[y.row_id].push(y);
  });

  return rowRecords.map(function (r) {
    var project = r.project_id ? projectById[r.project_id] : null;
    var alt = r.alternative_financing ? projectById[r.alternative_financing] : null;
    var employee = r.employee_id ? employeeById[r.employee_id] : null;
    var packed = flattenYearRecords_(yearsByRow[r.id] || [], r.source || SOURCES.DA);
    var source = r.source || SOURCES.DA;
    return {
      id: r.id,
      sort_order: r.sort_order,
      source: source,
      version: r.version || 'Cur',
      project_id: r.project_id,
      project_topic: project ? project.project_topic : null,
      alternative_financing: r.alternative_financing,
      alternative_financing_topic: alt ? alt.project_topic : null,
      workpackage: r.workpackage,
      role: r.role || (employee ? employee.role : null),
      employee_id: r.employee_id,
      name: employee ? employee.name : null,
      notes_project_lead: r.notes_project_lead,
      notes_people_lead: r.notes_people_lead,
      notes_employee: r.notes_employee,
      comments: r.comments,
      changed_by: r.changed_by,
      updated_at: r.updated_at,
      change_history: r.change_history || [],
      is_availability: source === SOURCES.AVAIL,
      is_header: source === SOURCES.HEADER,
      months: packed.months,
      phases: packed.phases
    };
  });
}

function flattenYearRecords_(records, source) {
  var months = {};
  var phases = {};
  source = source || SOURCES.DA;
  records.forEach(function (rec) {
    var year = rec.year;
    var dt = rec.data_type;
    // Workbook stores capacity/timeline under Sc=Dm; reinterpret using parent Source.
    if (source === SOURCES.HEADER && (dt === DATA_TYPES.DEMAND || dt === DATA_TYPES.TIMELINE)) {
      dt = DATA_TYPES.TIMELINE;
    } else if (source === SOURCES.AVAIL && dt === DATA_TYPES.DEMAND) {
      dt = DATA_TYPES.AVAILABILITY;
    }
    MONTH_MM.forEach(function (mm, i) {
      var key = year + '-' + mm + '-01';
      var mh = MONTH_HEADERS[i];
      var val = rec[mh];
      if (val == null || val === '') return;
      if (dt === DATA_TYPES.TIMELINE) {
        phases[key] = String(val);
        return;
      }
      if (!months[key]) months[key] = { demand: null, allocation: null, employee: null, budget: null };
      if (dt === DATA_TYPES.DEMAND || dt === DATA_TYPES.AVAILABILITY) {
        months[key].demand = toNum_(val);
      } else if (dt === DATA_TYPES.ALLOCATION) {
        months[key].allocation = toNum_(val);
      } else if (dt === DATA_TYPES.EMPLOYEE) {
        months[key].employee = toNum_(val);
      } else if (dt === DATA_TYPES.BUDGET) {
        months[key].budget = toNum_(val);
      }
    });
  });
  return { months: months, phases: phases };
}

function expandRowToYearRecords_(row) {
  var out = [];
  var byYear = {};

  function ensure(year, dataType) {
    var k = year + '|' + dataType;
    if (!byYear[k]) {
      byYear[k] = emptyYearRec_(row.id, year, dataType);
      out.push(byYear[k]);
    }
    return byYear[k];
  }

  if (row.is_header || row.source === SOURCES.HEADER) {
    Object.keys(row.phases || {}).forEach(function (monthKey) {
      var parsed = parseMonthKey_(monthKey);
      if (!parsed) return;
      var text = row.phases[monthKey];
      if (text == null || String(text).trim() === '') return;
      var rec = ensure(parsed.year, DATA_TYPES.TIMELINE);
      rec['m' + parsed.mm] = String(text).trim();
    });
    return out;
  }

  var demandType = (row.is_availability || row.source === SOURCES.AVAIL)
    ? DATA_TYPES.AVAILABILITY
    : DATA_TYPES.DEMAND;

  Object.keys(row.months || {}).forEach(function (monthKey) {
    var parsed = parseMonthKey_(monthKey);
    if (!parsed) return;
    var cell = row.months[monthKey] || {};
    if (cell.demand != null && cell.demand !== '') {
      var dRec = ensure(parsed.year, demandType);
      dRec['m' + parsed.mm] = Number(cell.demand);
    }
    if (cell.allocation != null && cell.allocation !== '') {
      var aRec = ensure(parsed.year, DATA_TYPES.ALLOCATION);
      aRec['m' + parsed.mm] = Number(cell.allocation);
    }
    if (cell.employee != null && cell.employee !== '') {
      var eRec = ensure(parsed.year, DATA_TYPES.EMPLOYEE);
      eRec['m' + parsed.mm] = Number(cell.employee);
    }
    if (cell.budget != null && cell.budget !== '') {
      var bRec = ensure(parsed.year, DATA_TYPES.BUDGET);
      bRec['m' + parsed.mm] = Number(cell.budget);
    }
  });
  return out;
}

function emptyYearRec_(rowId, year, dataType) {
  var rec = { row_id: rowId, year: year, data_type: dataType };
  MONTH_HEADERS.forEach(function (mh) { rec[mh] = null; });
  return rec;
}

function parseMonthKey_(key) {
  var m = /^(\d{4})-(\d{2})(?:-01)?/.exec(String(key || ''));
  if (!m) return null;
  return { year: Number(m[1]), mm: m[2] };
}

/* ========== Delta write ========== */

function writeDelta_(incoming, originalById, existingById, deletedIds, changedBy) {
  var del = {};
  (deletedIds || []).forEach(function (id) { del[id] = true; });

  var dirtyRows = [];
  var dirtyYears = [];
  var touchedYearRowIds = {};

  incoming.forEach(function (row) {
    if (del[row.id]) return;
    if (String(row.id || '').indexOf('__') === 0) return;

    var orig = originalById[row.id] || originalById[normalizeSheetId_(row.id)] ||
      existingById[row.id] || existingById[normalizeSheetId_(row.id)];
    var recordDirty = !orig || rowRecordChanged_(row, orig);
    var monthsDirty = !orig || monthCellsChanged_(row.months, orig.months) ||
      phasesChanged_(row.phases, orig.phases);

    // Stamp Changed by / Updated at / history whenever identity OR month/phase data changes
    if (recordDirty || monthsDirty) dirtyRows.push(row);
    if (monthsDirty) {
      dirtyYears = dirtyYears.concat(expandRowToYearRecords_(row));
      touchedYearRowIds[normalizeSheetId_(row.id)] = true;
      touchedYearRowIds[String(row.id)] = true;
    }
  });

  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  dirtyRows.forEach(function (row) {
    applyChangeStamp_(row, changedBy, stamp);
  });
  var rowsUpserted = upsertRowRecords_(dirtyRows, changedBy, stamp);
  var yearsUpserted = upsertYearRecords_(dirtyYears, touchedYearRowIds);
  var rowsDeleted = deleteRowIds_(deletedIds);

  return {
    rowsUpserted: rowsUpserted,
    yearsUpserted: yearsUpserted,
    rowsDeleted: rowsDeleted,
    stampedAt: stamp,
    stampedBy: changedBy || '',
    stampedRowIds: dirtyRows.map(function (r) { return r.id; }),
    changeHistoryById: dirtyRows.reduce(function (o, r) {
      o[r.id] = r.change_history || [];
      return o;
    }, {})
  };
}

function upsertRowRecords_(rows, changedBy, stamp) {
  if (!rows || !rows.length) return 0;
  var sh = sheet_(SHEETS.DATA);
  var lastRow = sh.getLastRow();
  var lastCol = Math.max(sh.getLastColumn(), DATA_HEADERS.length);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (!headers || !headers[0]) {
    headers = DATA_HEADERS.slice();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    lastCol = headers.length;
  }
  var idx = headerIndex_(headers);
  var width = headers.length;
  var idCol = idx['Row_ID'] != null ? idx['Row_ID'] : idx['ID'];
  var idToSheetRow = {};
  if (lastRow >= 2 && idCol != null) {
    // Read only the ID column — much faster than getDataRange on wide sheets
    var idVals = sh.getRange(2, idCol + 1, lastRow, idCol + 1).getValues();
    for (var i = 0; i < idVals.length; i++) {
      var key = normalizeSheetId_(idVals[i][0]);
      if (key) idToSheetRow[key] = i + 2;
    }
  }

  var updates = [];
  var toAppend = [];
  rows.forEach(function (row) {
    var line = rowToLine_(row, headers, idx, changedBy, stamp);
    var sheetRow = idToSheetRow[normalizeSheetId_(row.id)];
    if (sheetRow) updates.push({ sheetRow: sheetRow, line: line });
    else toAppend.push(line);
  });
  writeContiguousRowBlocks_(sh, updates, width);
  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, width).setValues(toAppend);
  }
  return rows.length;
}

/** Write sparse row updates as contiguous setValues blocks (fewer service calls). */
function writeContiguousRowBlocks_(sh, updates, width) {
  if (!updates || !updates.length) return;
  updates.sort(function (a, b) { return a.sheetRow - b.sheetRow; });
  var i = 0;
  while (i < updates.length) {
    var startRow = updates[i].sheetRow;
    var block = [updates[i].line];
    while (i + 1 < updates.length && updates[i + 1].sheetRow === updates[i].sheetRow + 1) {
      i++;
      block.push(updates[i].line);
    }
    sh.getRange(startRow, 1, block.length, width).setValues(block);
    i++;
  }
}

function rowToLine_(row, headers, idx, changedBy, stamp) {
  var line = new Array(headers.length).fill('');
  setAny_(line, idx, ['Row_ID', 'ID'], row.id);
  setAny_(line, idx, ['Sort', 'Sort Order'], row.sort_order || 0);
  setAny_(line, idx, ['Source'], row.source || SOURCES.DA);
  setAny_(line, idx, ['Version'], row.version || 'Cur');
  setAny_(line, idx, ['ID Project', 'Project ID'], row.project_id || '');
  setAny_(line, idx, ['Alternative Financing'], row.alternative_financing || '');
  setAny_(line, idx, ['Workpackage (opt.)', 'Workpackage'], row.workpackage || '');
  setAny_(line, idx, ['Role'], row.role || '');
  setAny_(line, idx, ['ID Emp', 'Employee ID'], row.employee_id || '');
  setAny_(line, idx, ['Notes Project Lead'], row.notes_project_lead || '');
  setAny_(line, idx, ['Notes People Lead'], row.notes_people_lead || '');
  setAny_(line, idx, ['Notes Employee'], row.is_header ? '' : (row.notes_employee || ''));
  setAny_(line, idx, ['Comments'], row.comments || '');
  setAny_(line, idx, ['Changed by'], changedBy || row.changed_by || '');
  setAny_(line, idx, ['Updated at'], stamp || row.updated_at || '');
  setAny_(line, idx, ['Change history'], serializeChangeHistory_(row.change_history));
  return line;
}

function upsertYearRecords_(yearRecs, touchedRowIds) {
  var touched = touchedRowIds || {};
  var touchCount = Object.keys(touched).length;
  if (!touchCount && (!yearRecs || !yearRecs.length)) return 0;

  var sh = sheet_(SHEETS.YEARS);
  var lastRow = sh.getLastRow();
  var lastCol = Math.max(sh.getLastColumn(), YEAR_HEADERS.length);
  var headers = lastRow >= 1 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : null;
  if (!headers || !headers[0]) {
    headers = YEAR_HEADERS.slice();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    lastRow = 1;
    lastCol = headers.length;
  }
  var idx = headerIndex_(headers);
  var width = headers.length;

  var ridCol = firstIdx_(idx, ['Row_ID', 'R_ID', 'row_id']);
  var yearCol = firstIdx_(idx, ['Year', 'Y', 'year']);
  var scCol = firstIdx_(idx, ['Source', 'Sc', 'data_type']);
  var verCol = firstIdx_(idx, ['Version', 'V']);

  var keyToRow = {};
  var keysByRowId = {};
  if (lastRow >= 2 && ridCol != null) {
    // Read only key columns (not the full month grid)
    var minC = ridCol;
    var maxC = ridCol;
    [yearCol, scCol, verCol].forEach(function (c) {
      if (c != null) {
        if (c < minC) minC = c;
        if (c > maxC) maxC = c;
      }
    });
    var keyBlock = sh.getRange(2, minC + 1, lastRow, maxC + 1).getValues();
    for (var i = 0; i < keyBlock.length; i++) {
      var rowVals = keyBlock[i];
      var rid = normalizeSheetId_(rowVals[ridCol - minC]);
      if (!rid) continue;
      // Only index rows we might touch — cuts work on huge sheets
      if (touchCount && !touched[rid] && !touched[String(rid)]) continue;
      var y = yearCol != null ? Number(rowVals[yearCol - minC]) || 0 : 0;
      var sc = scCol != null ? String(rowVals[scCol - minC] || '') : '';
      var ver = verCol != null ? String(rowVals[verCol - minC] || 'Cur') : 'Cur';
      var key = rid + '|' + y + '|' + normalizeDataType_(sc) + '|' + ver;
      keyToRow[key] = i + 2;
      if (!keysByRowId[rid]) keysByRowId[rid] = [];
      keysByRowId[rid].push(key);
    }
  }

  var newKeysByRow = {};
  var updates = [];
  var toAppend = [];
  (yearRecs || []).forEach(function (y) {
    var sc = dataTypeToSc_(y.data_type);
    var ver = y.version || 'Cur';
    var rid = normalizeSheetId_(y.row_id);
    var key = rid + '|' + y.year + '|' + normalizeDataType_(sc) + '|' + ver;
    if (!newKeysByRow[rid]) newKeysByRow[rid] = {};
    newKeysByRow[rid][key] = true;
    var line = yearRecToLine_(y, headers, idx, sc, ver);
    var sheetRow = keyToRow[key];
    if (sheetRow) updates.push({ sheetRow: sheetRow, line: line });
    else toAppend.push(line);
  });

  var toDelete = [];
  Object.keys(touched).forEach(function (ridRaw) {
    var rid = normalizeSheetId_(ridRaw);
    (keysByRowId[rid] || []).forEach(function (key) {
      if (!newKeysByRow[rid] || !newKeysByRow[rid][key]) {
        if (keyToRow[key]) toDelete.push(keyToRow[key]);
      }
    });
  });

  // In-place updates + appends do not shift existing row numbers; delete last.
  writeContiguousRowBlocks_(sh, updates, width);
  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, width).setValues(toAppend);
  }
  deleteSheetRows_(sh, toDelete);
  return (yearRecs || []).length;
}

function firstIdx_(idx, names) {
  for (var i = 0; i < names.length; i++) {
    if (idx[names[i]] != null) return idx[names[i]];
  }
  return null;
}

function yearRecToLine_(y, headers, idx, sc, ver) {
  var line = new Array(headers.length).fill('');
  sc = sc || dataTypeToSc_(y.data_type);
  ver = ver || y.version || 'Cur';
  setAny_(line, idx, ['R_ID', 'row_id', 'Row_ID'], y.row_id);
  setAny_(line, idx, ['Y', 'year', 'Year'], y.year);
  setAny_(line, idx, ['Sc', 'data_type', 'Source'], sc);
  setAny_(line, idx, ['V', 'Version'], ver);
  setAny_(line, idx, ['Monthrow_ID', 'MR_ID'], monthrowCompositeId_(y.row_id, y.year, sc, ver));
  MONTH_HEADERS.forEach(function (mh, i) {
    var v = y[mh];
    setAny_(line, idx, [mh, MONTH_HEADERS_SHORT[i], MONTH_HEADERS_LONG[i]], v == null ? '' : v);
  });
  return line;
}

function sheetRowEquals_(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  for (var i = 0; i < a.length; i++) {
    if (String(a[i] == null ? '' : a[i]) !== String(b[i] == null ? '' : b[i])) return false;
  }
  return true;
}

function deleteSheetRows_(sh, rowNumbersDescOrAsc) {
  if (!rowNumbersDescOrAsc || !rowNumbersDescOrAsc.length) return;
  var rows = rowNumbersDescOrAsc.slice().sort(function (a, b) { return b - a; });
  // Merge contiguous blocks for fewer API calls
  var i = 0;
  while (i < rows.length) {
    var end = rows[i];
    var start = end;
    while (i + 1 < rows.length && rows[i + 1] === start - 1) {
      i++;
      start = rows[i];
    }
    sh.deleteRows(start, end - start + 1);
    i++;
  }
}

function normalizeSheetId_(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number' && isFinite(v)) {
    if (Math.floor(v) === v) return String(Math.floor(v));
    return String(v);
  }
  var s = String(v).trim();
  if (/^\d+\.0$/.test(s)) return s.slice(0, -2);
  return s;
}

function deleteRowIds_(deletedIds) {
  if (!deletedIds || !deletedIds.length) return 0;
  var del = {};
  deletedIds.forEach(function (id) {
    del[String(id)] = true;
    del[normalizeSheetId_(id)] = true;
  });

  var sh = sheet_(SHEETS.DATA);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = headerIndex_(headers);
  var idCol = firstIdx_(idx, ['Row_ID', 'ID']);
  if (idCol == null) return 0;
  var idVals = sh.getRange(2, idCol + 1, lastRow, idCol + 1).getValues();
  var toDelete = [];
  for (var i = 0; i < idVals.length; i++) {
    var id = normalizeSheetId_(idVals[i][0]);
    if (del[id]) toDelete.push(i + 2);
  }
  var removed = toDelete.length;
  deleteSheetRows_(sh, toDelete);

  var ysh = sheet_(SHEETS.YEARS);
  var yLast = ysh.getLastRow();
  if (yLast >= 2) {
    var yLastCol = Math.max(ysh.getLastColumn(), 1);
    var yh = ysh.getRange(1, 1, 1, yLastCol).getValues()[0];
    var yidx = headerIndex_(yh);
    var ridCol = firstIdx_(yidx, ['Row_ID', 'R_ID', 'row_id']);
    if (ridCol != null) {
      var rids = ysh.getRange(2, ridCol + 1, yLast, ridCol + 1).getValues();
      var yDel = [];
      for (var j = 0; j < rids.length; j++) {
        var rid = normalizeSheetId_(rids[j][0]);
        if (del[rid]) yDel.push(j + 2);
      }
      deleteSheetRows_(ysh, yDel);
    }
  }
  return removed;
}

function rowRecordChanged_(current, original) {
  if (!original) return true;
  return !(
    sameText_(current.source, original.source) &&
    sameText_(current.version, original.version) &&
    Math.trunc(current.sort_order || 0) === Math.trunc(original.sort_order || 0) &&
    sameText_(current.project_id, original.project_id) &&
    sameText_(current.alternative_financing, original.alternative_financing) &&
    sameText_(current.employee_id, original.employee_id) &&
    sameText_(current.workpackage, original.workpackage) &&
    sameText_(current.role, original.role) &&
    sameText_(current.notes_project_lead, original.notes_project_lead) &&
    sameText_(current.notes_people_lead, original.notes_people_lead) &&
    sameText_(current.notes_employee, original.notes_employee) &&
    sameText_(current.comments, original.comments)
  );
}

function monthCellsChanged_(current, original) {
  var keys = {};
  Object.keys(current || {}).forEach(function (k) { keys[k] = true; });
  Object.keys(original || {}).forEach(function (k) { keys[k] = true; });
  return Object.keys(keys).some(function (key) {
    var a = (current || {})[key] || {};
    var b = (original || {})[key] || {};
    return !sameNum_(a.demand, b.demand) || !sameNum_(a.allocation, b.allocation) ||
      !sameNum_(a.employee, b.employee) || !sameNum_(a.budget, b.budget);
  });
}

function phasesChanged_(current, original) {
  var keys = {};
  Object.keys(current || {}).forEach(function (k) { keys[k] = true; });
  Object.keys(original || {}).forEach(function (k) { keys[k] = true; });
  return Object.keys(keys).some(function (key) {
    return sameText_(current && current[key], original && original[key]) === false;
  });
}

/* ========== Snapshots ========== */

function todaySnapshotDate_() {
  var d = new Date();
  var y = d.getFullYear();
  var mo = ('0' + (d.getMonth() + 1)).slice(-2);
  var day = ('0' + d.getDate()).slice(-2);
  return Number('' + y + mo + day);
}

function formatSnapshotLabel_(dateNum) {
  var s = String(dateNum).padStart(8, '0');
  return s.slice(0, 4) + '-' + s.slice(4, 6) + '-' + s.slice(6, 8);
}

function listSnapshotDates_() {
  var ss = ss_();
  var dates = {};
  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    var n = parseSnapshotDateFromName_(name);
    if (n != null) dates[n] = true;
  });
  return Object.keys(dates).map(Number).sort(function (a, b) { return b - a; });
}

function parseSnapshotDateFromName_(name) {
  var m = /^(\d{8}) data$/.exec(name);
  if (m) return Number(m[1]);
  var prefixes = [
    SNAP_DATA_PREFIX_LEGACY, SNAP_ROW_PREFIX_LEGACY,
    SNAP_YEARS_PREFIX_LEGACY, SNAP_ROW_YEARS_PREFIX_LEGACY
  ];
  for (var i = 0; i < prefixes.length; i++) {
    if (name.indexOf(prefixes[i]) === 0) {
      var n = Number(name.slice(prefixes[i].length));
      if (!isNaN(n)) return n;
    }
  }
  return null;
}

function snapshotSheetNames_(dateNum) {
  var label = String(dateNum);
  var ss = ss_();
  var dataName = label + SNAP_DATA_SUFFIX;
  var yearsName = label + SNAP_YEARS_SUFFIX;
  if (ss.getSheetByName(dataName)) {
    return { data: dataName, years: yearsName };
  }
  var legacy = [
    [SNAP_DATA_PREFIX_LEGACY + label, SNAP_YEARS_PREFIX_LEGACY + label],
    [SNAP_ROW_PREFIX_LEGACY + label, SNAP_ROW_YEARS_PREFIX_LEGACY + label]
  ];
  for (var i = 0; i < legacy.length; i++) {
    if (ss.getSheetByName(legacy[i][0])) {
      return { data: legacy[i][0], years: legacy[i][1] };
    }
  }
  return { data: dataName, years: yearsName };
}

function defaultCompareSnapshotDate_(dates, preferred) {
  if (preferred) {
    var pref = Number(preferred);
    if (dates.indexOf(pref) >= 0) return pref;
  }
  var today = todaySnapshotDate_();
  var before = null;
  var todaySnap = null;
  dates.forEach(function (d) {
    if (d < today && (before == null || d > before)) before = d;
    else if (d === today) todaySnap = d;
  });
  return before != null ? before : todaySnap;
}

function loadSnapshotRowsPack_(dateNum, projects, employees) {
  var names = snapshotSheetNames_(dateNum);
  var rowSh = ss_().getSheetByName(names.data);
  if (!rowSh) return { rows: [], rowSnapshots: {} };

  var values = rowSh.getDataRange().getValues();
  if (values.length < 2) return { rows: [], rowSnapshots: {} };
  var h = values[0];
  var idx = headerIndex_(h);
  var rowRecords = values.slice(1).filter(function (r) {
    return isPlanningId_(col_(r, idx, ['Row_ID', 'ID']));
  }).map(function (r) {
    var pid = normalizeFk_(col_(r, idx, ['ID Project', 'Project ID']));
    var eid = normalizeFk_(col_(r, idx, ['ID Emp', 'Employee ID']));
    return {
      id: String(col_(r, idx, ['Row_ID', 'ID'])),
      sort_order: Number(col_(r, idx, ['Sort', 'Sort Order'])) || 0,
      source: String(col_(r, idx, ['Source']) || SOURCES.DA),
      version: String(col_(r, idx, ['Version']) || 'Cur'),
      project_id: blankToNull_(pid),
      alternative_financing: blankToNull_(col_(r, idx, ['Alternative Financing'])),
      workpackage: blankToNull_(col_(r, idx, ['Workpackage (opt.)', 'Workpackage'])),
      role: blankToNull_(col_(r, idx, ['Role'])),
      employee_id: blankToNull_(eid),
      notes_project_lead: blankToNull_(col_(r, idx, ['Notes Project Lead'])),
      notes_people_lead: blankToNull_(col_(r, idx, ['Notes People Lead'])),
      notes_employee: blankToNull_(col_(r, idx, ['Notes Employee'])),
      comments: blankToNull_(col_(r, idx, ['Comments'])),
      changed_by: blankToNull_(col_(r, idx, ['Changed by'])),
      updated_at: blankToNull_(col_(r, idx, ['Updated at'])),
      change_history: parseChangeHistory_(col_(r, idx, ['Change history']))
    };
  });

  var rows = assembleRows_(rowRecords, [], projects, employees);
  var rowSnapshots = {};
  rows.forEach(function (r) {
    var snap = {
      project_id: r.project_id,
      alternative_financing: r.alternative_financing,
      employee_id: r.employee_id,
      workpackage: r.workpackage,
      role: r.role,
      notes_project_lead: r.notes_project_lead,
      notes_people_lead: r.notes_people_lead,
      notes_employee: r.notes_employee,
      comments: r.comments,
      project_topic: r.project_topic,
      alternative_financing_topic: r.alternative_financing_topic,
      name: r.name
    };
    rowSnapshots[r.id] = snap;
    rowSnapshots[String(r.id)] = snap;
  });
  return { rows: rows, rowSnapshots: rowSnapshots };
}

function loadSnapshotPack_(dateNum, projects, employees) {
  var names = snapshotSheetNames_(dateNum);
  var rowSh = ss_().getSheetByName(names.data);
  var yearSh = ss_().getSheetByName(names.years);
  if (!rowSh) return { rows: [], rowSnapshots: {} };

  var values = rowSh.getDataRange().getValues();
  if (values.length < 2) return { rows: [], rowSnapshots: {} };
  var h = values[0];
  var idx = headerIndex_(h);
  var rowRecords = values.slice(1).filter(function (r) {
    return isPlanningId_(col_(r, idx, ['Row_ID', 'ID']));
  }).map(function (r) {
    var pid = normalizeFk_(col_(r, idx, ['ID Project', 'Project ID']));
    var eid = normalizeFk_(col_(r, idx, ['ID Emp', 'Employee ID']));
    return {
      id: String(col_(r, idx, ['Row_ID', 'ID'])),
      sort_order: Number(col_(r, idx, ['Sort', 'Sort Order'])) || 0,
      source: String(col_(r, idx, ['Source']) || SOURCES.DA),
      version: String(col_(r, idx, ['Version']) || 'Cur'),
      project_id: blankToNull_(pid),
      alternative_financing: blankToNull_(col_(r, idx, ['Alternative Financing'])),
      workpackage: blankToNull_(col_(r, idx, ['Workpackage (opt.)', 'Workpackage'])),
      role: blankToNull_(col_(r, idx, ['Role'])),
      employee_id: blankToNull_(eid),
      notes_project_lead: blankToNull_(col_(r, idx, ['Notes Project Lead'])),
      notes_people_lead: blankToNull_(col_(r, idx, ['Notes People Lead'])),
      notes_employee: blankToNull_(col_(r, idx, ['Notes Employee'])),
      comments: blankToNull_(col_(r, idx, ['Comments'])),
      changed_by: blankToNull_(col_(r, idx, ['Changed by'])),
      updated_at: blankToNull_(col_(r, idx, ['Updated at'])),
      change_history: parseChangeHistory_(col_(r, idx, ['Change history']))
    };
  });

  var yearRecords = yearSh ? readYearRecords_(names.years) : [];
  var rows = assembleRows_(rowRecords, yearRecords, projects, employees);
  var rowSnapshots = {};
  rows.forEach(function (r) {
    var snap = {
      project_id: r.project_id,
      alternative_financing: r.alternative_financing,
      employee_id: r.employee_id,
      workpackage: r.workpackage,
      role: r.role,
      notes_project_lead: r.notes_project_lead,
      notes_people_lead: r.notes_people_lead,
      notes_employee: r.notes_employee,
      comments: r.comments,
      project_topic: r.project_topic,
      alternative_financing_topic: r.alternative_financing_topic,
      name: r.name
    };
    rowSnapshots[r.id] = snap;
    rowSnapshots[String(r.id)] = snap;
  });
  return { rows: rows, rowSnapshots: rowSnapshots };
}

/* ========== Domain rules ========== */

function syncAvailabilityTotals_(rows) {
  var assigned = {};
  rows.forEach(function (r) {
    if (r.is_availability || r.is_header || r.source === SOURCES.HEADER || r.source === SOURCES.AVAIL) return;
    var emp = (r.employee_id || '').trim();
    if (!emp) return;
    Object.keys(r.months || {}).forEach(function (m) {
      var a = r.months[m] && r.months[m].allocation;
      if (a == null || a === '') return;
      var k = emp + '|' + normalizeMonth_(m);
      assigned[k] = (assigned[k] || 0) + Number(a);
    });
  });

  rows.forEach(function (r) {
    if (!(r.is_availability || r.source === SOURCES.AVAIL)) return;
    var emp = (r.employee_id || '').trim();
    if (!emp) return;
    r.months = r.months || {};
    var monthKeys = {};
    Object.keys(r.months).forEach(function (m) { monthKeys[normalizeMonth_(m)] = true; });
    Object.keys(assigned).forEach(function (k) {
      if (k.indexOf(emp + '|') === 0) monthKeys[k.split('|')[1]] = true;
    });
    Object.keys(monthKeys).forEach(function (m) {
      var key = m.indexOf('-01') >= 0 ? m : m + '-01';
      if (key.length === 7) key = key + '-01';
      var cell = r.months[key] || r.months[m] || { demand: null, allocation: null };
      var demand = cell.demand;
      var prevAlloc = cell.allocation;
      var asg = assigned[emp + '|' + normalizeMonth_(m)];
      if ((demand == null || demand === '') && prevAlloc != null &&
          (asg === undefined || Math.abs(Number(prevAlloc) - asg) > 1e-6)) {
        demand = Number(prevAlloc);
      }
      var nextAlloc = asg === undefined ? null : Math.round(asg * 100) / 100;
      r.months[key] = { demand: demand == null || demand === '' ? null : Number(demand), allocation: nextAlloc };
    });
  });
}

function levelAllocations_(rows, months, delta) {
  var capacity = {};
  rows.forEach(function (r) {
    if (!(r.is_availability || r.source === SOURCES.AVAIL)) return;
    var emp = (r.employee_id || '').trim();
    if (!emp) return;
    Object.keys(r.months || {}).forEach(function (m) {
      var d = r.months[m] && r.months[m].demand;
      if (d == null) return;
      capacity[emp + '|' + normalizeMonth_(m)] = Number(d);
    });
  });

  var changed = 0;
  months.forEach(function (ym) {
    var m = ym.length === 7 ? ym + '-01' : ym;
    var byEmp = {};
    rows.forEach(function (r) {
      if (r.is_availability || r.is_header || r.source === SOURCES.AVAIL || r.source === SOURCES.HEADER) return;
      var emp = (r.employee_id || '').trim();
      if (!emp) return;
      var cell = r.months && (r.months[m] || r.months[ym]);
      var a = cell && cell.allocation;
      if (a == null) return;
      if (!byEmp[emp]) byEmp[emp] = [];
      byEmp[emp].push({ row: r, month: m, value: Number(a) });
    });
    Object.keys(byEmp).forEach(function (emp) {
      var list = byEmp[emp];
      var sum = list.reduce(function (s, x) { return s + x.value; }, 0);
      var cap = capacity[emp + '|' + normalizeMonth_(m)];
      if (cap == null) cap = 1;
      if (!(sum > cap + (delta || 0) + 1e-9)) return;
      var scale = cap / sum;
      list.forEach(function (x) {
        var next = Math.round(x.value * scale * 100) / 100;
        if (!x.row.months[x.month]) x.row.months[x.month] = { demand: null, allocation: null };
        if (x.row.months[x.month].allocation !== next) {
          x.row.months[x.month].allocation = next;
          changed++;
        }
      });
    });
  });
  return changed;
}

/** Compact unique key for data_years rows (display / legacy lookup). Upsert uses Row_ID+Year+Source+Version. */
function monthrowCompositeId_(rowId, year, sc, ver) {
  var rid = String(normalizeSheetId_(rowId)).replace(/-/g, '');
  if (rid.length > 12) rid = rid.slice(-12);
  return rid + String(year) + String(sc || '') + String(ver || 'Cur');
}

function displayNameFromEmail_(email) {
  var e = String(email || '').trim();
  if (!e) return '';
  if (e.indexOf('@') < 0) return e;
  var local = e.split('@')[0];
  return local.split(/[._+\-]+/).filter(function (w) { return w; }).map(function (w) {
    return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
  }).join(' ');
}

function rowInProjectLeadScopeServer_(row, actorName, role, access, projects) {
  if (role === 'Admin') return true;
  if (role !== 'Project Lead' && role !== 'People+Project Lead') return true;
  if (row.is_availability || row.source === SOURCES.AVAIL) return true;
  return rowProjectInLeadScopeServer_(row.project_id, actorName, access, projects);
}

function rowProjectInLeadScopeServer_(projectId, actorName, access, projects) {
  if (!projectId) return false;
  // actorName = registered end-user (Execute-as-Me means Session email is the owner).
  var me = String(actorName || '').trim().toLowerCase();
  if (!me) return false;
  var ids = {};
  var programs = {};
  var hasGlobal = false;
  var matched = false;
  (access || []).forEach(function (a) {
    var rt = normalizeRole_(a.roleType || a.role_type || a.role);
    if (rt !== 'Project Lead') return;
    var un = String(a.userName || a.user_name || '').trim().toLowerCase();
    if (un !== me) return;
    matched = true;
    var st = normalizeAccessScopeType_(a.scopeType || a.scope_type || 'global');
    var sv = String(a.scopeValue != null ? a.scopeValue : (a.scope_value || '')).trim();
    if (st === 'global') hasGlobal = true;
    else if (st === 'project' && sv) ids[sv] = true;
    else if (st === 'program' && sv) programs[sv] = true;
  });
  (projects || []).forEach(function (p) {
    var lead = String(p.project_lead || '').trim().toLowerCase();
    if (lead && lead === me) {
      ids[p.id] = true;
      matched = true;
    }
  });
  if (hasGlobal) return true;
  if (!matched) return false;
  if (ids[String(projectId)]) return true;
  var proj = null;
  for (var i = 0; i < (projects || []).length; i++) {
    if (String(projects[i].id) === String(projectId)) { proj = projects[i]; break; }
  }
  if (proj && proj.program && programs[String(proj.program)]) return true;
  return false;
}

function roleUnitFromRole_(role) {
  var s = String(role || '').trim();
  if (!s) return '';
  var slash = s.indexOf('/');
  var beforeSlash = slash < 0 ? s : s.slice(0, slash).trim();
  if (slash >= 0 && beforeSlash.indexOf(' ') < 0) return beforeSlash;
  var sp = s.indexOf(' ');
  return sp < 0 ? s : s.slice(0, sp);
}

function roleDepartmentFromRole_(role) {
  var s = String(role || '').trim();
  if (!s) return '';
  var slash = s.indexOf('/');
  return slash < 0 ? s : s.slice(0, slash).trim();
}

function redactEmployeeCommentsForRole_(employees, role, actorName, access) {
  if (role === 'Admin') return employees || [];
  var canPeople = role === 'People Lead' || role === 'People+Project Lead';
  return (employees || []).map(function (e) {
    var out = Object.assign({}, e);
    if (!canPeople) {
      out.comments = '';
      return out;
    }
    if (!rowInPeopleLeadScopeServer_({
      employee_id: e.id,
      role: e.role,
      department: e.department || e.group_name,
      group_name: e.group_name
    }, actorName, access, employees)) {
      out.comments = '';
    }
    return out;
  });
}

function employeeDeptUnitKeys_(emp, row) {
  var keys = [];
  function add(v) {
    var s = String(v || '').trim().toLowerCase();
    if (s && s !== '—') keys.push(s);
  }
  add(emp && emp.group_name);
  add(emp && emp.department);
  add(emp && emp.unit);
  add(row && row.group_name);
  add(row && row.department);
  add(emp && roleDepartmentFromRole_(emp.role));
  add(emp && roleUnitFromRole_(emp.role));
  add(row && roleDepartmentFromRole_(row.role));
  add(row && roleUnitFromRole_(row.role));
  return keys;
}

function rowInPeopleLeadScopeServer_(row, actorName, access, employees) {
  // actorName = registered end-user (Execute-as-Me means Session email is the owner).
  var me = String(actorName || '').trim().toLowerCase();
  if (!me) return false;
  var depts = {};
  var units = {};
  var hasGlobal = false;
  var matched = false;
  (access || []).forEach(function (a) {
    var rt = normalizeRole_(a.roleType || a.role_type || a.role);
    if (rt !== 'People Lead') return;
    var un = String(a.userName || a.user_name || '').trim().toLowerCase();
    if (un !== me) return;
    matched = true;
    var st = normalizeAccessScopeType_(a.scopeType || a.scope_type || 'global');
    var sv = String(a.scopeValue != null ? a.scopeValue : (a.scope_value || '')).trim();
    if (st === 'global') hasGlobal = true;
    else if (st === 'department' && sv) depts[sv.toLowerCase()] = true;
    else if (st === 'unit' && sv) units[sv.toLowerCase()] = true;
  });
  if (hasGlobal) return true;
  if (!matched || (!Object.keys(depts).length && !Object.keys(units).length)) return false;
  var empId = row.employee_id;
  var emp = null;
  for (var i = 0; i < (employees || []).length; i++) {
    if (String(employees[i].id) === String(empId)) { emp = employees[i]; break; }
  }
  var keys = employeeDeptUnitKeys_(emp, row);
  return keys.some(function (k) { return depts[k] || units[k]; });
}

function sanitizeEmployeesForRole_(incoming, existing, role, actorName, access) {
  var byId = {};
  (existing || []).forEach(function (e) { byId[String(e.id)] = e; });
  return (incoming || []).map(function (e) {
    var prev = byId[String(e.id)];
    if (!prev) return e;
    var out = Object.assign({}, e);
    if (role === 'Admin') return out;
    // People Lead / combined: comments only inside assigned dept/unit
    if (!rowInPeopleLeadScopeServer_({ employee_id: e.id, role: e.role, department: e.department || e.group_name, group_name: e.group_name }, actorName, access, existing)) {
      out.comments = prev.comments;
      // also block master edits outside scope
      return Object.assign({}, prev);
    }
    return out;
  });
}

function applyOutOfProjectLeadLocks_(row, prev, months) {
  var out = JSON.parse(JSON.stringify(row));
  out.notes_project_lead = prev.notes_project_lead;
  out.workpackage = prev.workpackage;
  out.alternative_financing = prev.alternative_financing;
  out.phases = prev.phases || out.phases;
  (months || []).forEach(function (ym) {
    var m = ym.length === 7 ? ym + '-01' : ym;
    var cur = (out.months && (out.months[m] || out.months[ym])) || {};
    var old = (prev.months && (prev.months[m] || prev.months[ym])) || {};
    cur.demand = old.demand;
    cur.budget = old.budget;
    if (!out.months) out.months = {};
    out.months[m] = cur;
  });
  return out;
}

function applyOutOfPeopleLeadLocks_(row, prev, months) {
  var out = JSON.parse(JSON.stringify(row));
  out.notes_people_lead = prev.notes_people_lead;
  out.employee_id = prev.employee_id;
  (months || []).forEach(function (ym) {
    var m = ym.length === 7 ? ym + '-01' : ym;
    var cur = (out.months && (out.months[m] || out.months[ym])) || {};
    var old = (prev.months && (prev.months[m] || prev.months[ym])) || {};
    if (out.is_availability || out.source === SOURCES.AVAIL) {
      cur.demand = old.demand;
    } else {
      cur.allocation = old.allocation;
    }
    if (!out.months) out.months = {};
    out.months[m] = cur;
  });
  return out;
}

function isEmployeeEditablePlaceholderName_(name) {
  var s = String(name || '').trim();
  if (!s) return false;
  if (/^(nn|n\.n\.?|tbd|open|vacant|\?+)$/i.test(s)) return true;
  if (/^external\s+employee$/i.test(s)) return true;
  if (/^internal\s+employee$/i.test(s)) return true;
  return false;
}

function isFixedRoleStaffName_(name) {
  var s = String(name || '').trim();
  if (!s) return false;
  if (/^(nn|n\.n\.?|tbd|open|vacant|\?+)$/i.test(s)) return false;
  if (/^external\s+employee$/i.test(s)) return false;
  if (/^internal\s+employee$/i.test(s)) return false;
  if (/new\s*hire|onboard|starter|to\s*hire|open\s*req/i.test(s)) return false;
  return true;
}

function applyRoleLocksOnSave_(row, prev, role, months) {
  if (!prev) return row;
  if (role === 'Viewer') return JSON.parse(JSON.stringify(prev));
  var out = JSON.parse(JSON.stringify(row));

  // Project Lead: never mutate Avail / ∑ Alloc rows.
  if (role === 'Project Lead' && (out.is_availability || out.source === SOURCES.AVAIL ||
      prev.is_availability || prev.source === SOURCES.AVAIL)) {
    return JSON.parse(JSON.stringify(prev));
  }

  if (role === 'Employee') {
    var empName = String(prev.name || out.name || '').trim();
    if (!isEmployeeEditablePlaceholderName_(empName)) {
      // Named colleagues: Employees may not change project / notes / Emp months.
      return JSON.parse(JSON.stringify(prev));
    }
    out.employee_id = prev.employee_id;
    out.role = prev.role;
    out.workpackage = prev.workpackage;
    out.notes_project_lead = prev.notes_project_lead;
    out.notes_people_lead = prev.notes_people_lead;
    out.comments = prev.comments;
    out.alternative_financing = prev.alternative_financing;
  } else if (role === 'Project Lead') {
    out.employee_id = prev.employee_id;
    out.notes_people_lead = prev.notes_people_lead;
    out.notes_employee = prev.notes_employee;
    out.comments = prev.comments;
    if (isFixedRoleStaffName_(prev.name || out.name)) {
      out.role = prev.role;
    }
  } else if (role === 'People Lead') {
    out.notes_project_lead = prev.notes_project_lead;
    out.notes_employee = prev.notes_employee;
  } else if (role === 'People+Project Lead') {
    out.notes_employee = prev.notes_employee;
    if (isFixedRoleStaffName_(prev.name || out.name)) {
      // Keep named-person role stable unless People Lead side is also changing via scope (allow role on people side)
      // Only lock when the role field changed without people-scope intent — keep simple: lock role for named.
      out.role = prev.role;
    }
  }

  months.forEach(function (ym) {
    var m = ym.length === 7 ? ym + '-01' : ym;
    var cur = (out.months && (out.months[m] || out.months[ym])) || { demand: null, allocation: null, employee: null, budget: null };
    var old = (prev.months && (prev.months[m] || prev.months[ym])) || { demand: null, allocation: null, employee: null, budget: null };
    if (out.is_header || out.source === SOURCES.HEADER) {
      if (role !== 'Admin' && role !== 'Project Lead' && role !== 'People+Project Lead') {
        out.phases = prev.phases || {};
      }
      return;
    }
    if (out.is_availability || out.source === SOURCES.AVAIL) {
      if (role !== 'Admin' && role !== 'People Lead' && role !== 'People+Project Lead') {
        cur.demand = old.demand;
      }
      cur.allocation = old.allocation;
      cur.employee = old.employee;
      cur.budget = old.budget;
    } else {
      if (role === 'Project Lead') {
        cur.allocation = old.allocation;
        cur.employee = old.employee; // PL never edits Emp month columns
      } else if (role === 'People Lead') {
        cur.demand = old.demand;
        cur.budget = old.budget;
      } else if (role === 'Employee') {
        cur.demand = old.demand;
        cur.allocation = old.allocation;
        cur.budget = old.budget;
        // Emp months allowed only for placeholder names (already gated above)
      } else if (role === 'People+Project Lead') {
        // both demand and allocation allowed (further scoped upstream)
      }
      if (role !== 'Admin' && role !== 'Project Lead' && role !== 'People+Project Lead') {
        cur.budget = old.budget;
      }
    }
    if (!out.months) out.months = {};
    out.months[m] = cur;
  });
  return out;
}

/* ========== helpers ========== */

function ix_(headers, name) {
  var i = headers.indexOf(name);
  return i < 0 ? -1 : i;
}

function headerIndex_(headers) {
  var o = {};
  headers.forEach(function (h, i) { o[String(h)] = i; });
  return o;
}

function col_(row, idx, names) {
  for (var i = 0; i < names.length; i++) {
    var key = names[i];
    if (idx[key] != null && idx[key] >= 0) return row[idx[key]];
  }
  return null;
}

function strCol_(row, idx, names) {
  var v = col_(row, idx, names);
  if (v == null || v === '') return '';
  return String(v);
}

function set_(line, idx, name, value) {
  if (idx[name] == null) return;
  line[idx[name]] = value;
}

function setAny_(line, idx, names, value) {
  for (var i = 0; i < names.length; i++) {
    if (idx[names[i]] != null) {
      line[idx[names[i]]] = value;
      return;
    }
  }
}

function normalizeDataType_(raw) {
  var s = String(raw || '').trim().toLowerCase();
  if (s === 'alc' || s === 'al' || s === 'allocation') return DATA_TYPES.ALLOCATION;
  if (s === 'av' || s === 'availability') return DATA_TYPES.AVAILABILITY;
  if (s === 'tl' || s === 'timeline') return DATA_TYPES.TIMELINE;
  if (s === 'emp' || s === 'em' || s === 'employee') return DATA_TYPES.EMPLOYEE;
  if (s === 'bud' || s === 'bg' || s === 'budget') return DATA_TYPES.BUDGET;
  if (s === 'dm' || s === 'demand') return DATA_TYPES.DEMAND;
  return s || DATA_TYPES.DEMAND;
}

function isMasterId_(v, prefix) {
  if (v == null || v === '') return false;
  var s = String(v).trim();
  if (!s) return false;
  var upper = s.toUpperCase();
  // Skip header leftovers / empty placeholders
  if (upper === 'PROJECT_ID' || upper === 'EMPLOYEE_ID' || upper === 'ID') return false;
  if (s === prefix || s === prefix + 'ID' || upper === String(prefix || '').toUpperCase() + 'ID') return false;
  if (s === '#' || s === '#N/A' || s === 'None') return false;
  // Prefer prefixed ids when present; also accept numeric / uuid / custom keys from the databank
  if (prefix && s.indexOf(prefix) === 0) return true;
  if (/^\d+(\.0)?$/.test(s)) return true;
  if (/^[0-9a-f-]{8,}$/i.test(s)) return true;
  return s.length >= 1;
}

function isPlanningId_(v) {
  if (v == null || v === '') return false;
  var s = String(v).trim();
  if (s === 'R_ID' || s === 'Row_ID' || s === 'ID' || s === 'Monthrow_ID') return false;
  // Numeric workbook IDs or UUID-like
  if (/^\d+(\.0)?$/.test(s)) return true;
  if (/^[0-9a-f-]{8,}$/i.test(s)) return true;
  if (/^R_/i.test(s)) return true;
  return false;
}

function normalizeFk_(v) {
  if (v == null || v === '') return null;
  var s = String(v).trim();
  if (s === 'P_' || s === '#N/A' || s === 'None') return null;
  if (typeof v === 'number' && v === Math.floor(v)) return String(Math.floor(v));
  return s;
}

function blankToNull_(v) {
  if (v == null || v === '') return null;
  return String(v);
}

function parseChangeHistory_(raw) {
  if (raw == null || raw === '') return [];
  if (Object.prototype.toString.call(raw) === '[object Array]') {
    return normalizeChangeHistoryList_(raw);
  }
  var s = String(raw).trim();
  if (!s) return [];
  try {
    var parsed = JSON.parse(s);
    return normalizeChangeHistoryList_(parsed);
  } catch (e) {
    return [];
  }
}

function normalizeChangeHistoryList_(list) {
  if (!list || !list.length) return [];
  return list.slice(0, 3).map(function (ev) {
    return {
      by: String((ev && (ev.by || ev.user || ev.changed_by)) || '').trim(),
      at: String((ev && (ev.at || ev.when || ev.updated_at)) || '').trim()
    };
  }).filter(function (ev) { return ev.by || ev.at; });
}

function serializeChangeHistory_(list) {
  var norm = normalizeChangeHistoryList_(list);
  return norm.length ? JSON.stringify(norm) : '';
}

function applyChangeStamp_(row, changedBy, stamp) {
  var by = String(changedBy || row.changed_by || '').trim();
  var at = String(stamp || '').trim();
  var hist = parseChangeHistory_(row.change_history);
  // Avoid duplicate consecutive identical stamps
  if (!(hist[0] && hist[0].by === by && hist[0].at === at)) {
    hist.unshift({ by: by, at: at });
  }
  row.change_history = hist.slice(0, 3);
  row.changed_by = by || row.changed_by || null;
  row.updated_at = at || row.updated_at || null;
}

function toNum_(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v;
  var n = Number(String(v).replace(',', '.'));
  return isNaN(n) ? null : n;
}

function cellValue_(v, asText) {
  if (v == null || v === '') return null;
  if (asText) return String(v);
  var n = toNum_(v);
  if (n != null) return n;
  // Keep non-numeric text (timeline labels often stored under Sc=Dm)
  return String(v);
}

function dataTypeToSc_(dt) {
  var s = String(dt || '').toLowerCase();
  if (s === DATA_TYPES.ALLOCATION) return 'Alc';
  if (s === DATA_TYPES.EMPLOYEE) return 'Emp';
  if (s === DATA_TYPES.BUDGET) return 'Bud';
  if (s === DATA_TYPES.TIMELINE) return 'Tl';
  return 'Dm';
}

function normalizeMonth_(m) {
  var s = String(m || '');
  var match = /^(\d{4}-\d{2})/.exec(s);
  return match ? match[1] : s;
}

function indexById_(rows) {
  var o = {};
  (rows || []).forEach(function (r) {
    if (!r || r.id == null || r.id === '') return;
    o[r.id] = r;
    o[normalizeSheetId_(r.id)] = r;
    o[String(r.id)] = r;
  });
  return o;
}

function sameText_(a, b) {
  var na = a == null || String(a).trim() === '' ? null : String(a).trim();
  var nb = b == null || String(b).trim() === '' ? null : String(b).trim();
  return na === nb;
}

function sameNum_(a, b) {
  if ((a == null || a === '') && (b == null || b === '')) return true;
  if (a == null || b == null || a === '' || b === '') return false;
  return Number(a) === Number(b);
}
