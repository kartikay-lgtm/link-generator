/**
 * Link Generator - referral signup -> Google Sheet + Linkrunner link
 *
 * Paste this into Extensions -> Apps Script on the sheet that should hold the
 * signups, then deploy it as a web app. SETUP.md walks through it.
 *
 * The Linkrunner API key is NOT in this file and must never be. It lives in
 * Project Settings -> Script Properties, so that copying this code around does
 * not copy the key with it.
 */

var SHEET_NAME = 'Signups';

// The spreadsheet this belongs to. Not used to open the file - the script is
// bound to it and reads it directly, which keeps the permission it asks for
// narrow. It is here so that pasting this into the wrong spreadsheet's editor
// fails loudly instead of quietly filling it up.
var EXPECTED_SPREADSHEET_ID = '19FVvUrE5QeGjr-PtIlI9SCbetmB2VL_xWnDHFy5AJuU';

var HEADERS = [
  'Timestamp',
  'Name',
  'Company',
  'Phone',
  'Referral link',
  'Referral code',
  'Referred by',
  'Claim ID',
  'Page'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    // tryLock, and inside the try, on purpose. waitLock THROWS on timeout, and
    // a throw from out here hands the browser an HTML error page instead of
    // JSON. The page cannot read that, so it reports a failure - for a signup
    // that may well have been written by whichever request held the lock.
    locked = lock.tryLock(15000);
    if (!locked) return json_({ ok: false, error: 'busy' });

    if (!e || !e.postData || !e.postData.contents) {
      return json_({ ok: false, error: 'empty request' });
    }

    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet_();
    var claimId = String(data.claimId || '');

    // Dedupe BEFORE minting. The page re-sends when a reply goes missing, and
    // minting above this line would spend a second link on that retry and show
    // the same person a different one each time.
    var existing = claimId ? findRow_(sheet, claimId) : 0;
    if (existing) {
      return json_({
        ok: true,
        duplicate: true,
        link: String(sheet.getRange(existing, HEADERS.indexOf('Referral link') + 1).getValue() || '')
      });
    }

    // Minting is wrapped on its own. If Linkrunner is down the signup still has
    // to land: a row with a blank link can be backfilled, a lost row is someone
    // who believes they signed up and did not.
    var link = '', code = '', linkError = '';
    try {
      var minted = mintReferralLink_(data);
      link = minted.link || '';
      code = minted.code || '';
    } catch (mintErr) {
      // Reported back to the caller as well as logged. Without this, a link
      // that fails to mint is invisible from outside - the row saves, the page
      // says "being prepared", and the reason sits in an execution log nobody
      // opens. The message never contains the key: it is only ever our own
      // "not set" text or Linkrunner's own reply.
      linkError = String(mintErr && mintErr.message ? mintErr.message : mintErr).slice(0, 300);
      Logger.log('link minting failed: ' + linkError);
    }

    var row = firstFreeRow_(sheet);

    // Format the phone cell as text BEFORE anything is written to it. Sheets
    // reads a bare run of digits as a number, and a leading + as the start of a
    // formula. Setting the format first means digits land exactly as typed.
    sheet.getRange(row, HEADERS.indexOf('Phone') + 1).setNumberFormat('@');

    sheet.getRange(row, 1, 1, HEADERS.length).setValues([[
      new Date(),
      String(data.name || ''),
      String(data.company || ''),
      String(data.phone || ''),
      link,
      code,
      String(data.referredBy || ''),
      claimId,
      String(data.page || '')
    ]]);

    var reply = { ok: true, link: link };
    if (linkError) reply.linkError = linkError;
    return json_(reply);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    if (locked) lock.releaseLock();
  }
}

/**
 * GET ?check=<claimId> answers whether that signup is already recorded, and
 * hands back the link it was given.
 *
 * This is what lets the page stop guessing. A POST whose reply goes missing
 * looks identical to one that never arrived, and assuming the worst is what
 * makes a saved entry show up as an error.
 */
function doGet(e) {
  var id = e && e.parameter && e.parameter.check;
  if (id) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss ? ss.getSheetByName(SHEET_NAME) : null;
    var row = sheet ? findRow_(sheet, String(id)) : 0;
    if (!row) return json_({ ok: true, saved: false });
    return json_({
      ok: true,
      saved: true,
      link: String(sheet.getRange(row, HEADERS.indexOf('Referral link') + 1).getValue() || '')
    });
  }
  return json_({ ok: true, message: 'Link Generator endpoint is live' });
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'No spreadsheet attached. This script has to be created from inside the ' +
      'sheet itself (Extensions -> Apps Script), not from script.google.com.'
    );
  }
  if (EXPECTED_SPREADSHEET_ID && ss.getId() !== EXPECTED_SPREADSHEET_ID) {
    throw new Error(
      'Wrong spreadsheet. This script expects ' + EXPECTED_SPREADSHEET_ID +
      ' but is attached to ' + ss.getId() +
      '. Either move it, or update EXPECTED_SPREADSHEET_ID.'
    );
  }

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(5, 280);
  }
  return sheet;
}

/**
 * The first genuinely free row, found by walking UP the Timestamp column.
 *
 * Not getLastRow() + 1. getLastRow() reports the last row holding content
 * ANYWHERE on the tab, so a single stray character far below the data - a
 * space left behind when someone cleared cells instead of deleting rows, a
 * note typed in an unused column - makes every new entry land beneath it and
 * leaves a block of blank rows in between. Anchoring on the Timestamp column
 * means only real entries can move the insertion point.
 */
function firstFreeRow_(sheet) {
  var last = sheet.getLastRow();
  if (last < 1) return 1;

  var stamps = sheet.getRange(1, 1, last, 1).getValues();
  for (var i = stamps.length - 1; i >= 0; i--) {
    var v = stamps[i][0];
    if (v !== null && v !== undefined && String(v).trim() !== '') return i + 2;
  }
  return 2;   // header only
}

/**
 * Deletes rows with no Timestamp, closing up any gaps. Safe to run repeatedly;
 * it never touches a row that has data.
 */
function tidyBlankRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;

  var last = sheet.getLastRow();
  if (last < 2) { Logger.log('Nothing to tidy.'); return; }

  var stamps = sheet.getRange(1, 1, last, 1).getValues();
  var removed = 0;

  // Bottom up, so deleting a row cannot shift the ones still to be checked.
  for (var i = stamps.length - 1; i >= 1; i--) {
    var v = stamps[i][0];
    if (v === null || v === undefined || String(v).trim() === '') {
      sheet.deleteRow(i + 1);
      removed++;
    }
  }
  Logger.log('Removed ' + removed + ' blank row(s). ' +
             (sheet.getLastRow() - 1) + ' entries remain.');
}

/** Row number for an existing claim id, or 0 if it has not been seen. */
function findRow_(sheet, claimId) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;

  var col = HEADERS.indexOf('Claim ID') + 1;
  var values = sheet.getRange(2, col, last - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === claimId) return i + 2;
  }
  return 0;
}

/**
 * Creates one referral link via Linkrunner.
 *
 *   POST https://api.linkrunner.io/api/v1/create-campaign
 *   header: linkrunner-key: <key>
 *   body:   { name, custom_display_id, is_shortlink, deeplink? }
 *   201  -> { msg, status, data: { link } }
 *
 * Each referral link is a Linkrunner CAMPAIGN. That is how their API models it:
 * there is no per-user link primitive, so one referrer means one campaign, and
 * custom_display_id carries our slug so the link reads as theirs.
 *
 * The key is read from Script Properties, never from this file:
 *   Project Settings -> Script Properties -> add LINKRUNNER_API_KEY
 *
 * Optional properties, both passed straight through to Linkrunner:
 *   REFERRAL_DESTINATION   where the link should send people (deeplink)
 *   REFERRAL_DESKTOP_URL   desktop fallback
 * Leave them unset and Linkrunner falls back to the project's own defaults.
 *
 * Returns { link, code }. Throwing is fine - doPost catches it and saves the
 * signup regardless, so a Linkrunner outage cannot cost a submission.
 */
function mintReferralLink_(data) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('LINKRUNNER_API_KEY');
  if (!key) throw new Error('LINKRUNNER_API_KEY is not set in Script Properties');

  var destination = props.getProperty('REFERRAL_DESTINATION');
  var desktop = props.getProperty('REFERRAL_DESKTOP_URL');
  var person = String(data.name || '').trim() || 'unknown';

  var lastBody = '';

  // custom_display_id has to be unique across the account, and two digits is
  // only a hundred of them - so two people sharing a first name can collide.
  // Retry with fresh digits rather than let that cost someone their link.
  for (var attempt = 0; attempt < 5; attempt++) {
    var slug = referralSlug_(person);

    var payload = {
      name: 'Referral - ' + person,
      custom_display_id: slug,
      is_shortlink: true
    };
    if (destination) payload.deeplink = destination;
    if (desktop) payload.link_for_desktop_users = desktop;

    var res = UrlFetchApp.fetch('https://api.linkrunner.io/api/v1/create-campaign', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'linkrunner-key': key },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true   // read the body ourselves rather than throw on 4xx
    });

    var status = res.getResponseCode();
    var body = res.getContentText();

    if (status >= 200 && status < 300) {
      var out = JSON.parse(body);
      var link = out && out.data && out.data.link;
      if (!link) throw new Error('Linkrunner gave no link back: ' + body);
      return { link: String(link), code: slug };
    }

    // A taken display id is the one failure worth another go. A bad key or a
    // rate limit will not fix itself by trying again with different digits.
    //
    // The wording matters: Linkrunner's own docs say "Duplicate ids are not
    // allowed!", which none of the obvious keywords catch, so match that
    // explicitly. The status is not documented for this case either, hence 409
    // alongside 400.
    if ((status === 400 || status === 409) &&
        /duplicate|display[_ ]?id|unique|exists|taken|already/i.test(body)) {
      lastBody = body;
      continue;
    }

    throw new Error('Linkrunner ' + status + ': ' + body);
  }

  throw new Error('no free referral code after 5 tries: ' + lastBody);
}

/**
 * The person's first name run straight into a two-digit number: "aditi42".
 *
 * Two digits is only a hundred codes per name, and custom_display_id has to be
 * unique across the whole Linkrunner account - so two people called Aditi have
 * a real chance of clashing. mintReferralLink_ retries on a clash, and with
 * five attempts the odds of nobody finding a free code are negligible until
 * you have dozens of people sharing a first name.
 */
function referralSlug_(name) {
  var first = String(name || '').trim().split(/\s+/)[0] || 'tal';
  first = first.toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 14) || 'tal';

  var digits = '';
  for (var i = 0; i < 2; i++) digits += Math.floor(Math.random() * 10);
  return first + digits;
}

/** Every reply the web app sends goes out through here, as JSON. */
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ===================================================================== *
 *  Link stats dashboard
 *
 *  Pulls every campaign from Linkrunner and writes one row per referral
 *  link into its own tab, sorted by installs so the people worth paying a
 *  bonus to are at the top.
 *
 *  On clicks: Linkrunner's dashboard shows them, but their API does not
 *  expose a click count anywhere - a campaign carries attributed_users and
 *  nothing else countable. So this tracks INSTALLS.
 * ===================================================================== */

var STATS_SHEET_NAME = 'Link stats';

var STATS_HEADERS = [
  'Referral code',
  'Name',
  'Company',
  'Phone',
  'Referral link',
  'Installs',
  'Active',
  'Link created',
  'Signed up at',
  'Last synced'
];

/** Rebuilds the Link stats tab. Safe to run as often as you like. */
function syncLinkStats() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;   // a sync is already running

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var signups = ss.getSheetByName(SHEET_NAME);
    if (!signups || signups.getLastRow() < 2) return;

    var rows = signups.getRange(2, 1, signups.getLastRow() - 1, HEADERS.length).getValues();
    var codeCol = HEADERS.indexOf('Referral code');
    var people = [];
    for (var i = 0; i < rows.length; i++) {
      var code = String(rows[i][codeCol] || '').trim();
      if (!code) continue;             // minting failed for this one; nothing to track yet
      people.push({
        code: code,
        name: rows[i][HEADERS.indexOf('Name')],
        company: rows[i][HEADERS.indexOf('Company')],
        phone: rows[i][HEADERS.indexOf('Phone')],
        link: rows[i][HEADERS.indexOf('Referral link')],
        signedUpAt: rows[i][0]
      });
    }
    if (!people.length) return;

    var campaigns = fetchAllCampaigns_();   // display_id -> campaign
    var now = new Date();

    var out = people.map(function (p) {
      // Match on the normalised code, then fall back to the code inside the
      // link we actually handed this person.
      var c = campaigns[normaliseCode_(p.code)] ||
              campaigns[codeFromLink_(p.link)] || null;
      return [
        p.code, p.name, p.company, p.phone, p.link,
        c ? Number(c.attributed_users || 0) : '',
        c ? (c.active ? 'yes' : 'no') : 'not found',
        c && c.created_at ? new Date(c.created_at) : '',
        p.signedUpAt,
        now
      ];
    });

    // Most installs first: this tab exists to answer "who gets a bonus".
    out.sort(function (a, b) { return (Number(b[5]) || 0) - (Number(a[5]) || 0); });

    writeStats_(ss, out);
  } finally {
    lock.releaseLock();
  }
}

/**
 * The code as used for lookups: lowercased and trimmed.
 *
 * Our slugs are minted lowercase ("megh41"), but a display_id is only ever
 * echoed back by Linkrunner, not guaranteed to come back in the same case.
 * Comparing raw strings makes a single case difference look exactly like a
 * campaign that does not exist, which reads on the tab as "not found".
 */
function normaliseCode_(v) {
  return String(v == null ? '' : v).trim().toLowerCase();
}

/** The ?c=CODE part of a campaign link, or '' if there isn't one. */
function codeFromLink_(link) {
  var m = /[?&]c=([^&#]+)/.exec(String(link || ''));
  return m ? normaliseCode_(decodeURIComponent(m[1])) : '';
}

/**
 * Every campaign in the project, keyed by display_id (lowercased). Follows
 * pagination.
 *
 * Also keyed by the code embedded in the campaign's own link, so a campaign
 * still resolves if Linkrunner ever reports a display_id that differs from
 * the code sitting in the link we handed out.
 */
function fetchAllCampaigns_() {
  var key = PropertiesService.getScriptProperties().getProperty('LINKRUNNER_API_KEY');
  if (!key) throw new Error('LINKRUNNER_API_KEY is not set in Script Properties');

  var byId = {};
  var page = 1;

  // Bounded rather than while(true): a pagination bug on their side should not
  // turn into a script that runs until Apps Script kills it.
  for (var guard = 0; guard < 50; guard++) {
    var res = UrlFetchApp.fetch(
      'https://api.linkrunner.io/api/v1/campaigns?filter=ALL&limit=1000&page=' + page,
      { method: 'get', headers: { 'linkrunner-key': key }, muteHttpExceptions: true });

    var status = res.getResponseCode();
    var body = res.getContentText();
    if (status < 200 || status >= 300) throw new Error('Linkrunner ' + status + ': ' + body);

    var parsed = JSON.parse(body);
    var data = parsed && parsed.data;
    var list = (data && data.campaigns) || [];
    for (var i = 0; i < list.length; i++) {
      var c = list[i];
      var id = normaliseCode_(c.display_id);
      if (id) byId[id] = c;

      // Secondary key. Never allowed to clobber a real display_id match.
      var viaLink = codeFromLink_(c.link || c.shareable_link);
      if (viaLink && !byId[viaLink]) byId[viaLink] = c;
    }

    var pg = data && data.pagination;
    if (!pg || !pg.pages || page >= pg.pages) break;
    page++;
  }
  return byId;
}

/**
 * Run this from the editor when the Installs column looks wrong. It answers
 * the one question that matters: are we failing to FIND the campaigns, or
 * finding them and being told the count is zero?
 *
 * Those two have completely different fixes, and the tab alone cannot tell
 * them apart. Nothing here prints the API key.
 */
function diagnoseStats() {
  var out = [];

  var campaigns = fetchAllCampaigns_();
  var ids = Object.keys(campaigns);

  // Distinct objects, since each campaign is indexed under up to two keys.
  var seen = [], total = 0;
  ids.forEach(function (k) {
    if (seen.indexOf(campaigns[k]) === -1) {
      seen.push(campaigns[k]);
      total += Number(campaigns[k].attributed_users || 0);
    }
  });

  out.push('Campaigns returned by Linkrunner: ' + seen.length);
  out.push('Sum of attributed_users across ALL of them: ' + total);
  out.push('');
  out.push('First few, exactly as Linkrunner reports them:');
  seen.slice(0, 8).forEach(function (c) {
    out.push('   display_id=' + c.display_id +
             '  attributed_users=' + c.attributed_users +
             '  active=' + c.active +
             '  link=' + (c.link || ''));
  });

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    out.push('');
    out.push('No signups on the ' + SHEET_NAME + ' tab yet.');
    Logger.log(out.join('\n'));
    return;
  }

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  var codeCol = HEADERS.indexOf('Referral code');
  var linkCol = HEADERS.indexOf('Referral link');

  var matched = 0, unmatched = [], nonzero = 0;
  rows.forEach(function (r) {
    var code = String(r[codeCol] || '').trim();
    if (!code) return;
    var c = campaigns[normaliseCode_(code)] || campaigns[codeFromLink_(r[linkCol])] || null;
    if (c) {
      matched++;
      if (Number(c.attributed_users || 0) > 0) nonzero++;
    } else {
      unmatched.push(code);
    }
  });

  out.push('');
  out.push('Sheet codes matched to a campaign: ' + matched);
  out.push('Of those, with installs above zero: ' + nonzero);
  out.push('Sheet codes with NO matching campaign: ' + unmatched.length +
           (unmatched.length ? '  -> ' + unmatched.slice(0, 15).join(', ') : ''));

  out.push('');
  if (unmatched.length && !matched) {
    out.push('READ: nothing matched. The codes on the sheet do not correspond to');
    out.push('any campaign on this Linkrunner account - most likely the key here');
    out.push('belongs to a different Linkrunner project than the dashboard you');
    out.push('are looking at.');
  } else if (matched && !nonzero && total === 0) {
    out.push('READ: the campaigns were all found, and Linkrunner itself reports');
    out.push('attributed_users = 0 for every one of them. The sheet is showing');
    out.push('exactly what the API returns, so the zeros are not a bug here.');
    out.push('attributed_users counts APP INSTALLS attributed through the');
    out.push('Linkrunner SDK. Link clicks are a different number, and their');
    out.push('public API does not expose clicks at all - so if the dashboard');
    out.push('figure you are comparing against is clicks or visits, it will');
    out.push('never match this column.');
  } else if (matched && nonzero) {
    out.push('READ: matching works and real install counts are coming through.');
    out.push('Run syncLinkStats() and the tab should agree with this.');
  }

  Logger.log(out.join('\n'));
}

/** Replaces the tab's contents, creating it and its header row if needed. */
function writeStats_(ss, rows) {
  var sheet = ss.getSheetByName(STATS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(STATS_SHEET_NAME);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, STATS_HEADERS.length).setValues([STATS_HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  if (rows.length) {
    // Phone as text before writing, same reason as the Signups tab.
    sheet.getRange(2, 4, rows.length, 1).setNumberFormat('@');
    sheet.getRange(2, 1, rows.length, STATS_HEADERS.length).setValues(rows);
  }

  sheet.setColumnWidth(5, 280);
  sheet.setColumnWidth(2, 160);
}

/** Adds a "Link Generator" menu whenever the spreadsheet is opened. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Link Generator')
    .addItem('Refresh stats now', 'refreshNow')
    .addToUi();
}

/** Menu handler: sync, then say so, so a click never feels like nothing. */
function refreshNow() {
  var ui = SpreadsheetApp.getUi();
  try {
    syncLinkStats();
    SpreadsheetApp.getActiveSpreadsheet().toast('Link stats updated.', 'Link Generator', 5);
  } catch (err) {
    ui.alert('Could not refresh', String(err), ui.ButtonSet.OK);
  }
}

/**
 * Run this ONCE from the editor. After that the tab refreshes on its own.
 *
 * Every 5 minutes by default. Apps Script cannot go below 1 minute, and a
 * spreadsheet can never be live the way a dashboard is: it has to go and ask
 * Linkrunner each time. installSync(1) is allowed, but each run counts against
 * a daily trigger-runtime budget, so 5 buys most of the freshness for a fifth
 * of the spend. The menu covers the moments you want it now rather than soon.
 *
 * Clears its own previous trigger first, so running it twice cannot stack two.
 */
function installSync(everyMinutes) {
  var mins = everyMinutes || 5;
  var allowed = [1, 5, 10, 15, 30];
  if (allowed.indexOf(mins) === -1) {
    throw new Error('Apps Script only allows ' + allowed.join(', ') + ' minute intervals.');
  }

  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === 'syncLinkStats') {
      ScriptApp.deleteTrigger(existing[i]);
    }
  }

  ScriptApp.newTrigger('syncLinkStats').timeBased().everyMinutes(mins).create();
  syncLinkStats();   // fill it in now rather than waiting for the first tick
  Logger.log('Syncing every ' + mins + ' minute(s). Tab filled in just now.');
}

/**
 * Run this to see what the script can actually find. Reports which Script
 * Properties exist, whether the spreadsheet is attached, and which tabs exist.
 *
 * Property VALUES are never printed - only whether each is set, and how long
 * it is. An execution log is not a safe place for an API key.
 */
function checkSetup() {
  var lines = [];

  var props = PropertiesService.getScriptProperties().getProperties();
  var names = Object.keys(props);
  lines.push('Script Properties found: ' + (names.length ? names.join(', ') : 'NONE'));

  ['LINKRUNNER_API_KEY', 'REFERRAL_DESTINATION', 'REFERRAL_DESKTOP_URL'].forEach(function (k) {
    var v = props[k];
    if (k === 'LINKRUNNER_API_KEY') {
      lines.push('  ' + k + ': ' + (v ? 'set (' + String(v).length + ' characters)' : 'NOT SET  <-- this is the problem'));
    } else {
      lines.push('  ' + k + ': ' + (v ? v : '(not set, optional)'));
    }
  });

  // A key with stray whitespace looks set but fails on the wire, and is
  // invisible in the properties table.
  var key = props.LINKRUNNER_API_KEY;
  if (key && key !== key.trim()) {
    lines.push('  WARNING: the key has leading or trailing whitespace. Re-paste it.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  lines.push('Spreadsheet attached: ' + (ss ? '"' + ss.getName() + '" (' + ss.getId() + ')' : 'NO'));
  if (ss) {
    lines.push('Matches EXPECTED_SPREADSHEET_ID: ' +
      (!EXPECTED_SPREADSHEET_ID ? 'not set yet - fill it in' :
       (ss.getId() === EXPECTED_SPREADSHEET_ID ? 'yes' : 'NO - wrong sheet')));
    lines.push('Tabs: ' + ss.getSheets().map(function (sh) { return sh.getName(); }).join(', '));
  }

  var triggers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  lines.push('Triggers installed: ' + (triggers.length ? triggers.join(', ') : 'none'));

  Logger.log(lines.join('\n'));
}

/**
 * Run this once from the editor to check the wiring end to end. It writes a row
 * you can then delete.
 *
 * Throws on failure on purpose. doPost swallows its own errors and answers with
 * {ok:false}, which is right for the website but means a run can "complete"
 * having saved nothing. Better to go red here than to look fine and be wrong.
 */
function testAppendRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error(
      'This script is not attached to any spreadsheet, so it cannot write ' +
      'anything. Open the sheet, choose Extensions -> Apps Script, and paste ' +
      'the code into the editor that opens from there.'
    );
  }

  var out = doPost({
    postData: {
      contents: JSON.stringify({
        name: 'Test Boss',
        company: 'Test Co',
        phone: '9876543210',
        claimId: 'test-' + Date.now(),
        referredBy: '',
        page: 'manual test from the Apps Script editor'
      })
    }
  });

  var res = JSON.parse(out.getContent());
  if (!res.ok) throw new Error('Nothing was saved. Reason: ' + res.error);

  Logger.log('Saved. Spreadsheet: "' + ss.getName() + '"');
  Logger.log('Spreadsheet id: ' + ss.getId());
  Logger.log('Tab: "' + SHEET_NAME + '", rows now: ' + ss.getSheetByName(SHEET_NAME).getLastRow());
  Logger.log('Referral link: ' + (res.link || '(none - check LINKRUNNER_API_KEY)'));
  Logger.log('Open it: ' + ss.getUrl());
}
