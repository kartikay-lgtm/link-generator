# Tracking clicks / installs / sign-ups per referral link (Linkrunner)

Context handoff from the **link-generator** build (link-generator-tan.vercel.app),
for reuse on the gift.tal.club flow. This is working in production: a Google
Sheet tab with one row per referral link and live Clicks / Installs / Sign-ups
columns, refreshed every 5 minutes by an Apps Script trigger.

Everything below was learned the hard way. The short version: **use the
Reporting API, not the Data API**, and there are four gotchas that each cost a
round of debugging.

---

## 1. Use the Reporting API. The Data API cannot do this.

```
GET https://api.linkrunner.io/api/v1/reporting/campaigns
Header: linkrunner-key: <server key from Settings → Data APIs>
```

This is the endpoint behind the dashboard's campaign table, so its numbers are
the ones on screen. It returns `clicks`, `installs`, `"sign-ups"`, plus spend,
revenue, ROAS, retention, ad sets.

**Two dead ends worth not repeating:**

| Endpoint | Why it doesn't work |
|---|---|
| `GET /campaigns` → `attributed_users` | **Lags badly.** Read `0` for a link that genuinely had activity. Do not trust it. |
| `GET /attributed-users` → `pagination.total` | Counts **sign-ups only**, not installs. An "attributed user" only exists once the app reports a signup, so someone who installs and never signs up is invisible. A link showing 2 installs / 1 sign-up returns 1. |

I spent two debugging rounds concluding "clicks and installs aren't available
anywhere" based on the Data APIs doc page, which lists only three endpoints
(`/campaigns`, `/attributed-users`, `/get-attribution-result`). The Reporting
API is a **separate doc page**. It exists. Go straight there.

---

## 2. Rate limit is 1 request per MINUTE per key

Not the 30/sec figure — that's a per-IP limit shared across other endpoints.

- `limit` caps at **100** campaigns per page.
- So an account with >100 campaigns needs pagination, and requesting page 2
  immediately returns **429** with `Retry-After: 60`.
- Sleep ~61s between pages, and honour `Retry-After` when a 429 arrives anyway.

**Scaling cliff to plan for:** each page costs a minute, and Apps Script kills
any execution at **6 minutes**. So this breaks somewhere around **500–700
campaigns**. At 175 campaigns (2 pages, ~1 min) it's fine. If gift.tal.club will
generate thousands of links, the sync needs to be resumable across trigger runs
— store the page cursor in Script Properties and pick up where it left off.

---

## 3. Every number is a formatted display string

The API returns `"3,201"`, `"$12,540.50"`, `"95.38%"` — not numbers.

Written straight into Sheets they land as **text**, and sorting ranks `"9"`
above `"1,200"`. Strip before use:

```js
function reportingNumber_(v) {
  if (v === null || v === undefined || v === '') return '';
  var n = Number(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? '' : n;
}
```

Also: the sign-ups key is **`"sign-ups"`** — hyphenated, not `signups` or
`sign_ups`. Reading the wrong spelling yields `undefined`, which writes a blank
cell rather than throwing, so it fails silently.

---

## 4. `from` / `to` set the metrics window, they don't filter the list

A campaign created after `to` still appears in the response, with zeros. This
matches the dashboard, where the date picker changes the numbers, not the rows.

Omit both for lifetime totals — which is what a referral programme should be
paid on.

---

## Matching campaigns back to your sheet rows

Two traps:

1. **`display_id` case is not guaranteed to survive the round trip.** Our slugs
   are minted lowercase, but comparing raw strings made a live campaign look
   like it didn't exist. Normalise both sides.
2. **A campaign's `display_id` and the code in its link can differ.** Index by
   both, with `display_id` winning.

```js
function normaliseCode_(v) {
  return String(v == null ? '' : v).trim().toLowerCase();
}

function codeFromLink_(link) {
  var m = /[?&]c=([^&#]+)/.exec(String(link || ''));
  return m ? normaliseCode_(decodeURIComponent(m[1])) : '';
}
```

Then on lookup: `reporting[normaliseCode_(code)] || reporting[codeFromLink_(link)]`.

---

## The fetch function (lift this directly)

```js
var REPORTING_URL = 'https://api.linkrunner.io/api/v1/reporting/campaigns';

/** Returns { normalisedCode: {clicks, installs, signups, active, created_at, link} } */
function fetchReportingCampaigns_() {
  var key = PropertiesService.getScriptProperties().getProperty('LINKRUNNER_API_KEY');
  if (!key) throw new Error('LINKRUNNER_API_KEY is not set in Script Properties');

  var byId = {};
  var page = 1, pages = 1, fetched = 0;

  for (var guard = 0; guard < 12 && page <= pages; guard++) {
    if (fetched > 0) Utilities.sleep(61000);           // 1 req/min per key

    var url = REPORTING_URL + '?limit=100&page=' + page +
              '&sort_field=installs&sort_order=descending';

    var res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'linkrunner-key': key },
      muteHttpExceptions: true
    });

    var status = res.getResponseCode();
    var body = res.getContentText();

    if (status === 429) {
      var hdrs = res.getHeaders() || {};
      var wait = Number(hdrs['Retry-After'] || hdrs['retry-after'] || 60);
      if (!wait || isNaN(wait)) wait = 60;
      Utilities.sleep((wait + 2) * 1000);
      fetched = 0;
      continue;                                        // retry the same page
    }
    if (status < 200 || status >= 300) {
      throw new Error('Linkrunner reporting ' + status + ': ' + body.slice(0, 300));
    }

    fetched++;

    var parsed = JSON.parse(body);
    var data = parsed && parsed.data;
    var list = (data && data.campaigns) || [];

    for (var i = 0; i < list.length; i++) {
      var c = list[i];
      var rec = {
        display_id: c.display_id,
        clicks:   reportingNumber_(c.clicks),
        installs: reportingNumber_(c.installs),
        signups:  reportingNumber_(c['sign-ups']),     // hyphenated key
        active:   c.active,
        created_at: c.created_at,
        link: c.link
      };

      var id = normaliseCode_(c.display_id);
      if (id) byId[id] = rec;

      var viaLink = codeFromLink_(c.link);
      if (viaLink && !byId[viaLink]) byId[viaLink] = rec;
    }

    var pg = data && data.pagination;
    pages = (pg && Number(pg.pages)) || 1;
    page++;
  }

  return byId;
}
```

---

## Apps Script traps this exposed (both cost real downtime)

### Never hold the script lock across the 60-second waits

The sync originally took `LockService.getScriptLock()` for its whole run, with
`Utilities.sleep(61000)` inside it. The form's `doPost` waits only 15s for that
same lock — so **every signup during a sync was refused with "busy"**: a
~60-second dead window every 5 minutes, at zero traffic.

Do all network I/O **outside** the lock. Take it only for the sheet write
(sub-second).

### `clearContents()` leaves number formats behind

Rebuilding the stats tab with `clearContents()` preserved each column's old
format. When a new column landed where a date column used to be, Sheets
rendered `1` as `12/31/1899` and `0` as `12/30/1899`.

Set formats explicitly on every write, **keyed by header name, not position** —
a positional map desynchronises the moment a column is inserted, which is
exactly how it broke:

```js
var STATS_FORMATS = {
  'Clicks': '0', 'Installs': '0', 'Sign-ups': '0',
  'Phone': '@', 'Referral link': '@',
  'Link created': 'dd mmm yyyy', 'Last synced': 'dd mmm yyyy hh:mm'
};
// after clearContents(), before writing values:
for (var i = 0; i < STATS_HEADERS.length; i++) {
  var f = STATS_FORMATS[STATS_HEADERS[i]];
  if (f) sheet.getRange(2, i + 1, depth, 1).setNumberFormat(f);
}
```

---

## Which metric to pay on

Clicks → Installs → Sign-ups is the funnel. A referral programme paying for
"approved" referrals should key on **sign-ups**: an install that never signs up
isn't a qualified referral. We sort the tab by sign-ups, installs as tiebreak.

---

## Reference

- Reporting API doc: `docs.linkrunner.io` → Reporting API (separate from "Data APIs")
- Working implementation: `github.com/kartikay-lgtm/link-generator` →
  `google-apps-script/Code.gs` — see `fetchReportingCampaigns_`, `syncLinkStats`,
  `writeStats_`, and `inspectCampaign` (a diagnostic that prints the Reporting
  API numbers next to both Data API sources for one code, for when a row
  disagrees with the dashboard).
