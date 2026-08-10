/* Link Generator — form handling, submission, and referral link display.
   Plain browser JS. No build step, no framework, nothing fetched at runtime. */
(function () {
  'use strict';

  var ENDPOINT = (window.LINKGEN_ENDPOINT || '').trim();

  var form    = document.getElementById('form');
  var alertEl = document.getElementById('alert');
  var btn     = document.getElementById('submit');
  var label   = document.getElementById('btn-label');

  var FIELDS = ['name', 'company', 'phone'];
  var ERRORS = {
    name:    'Enter your full name',
    company: 'Enter your company name',
    phone:   'Enter a 10-digit mobile number'
  };

  function el(id) { return document.getElementById(id); }
  function input(n) { return el('f-' + n); }
  function errBox(n) { return el('e-' + n); }

  function digits(v) { return String(v == null ? '' : v).replace(/[^0-9]/g, ''); }

  // Accepts what people actually type: +91 98765 43210, 098765 43210, and so on.
  function normalisePhone(raw) {
    var d = digits(raw);
    if (d.length === 12 && d.slice(0, 2) === '91') d = d.slice(2);
    if (d.length === 11 && d[0] === '0') d = d.slice(1);
    return d;
  }

  function fieldError(name, raw) {
    var v = String(raw == null ? '' : raw).trim();
    if (name === 'name')    return v.length >= 2 ? '' : ERRORS.name;
    if (name === 'company') return v.length >= 2 ? '' : ERRORS.company;
    if (name === 'phone')   return /^[6-9][0-9]{9}$/.test(normalisePhone(v)) ? '' : ERRORS.phone;
    return '';
  }

  function showError(name, msg) {
    var box = errBox(name), field = input(name);
    if (msg) { box.textContent = msg; box.classList.remove('hidden'); field.classList.add('bad'); }
    else     { box.textContent = '';  box.classList.add('hidden');    field.classList.remove('bad'); }
  }

  function hideAlert() { alertEl.classList.add('hidden'); alertEl.textContent = ''; }
  function showAlert(m) { alertEl.textContent = m; alertEl.classList.remove('hidden'); }

  // Clear a shown error as soon as someone starts fixing it, but never raise a
  // new one mid-keystroke - that scolds people while they are still typing.
  FIELDS.forEach(function (name) {
    var field = input(name);
    field.addEventListener('input', function () {
      if (!errBox(name).classList.contains('hidden')) showError(name, '');
      hideAlert();
    });
    field.addEventListener('blur', function () {
      showError(name, fieldError(name, field.value));
    });
  });

  function setSending(on) {
    btn.disabled = on;
    label.textContent = on ? 'Generating your link' : 'Generate my referral link';
    var spinner = btn.querySelector('.spin');
    if (on && !spinner) {
      var s = document.createElement('span');
      s.className = 'spin';
      s.setAttribute('aria-hidden', 'true');
      btn.insertBefore(s, label);
    }
    if (!on && spinner) spinner.remove();
  }

  /* One id per submission. It rides along so a submission re-sent after a lost
     reply is recognised as the same one, rather than minting a second link. */
  function makeClaimId() {
    try {
      if (window.crypto && window.crypto.randomUUID) return window.crypto.randomUUID();
    } catch (e) {}
    return 'g-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 10);
  }

  /* Whoever referred THIS visitor, if they arrived on someone else's link.
     Generating links is only half of attribution; without reading the inbound
     code there is no way to trace a chain back. */
  function referredBy() {
    try { return new URLSearchParams(window.location.search).get('ref') || ''; }
    catch (e) { return ''; }
  }

  /* Sends one submission to the Apps Script web app.
     Resolves { ok, link }.

     Content-Type is deliberately text/plain: that keeps this a "simple"
     request, and simple requests skip the CORS preflight. Apps Script has no
     doOptions, so anything that preflights is refused before doPost runs. The
     body is still JSON; the script parses it itself. */
  function save(body, claimId) {
    if (!ENDPOINT) {
      console.warn('[link-generator] No endpoint set. Nothing was saved.\n' +
                   'Paste the Apps Script /exec URL into config.js (see SETUP.md).', body);
      return Promise.resolve({ ok: false, link: '' });
    }

    function postOpts() {
      return {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: body
      };
    }

    // Apps Script answers a POST with a redirect and routinely takes a few
    // seconds, more when it also has to mint a link. Bound the wait here rather
    // than leaving it to the browser, which on a poor connection hangs far
    // longer than anyone will sit through.
    function post() {
      var opts = postOpts(), timer = null;
      if (typeof AbortController === 'function') {
        var ctrl = new AbortController();
        opts.signal = ctrl.signal;
        timer = setTimeout(function () { ctrl.abort(); }, 25000);
      }
      function clear() { if (timer) clearTimeout(timer); }
      return fetch(ENDPOINT, opts)
        .then(function (r) { return r.json(); })
        .then(function (r) {
          clear();
          // Surfaced in the console, never to the visitor: if a link failed to
          // mint, the reason should be one glance away rather than buried in an
          // Apps Script execution log.
          if (r && r.linkError) console.warn('[link-generator] link not minted:', r.linkError);
          return { ok: !!(r && r.ok), link: (r && r.link) || '' };
        }, function (e) { clear(); throw e; });
    }

    /* The bit that stops the guessing: ask the sheet whether this submission is
       already recorded, and get back the link it was given. An unreadable reply
       is not the same as a failed write, and assuming the worst is what makes a
       saved entry show up as an error. */
    function lookUp() {
      var sep = ENDPOINT.indexOf('?') === -1 ? '?' : '&';
      return fetch(ENDPOINT + sep + 'check=' + encodeURIComponent(claimId))
        .then(function (r) { return r.json(); })
        .then(function (j) {
          return (j && j.saved) ? { state: 'saved', link: j.link || '' }
                                : { state: 'absent', link: '' };
        })
        .catch(function () { return { state: 'unreadable', link: '' }; });
    }

    return post().catch(function () {
      return lookUp().then(function (res) {
        if (res.state === 'saved') return { ok: true, link: res.link };

        // Nothing readable at all. Without a reply we cannot learn the link, and
        // a link we cannot show is no use to the person waiting - so report the
        // failure rather than send them to a blank result page.
        if (res.state === 'unreadable') return { ok: false, link: '' };

        // Genuinely absent: one more try, then look again before giving up.
        return post().catch(function () {
          return lookUp().then(function (r2) {
            return { ok: r2.state === 'saved', link: r2.link };
          });
        });
      });
    });
  }

  function showResult(link, name) {
    el('link').value = link;
    if (name) {
      el('done-title').textContent =
        'here is your link, ' + name.trim().split(/\s+/)[0].toLowerCase();
    }
    el('view-form').classList.add('hidden');
    el('view-done').classList.remove('hidden');
    window.scrollTo(0, 0);
  }

  el('copy').addEventListener('click', function () {
    var field = el('link'), btnCopy = el('copy');
    function flash() {
      btnCopy.textContent = 'Copied';
      setTimeout(function () { btnCopy.textContent = 'Copy'; }, 1600);
    }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(field.value).then(flash, function () {
        field.select(); document.execCommand('copy'); flash();
      });
    } else {
      field.select(); document.execCommand('copy'); flash();
    }
  });

  /* The link rides inside the text itself rather than a separate url param -
     X folds a bare url param into the text anyway, and this way the message
     still reads right if a platform only keeps one of the two. */
  function shareText() {
    var link = el('link').value.trim();
    var msg = "I'm hiring on Tal. It's the app where bosses of Bangalore can hire great talent.";
    return link ? msg + '\n' + link : msg;
  }

  /* A plain new tab, not a popup window: no width/height/left/top means the
     browser opens it exactly like a normal clicked link, which is what looks
     right and behaves right on a phone - a fixed-size popup either gets
     ignored in favour of a full tab there anyway, or renders too small to
     use. */
  function openShare(url) {
    window.open(url, '_blank', 'noopener,noreferrer');
  }

  el('share-x').addEventListener('click', function () {
    openShare('https://twitter.com/intent/tweet?text=' + encodeURIComponent(shareText()));
  });

  /* Synchronous on purpose - document.execCommand('copy') runs and finishes
     in the same tick, unlike navigator.clipboard.writeText() which resolves
     as a microtask. That distinction mattered here: writing to the clipboard
     *after* kicking off window.open() raced the new tab taking focus, and
     losing that race makes Clipboard API throw "document is not focused" -
     silently, since nothing awaited it. That silent failure is exactly what
     "no pre-filled text at all" looks like. Copying first and only opening
     the tab once the copy has actually finished removes the race. */
  function copyToClipboard(text) {
    var area = document.createElement('textarea');
    area.value = text;
    area.style.position = 'fixed';
    area.style.opacity = '0';
    document.body.appendChild(area);
    area.focus();
    area.select();
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (e) {}
    document.body.removeChild(area);
    return ok;
  }

  function flashShareNote(msg) {
    var note = el('share-note');
    note.textContent = msg;
    note.classList.remove('hidden');
    clearTimeout(flashShareNote._t);
    flashShareNote._t = setTimeout(function () { note.classList.add('hidden'); }, 4000);
  }

  el('share-linkedin').addEventListener('click', function () {
    var link = el('link').value.trim();

    // LinkedIn stopped honouring title/summary params years ago - the old
    // shareArticle endpoint silently drops them. Their only supported param
    // today is url, and the post body is whatever the person types
    // themselves. Best we can do without that param is put the caption on
    // the clipboard first, so a paste finishes the job LinkedIn won't let a
    // link do.
    var copied = copyToClipboard(shareText());
    flashShareNote(copied
      ? 'caption copied — paste it into your LinkedIn post'
      : "couldn't copy automatically — write your own line before posting");

    openShare('https://www.linkedin.com/sharing/share-offsite/?url=' + encodeURIComponent(link || 'https://tal.club/'));
  });

  form.addEventListener('submit', function (e) {
    e.preventDefault();
    hideAlert();

    var values = {};
    FIELDS.forEach(function (n) { values[n] = input(n).value; });

    var bad = false;
    FIELDS.forEach(function (n) {
      var msg = fieldError(n, values[n]);
      showError(n, msg);
      if (msg) bad = true;
    });
    if (bad) return;

    // Bot filled the hidden field: swallow it silently rather than tell them.
    if (document.querySelector('input[name=website]').value) {
      showResult('https://tal.club/', '');
      return;
    }

    setSending(true);

    var claimId = makeClaimId();
    var payload = {
      name:       values.name.trim(),
      company:    values.company.trim(),
      phone:      normalisePhone(values.phone),
      claimId:    claimId,
      referredBy: referredBy(),
      page: (function () { try { return window.location.href; } catch (e) { return ''; } })()
    };

    var started = Date.now();
    save(JSON.stringify(payload), claimId).then(function (res) {
      // Hold the spinner briefly even on an instant reply, so a success state
      // never appears before the click has visibly registered.
      var wait = Math.max(0, 450 - (Date.now() - started));
      setTimeout(function () {
        setSending(false);
        if (res.ok && res.link) {
          showResult(res.link, payload.name);
        } else if (res.ok) {
          // Saved, but no link came back. Not a failure - the entry is safe and
          // someone can follow up, so do not send them round again.
          showResult('', payload.name);
          el('link').value = '(your link is being prepared, we will send it across)';
          el('copy').classList.add('hidden');
        } else {
          showAlert('We could not generate your link just now. Nothing was lost, try again.');
        }
      }, wait);
    });
  });
})();
