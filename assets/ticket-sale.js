/* Ticket sale window — edit constants below (Asia/Hong_Kong wall time). */
window.NOAHS_TICKET_SALE = (function () {
  /* Flip Metal/早鳥 → Last Call/預售 at this instant. */
  var PRESALE_START = '2026-11-21T00:00:00+08:00';
  /* After the flip, embed this Tally form ID. null = keep current form. */
  var TALLY_FORM_AFTER = null;
  var TALLY_FORM_BEFORE = 'J9b2zY';
  var TALLY_EMBED_QS = 'alignLeft=1&hideTitle=1&transparentBackground=1&dynamicHeight=1';
  var SAFETY_MS = 60000;

  /* ?now=ISO starts a simulated clock that advances with real elapsed time,
     so ?now=…23:59:50+08:00 flips ~10s later without reload. */
  var nowAnchorMs = null;   /* overridden wall-clock at page load */
  var realAnchorMs = null;  /* performance/real ms when override was read */

  function initNowOverride() {
    try {
      var raw = new URLSearchParams(window.location.search).get('now');
      if (!raw) return;
      var d = new Date(raw);
      if (isNaN(d.getTime())) return;
      nowAnchorMs = d.getTime();
      realAnchorMs = Date.now();
    } catch (e) { /* ignore */ }
  }
  initNowOverride();

  function getNow() {
    if (nowAnchorMs != null && realAnchorMs != null) {
      return new Date(nowAnchorMs + (Date.now() - realAnchorMs));
    }
    return new Date();
  }

  function getPresaleStart() {
    return new Date(PRESALE_START);
  }

  function isPresaleOpen(now) {
    var n = now || getNow();
    return n.getTime() >= getPresaleStart().getTime();
  }

  function selectableKeys(now) {
    return isPresaleOpen(now)
      ? ['lastcall', 'presale']
      : ['metal', 'earlybird'];
  }

  function isSelectable(key, now) {
    return selectableKeys(now).indexOf(key) !== -1;
  }

  function getTallyFormId(now) {
    if (isPresaleOpen(now) && TALLY_FORM_AFTER) return TALLY_FORM_AFTER;
    return TALLY_FORM_BEFORE;
  }

  /**
   * Build Tally qty query. Only currently-selectable types are included;
   * selected key → 1, other selectable → 0. Disabled types omitted.
   * selectedKey null/undefined → all selectable = 0.
   */
  function buildQtyParams(selectedKey, now) {
    var keys = selectableKeys(now);
    var map = {
      metal: 'qty_metal',
      earlybird: 'qty_earlybird',
      lastcall: 'qty_metal_lastcall',
      presale: 'qty_presale'
    };
    var parts = [];
    keys.forEach(function (k) {
      parts.push(encodeURIComponent(map[k]) + '=' + (k === selectedKey ? '1' : '0'));
    });
    return parts.join('&');
  }


  function readMemberPrefill() {
    /* Prefer logged-in session; else registration prefill. */
    try {
      var raw = sessionStorage.getItem('na_member');
      if (raw) {
        var s = JSON.parse(raw);
        if (s && s.profile) return s.profile;
      }
    } catch (e) {}
    try {
      var raw2 = sessionStorage.getItem('na_prefill');
      if (raw2) return JSON.parse(raw2);
    } catch (e2) {}
    return null;
  }

  function appendMemberPrefill(src) {
    var p = readMemberPrefill();
    if (!p) return src;
    var parts = [];
    if (p.member_no) parts.push('member_no=' + encodeURIComponent(p.member_no));
    if (p.name) parts.push('name=' + encodeURIComponent(p.name));
    if (p.phone) parts.push('phone=' + encodeURIComponent(p.phone));
    if (p.email) parts.push('email=' + encodeURIComponent(p.email));
    if (!parts.length) return src;
    return src + (src.indexOf('?') >= 0 ? '&' : '?') + parts.join('&');
  }

  function buildEmbedSrc(selectedKey, now) {
    var n = now || getNow();
    var phase = isPresaleOpen(n) ? 'post' : 'pre';
    var base = 'https://tally.so/embed/' + getTallyFormId(n) + '?' + TALLY_EMBED_QS + '&phase=' + phase;
    var qs = buildQtyParams(selectedKey, n);
    var src = qs ? base + '&' + qs : base;
    return appendMemberPrefill(src);
  }

  function stickyPricesHtml(now) {
    if (isPresaleOpen(now)) {
      return (
        '<span class="lang-zh">會員 $420 · 預售 $450</span>' +
        '<span class="lang-en lang-hidden">Member $420 · Presale $450</span>'
      );
    }
    return (
      '<span class="lang-zh">Metal $350 · 早鳥 $380 · 11/21 開售 $420／$450</span>' +
      '<span class="lang-en lang-hidden">Metal $350 · Early bird $380 · On sale 21 Nov $420/$450</span>'
    );
  }

  function syncLangVisibility(root) {
    if (!root) return;
    var lang = (document.documentElement.lang === 'en') ? 'en' : 'zh';
    root.querySelectorAll('.lang-zh').forEach(function (el) {
      el.classList.toggle('lang-hidden', lang === 'en');
    });
    root.querySelectorAll('.lang-en').forEach(function (el) {
      el.classList.toggle('lang-hidden', lang === 'zh');
    });
  }

  function applyStickyBar(el, now) {
    if (!el) return;
    el.innerHTML = stickyPricesHtml(now || getNow());
    syncLangVisibility(el);
  }

  function applyTicketCards(cards, now) {
    var n = now || getNow();
    var open = isPresaleOpen(n);
    Array.prototype.forEach.call(cards, function (card) {
      var key = card.getAttribute('data-ticket');
      var on = isSelectable(key, n);
      card.classList.toggle('is-disabled', !on);
      card.classList.toggle('is-primary', open && key === 'presale');
      card.setAttribute('aria-disabled', on ? 'false' : 'true');
      if (!on) {
        card.classList.remove('is-selected');
        card.setAttribute('aria-selected', 'false');
      }
      var status = card.querySelector('.ticket-pick__status');
      if (!status) {
        status = document.createElement('div');
        status.className = 'ticket-pick__status';
        card.appendChild(status);
      }
      if (!on) {
        if (open) {
          status.innerHTML =
            '<span class="lang-zh">已停售</span><span class="lang-en lang-hidden">Sold out</span>';
        } else {
          status.innerHTML =
            '<span class="lang-zh">11月21日開售</span><span class="lang-en lang-hidden">On sale 21 Nov</span>';
        }
        status.hidden = false;
      } else {
        status.innerHTML = '';
        status.hidden = true;
      }
      syncLangVisibility(card);
    });
  }

  /* ---- live phase watcher ---- */
  var listeners = [];
  var lastOpen = null;
  var flipTimer = null;
  var safetyTimer = null;
  var watching = false;

  function notify(phaseChanged) {
    var now = getNow();
    var open = isPresaleOpen(now);
    var payload = { phaseChanged: !!phaseChanged, open: open, now: now };
    listeners.forEach(function (fn) {
      try { fn(payload); } catch (e) { /* page handler error */ }
    });
  }

  function evaluate() {
    var open = isPresaleOpen();
    if (lastOpen === null) {
      lastOpen = open;
      return false;
    }
    if (open !== lastOpen) {
      lastOpen = open;
      notify(true);
      return true;
    }
    return false;
  }

  function clearFlipTimer() {
    if (flipTimer != null) {
      clearTimeout(flipTimer);
      flipTimer = null;
    }
  }

  function scheduleExactFlip() {
    clearFlipTimer();
    var ms = getPresaleStart().getTime() - getNow().getTime();
    if (ms <= 0) {
      evaluate();
      return;
    }
    /* setTimeout max ~24.8 days; beyond that the 60s safety net covers it. */
    if (ms > 2147483647) return;
    flipTimer = setTimeout(function () {
      flipTimer = null;
      evaluate();
      /* if still before (clock skew / sim), reschedule */
      if (!isPresaleOpen()) scheduleExactFlip();
    }, ms + 30);
  }

  function startWatching(onPhaseChange) {
    if (typeof onPhaseChange === 'function') listeners.push(onPhaseChange);
    if (watching) {
      scheduleExactFlip();
      return;
    }
    watching = true;
    lastOpen = isPresaleOpen();
    scheduleExactFlip();
    safetyTimer = setInterval(function () {
      evaluate();
      /* keep exact timer honest if ?now= is advancing or tab slept */
      scheduleExactFlip();
    }, SAFETY_MS);
    document.addEventListener('visibilitychange', function () {
      if (document.visibilityState === 'visible') {
        evaluate();
        scheduleExactFlip();
      }
    });
  }

  /** ms until flip from current getNow() — handy for tests */
  function msUntilFlip() {
    return getPresaleStart().getTime() - getNow().getTime();
  }

  return {
    PRESALE_START: PRESALE_START,
    TALLY_FORM_AFTER: TALLY_FORM_AFTER,
    TALLY_FORM_BEFORE: TALLY_FORM_BEFORE,
    getNow: getNow,
    getPresaleStart: getPresaleStart,
    isPresaleOpen: isPresaleOpen,
    selectableKeys: selectableKeys,
    isSelectable: isSelectable,
    getTallyFormId: getTallyFormId,
    buildQtyParams: buildQtyParams,
    buildEmbedSrc: buildEmbedSrc,
    readMemberPrefill: readMemberPrefill,
    appendMemberPrefill: appendMemberPrefill,
    stickyPricesHtml: stickyPricesHtml,
    applyStickyBar: applyStickyBar,
    applyTicketCards: applyTicketCards,
    syncLangVisibility: syncLangVisibility,
    startWatching: startWatching,
    msUntilFlip: msUntilFlip
  };
})();
