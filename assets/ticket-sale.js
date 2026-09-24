/* Ticket sale window — edit ONE constant below (Asia/Hong_Kong wall time). */
window.NOAHS_TICKET_SALE = (function () {
  var PRESALE_START = '2026-11-21T00:00:00+08:00';

  function parseNowOverride() {
    try {
      var raw = new URLSearchParams(window.location.search).get('now');
      if (!raw) return null;
      var d = new Date(raw);
      return isNaN(d.getTime()) ? null : d;
    } catch (e) {
      return null;
    }
  }

  function getNow() {
    return parseNowOverride() || new Date();
  }

  function getPresaleStart() {
    return new Date(PRESALE_START);
  }

  /** true once clock (or ?now=) reaches PRESALE_START */
  function isPresaleOpen(now) {
    var n = now || getNow();
    return n.getTime() >= getPresaleStart().getTime();
  }

  /**
   * Selectable ticket keys for the current phase.
   * Before: Metal $350 + 早鳥 $380
   * After:  Metal Last Call $420 + 預售 $450
   */
  function selectableKeys(now) {
    return isPresaleOpen(now)
      ? ['lastcall', 'presale']
      : ['metal', 'earlybird'];
  }

  function isSelectable(key, now) {
    return selectableKeys(now).indexOf(key) !== -1;
  }

  /**
   * Build Tally qty query. Only currently-selectable types are included;
   * selected key → 1, other selectable → 0. Disabled types omitted.
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

  /** Sticky-bar price line HTML (uses .lang-zh / .lang-en) */
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
    el.innerHTML = stickyPricesHtml(now);
    syncLangVisibility(el);
  }

  /**
   * Apply disabled / primary / status badge state to .ticket-pick__card nodes.
   * Optional .ticket-pick__status created if missing.
   */
  function applyTicketCards(cards, now) {
    var open = isPresaleOpen(now);
    Array.prototype.forEach.call(cards, function (card) {
      var key = card.getAttribute('data-ticket');
      var on = isSelectable(key, now);
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

  return {
    PRESALE_START: PRESALE_START,
    getNow: getNow,
    getPresaleStart: getPresaleStart,
    isPresaleOpen: isPresaleOpen,
    selectableKeys: selectableKeys,
    isSelectable: isSelectable,
    buildQtyParams: buildQtyParams,
    stickyPricesHtml: stickyPricesHtml,
    applyStickyBar: applyStickyBar,
    applyTicketCards: applyTicketCards,
    syncLangVisibility: syncLangVisibility
  };
})();
