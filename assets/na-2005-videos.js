/*
 * Noah's Ark 2005 — live videos (single source of truth).
 *
 * HOW TO ADD THE NEXT VIDEO
 *   1. Add one object to NA_2005_VIDEOS below:
 *        { band: 'deathguy',            // must match the band page slug (/live/2005/<band>)
 *          id: 'XXXXXXXXXXX',           // YouTube video id (the part after youtu.be/)
 *          zh: "Deathguy｜Noah's Ark 2005",        // heading shown in 中文
 *          en: "Deathguy | Noah's Ark 2005" }      // heading shown in English
 *      Order in this array = order on the band page and on /videos.
 *      hero: true = event-wide clip (no band): shown at the top of /live/2005.
 *   2. Bump ?v= on the <script src=".../na-2005-videos.js?v=..."> tags
 *      (live/2005.html, live/2005/*.html, videos.html) so browsers refetch.
 *   That's it: the band page shows the player, the /live/2005 tile gets a ▶ badge,
 *   and /videos lists it in the "Noah's Ark 2005" group automatically.
 *   Bands with no entry show nothing extra (no placeholder text).
 *   If the band has NO 2005 band page (not in BAND_PAGES below), the video is listed
 *   in the "視頻 / Videos" section at the bottom of /live/2005 instead. If a band page
 *   is added later (live/2005/<slug>.html + _redirects + tile), add its slug to BAND_PAGES.
 */
window.NA_2005_BAND_PAGES = ['bereavement', 'hermetic-silence', 'deathguy', 'departing-cross', 'cadaver'];
window.NA_2005_VIDEOS = [
  /* "Noahs Ark 2005 11 02m41s" — belongs to no band: hero video at the top of /live/2005,
     and listed first on /videos. (hero: true → not tied to a band page) */
  { band: '', hero: true, id: '7pnPY4vzt7Y',
    zh: "Noah's Ark 2005", en: "Noah's Ark 2005" },
  /* Official channel @noahsark_2005 — "Noahs Ark 2005 01 05m16s" */
  { band: 'hermetic-silence', id: 'gXsVsvZMU24',
    zh: "Hermetic Silence｜Noah's Ark 2005（一）", en: "Hermetic Silence | Noah's Ark 2005 Part 1" },
  /* "Noahs Ark 2005 02 07m42s" */
  { band: 'hermetic-silence', id: 'sKTAmFO6n7M',
    zh: "Hermetic Silence｜Noah's Ark 2005（二）", en: "Hermetic Silence | Noah's Ark 2005 Part 2" },
  /* "Noahs Ark 2005 03 05m58s" */
  { band: 'cadaver', id: 'x1a-ksNRUz0',
    zh: "Cadaver｜Noah's Ark 2005（一）", en: "Cadaver | Noah's Ark 2005 Part 1" },
  /* "Noahs Ark 2005 04 05m02s" */
  { band: 'cadaver', id: 'rdFSwvL4f_U',
    zh: "Cadaver｜Noah's Ark 2005（二）", en: "Cadaver | Noah's Ark 2005 Part 2" },
  /* "Noahs Ark 2005 05 14m53s" */
  { band: 'bereavement', id: 'c3iM1PZqPgM',
    zh: "Bereavement｜Noah's Ark 2005", en: "Bereavement | Noah's Ark 2005" },
  /* "Noahs Ark 2005 06 06m17s" — no 2005 band page/tile for Evocation (yet),
     so it is listed in the "Videos" section at the bottom of /live/2005. */
  { band: 'evocation', id: 'Tx-59Kuqz2U',
    zh: "招魂 Evocation｜Noah's Ark 2005（一）", en: "Evocation | Noah's Ark 2005 Part 1" },
  /* "Noahs Ark 2005 07 06m57s" */
  { band: 'evocation', id: 'ExJogtiSs1c',
    zh: "招魂 Evocation｜Noah's Ark 2005（二）", en: "Evocation | Noah's Ark 2005 Part 2" },
  /* "Noahs Ark 2005 08 05m34s" */
  { band: 'deathguy', id: 'p15e95AFolI',
    zh: "Deathguy｜Noah's Ark 2005（一）", en: "Deathguy | Noah's Ark 2005 Part 1" },
  /* "Noahs Ark 2005 09 05m03s" */
  { band: 'deathguy', id: 'us-tjJx47U0',
    zh: "Deathguy｜Noah's Ark 2005（二）", en: "Deathguy | Noah's Ark 2005 Part 2" },
  /* "Noahs Ark 2005 10 03m05s" */
  { band: 'deathguy', id: '8YmhkjIB7mg',
    zh: "Deathguy｜Noah's Ark 2005（三）", en: "Deathguy | Noah's Ark 2005 Part 3" },
  /* "Noahs Ark 2005 12 11m02s" (no. 11 not published yet — insert it above this line when it is) */
  { band: 'departing-cross', id: '0xu1yudwl00',
    zh: "Departing Cross｜Noah's Ark 2005", en: "Departing Cross | Noah's Ark 2005" }
];

(function () {
  var VIDEOS = window.NA_2005_VIDEOS || [];
  var BAND_PAGES = window.NA_2005_BAND_PAGES || [];
  function hasPage(slug) { return BAND_PAGES.indexOf(slug) !== -1; }

  function injectCss() {
    if (document.getElementById('na2005-video-css')) return;
    var css =
      '.na-yt{position:relative;width:100%;aspect-ratio:16/9;background:#000;border-radius:inherit;overflow:hidden}' +
      '.na-yt__btn{position:absolute;inset:0;width:100%;height:100%;padding:0;border:0;margin:0;cursor:pointer;background:#000;display:block}' +
      '.na-yt__btn img{position:absolute;inset:0;width:100%;height:100%;max-width:none;object-fit:cover;opacity:.9;transition:opacity .2s ease,transform .3s ease}' +
      '.na-yt__btn:hover img,.na-yt__btn:focus-visible img{opacity:1;transform:scale(1.02)}' +
      '.na-yt__play{position:absolute;left:50%;top:50%;width:68px;height:48px;transform:translate(-50%,-50%);' +
        'border-radius:14px;background:rgba(185,28,28,.92);box-shadow:0 6px 22px rgba(0,0,0,.55);display:flex;align-items:center;justify-content:center}' +
      '.na-yt__play::after{content:"";display:block;margin-left:4px;border-style:solid;border-width:10px 0 10px 17px;border-color:transparent transparent transparent #fff}' +
      '.na-yt__btn:focus-visible{outline:2px solid #ef4444;outline-offset:-2px}' +
      '.na-yt iframe{position:absolute;inset:0;width:100%;height:100%;border:0}' +
      '.na2005-video{margin:0 0 1.25rem;min-width:0}' +
      '@media (min-width:768px){.band-detail__video [data-na2005-videos]{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:0 1rem}}' +
      '.na2005-video__title{color:#f4f4f5;font-size:1rem;font-weight:800;line-height:1.35;margin:0 0 .6rem}' +
      '.na2005-video .na-yt{border:1px solid #3f3f46;border-radius:.75rem}' +
      '.na2005-badge{position:absolute;left:.45rem;top:.45rem;z-index:3;display:inline-flex;align-items:center;gap:.25rem;' +
        'padding:.18rem .45rem;border-radius:999px;background:rgba(185,28,28,.92);color:#fff;font-size:.62rem;font-weight:800;' +
        'letter-spacing:.04em;line-height:1.2;pointer-events:none;box-shadow:0 2px 8px rgba(0,0,0,.5)}';
    var s = document.createElement('style');
    s.id = 'na2005-video-css';
    s.textContent = css;
    document.head.appendChild(s);
  }

  function isEn() { return (document.documentElement.lang || '').toLowerCase().indexOf('en') === 0; }
  function bi(zh, en) {
    var e = isEn();
    return '<span class="lang-zh' + (e ? ' lang-hidden' : '') + '">' + zh + '</span>' +
           '<span class="lang-en' + (e ? '' : ' lang-hidden') + '">' + en + '</span>';
  }
  function esc(t) {
    return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  /* Lightweight click-to-load facade: thumbnail + play button; click swaps in the iframe. */
  function facade(v) {
    var wrap = document.createElement('div');
    wrap.className = 'na-yt';
    wrap.innerHTML =
      '<button type="button" class="na-yt__btn" aria-label="' + esc('Play: ' + v.en) + '">' +
        '<img src="https://i.ytimg.com/vi/' + v.id + '/hqdefault.jpg" alt="" loading="lazy" decoding="async" width="480" height="360" />' +
        '<span class="na-yt__play" aria-hidden="true"></span>' +
      '</button>';
    wrap.querySelector('button').addEventListener('click', function () {
      var f = document.createElement('iframe');
      f.src = 'https://www.youtube-nocookie.com/embed/' + v.id + '?autoplay=1&rel=0&playsinline=1';
      f.title = v.en;
      f.setAttribute('loading', 'lazy');
      f.setAttribute('allow', 'accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share');
      f.setAttribute('referrerpolicy', 'strict-origin-when-cross-origin');
      f.setAttribute('allowfullscreen', '');
      wrap.innerHTML = '';
      wrap.appendChild(f);
    });
    return wrap;
  }

  function forBand(slug) { return VIDEOS.filter(function (v) { return v.band === slug; }); }

  function render() {
    if (!VIDEOS.length) return;
    injectCss();

    /* 1) Band page: <div class="band-detail__video" data-na2005-band="slug" hidden> … <div data-na2005-videos></div> */
    function fill(sec, list) {
      var host = sec.querySelector('[data-na2005-videos]');
      if (!list.length || !host) return;
      list.forEach(function (v) {
        var item = document.createElement('div');
        item.className = 'na2005-video';
        item.innerHTML = '<h3 class="na2005-video__title">' + bi(esc(v.zh), esc(v.en)) + '</h3>';
        item.appendChild(facade(v));
        host.appendChild(item);
      });
      sec.hidden = false;
    }
    Array.prototype.forEach.call(document.querySelectorAll('[data-na2005-band]'), function (sec) {
      fill(sec, forBand(sec.getAttribute('data-na2005-band')));
    });
    /* 1b) /live/2005: <section data-na2005-extra hidden> — videos of bands without a 2005 page */
    Array.prototype.forEach.call(document.querySelectorAll('[data-na2005-extra]'), function (sec) {
      fill(sec, VIDEOS.filter(function (v) { return !v.hero && !hasPage(v.band); }));
    });
    /* 1c) /live/2005 top: <section data-na2005-hero hidden> — event-wide (hero) videos */
    Array.prototype.forEach.call(document.querySelectorAll('[data-na2005-hero]'), function (sec) {
      fill(sec, VIDEOS.filter(function (v) { return v.hero; }));
    });

    /* 2) /live/2005 tiles: <a class="band-card" data-na2005-tile="slug"> → small ▶ badge on the photo */
    Array.prototype.forEach.call(document.querySelectorAll('[data-na2005-tile]'), function (tile) {
      if (!forBand(tile.getAttribute('data-na2005-tile')).length) return;
      var logo = tile.querySelector('.band-card__logo') || tile;
      var b = document.createElement('span');
      b.className = 'na2005-badge';
      b.innerHTML = '▶ ' + bi('影片', 'Video');
      logo.appendChild(b);
      tile.setAttribute('aria-label', (tile.textContent || '').trim() + ' — video');
    });

    /* 3) /videos: <div data-na2005-video-grid> inside a hidden group → one .video-card per video */
    Array.prototype.forEach.call(document.querySelectorAll('[data-na2005-video-grid]'), function (grid) {
      VIDEOS.forEach(function (v) {
        var card = document.createElement('article');
        card.className = 'video-card';
        var emb = document.createElement('div');
        emb.className = 'video-card__embed';
        emb.appendChild(facade(v));
        card.appendChild(emb);
        var body = document.createElement('div');
        body.className = 'video-card__body';
        body.innerHTML =
          '<h3 class="video-card__title">' + bi(esc(v.zh), esc(v.en)) + '</h3>' +
          '<p class="video-card__meta">' +
            (v.hero
              ? '<a class="text-red-400 hover:text-red-300" href="/live/2005">' + bi('現場相冊', 'Live photos') + '</a>'
              : hasPage(v.band)
              ? '<a class="text-red-400 hover:text-red-300" href="/live/2005/' + esc(v.band) + '">' + bi('樂隊頁', 'Band page') + '</a>'
              : '<a class="text-red-400 hover:text-red-300" href="/live/2005#videos">Noah\'s Ark 2005</a>') +
            ' · <a class="text-red-400 hover:text-red-300" href="https://www.youtube.com/watch?v=' + esc(v.id) + '" target="_blank" rel="noopener">YouTube</a>' +
          '</p>';
        card.appendChild(body);
        grid.appendChild(card);
      });
      var group = grid.closest('[data-na2005-group]');
      if (group) group.hidden = false;
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', render);
  else render();
})();
