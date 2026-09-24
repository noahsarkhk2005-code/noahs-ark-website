/* Noah's Ark member centre
 *
 * ===== IT狗 API contract =====
 * POST MEMBER_ENDPOINT
 *   Content-Type: text/plain;charset=utf-8  (only header)
 *   body: JSON; always include device_id
 *
 * {action:'register', type:'existing'|'new', name, phone, email, consent, device_id
 *   // phone optional free-form ("" if blank); no 8-digit / +852 normalisation
 *   [, member_no]}  // member_no ONLY for existing; omit entirely for new
 *   → {status:'pending'|'invalid'|'locked'}
 *
 * {action:'login', member_no, password, device_id}
 *   → {status:'ok', token, must_change_password, profile:{member_no,name,phone,email}}
 *     | {status:'invalid'|'locked'|'pending'}
 *   pending = not approved OR password not yet set; token ~12h
 *
 * {action:'setup_password', setup_token, new_password, device_id}
 *   → {status:'ok'|'invalid'|'expired'}
 *   Link: /member?setup=<setup_token> (~72h)
 *
 * {action:'change_password', token, old_password, new_password, device_id}
 *   → {status:'ok'|'invalid'|'expired'}
 *
 * {action:'reset_request', member_no, device_id}
 *   → always {status:'pending'}  // no enumeration
 *
 * {action:'reset_password', reset_token, new_password, device_id}
 *   → {status:'ok'|'invalid'|'expired'}
 *   Link: /member?reset=<reset_token> (~1h)
 *
 * {action:'profile', token, device_id}
 *   → {status:'ok', profile} | {status:'expired'}
 *
 * {action:'logout', token, device_id} → {status:'ok'}
 *
 * {action:'google_auth', id_token, device_id}
 *   → {status:'ok', token, must_change_password, profile}
 *     | {status:'need_register', name, email}
 *     | {status:'pending'|'invalid'|'locked'}
 *
 * {action:'register', ..., id_token?}  // Google path: omit name/email; include id_token
 *   → {status:'pending'|'invalid'|'locked'}
 *
 * sessionStorage `na_member`  = {token, profile}
 * sessionStorage `na_prefill` = {member_no?, name, phone, email}  // after register
 * localStorage   `na_device_id`, `na_last_member_no` (number only)
 * GOOGLE_CLIENT_ID null = no Google button; localhost ?gclient= override
 * ============================================= */
(function () {
  'use strict';

  var MEMBER_ENDPOINT = null;
  var GOOGLE_CLIENT_ID = null;
  var MEMBER_NO_RE = /^NA\d{6}$/;
  var DEVICE_KEY = 'na_device_id';
  var SESSION_KEY = 'na_member';
  var PREFILL_KEY = 'na_prefill';
  var LAST_NO_KEY = 'na_last_member_no';
  var MIN_PW = 8;

  var params = new URLSearchParams(location.search);
  var session = null;

  var viewGuest  = document.getElementById('view-guest');
  var viewLogged = document.getElementById('view-logged');
  var viewSetup  = document.getElementById('view-setup');
  var viewReset  = document.getElementById('view-reset');
  var viewChange = document.getElementById('view-change');
  var resultEl   = document.getElementById('member-result');
  var tabsBar    = document.getElementById('member-tabs');
  var postRegEl  = document.getElementById('post-reg-cta');

  var tabLogin    = document.getElementById('tab-login');
  var tabExisting = document.getElementById('tab-existing');
  var tabNew      = document.getElementById('tab-new');

  var panelLogin    = document.getElementById('panel-login');
  var panelExisting = document.getElementById('panel-register-existing');
  var panelNew      = document.getElementById('panel-register-new');
  var panelForgot   = document.getElementById('panel-forgot');
  var panelGoogleReg = document.getElementById('panel-google-register');

  var pendingGoogleIdToken = null;
  var pendingGoogleProfile = null; /* {name, email} from need_register */
  var gsiReady = false;

  var MSG = {
    pending_reg: {
      zh: '已收到，我哋核對後會電郵通知你',
      en: "Received — we'll email you once it's verified",
      cls: 'is-pending'
    },
    pending_login: {
      zh: '你嘅登記仍在核對中，或者未設定密碼。批核後請用電郵入面嘅連結設定密碼。',
      en: 'Your registration is still being verified, or you have not set a password yet. After approval, use the link in your email to set a password.',
      cls: 'is-pending'
    },
    invalid_reg: {
      zh: '格式唔啱，請再填',
      en: "Something's not in the right format, please check and try again",
      cls: 'is-invalid'
    },
    invalid_login: {
      zh: '號碼或密碼唔啱',
      en: 'Member number or password is incorrect',
      cls: 'is-invalid'
    },
    invalid_pw: {
      zh: '密碼唔啱，或者新密碼唔合規格',
      en: 'Password incorrect, or new password does not meet requirements',
      cls: 'is-invalid'
    },
    locked: {
      zh: '試太多次，請稍後再試',
      en: 'Too many attempts, please try again later',
      cls: 'is-locked'
    },
    expired_session: {
      zh: '登入已過期，請重新登入',
      en: 'Session expired — please log in again',
      cls: 'is-locked'
    },
    expired_setup: {
      zh: '連結已過期，請電郵或 IG 話我哋知',
      en: 'This link has expired. Please email us or message us on IG.',
      cls: 'is-locked'
    },
    expired_reset: {
      zh: '連結已過期，請再撳「忘記密碼」',
      en: 'This link has expired. Please tap “Forgot password?” again.',
      cls: 'is-locked'
    },
    setup_ok: {
      zh: '密碼設定好，請用會員號碼同密碼登入',
      en: 'Password set — please log in with your member number and password',
      cls: 'is-pending'
    },
    reset_ok: {
      zh: '密碼已重設，請用新密碼登入',
      en: 'Password reset — please log in with your new password',
      cls: 'is-pending'
    },
    change_ok: {
      zh: '密碼已更改',
      en: 'Password updated',
      cls: 'is-pending'
    },
    forgot_pending: {
      zh: '如果資料吻合，我哋會寄重設連結去你登記嘅電郵',
      en: "If the details match, we'll email a reset link to your registered address",
      cls: 'is-pending'
    },
    soon: {
      zh: '會員功能即將開放，請稍後再試。',
      en: 'Member features open soon — please check back later.',
      cls: 'is-soon'
    },
    error: {
      zh: '暫時未能提交，請稍後再試',
      en: 'Unable to submit right now — please try again later',
      cls: 'is-locked'
    },
    pw_mismatch: {
      zh: '兩次輸入嘅新密碼唔一致',
      en: 'New password entries do not match',
      cls: 'is-invalid'
    },
    pw_short: {
      zh: '新密碼最少 8 個字元',
      en: 'New password must be at least 8 characters',
      cls: 'is-invalid'
    }
  };

  function syncLang(root) {
    if (!root) return;
    var lang = document.documentElement.lang === 'en' ? 'en' : 'zh';
    root.querySelectorAll('.lang-zh').forEach(function (el) {
      el.classList.toggle('lang-hidden', lang === 'en');
    });
    root.querySelectorAll('.lang-en').forEach(function (el) {
      el.classList.toggle('lang-hidden', lang === 'zh');
    });
  }

  function showState(key, extraHtml) {
    var m = MSG[key] || MSG.error;
    resultEl.className = 'member-result is-on ' + m.cls;
    resultEl.innerHTML =
      '<span class="lang-zh">' + m.zh + '</span>' +
      '<span class="lang-en lang-hidden">' + m.en + '</span>' +
      (extraHtml || '');
    syncLang(resultEl);
  }

  function clearResult() {
    resultEl.className = 'member-result';
    resultEl.innerHTML = '';
    if (postRegEl) postRegEl.hidden = true;
  }

  function applyPlaceholders() {
    var lang = document.documentElement.lang === 'en' ? 'en' : 'zh';
    document.querySelectorAll('[data-ph-zh]').forEach(function (el) {
      var ph = el.getAttribute(lang === 'en' ? 'data-ph-en' : 'data-ph-zh');
      if (ph != null) el.setAttribute('placeholder', ph);
    });
  }

  function getDeviceId() {
    try {
      var id = localStorage.getItem(DEVICE_KEY);
      if (id && /^[0-9a-fA-F-]{16,}$/.test(id)) return id;
      id = (window.crypto && crypto.randomUUID)
        ? crypto.randomUUID()
        : ('na-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 12));
      localStorage.setItem(DEVICE_KEY, id);
      return id;
    } catch (e) {
      return 'na-anon-' + Date.now().toString(36);
    }
  }

  function normalizeMemberNo(v) {
    return String(v || '').replace(/\s+/g, '').toUpperCase();
  }

  /* Phone is optional + free-form (HK / overseas / blank). Trim only. */
  function trimPhone(v) {
    return String(v || '').replace(/^\s+|\s+$/g, '');
  }

  function readSession() {
    try {
      var raw = sessionStorage.getItem(SESSION_KEY);
      if (!raw) return null;
      var o = JSON.parse(raw);
      if (!o || !o.token || !o.profile) return null;
      return o;
    } catch (e) { return null; }
  }

  function writeSession(token, profile) {
    session = { token: token, profile: profile };
    try { sessionStorage.setItem(SESSION_KEY, JSON.stringify(session)); } catch (e) {}
  }

  function clearSession() {
    session = null;
    try { sessionStorage.removeItem(SESSION_KEY); } catch (e) {}
  }

  function savePrefill(obj) {
    try { sessionStorage.setItem(PREFILL_KEY, JSON.stringify(obj)); } catch (e) {}
  }

  function rememberMemberNo(no) {
    try {
      if (no && MEMBER_NO_RE.test(no)) localStorage.setItem(LAST_NO_KEY, no);
    } catch (e) {}
  }

  function loadLastMemberNo() {
    try { return localStorage.getItem(LAST_NO_KEY) || ''; } catch (e) { return ''; }
  }

  function resolveEndpoint() {
    var host = location.hostname;
    if (host === '127.0.0.1' || host === 'localhost') {
      var ep = params.get('endpoint');
      if (ep) return ep;
    }
    return MEMBER_ENDPOINT;
  }

  function resolveGoogleClientId() {
    var host = location.hostname;
    if (host === '127.0.0.1' || host === 'localhost') {
      var g = params.get('gclient');
      if (g) return g;
    }
    return GOOGLE_CLIENT_ID;
  }

  function isInAppBrowser() {
    var ua = navigator.userAgent || '';
    return /Instagram|FBAN|FBAV|Messenger|Threads/i.test(ua);
  }


  function apiPost(payload) {
    var url = resolveEndpoint();
    if (!url) return Promise.resolve({ _noEndpoint: true });
    payload.device_id = getDeviceId();
    return fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
      redirect: 'follow'
    }).then(function (res) {
      return res.json().catch(function () { return {}; });
    });
  }

  function hideAllViews() {
    [viewGuest, viewLogged, viewSetup, viewReset, viewChange].forEach(function (el) {
      if (el) el.hidden = true;
    });
  }

  function showView(el) {
    hideAllViews();
    if (el) el.hidden = false;
  }

  function setGuestTab(tab) {
    if (tabsBar) tabsBar.hidden = (tab === 'forgot' || tab === 'google_reg');
    [tabLogin, tabExisting, tabNew].forEach(function (t) {
      if (!t) return;
      var on = (tab === 'login' && t === tabLogin) ||
               (tab === 'existing' && t === tabExisting) ||
               (tab === 'new' && t === tabNew);
      t.classList.toggle('is-active', on);
      t.setAttribute('aria-selected', on ? 'true' : 'false');
    });
    [panelLogin, panelExisting, panelNew, panelForgot, panelGoogleReg].forEach(function (p) {
      if (p) p.hidden = true;
    });
    if (tab === 'login' && panelLogin) panelLogin.hidden = false;
    if (tab === 'existing' && panelExisting) panelExisting.hidden = false;
    if (tab === 'new' && panelNew) panelNew.hidden = false;
    if (tab === 'forgot' && panelForgot) panelForgot.hidden = false;
    if (tab === 'google_reg' && panelGoogleReg) panelGoogleReg.hidden = false;
    clearResult();
    refreshGoogleSlots();
  }

  function escapeHtml(s) {
    return String(s || '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function showPostRegCta() {
    if (!postRegEl) return;
    postRegEl.hidden = false;
  }

  function paintLoggedIn() {
    if (!session || !session.profile) return;
    var p = session.profile;
    var greet = document.getElementById('logged-greet');
    if (greet) {
      greet.innerHTML =
        '<span class="lang-zh">你好，' + escapeHtml(p.name) + '</span>' +
        '<span class="lang-en lang-hidden">Hello, ' + escapeHtml(p.name) + '</span>' +
        '<div class="member-no-line">' + escapeHtml(p.member_no) + '</div>';
      syncLang(greet);
    }
    if (postRegEl) postRegEl.hidden = true;
    showView(viewLogged);
  }

  function wirePwToggle(btnId, inputId) {
    var btn = document.getElementById(btnId);
    var input = document.getElementById(inputId);
    if (!btn || !input) return;
    btn.addEventListener('click', function () {
      var show = input.type === 'password';
      input.type = show ? 'text' : 'password';
      btn.setAttribute('aria-pressed', show ? 'true' : 'false');
      var a = btn.querySelector('[data-when="show"]');
      var b = btn.querySelector('[data-when="hide"]');
      if (a) a.hidden = !show;
      if (b) b.hidden = show;
    });
  }

  function bindMemberNoInput(id) {
    var el = document.getElementById(id);
    if (!el) return;
    el.addEventListener('input', function () {
      var start = el.selectionStart;
      el.value = normalizeMemberNo(el.value);
      if (typeof start === 'number') {
        try { el.setSelectionRange(start, start); } catch (e) {}
      }
    });
  }

  function bindPhoneInput(id) {
    var el = document.getElementById(id);
    if (!el) return;
    el.addEventListener('blur', function () {
      el.value = trimPhone(el.value);
    });
  }

  function validateNewPwPair(newId, confirmId) {
    var np = document.getElementById(newId).value;
    var cp = document.getElementById(confirmId).value;
    if (np.length < MIN_PW) { showState('pw_short'); return null; }
    if (np !== cp) { showState('pw_mismatch'); return null; }
    return np;
  }

  function handleRegisterSuccess(payload) {
    var pref = { name: payload.name, phone: payload.phone, email: payload.email };
    if (payload.member_no) pref.member_no = payload.member_no;
    savePrefill(pref);
    showState('pending_reg');
    showPostRegCta();
  }

  function onLoginSubmit(e) {
    e.preventDefault();
    var no = normalizeMemberNo(document.getElementById('login-member-no').value);
    var pw = document.getElementById('login-password').value;
    document.getElementById('login-member-no').value = no;
    if (!MEMBER_NO_RE.test(no) || !pw) { showState('invalid_login'); return; }
    if (!resolveEndpoint()) { showState('soon'); return; }
    var btn = document.getElementById('login-submit');
    btn.disabled = true;
    apiPost({ action: 'login', member_no: no, password: pw }).then(function (data) {
      if (data._noEndpoint) { showState('soon'); return; }
      if (data.status === 'ok') {
        writeSession(data.token, data.profile);
        rememberMemberNo((data.profile && data.profile.member_no) || no);
        if (data.profile) {
          savePrefill({
            member_no: data.profile.member_no,
            name: data.profile.name,
            phone: data.profile.phone,
            email: data.profile.email
          });
        }
        if (data.must_change_password) {
          showView(viewChange);
          clearResult();
        } else {
          paintLoggedIn();
          clearResult();
        }
      } else if (data.status === 'pending') showState('pending_login');
      else if (data.status === 'locked') showState('locked');
      else if (data.status === 'invalid') showState('invalid_login');
      else showState('error');
    }).catch(function () { showState('error'); })
      .finally(function () { btn.disabled = false; });
  }

  function buildRegisterPayload(type) {
    var prefix = type === 'new' ? 'new' : 'ex';
    var name = document.getElementById(prefix + '-name').value.trim();
    var phone = trimPhone(document.getElementById(prefix + '-phone').value);
    document.getElementById(prefix + '-phone').value = phone;
    var email = document.getElementById(prefix + '-email').value.trim();
    var consent = document.getElementById(prefix + '-consent').checked;
    if (!name || !email || !consent) {
      showState('invalid_reg');
      return null;
    }
    var payload = {
      action: 'register',
      type: type,
      name: name,
      phone: phone,
      email: email,
      consent: true
    };
    if (type === 'existing') {
      var no = normalizeMemberNo(document.getElementById('ex-member-no').value);
      document.getElementById('ex-member-no').value = no;
      if (!MEMBER_NO_RE.test(no)) { showState('invalid_reg'); return null; }
      payload.member_no = no;
    }
    return payload;
  }

  function onRegisterSubmit(type, e) {
    e.preventDefault();
    var payload = buildRegisterPayload(type);
    if (!payload) return;
    if (!resolveEndpoint()) {
      handleRegisterSuccess(payload);
      return;
    }
    var btn = e.target.querySelector('[type="submit"]');
    if (btn) btn.disabled = true;
    apiPost(payload).then(function (data) {
      if (data._noEndpoint) { handleRegisterSuccess(payload); return; }
      if (data.status === 'pending') handleRegisterSuccess(payload);
      else if (data.status === 'invalid') showState('invalid_reg');
      else if (data.status === 'locked') showState('locked');
      else showState('error');
    }).catch(function () { showState('error'); })
      .finally(function () { if (btn) btn.disabled = false; });
  }

  function onForgotSubmit(e) {
    e.preventDefault();
    var no = normalizeMemberNo(document.getElementById('forgot-member-no').value);
    document.getElementById('forgot-member-no').value = no;
    if (!resolveEndpoint()) { showState('soon'); return; }
    var btn = document.getElementById('forgot-submit');
    btn.disabled = true;
    apiPost({ action: 'reset_request', member_no: no || 'UNKNOWN' })
      .then(function () { showState('forgot_pending'); })
      .catch(function () { showState('forgot_pending'); })
      .finally(function () { btn.disabled = false; });
  }

  function onSetupSubmit(e) {
    e.preventDefault();
    var np = validateNewPwPair('setup-new', 'setup-confirm');
    if (!np) return;
    var token = params.get('setup');
    if (!resolveEndpoint()) { showState('soon'); return; }
    var btn = document.getElementById('setup-submit');
    btn.disabled = true;
    apiPost({ action: 'setup_password', setup_token: token, new_password: np }).then(function (data) {
      if (data.status === 'ok') {
        try {
          var u = new URL(location.href);
          u.searchParams.delete('setup');
          var q = u.searchParams.toString();
          history.replaceState(null, '', u.pathname + (q ? '?' + q : ''));
        } catch (err) {}
        params = new URLSearchParams(location.search);
        showView(viewGuest);
        setGuestTab('login');
        showState('setup_ok');
      } else if (data.status === 'expired') {
        showState('expired_setup',
          '<div class="member-result-extra">' +
          '<a href="mailto:noahsarkhk2005@gmail.com">noahsarkhk2005@gmail.com</a>' +
          ' · <a href="https://www.instagram.com/noahs_ark_hk/" target="_blank" rel="noopener">IG @noahs_ark_hk</a>' +
          '</div>');
      } else if (data.status === 'invalid') showState('invalid_pw');
      else showState('error');
    }).catch(function () { showState('error'); })
      .finally(function () { btn.disabled = false; });
  }

  function onResetSubmit(e) {
    e.preventDefault();
    var np = validateNewPwPair('reset-new', 'reset-confirm');
    if (!np) return;
    var token = params.get('reset');
    if (!resolveEndpoint()) { showState('soon'); return; }
    var btn = document.getElementById('reset-submit');
    btn.disabled = true;
    apiPost({ action: 'reset_password', reset_token: token, new_password: np }).then(function (data) {
      if (data.status === 'ok') {
        try {
          var u = new URL(location.href);
          u.searchParams.delete('reset');
          var q = u.searchParams.toString();
          history.replaceState(null, '', u.pathname + (q ? '?' + q : ''));
        } catch (err) {}
        params = new URLSearchParams(location.search);
        showView(viewGuest);
        setGuestTab('login');
        showState('reset_ok');
      } else if (data.status === 'expired') showState('expired_reset');
      else if (data.status === 'invalid') showState('invalid_pw');
      else showState('error');
    }).catch(function () { showState('error'); })
      .finally(function () { btn.disabled = false; });
  }

  function onChangeSubmit(e) {
    e.preventDefault();
    if (!session || !session.token) {
      showView(viewGuest); setGuestTab('login'); showState('expired_session'); return;
    }
    var oldPw = document.getElementById('change-old').value;
    var np = validateNewPwPair('change-new', 'change-confirm');
    if (!np) return;
    if (!oldPw) { showState('invalid_pw'); return; }
    if (!resolveEndpoint()) { showState('soon'); return; }
    var btn = document.getElementById('change-submit');
    btn.disabled = true;
    apiPost({
      action: 'change_password',
      token: session.token,
      old_password: oldPw,
      new_password: np
    }).then(function (data) {
      if (data.status === 'ok') {
        paintLoggedIn();
        showState('change_ok');
      } else if (data.status === 'expired') {
        clearSession();
        showView(viewGuest); setGuestTab('login');
        showState('expired_session');
      } else if (data.status === 'invalid') showState('invalid_pw');
      else showState('error');
    }).catch(function () { showState('error'); })
      .finally(function () { btn.disabled = false; });
  }

  function doLogout() {
    var tok = session && session.token;
    clearSession();
    if (tok && resolveEndpoint()) {
      apiPost({ action: 'logout', token: tok }).catch(function () {});
    }
    showView(viewGuest);
    setGuestTab('login');
    clearResult();
  }

  function validateProfileOnLoad() {
    session = readSession();
    if (!session) return Promise.resolve(false);
    if (!resolveEndpoint()) return Promise.resolve(true);
    return apiPost({ action: 'profile', token: session.token }).then(function (data) {
      if (data.status === 'ok' && data.profile) {
        writeSession(session.token, data.profile);
        return true;
      }
      clearSession();
      return false;
    }).catch(function () {
      clearSession();
      return false;
    });
  }

  function hideAllGoogleButtons() {
    document.querySelectorAll('[data-google-slot]').forEach(function (slot) {
      slot.hidden = true;
      var host = slot.querySelector('[data-google-btn]');
      if (host) host.innerHTML = '';
      var div = slot.querySelector('[data-google-divider]');
      if (div) div.hidden = true;
      var hint = slot.querySelector('[data-inapp-hint]');
      if (hint) hint.hidden = true;
    });
  }

  function refreshGoogleSlots() {
    var clientId = resolveGoogleClientId();
    var slots = document.querySelectorAll('[data-google-slot]');
    if (!clientId) {
      hideAllGoogleButtons();
      return;
    }
    if (isInAppBrowser()) {
      slots.forEach(function (slot) {
        /* Only show hint on visible guest panels (parent not hidden) */
        var panel = slot.parentElement;
        if (panel && panel.hidden) {
          slot.hidden = true;
          return;
        }
        slot.hidden = false;
        var host = slot.querySelector('[data-google-btn]');
        if (host) host.innerHTML = '';
        var div = slot.querySelector('[data-google-divider]');
        if (div) div.hidden = true;
        var hint = slot.querySelector('[data-inapp-hint]');
        if (hint) {
          hint.hidden = false;
          syncLang(hint);
        }
      });
      return;
    }
    slots.forEach(function (slot) {
      var panel = slot.parentElement;
      if (panel && panel.hidden) {
        slot.hidden = true;
        return;
      }
      slot.hidden = false;
      var hint = slot.querySelector('[data-inapp-hint]');
      if (hint) hint.hidden = true;
      var div = slot.querySelector('[data-google-divider]');
      if (div) div.hidden = false;
      syncLang(slot);
      var host = slot.querySelector('[data-google-btn]');
      if (!host) return;
      host.innerHTML = '';
      if (gsiReady && window.google && google.accounts && google.accounts.id) {
        try {
          google.accounts.id.renderButton(host, {
            type: 'standard',
            theme: 'outline',
            size: 'large',
            text: 'continue_with',
            shape: 'rectangular',
            logo_alignment: 'left',
            width: Math.min(320, Math.max(240, host.clientWidth || 280)),
            locale: (document.documentElement.lang === 'en') ? 'en' : 'zh-HK'
          });
        } catch (e) {
          slot.hidden = true;
        }
      } else if (location.hostname === '127.0.0.1' || location.hostname === 'localhost') {
        /* Localhost placeholder when GSI script not loaded — for layout/QA only */
        var btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'na-google-placeholder';
        btn.setAttribute('aria-label', 'Continue with Google');
        btn.innerHTML = '<span class="na-g-icon" aria-hidden="true">G</span><span class="lang-zh">透過 Google 繼續</span><span class="lang-en lang-hidden">Continue with Google</span>';
        btn.addEventListener('click', function () {
          handleGoogleCredential({ credential: 'mock:new-sub:新谷:new@g.co' });
        });
        host.appendChild(btn);
        syncLang(host);
      }
    });
  }

  function handleGoogleCredential(response) {
    var idToken = response && response.credential;
    if (!idToken) { showState('error'); return; }
    if (!resolveEndpoint()) { showState('soon'); return; }
    apiPost({ action: 'google_auth', id_token: idToken }).then(function (data) {
      if (data._noEndpoint) { showState('soon'); return; }
      if (data.status === 'ok') {
        writeSession(data.token, data.profile);
        if (data.profile && data.profile.member_no) rememberMemberNo(data.profile.member_no);
        if (data.profile) {
          savePrefill({
            member_no: data.profile.member_no,
            name: data.profile.name,
            phone: data.profile.phone,
            email: data.profile.email
          });
        }
        pendingGoogleIdToken = null;
        pendingGoogleProfile = null;
        if (data.must_change_password) {
          showView(viewChange);
          clearResult();
        } else {
          paintLoggedIn();
          clearResult();
        }
      } else if (data.status === 'need_register') {
        pendingGoogleIdToken = idToken;
        pendingGoogleProfile = { name: data.name || '', email: data.email || '' };
        showGoogleRegisterForm();
      } else if (data.status === 'pending') {
        showState('pending_login');
      } else if (data.status === 'locked') {
        showState('locked');
      } else if (data.status === 'invalid') {
        showState('invalid_login');
      } else {
        showState('error');
      }
    }).catch(function () { showState('error'); });
  }

  function showGoogleRegisterForm() {
    showView(viewGuest);
    setGuestTab('google_reg');
    document.getElementById('g-reg-name').textContent = (pendingGoogleProfile && pendingGoogleProfile.name) || '';
    document.getElementById('g-reg-email').textContent = (pendingGoogleProfile && pendingGoogleProfile.email) || '';
    document.getElementById('g-member-no').value = '';
    document.getElementById('g-phone').value = '';
    document.getElementById('g-consent').checked = false;
    setGoogleRegType('existing');
    clearResult();
  }

  function setGoogleRegType(type) {
    var isNew = type === 'new';
    var exLabel = document.getElementById('g-choice-existing');
    var newLabel = document.getElementById('g-choice-new');
    if (exLabel) exLabel.classList.toggle('is-on', !isNew);
    if (newLabel) newLabel.classList.toggle('is-on', isNew);
    var exRadio = document.querySelector('input[name="g-type"][value="existing"]');
    var newRadio = document.querySelector('input[name="g-type"][value="new"]');
    if (exRadio) exRadio.checked = !isNew;
    if (newRadio) newRadio.checked = isNew;
    var field = document.getElementById('g-field-member-no');
    if (field) field.hidden = isNew;
    var input = document.getElementById('g-member-no');
    if (input) input.required = !isNew;
  }

  function onGoogleRegisterSubmit(e) {
    e.preventDefault();
    if (!pendingGoogleIdToken) { showState('error'); return; }
    var typeEl = document.querySelector('input[name="g-type"]:checked');
    var type = typeEl ? typeEl.value : 'existing';
    var phone = trimPhone(document.getElementById('g-phone').value);
    document.getElementById('g-phone').value = phone;
    var consent = document.getElementById('g-consent').checked;
    if (!consent) { showState('invalid_reg'); return; }
    var payload = {
      action: 'register',
      type: type,
      phone: phone,
      consent: true,
      id_token: pendingGoogleIdToken
    };
    if (type === 'existing') {
      var no = normalizeMemberNo(document.getElementById('g-member-no').value);
      document.getElementById('g-member-no').value = no;
      if (!MEMBER_NO_RE.test(no)) { showState('invalid_reg'); return; }
      payload.member_no = no;
    }
    if (!resolveEndpoint()) {
      /* preview: treat as pending with prefill from Google profile */
      var pref = {
        name: (pendingGoogleProfile && pendingGoogleProfile.name) || '',
        email: (pendingGoogleProfile && pendingGoogleProfile.email) || '',
        phone: phone
      };
      if (payload.member_no) pref.member_no = payload.member_no;
      savePrefill(pref);
      showState('pending_reg');
      showPostRegCta();
      return;
    }
    var btn = document.getElementById('g-reg-submit');
    if (btn) btn.disabled = true;
    apiPost(payload).then(function (data) {
      if (data.status === 'pending' || data._noEndpoint) {
        var pref2 = {
          name: (pendingGoogleProfile && pendingGoogleProfile.name) || '',
          email: (pendingGoogleProfile && pendingGoogleProfile.email) || '',
          phone: phone
        };
        if (payload.member_no) pref2.member_no = payload.member_no;
        savePrefill(pref2);
        showState('pending_reg');
        showPostRegCta();
      } else if (data.status === 'invalid') showState('invalid_reg');
      else if (data.status === 'locked') showState('locked');
      else showState('error');
    }).catch(function () { showState('error'); })
      .finally(function () { if (btn) btn.disabled = false; });
  }

  function loadGoogleIdentity() {
    var clientId = resolveGoogleClientId();
    if (!clientId) {
      hideAllGoogleButtons();
      return;
    }
    if (isInAppBrowser()) {
      refreshGoogleSlots();
      return;
    }
    function initGsi() {
      try {
        google.accounts.id.initialize({
          client_id: clientId,
          callback: handleGoogleCredential,
          auto_select: false,
          cancel_on_tap_outside: true
        });
        gsiReady = true;
        refreshGoogleSlots();
      } catch (e) {
        gsiReady = false;
        hideAllGoogleButtons();
      }
    }
    if (window.google && google.accounts && google.accounts.id) {
      initGsi();
      return;
    }
    var existing = document.querySelector('script[data-na-gsi]');
    if (existing) return;
    var s = document.createElement('script');
    s.src = 'https://accounts.google.com/gsi/client';
    s.async = true;
    s.defer = true;
    s.setAttribute('data-na-gsi', '1');
    s.onload = initGsi;
    s.onerror = function () {
      gsiReady = false;
      hideAllGoogleButtons();
    };
    document.head.appendChild(s);
  }

  function maybeMockGoogle() {
    var host = location.hostname;
    if (host !== '127.0.0.1' && host !== 'localhost') return;
    var mock = params.get('mockgoogle');
    if (!mock) return;
    /* mock format: linked | new | pending | locked | invalid | mock:<sub>:<name>:<email> */
    var token = mock;
    if (mock === 'linked') token = 'mock:linked-sub:阿谷:g@linked.co';
    else if (mock === 'new') token = 'mock:new-sub:新谷:new@g.co';
    else if (mock === 'pending') token = 'mock:pending-sub:待核:p@g.co';
    else if (mock === 'locked') token = 'mock:locked-sub:鎖住:l@g.co';
    else if (mock === 'invalid') token = 'mock:bad';
    /* Defer until endpoint override is usable */
    setTimeout(function () {
      handleGoogleCredential({ credential: token });
    }, 200);
  }


  function bind() {
    document.getElementById('form-login').addEventListener('submit', onLoginSubmit);
    document.getElementById('form-register-existing').addEventListener('submit', function (e) {
      onRegisterSubmit('existing', e);
    });
    document.getElementById('form-register-new').addEventListener('submit', function (e) {
      onRegisterSubmit('new', e);
    });
    document.getElementById('form-forgot').addEventListener('submit', onForgotSubmit);
    document.getElementById('form-setup').addEventListener('submit', onSetupSubmit);
    document.getElementById('form-reset').addEventListener('submit', onResetSubmit);
    document.getElementById('form-change').addEventListener('submit', onChangeSubmit);

    tabLogin.addEventListener('click', function () { setGuestTab('login'); });
    tabExisting.addEventListener('click', function () { setGuestTab('existing'); });
    tabNew.addEventListener('click', function () { setGuestTab('new'); });

    document.getElementById('link-forgot').addEventListener('click', function (e) {
      e.preventDefault(); setGuestTab('forgot');
    });
    document.getElementById('link-back-login').addEventListener('click', function (e) {
      e.preventDefault(); setGuestTab('login');
    });

    function goTickets() { location.href = '/tickets'; }
    document.getElementById('btn-buy').addEventListener('click', goTickets);
    var postBuy = document.getElementById('post-reg-buy');
    if (postBuy) postBuy.addEventListener('click', goTickets);

    document.getElementById('btn-change-pw').addEventListener('click', function () {
      showView(viewChange); clearResult();
    });
    document.getElementById('btn-logout').addEventListener('click', doLogout);
    document.getElementById('change-cancel').addEventListener('click', function () {
      if (session) paintLoggedIn();
      else { showView(viewGuest); setGuestTab('login'); }
    });


    document.getElementById('form-google-register').addEventListener('submit', onGoogleRegisterSubmit);
    document.getElementById('g-reg-cancel').addEventListener('click', function () {
      pendingGoogleIdToken = null;
      pendingGoogleProfile = null;
      setGuestTab('login');
    });
    document.querySelectorAll('input[name="g-type"]').forEach(function (r) {
      r.addEventListener('change', function () { setGoogleRegType(r.value); });
    });
    var gEx = document.getElementById('g-choice-existing');
    var gNew = document.getElementById('g-choice-new');
    if (gEx) gEx.addEventListener('click', function (e) {
      e.preventDefault();
      setGoogleRegType('existing');
    });
    if (gNew) gNew.addEventListener('click', function (e) {
      e.preventDefault();
      setGoogleRegType('new');
    });
    document.querySelectorAll('[data-copy-link]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var url = location.href;
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(url).catch(function () {});
        } else {
          var ta = document.createElement('textarea');
          ta.value = url; document.body.appendChild(ta); ta.select();
          try { document.execCommand('copy'); } catch (e) {}
          document.body.removeChild(ta);
        }
      });
    });
    bindMemberNoInput('g-member-no');
    bindPhoneInput('g-phone');
    wirePwToggle('login-pw-toggle', 'login-password');
    ['login-member-no', 'ex-member-no', 'forgot-member-no'].forEach(bindMemberNoInput);
    ['ex-phone', 'new-phone'].forEach(bindPhoneInput);
  }

  function boot() {
    bind();
    applyPlaceholders();
    loadGoogleIdentity();

    var last = loadLastMemberNo();
    var loginNo = document.getElementById('login-member-no');
    if (last && loginNo && MEMBER_NO_RE.test(last)) loginNo.value = last;

    var setupTok = params.get('setup');
    var resetTok = params.get('reset');
    var preview = params.get('state');
    var tabParam = params.get('tab');

    if (preview === 'setup' || preview === 'setup_expired') {
      showView(viewSetup);
      if (preview === 'setup_expired') {
        showState('expired_setup',
          '<div class="member-result-extra">' +
          '<a href="mailto:noahsarkhk2005@gmail.com">noahsarkhk2005@gmail.com</a>' +
          ' · <a href="https://www.instagram.com/noahs_ark_hk/" target="_blank" rel="noopener">IG @noahs_ark_hk</a>' +
          '</div>');
      }
      return;
    }
    if (preview === 'reset' || preview === 'reset_expired') {
      showView(viewReset);
      if (preview === 'reset_expired') showState('expired_reset');
      return;
    }
    if (preview === 'change') {
      session = { token: 'preview', profile: { member_no: 'NA000099', name: '阿乙', phone: '98765432', email: 'b@c.co' } };
      showView(viewChange);
      return;
    }
    if (preview === 'logged') {
      session = { token: 'preview', profile: { member_no: 'NA000099', name: '阿乙', phone: '98765432', email: 'b@c.co' } };
      paintLoggedIn();
      return;
    }
    if (preview === 'locked' || preview === 'invalid' || preview === 'pending' || preview === 'expired' || preview === 'forgot') {
      showView(viewGuest);
      setGuestTab(preview === 'forgot' ? 'forgot' : 'login');
      if (preview === 'locked') showState('locked');
      else if (preview === 'invalid') showState('invalid_login');
      else if (preview === 'pending') showState('pending_login');
      else if (preview === 'expired') showState('expired_session');
      else if (preview === 'forgot') showState('forgot_pending');
      return;
    }

    if (setupTok) { showView(viewSetup); return; }
    if (resetTok) { showView(viewReset); return; }

    validateProfileOnLoad().then(function (ok) {
      if (ok && session) { paintLoggedIn(); return; }
      showView(viewGuest);
      if (tabParam === 'new') setGuestTab('new');
      else if (tabParam === 'existing' || tabParam === 'register') setGuestTab('existing');
      else setGuestTab('login');
      maybeMockGoogle();
    });
  }

  window.NOAHS_MEMBER = {
    MEMBER_ENDPOINT: MEMBER_ENDPOINT,
    GOOGLE_CLIENT_ID: GOOGLE_CLIENT_ID,
    MEMBER_NO_RE: MEMBER_NO_RE,
    SESSION_KEY: SESSION_KEY,
    PREFILL_KEY: PREFILL_KEY,
    getDeviceId: getDeviceId,
    readSession: readSession,
    applyPlaceholders: applyPlaceholders,
    showState: showState,
    setGuestTab: setGuestTab,
    refreshGoogleSlots: refreshGoogleSlots,
    handleGoogleCredential: handleGoogleCredential
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
})();
