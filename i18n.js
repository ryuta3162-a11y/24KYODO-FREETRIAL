/**
 * JOYFIT24 経堂 LP — 軽量 i18n（日本語はHTML原文を保持）
 */
(function (global) {
  const STORAGE_KEY = 'joyfit-kyodo-lang';
  const DEFAULT_LANG = 'ja';
  const LANG_META = {
    ja: { label: '日本語', htmlLang: 'ja', font: "'Noto Sans JP', sans-serif" },
    en: { label: 'English', htmlLang: 'en', font: "'Noto Sans JP', sans-serif" },
    'zh-CN': { label: '简体中文', htmlLang: 'zh-CN', font: "'Noto Sans SC', 'Noto Sans JP', sans-serif" },
    'zh-TW': { label: '繁體中文', htmlLang: 'zh-TW', font: "'Noto Sans TC', 'Noto Sans JP', sans-serif" },
    ko: { label: '한국어', htmlLang: 'ko', font: "'Noto Sans KR', 'Noto Sans JP', sans-serif" },
  };

  let currentLang = DEFAULT_LANG;
  let catalog = {};
  const defaults = new Map();

  function getPath(obj, path) {
    return path.split('.').reduce((acc, key) => (acc && acc[key] != null ? acc[key] : undefined), obj);
  }

  function captureDefaults() {
    document.querySelectorAll('[data-i18n], [data-i18n-html], [data-i18n-placeholder], [data-i18n-title]').forEach((el) => {
      if (el.dataset.i18nPlaceholder != null || el.hasAttribute('data-i18n-placeholder')) {
        defaults.set(el, { type: 'placeholder', value: el.getAttribute('placeholder') || '' });
      } else if (el.dataset.i18nTitle != null || el.hasAttribute('data-i18n-title')) {
        defaults.set(el, { type: 'title', value: el.getAttribute('title') || '' });
      } else if (el.hasAttribute('data-i18n-html')) {
        defaults.set(el, { type: 'html', value: el.innerHTML });
      } else {
        defaults.set(el, { type: 'text', value: el.textContent });
      }
    });

    document.querySelectorAll('option[data-i18n]').forEach((el) => {
      defaults.set(el, { type: 'text', value: el.textContent });
    });
  }

  function restoreDefaults() {
    defaults.forEach((item, el) => {
      if (item.type === 'placeholder') el.setAttribute('placeholder', item.value);
      else if (item.type === 'title') el.setAttribute('title', item.value);
      else if (item.type === 'html') el.innerHTML = item.value;
      else el.textContent = item.value;
    });
    if (defaults.has(document.documentElement)) {
      document.title = defaults.get(document.documentElement).value;
    }
  }

  function applyCatalog() {
    if (currentLang === DEFAULT_LANG) {
      restoreDefaults();
      return;
    }

    document.querySelectorAll('[data-i18n]').forEach((el) => {
      const value = getPath(catalog, el.getAttribute('data-i18n'));
      if (value != null) el.textContent = value;
    });

    document.querySelectorAll('[data-i18n-html]').forEach((el) => {
      const value = getPath(catalog, el.getAttribute('data-i18n-html'));
      if (value != null) el.innerHTML = value;
    });

    document.querySelectorAll('[data-i18n-placeholder]').forEach((el) => {
      const value = getPath(catalog, el.getAttribute('data-i18n-placeholder'));
      if (value != null) el.setAttribute('placeholder', value);
    });

    document.querySelectorAll('option[data-i18n]').forEach((el) => {
      const value = getPath(catalog, el.getAttribute('data-i18n'));
      if (value != null) el.textContent = value;
    });

    const title = getPath(catalog, 'meta.title');
    if (title) document.title = title;
  }

  function applyDocumentLang() {
    const meta = LANG_META[currentLang] || LANG_META.ja;
    document.documentElement.lang = meta.htmlLang;
    document.body.style.fontFamily = meta.font;
    document.body.dataset.lang = currentLang;
  }

  function updateSwitcherUI() {
    const current = document.getElementById('lang-current');
    const menu = document.getElementById('lang-menu');
    if (!current || !menu) return;

    current.textContent = (LANG_META[currentLang] || LANG_META.ja).label;
    menu.querySelectorAll('[data-lang]').forEach((btn) => {
      const active = btn.dataset.lang === currentLang;
      btn.classList.toggle('is-active', active);
      btn.setAttribute('aria-selected', active ? 'true' : 'false');
    });
  }

  async function loadCatalog(lang) {
    if (lang === DEFAULT_LANG) {
      catalog = {};
      return;
    }
    if (global.JOYFIT_LOCALES && global.JOYFIT_LOCALES[lang]) {
      catalog = global.JOYFIT_LOCALES[lang];
      return;
    }
    const res = await fetch(`locales/${lang}.json`);
    if (!res.ok) throw new Error(`Locale load failed: ${lang}`);
    catalog = await res.json();
  }

  async function setLanguage(lang) {
    if (!LANG_META[lang]) return;
    try {
      currentLang = lang;
      localStorage.setItem(STORAGE_KEY, lang);
      await loadCatalog(lang);
      applyDocumentLang();
      applyCatalog();
      updateSwitcherUI();
      global.dispatchEvent(new CustomEvent('joyfit:langchange', { detail: { lang } }));
    } catch (err) {
      console.error('[JoyfitI18n]', err);
    }
  }

  function t(key, fallback) {
    if (currentLang === DEFAULT_LANG) {
      const el = document.querySelector(`[data-i18n="${key}"]`);
      if (el && defaults.has(el)) return defaults.get(el).value;
      const elHtml = document.querySelector(`[data-i18n-html="${key}"]`);
      if (elHtml && defaults.has(elHtml)) return defaults.get(elHtml).innerHTML;
      return fallback != null ? fallback : key;
    }
    const value = getPath(catalog, key);
    return value != null ? value : (fallback != null ? fallback : key);
  }

  function planLabel(value) {
    if (value === '見学') return t('form.visit', '見学');
    if (value === '体験') return t('form.trial', '体験');
    return value;
  }

  function initSwitcher() {
    const wrap = document.getElementById('lang-switcher');
    const btn = document.getElementById('lang-toggle');
    const menu = document.getElementById('lang-menu');
    if (!wrap || !btn || !menu) return;

    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const open = wrap.classList.toggle('is-open');
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
    });

    menu.querySelectorAll('[data-lang]').forEach((item) => {
      item.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        setLanguage(item.dataset.lang);
        wrap.classList.remove('is-open');
        btn.setAttribute('aria-expanded', 'false');
      });
    });

    menu.addEventListener('click', (e) => e.stopPropagation());

    document.addEventListener('click', () => {
      wrap.classList.remove('is-open');
      btn.setAttribute('aria-expanded', 'false');
    });
  }

  async function init() {
    captureDefaults();
    defaults.set(document.documentElement, { type: 'title', value: document.title });

    const saved = localStorage.getItem(STORAGE_KEY);
    const initial = LANG_META[saved] ? saved : DEFAULT_LANG;
    initSwitcher();
    await setLanguage(initial);
  }

  global.JoyfitI18n = {
    init,
    setLanguage,
    t,
    planLabel,
    getLanguage: () => currentLang,
  };
})(window);
