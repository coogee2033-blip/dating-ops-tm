// ==UserScript==
// @name         Dating Ops - MEM
// @namespace    dating-ops-tm
// @version      2.3.0
// @description  MEM (mem44.com) サイト用返信アシストツール - list/personal対応 + CORS回避版
// @author       Dating Ops Team
// @match        https://mem44.com/staff/*
// @match        https://*.mem44.com/staff/*
// @grant        GM_xmlhttpRequest
// @connect      docs.google.com
// @connect      googleusercontent.com
// @connect      google.com
// @run-at       document-idle
// @noframes
// ==/UserScript==

/**
 * ============================================================
 * 実装メモ (v2.3.0)
 * ============================================================
 * 
 * 【ページモード判定】
 *   - /staff/index → mode="list"（一括送信）
 *   - /staff/personalbox → mode="personal"（個別送信）
 *   - その他 → mode="other"
 * 
 * 【行検出】
 *   - mode=list のみ実施
 *   - scope特定: td.chatview → tr.rowitem → checkbox領域 → fallback
 *   - isVisible は行検出では使わない（0化事故防止）
 * 
 * 【返信欄検出】
 *   - 全モードで実施
 *   - 優先順: message1 → msg.wd100 → msg → message → mail_body → content
 *   - 除外: memo/admin/note/futari/two_memo/staff
 * 
 * 【Google Sheet】
 *   - GM_xmlhttpRequest でCORS回避
 *   - エラー時は種別表示 + 直接入力フォールバック
 * ============================================================
 */

(function() {
  'use strict';

  // ============================================================
  // 定数
  // ============================================================
  const SITE_TYPE = 'mem';
  const SITE_NAME = 'MEM';
  const PANEL_ID = 'dating-ops-panel';
  const STORAGE_PREFIX = 'datingOps_mem_';
  const OBSERVER_DEBOUNCE_MS = 300;
  const HEALTH_CHECK_MS = 3000;
  const URL_POLL_MS = 500;

  // 段階的検出のタイミング（ms）
  const STAGED_DETECT_DELAYS = [0, 300, 800, 1500];

  // ============================================================
  // サイト設定
  // ============================================================
  const CONFIG = {
    // ページモード判定
    pageModes: {
      list: [/\/staff\/index/, /\/staff\/.*list/, /\/staff\/.*inbox/],
      personal: [/\/staff\/personalbox/, /\/staff\/personal/, /\/staff\/.*detail/],
    },

    // 返信欄セレクタ（優先順位順）
    textareaSelectors: [
      'textarea[name="message1"]',
      'textarea.msg.wd100',
      'textarea.msg',
      'textarea[name="message"]',
      'textarea[name="body"]',
      'textarea[name="mail_body"]',
      'textarea[name="content"]',
    ],

    // 除外キーワード
    excludeKeywords: ['memo', 'admin', 'note', 'futari', 'two_memo', 'staff'],
    excludeClassKeywords: ['admin', 'memo', 'note'],

    // 行検出セレクタ
    rowSelectors: ['tr.rowitem', 'tr[id^="row"]', 'tr[class*="row"]', 'tbody > tr', 'tr'],

    // 表示件数テキスト
    countTextSelectors: ['.count', '.result', '.paging', '.list_head', '[class*="count"]', '[class*="paging"]'],
    countPatterns: [/(\d+)\s*件/, /表示\s*[:\s]*(\d+)/, /全\s*(\d+)/],

    // トリガセレクタ
    clickTriggerSelectors: [
      'input[type="submit"][value*="更新"]', 'input[value="更新"]', '.reload-btn',
      'input[type="submit"][value*="切替"]', 'input[type="submit"][value*="表示"]',
      '.pagination a', '.paging a', 'a[href*="page="]',
      'input[value*="一括"]', 'input[value*="チェック"]',
    ],
    changeTriggerSelectors: ['select[name*="box"]', 'select[name*="folder"]', 'select[name*="type"]'],

    // 一覧領域候補（MutationObserver用）
    listAreaSelectors: [
      'table.list', 'table[class*="list"]', '#list_area', '.list_area',
      'form table', '#main_content table', '.content table',
    ],

    ui: {
      title: 'MEM 返信アシスト',
      primaryColor: '#48bb78',
      accentColor: '#276749',
    },
  };

  // ============================================================
  // 状態
  // ============================================================
  const state = {
    initialized: false,
    stopped: false,
    panel: null,
    observer: null,
    healthTimer: null,
    urlPollTimer: null,
    stagedDetectIndex: 0,
    stagedDetectTimer: null,
    lastUrl: '',

    // ページモード
    pageMode: 'other', // list, personal, other

    // 返信欄
    textarea: { status: 'pending', element: null, selector: null, where: null, iframeSrc: null },

    // 行検出
    rows: { count: 0, displayCountNum: null, scopeHint: null, usedSelector: null, mismatchWarning: null },

    // 検出メタ
    lastDetectTime: null,
    lastTrigger: null,

    // Sheet
    messages: [],
    sheetUrl: '',
    sheetStatus: 'idle', // idle, loading, success, error
    sheetError: null,

    // パネル
    panelPos: { x: 20, y: 20 },
    dragging: false,
    dragOffset: { x: 0, y: 0 },

    // 監視
    observeTargetHint: null,
  };

  // ============================================================
  // ロガー
  // ============================================================
  const LOG_PREFIX = `[DatingOps:${SITE_TYPE}]`;
  const LOG_STYLES = {
    info: 'background:#48bb78;color:#fff;padding:2px 6px;border-radius:3px;',
    warn: 'background:#f6ad55;color:#000;padding:2px 6px;border-radius:3px;',
    error: 'background:#fc8181;color:#000;padding:2px 6px;border-radius:3px;',
  };
  const logger = {
    info: (msg, data) => console.log(`%c${LOG_PREFIX}%c ${msg}`, LOG_STYLES.info, '', data !== undefined ? data : ''),
    warn: (msg, data) => console.log(`%c${LOG_PREFIX}%c ${msg}`, LOG_STYLES.warn, '', data !== undefined ? data : ''),
    error: (msg, data) => console.log(`%c${LOG_PREFIX}%c ${msg}`, LOG_STYLES.error, '', data !== undefined ? data : ''),
  };

  // ============================================================
  // ユーティリティ
  // ============================================================
  function $(sel, ctx) { try { return (ctx || document).querySelector(sel); } catch { return null; } }
  function $$(sel, ctx) { try { return Array.from((ctx || document).querySelectorAll(sel)); } catch { return []; } }
  function debounce(fn, ms) { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; }
  function storageGet(k, d) { try { const v = localStorage.getItem(STORAGE_PREFIX + k); return v ? JSON.parse(v) : d; } catch { return d; } }
  function storageSet(k, v) { try { localStorage.setItem(STORAGE_PREFIX + k, JSON.stringify(v)); } catch {} }
  function escapeHtml(s) { if (!s) return ''; const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }
  function formatTime(ts) { if (!ts) return '--:--:--'; return new Date(ts).toTimeString().substring(0, 8); }

  // ============================================================
  // ページモード判定
  // ============================================================
  function detectPageMode() {
    const path = location.pathname;
    for (const pattern of CONFIG.pageModes.list) {
      if (pattern.test(path)) return 'list';
    }
    for (const pattern of CONFIG.pageModes.personal) {
      if (pattern.test(path)) return 'personal';
    }
    return 'other';
  }

  // ============================================================
  // 除外判定
  // ============================================================
  function isExcludedTextarea(el) {
    if (!el) return true;
    const name = (el.name || '').toLowerCase();
    const id = (el.id || '').toLowerCase();
    const cls = (typeof el.className === 'string' ? el.className : '').toLowerCase();
    for (const kw of CONFIG.excludeKeywords) {
      if (name.includes(kw) || id.includes(kw)) return true;
    }
    for (const kw of CONFIG.excludeClassKeywords) {
      if (cls.includes(kw)) return true;
    }
    const dt = el.getAttribute('data-type');
    if (dt && (dt.includes('admin') || dt.includes('memo'))) return true;
    return false;
  }

  function isElementVisible(el, win) {
    if (!el) return false;
    try {
      const s = (win || window).getComputedStyle(el);
      return s.display !== 'none' && s.visibility !== 'hidden' && parseFloat(s.opacity) !== 0;
    } catch { return true; }
  }

  // ============================================================
  // scope特定（行検出用）
  // ============================================================
  function findRowScope() {
    // (1) td.chatview
    const cv = $('td.chatview');
    if (cv) {
      const t = cv.closest('table');
      if (t) return { el: $('tbody', t) || t, hint: 'td.chatview → table' };
    }
    // (2) tr.rowitem
    const ri = $('tr.rowitem');
    if (ri) {
      const t = ri.closest('table');
      if (t) return { el: $('tbody', t) || t, hint: 'tr.rowitem → table' };
    }
    // (3) checkbox複数
    const cbs = $$('table input[type="checkbox"]');
    if (cbs.length >= 2) {
      const t = cbs[0].closest('table');
      if (t) return { el: $('tbody', t) || t, hint: `checkbox(${cbs.length}) → table` };
    }
    // (4) listAreaから探す
    for (const sel of CONFIG.listAreaSelectors) {
      const el = $(sel);
      if (el) return { el, hint: `listArea: ${sel}` };
    }
    return { el: document.body, hint: 'fallback: body' };
  }

  function countRowsInScope(scopeEl) {
    for (const sel of CONFIG.rowSelectors) {
      try {
        const rows = $$(sel, scopeEl).filter(tr => !$('th', tr));
        if (rows.length > 0) return { count: rows.length, selector: sel };
      } catch {}
    }
    return { count: 0, selector: null };
  }

  function extractDisplayCount() {
    for (const sel of CONFIG.countTextSelectors) {
      for (const el of $$(sel)) {
        for (const pat of CONFIG.countPatterns) {
          const m = (el.textContent || '').match(pat);
          if (m && m[1]) { const n = parseInt(m[1], 10); if (!isNaN(n) && n >= 0) return n; }
        }
      }
    }
    return null;
  }

  // ============================================================
  // 行検出
  // ============================================================
  function detectRows() {
    if (state.pageMode !== 'list') {
      state.rows = { count: state.pageMode === 'personal' ? 1 : 0, displayCountNum: null, scopeHint: 'N/A (not list mode)', usedSelector: null, mismatchWarning: null };
      return;
    }
    const scope = findRowScope();
    const res = countRowsInScope(scope.el);
    const dispNum = extractDisplayCount();
    let warn = null;
    if (dispNum !== null && res.count === 0 && dispNum > 0) warn = '行セレクタ不一致';
    else if (dispNum !== null && dispNum > 0 && Math.abs(res.count - dispNum) >= 5) warn = 'scope誤認の可能性';
    state.rows = { count: res.count, displayCountNum: dispNum, scopeHint: scope.hint, usedSelector: res.selector, mismatchWarning: warn };
  }

  // ============================================================
  // 返信欄検出
  // ============================================================
  function findReplyTextarea() {
    for (const sel of CONFIG.textareaSelectors) {
      for (const el of $$(sel)) {
        if (!isExcludedTextarea(el) && isElementVisible(el)) {
          return { element: el, selector: sel, where: 'main', iframeSrc: null };
        }
      }
    }
    for (const iframe of $$('iframe')) {
      try {
        const doc = iframe.contentDocument || iframe.contentWindow?.document;
        if (!doc) continue;
        const win = iframe.contentWindow;
        const src = iframe.src || iframe.name || 'iframe';
        for (const sel of CONFIG.textareaSelectors) {
          for (const el of $$(sel, doc)) {
            if (!isExcludedTextarea(el) && isElementVisible(el, win)) {
              return { element: el, selector: sel, where: 'iframe', iframeSrc: src };
            }
          }
        }
        if (doc.body?.contentEditable === 'true') {
          return { element: doc.body, selector: 'body[contenteditable]', where: 'iframe', iframeSrc: src };
        }
      } catch {}
    }
    return null;
  }

  function detectTextarea() {
    const res = findReplyTextarea();
    if (res) {
      state.textarea = { status: 'found', element: res.element, selector: res.selector, where: res.where, iframeSrc: res.iframeSrc };
      return true;
    }
    state.textarea = { status: 'not_found', element: null, selector: null, where: null, iframeSrc: null };
    return false;
  }

  // ============================================================
  // 統合検出
  // ============================================================
  function runDetection(trigger) {
    if (state.stopped) return;
    state.pageMode = detectPageMode();
    detectRows();
    detectTextarea();
    state.lastDetectTime = Date.now();
    state.lastTrigger = trigger || 'manual';
    logger.info(`検出 [${trigger}] mode=${state.pageMode} rows=${state.rows.count} ta=${state.textarea.status}`);
    updateUI();
  }

  function runStagedDetection() {
    if (state.stopped) return;
    clearTimeout(state.stagedDetectTimer);
    state.stagedDetectIndex = 0;

    function next() {
      if (state.stopped || state.stagedDetectIndex >= STAGED_DETECT_DELAYS.length) return;
      const delay = STAGED_DETECT_DELAYS[state.stagedDetectIndex];
      state.stagedDetectTimer = setTimeout(() => {
        runDetection(`init-stage${state.stagedDetectIndex}`);
        state.stagedDetectIndex++;
        // 行が見つかったら終了、見つからなければ次へ
        if (state.pageMode === 'list' && state.rows.count === 0 && state.stagedDetectIndex < STAGED_DETECT_DELAYS.length) {
          next();
        }
      }, delay);
    }
    next();
  }

  // ============================================================
  // テキスト挿入
  // ============================================================
  function insertText(text) {
    const el = state.textarea.element;
    if (!el) { showNotify('返信欄が見つかりません', 'error'); return false; }
    if (isExcludedTextarea(el)) { showNotify('この欄には挿入できません', 'error'); return false; }
    const cur = (el.tagName === 'TEXTAREA' || el.tagName === 'INPUT') ? el.value : (el.textContent || el.innerHTML || '');
    if (cur && cur.length > 0) { showNotify('既にテキストが入力されています', 'warn'); return false; }
    try {
      if (el.tagName === 'TEXTAREA' || el.tagName === 'INPUT') {
        el.value = text;
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
      } else {
        el.innerHTML = escapeHtml(text).replace(/\n/g, '<br>');
        el.dispatchEvent(new Event('input', { bubbles: true }));
      }
      el.focus();
      showNotify('挿入しました', 'success');
      return true;
    } catch (e) { showNotify('挿入失敗', 'error'); return false; }
  }

  // ============================================================
  // Google Sheet（GM_xmlhttpRequest でCORS回避）
  // ============================================================
  function normalizeSheetUrl(url) {
    if (!url) return null;
    const u = url.trim();
    if (u.includes('/export?format=csv') || u.includes('output=csv') || u.includes('/gviz/tq')) return u;
    const m = u.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (m) {
      const gidM = u.match(/gid=(\d+)/);
      return `https://docs.google.com/spreadsheets/d/${m[1]}/export?format=csv&gid=${gidM ? gidM[1] : '0'}`;
    }
    return u;
  }

  function parseCsv(csv) {
    return csv.split('\n').map(line => {
      const t = line.trim();
      if (!t) return null;
      if (t.startsWith('"')) {
        const end = t.indexOf('",');
        if (end > 0) return t.substring(1, end).replace(/""/g, '"');
        if (t.endsWith('"')) return t.slice(1, -1).replace(/""/g, '"');
      }
      const c = t.indexOf(',');
      return c > 0 ? t.substring(0, c) : t;
    }).filter(Boolean).map(s => s.trim()).filter(Boolean);
  }

  function fetchSheet(url) {
    state.sheetStatus = 'loading';
    state.sheetError = null;
    updateUI();

    const normalized = normalizeSheetUrl(url);
    if (!normalized) {
      state.sheetStatus = 'error';
      state.sheetError = 'URLが無効です';
      updateUI();
      return;
    }

    logger.info('Sheet読み込み開始', normalized);

    // GM_xmlhttpRequest があれば使う（CORS回避）
    if (typeof GM_xmlhttpRequest !== 'undefined') {
      GM_xmlhttpRequest({
        method: 'GET',
        url: normalized,
        timeout: 15000,
        onload: function(res) {
          if (res.status >= 200 && res.status < 300) {
            const msgs = parseCsv(res.responseText);
            if (msgs.length === 0) {
              state.sheetStatus = 'error';
              state.sheetError = 'データがありません';
            } else {
              state.messages = msgs;
              state.sheetUrl = url;
              state.sheetStatus = 'success';
              storageSet('sheetUrl', url);
              storageSet('messages', msgs);
              showNotify(`${msgs.length}件読み込み`, 'success');
              logger.info(`Sheet読み込み成功: ${msgs.length}件`);
            }
          } else if (res.status === 403) {
            state.sheetStatus = 'error';
            state.sheetError = '権限エラー(403): シートを「リンクを知っている全員」に公開してください';
          } else if (res.status === 404) {
            state.sheetStatus = 'error';
            state.sheetError = 'シート未発見(404): URLを確認してください';
          } else {
            state.sheetStatus = 'error';
            state.sheetError = `HTTPエラー: ${res.status}`;
          }
          updateUI();
        },
        onerror: function() {
          state.sheetStatus = 'error';
          state.sheetError = 'ネットワークエラー';
          updateUI();
        },
        ontimeout: function() {
          state.sheetStatus = 'error';
          state.sheetError = 'タイムアウト';
          updateUI();
        },
      });
    } else {
      // fallback: fetch（CORSで失敗する可能性あり）
      fetch(normalized, { method: 'GET', mode: 'cors', credentials: 'omit' })
        .then(r => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.text(); })
        .then(csv => {
          const msgs = parseCsv(csv);
          if (msgs.length === 0) throw new Error('データがありません');
          state.messages = msgs;
          state.sheetUrl = url;
          state.sheetStatus = 'success';
          storageSet('sheetUrl', url);
          storageSet('messages', msgs);
          showNotify(`${msgs.length}件読み込み`, 'success');
        })
        .catch(e => {
          state.sheetStatus = 'error';
          state.sheetError = e.message.includes('Failed to fetch') ? 'CORSエラー: 直接入力を使用してください' : e.message;
        })
        .finally(() => updateUI());
    }
  }

  // ============================================================
  // 通知
  // ============================================================
  function showNotify(msg, type) {
    const old = $('#dating-ops-notify');
    if (old) old.remove();
    const colors = { info: '#48bb78', success: '#48bb78', warn: '#f6ad55', error: '#fc8181' };
    const el = document.createElement('div');
    el.id = 'dating-ops-notify';
    el.textContent = msg;
    Object.assign(el.style, {
      position: 'fixed', bottom: '20px', right: '20px', padding: '10px 16px',
      background: colors[type] || colors.info, color: '#fff', borderRadius: '6px',
      boxShadow: '0 4px 12px rgba(0,0,0,0.3)', zIndex: '2147483647',
      fontSize: '13px', fontFamily: 'system-ui, sans-serif', maxWidth: '300px',
    });
    document.body.appendChild(el);
    setTimeout(() => { el.style.opacity = '0'; el.style.transition = 'opacity 0.3s'; setTimeout(() => el.remove(), 300); }, 3500);
  }

  // ============================================================
  // パネル UI
  // ============================================================
  function createStyles() {
    if ($('#dating-ops-styles')) return;
    const p = CONFIG.ui.primaryColor, a = CONFIG.ui.accentColor;
    const css = `
      #${PANEL_ID}{position:fixed;width:350px;max-height:85vh;background:#1e1e2e;border:1px solid ${p};border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,0.5);z-index:2147483646;font-family:system-ui,sans-serif;font-size:13px;color:#cdd6f4;overflow:hidden}
      #${PANEL_ID} *{box-sizing:border-box}
      #${PANEL_ID} .header{display:flex;justify-content:space-between;align-items:center;padding:8px 12px;background:linear-gradient(135deg,${p},${a});cursor:move;user-select:none}
      #${PANEL_ID} .header-title{font-weight:600;color:#fff}
      #${PANEL_ID} .header-close{width:22px;height:22px;display:flex;align-items:center;justify-content:center;background:rgba(255,255,255,0.2);border:none;border-radius:4px;color:#fff;cursor:pointer;font-size:14px}
      #${PANEL_ID} .header-close:hover{background:rgba(255,255,255,0.3)}
      #${PANEL_ID} .body{padding:10px;overflow-y:auto;max-height:calc(85vh - 40px)}
      #${PANEL_ID} .section{margin-bottom:10px}
      #${PANEL_ID} .section-title{font-size:11px;font-weight:600;color:#89b4fa;text-transform:uppercase;margin-bottom:6px}
      #${PANEL_ID} .status-box{background:#313244;border-radius:6px;padding:8px;font-size:12px}
      #${PANEL_ID} .status-row{display:flex;justify-content:space-between;margin-bottom:4px}
      #${PANEL_ID} .status-row:last-child{margin-bottom:0}
      #${PANEL_ID} .status-label{color:#6c7086}
      #${PANEL_ID} .status-val{font-weight:500;text-align:right;max-width:180px;overflow:hidden;text-overflow:ellipsis}
      #${PANEL_ID} .st-ok{color:#a6e3a1}
      #${PANEL_ID} .st-warn{color:#f9e2af}
      #${PANEL_ID} .st-err{color:#f38ba8}
      #${PANEL_ID} .st-info{color:#89b4fa}
      #${PANEL_ID} input,#${PANEL_ID} textarea{width:100%;padding:7px 9px;background:#313244;border:1px solid #45475a;border-radius:4px;color:#cdd6f4;font-size:12px}
      #${PANEL_ID} input:focus,#${PANEL_ID} textarea:focus{outline:none;border-color:${p}}
      #${PANEL_ID} textarea{resize:vertical;min-height:60px;max-height:120px}
      #${PANEL_ID} .btn-row{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px}
      #${PANEL_ID} .btn{padding:6px 10px;border:none;border-radius:4px;font-size:11px;font-weight:500;cursor:pointer}
      #${PANEL_ID} .btn-pri{background:${p};color:#fff}
      #${PANEL_ID} .btn-pri:hover{background:${a}}
      #${PANEL_ID} .btn-sec{background:#45475a;color:#cdd6f4}
      #${PANEL_ID} .btn-sec:hover{background:#585b70}
      #${PANEL_ID} .btn-red{background:#f38ba8;color:#1e1e2e}
      #${PANEL_ID} .msg-list{max-height:120px;overflow-y:auto;background:#313244;border-radius:4px}
      #${PANEL_ID} .msg-item{padding:6px 8px;border-bottom:1px solid #45475a;cursor:pointer;font-size:11px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
      #${PANEL_ID} .msg-item:last-child{border-bottom:none}
      #${PANEL_ID} .msg-item:hover{background:#45475a}
      #${PANEL_ID} .error-box{background:rgba(243,139,168,0.15);border:1px solid #f38ba8;border-radius:4px;padding:8px;font-size:11px;color:#f38ba8;margin-bottom:10px}
      #${PANEL_ID} .warn-box{background:rgba(249,226,175,0.15);border:1px solid #f9e2af;border-radius:4px;padding:8px;font-size:11px;color:#f9e2af;margin-bottom:10px}
      #${PANEL_ID} .help{font-size:10px;color:#6c7086;margin-top:4px}
      #${PANEL_ID} details{margin-top:6px}
      #${PANEL_ID} summary{cursor:pointer;font-size:10px;color:#6c7086}
      #${PANEL_ID} .detail-box{background:#45475a;border-radius:4px;padding:6px;font-size:10px;color:#a6adc8;margin-top:4px}
      #${PANEL_ID} .detail-box .label{color:#6c7086}
    `;
    const s = document.createElement('style');
    s.id = 'dating-ops-styles';
    s.textContent = css;
    document.head.appendChild(s);
  }

  function buildPanelHtml() {
    const pm = state.pageMode;
    const r = state.rows;
    const t = state.textarea;

    // mode表示
    const modeText = pm === 'list' ? '一括送信' : pm === 'personal' ? '個別送信' : 'その他';
    const modeClass = pm === 'list' ? 'st-info' : pm === 'personal' ? 'st-ok' : 'st-warn';

    // 行数表示
    let rowText, rowClass;
    if (pm === 'personal') {
      rowText = '1 (個別送信ページ)';
      rowClass = 'st-ok';
    } else if (pm === 'list') {
      rowText = r.count > 0 ? `${r.count}行` : '0行';
      rowClass = r.count > 0 ? 'st-ok' : 'st-warn';
    } else {
      rowText = '-';
      rowClass = 'st-info';
    }

    // 返信欄
    let taText, taClass;
    if (t.status === 'found') {
      taText = `${t.selector} (${t.where})`;
      taClass = 'st-ok';
    } else {
      taText = '未検出';
      taClass = 'st-warn';
    }

    // Sheet
    const sheetMap = { idle: { t: '未設定', c: 'st-warn' }, loading: { t: '読込中...', c: 'st-info' }, success: { t: `${state.messages.length}件`, c: 'st-ok' }, error: { t: 'エラー', c: 'st-err' } };
    const sh = sheetMap[state.sheetStatus] || sheetMap.idle;

    // 詳細
    let detail = `<div><span class="label">最終検出:</span> ${formatTime(state.lastDetectTime)}</div>`;
    detail += `<div><span class="label">トリガ:</span> ${state.lastTrigger || '-'}</div>`;
    detail += `<div><span class="label">scope:</span> ${escapeHtml(r.scopeHint || '-')}</div>`;
    detail += `<div><span class="label">行セレクタ:</span> ${escapeHtml(r.usedSelector || '-')}</div>`;
    if (t.iframeSrc) detail += `<div><span class="label">iframe:</span> ${escapeHtml(t.iframeSrc)}</div>`;
    if (state.observeTargetHint) detail += `<div><span class="label">監視:</span> ${escapeHtml(state.observeTargetHint)}</div>`;

    // 警告
    let warn = '';
    if (r.mismatchWarning) warn = `<div class="warn-box">⚠️ ${escapeHtml(r.mismatchWarning)}</div>`;

    // メッセージ
    const msgHtml = state.messages.length > 0
      ? state.messages.map((m, i) => `<div class="msg-item" data-idx="${i}" title="${escapeHtml(m)}">${escapeHtml(m.length > 40 ? m.substring(0, 40) + '...' : m)}</div>`).join('')
      : '<div class="msg-item" style="color:#6c7086;">メッセージなし</div>';

    return `
      <div class="header"><span class="header-title">${escapeHtml(CONFIG.ui.title)}</span><button class="header-close" id="dp-close">✕</button></div>
      <div class="body">
        ${warn}
        <div class="section">
          <div class="section-title">ステータス</div>
          <div class="status-box">
            <div class="status-row"><span class="status-label">モード:</span><span class="status-val ${modeClass}">${modeText}</span></div>
            <div class="status-row"><span class="status-label">行数:</span><span class="status-val ${rowClass}">${rowText}</span></div>
            <div class="status-row"><span class="status-label">返信欄:</span><span class="status-val ${taClass}">${taText}</span></div>
            <div class="status-row"><span class="status-label">Sheet:</span><span class="status-val ${sh.c}">${sh.t}</span></div>
          </div>
          <details><summary>詳細情報</summary><div class="detail-box">${detail}</div></details>
        </div>
        ${state.sheetError ? `<div class="error-box">${escapeHtml(state.sheetError)}</div>` : ''}
        <div class="section">
          <div class="section-title">Google Sheet</div>
          <input type="text" id="dp-sheet-url" placeholder="https://docs.google.com/spreadsheets/d/..." value="${escapeHtml(state.sheetUrl)}">
          <div class="help">シートを「リンクを知っている全員」に公開</div>
          <div class="btn-row"><button class="btn btn-pri" id="dp-load-sheet">Sheet読込</button><button class="btn btn-sec" id="dp-rescan">再検出</button></div>
        </div>
        <div class="section">
          <div class="section-title">直接入力</div>
          <textarea id="dp-direct" placeholder="メッセージを1行ずつ入力..."></textarea>
          <div class="btn-row"><button class="btn btn-pri" id="dp-apply-direct">適用</button></div>
        </div>
        <div class="section">
          <div class="section-title">メッセージ (${state.messages.length}件)</div>
          <div class="msg-list" id="dp-msg-list">${msgHtml}</div>
        </div>
        <div class="btn-row"><button class="btn btn-red" id="dp-stop">${state.stopped ? '停止中' : '停止'}</button><button class="btn btn-sec" id="dp-diag">診断</button></div>
      </div>
    `;
  }

  function createPanel() {
    const old = $(`#${PANEL_ID}`);
    if (old) old.remove();
    createStyles();
    const panel = document.createElement('div');
    panel.id = PANEL_ID;
    panel.innerHTML = buildPanelHtml();
    const pos = storageGet('panelPos', null);
    if (pos) state.panelPos = pos;
    panel.style.left = `${state.panelPos.x}px`;
    panel.style.top = `${state.panelPos.y}px`;
    document.body.appendChild(panel);
    state.panel = panel;
    bindPanelEvents();
  }

  function bindPanelEvents() {
    const p = state.panel;
    if (!p) return;
    p.querySelector('.header')?.addEventListener('mousedown', onDragStart);
    p.querySelector('#dp-close')?.addEventListener('click', () => p.style.display = 'none');
    p.querySelector('#dp-load-sheet')?.addEventListener('click', () => {
      const url = p.querySelector('#dp-sheet-url')?.value?.trim();
      url ? fetchSheet(url) : showNotify('URLを入力', 'warn');
    });
    p.querySelector('#dp-rescan')?.addEventListener('click', () => { showNotify('再検出', 'info'); runDetection('manual'); });
    p.querySelector('#dp-apply-direct')?.addEventListener('click', () => {
      const txt = p.querySelector('#dp-direct')?.value?.trim();
      if (!txt) { showNotify('テキストを入力', 'warn'); return; }
      const msgs = txt.split('\n').map(l => l.trim()).filter(Boolean);
      if (msgs.length === 0) { showNotify('有効なメッセージなし', 'warn'); return; }
      state.messages = msgs;
      storageSet('messages', msgs);
      state.sheetStatus = 'success';
      showNotify(`${msgs.length}件適用`, 'success');
      updateUI();
    });
    p.querySelector('#dp-stop')?.addEventListener('click', () => {
      state.stopped = true;
      state.observer?.disconnect();
      clearInterval(state.healthTimer);
      clearInterval(state.urlPollTimer);
      clearTimeout(state.stagedDetectTimer);
      showNotify('停止しました', 'warn');
      updateUI();
    });
    p.querySelector('#dp-diag')?.addEventListener('click', runDiagnostic);
    p.querySelector('#dp-msg-list')?.addEventListener('click', e => {
      const item = e.target.closest('.msg-item');
      if (item) { const i = parseInt(item.dataset.idx, 10); if (!isNaN(i) && state.messages[i]) insertText(state.messages[i]); }
    });
    document.addEventListener('mousemove', onDragMove);
    document.addEventListener('mouseup', onDragEnd);
  }

  function onDragStart(e) {
    if (e.target.closest('.header-close')) return;
    state.dragging = true;
    state.dragOffset = { x: e.clientX - state.panelPos.x, y: e.clientY - state.panelPos.y };
    if (state.panel) state.panel.style.transition = 'none';
  }
  function onDragMove(e) {
    if (!state.dragging || !state.panel) return;
    state.panelPos = { x: Math.max(0, Math.min(window.innerWidth - 350, e.clientX - state.dragOffset.x)), y: Math.max(0, Math.min(window.innerHeight - 50, e.clientY - state.dragOffset.y)) };
    state.panel.style.left = `${state.panelPos.x}px`;
    state.panel.style.top = `${state.panelPos.y}px`;
  }
  function onDragEnd() {
    if (!state.dragging) return;
    state.dragging = false;
    storageSet('panelPos', state.panelPos);
    if (state.panel) state.panel.style.transition = '';
  }
  function updateUI() { if (state.panel) { state.panel.innerHTML = buildPanelHtml(); bindPanelEvents(); } }
  function runDiagnostic() {
    const d = { siteType: SITE_TYPE, version: '2.3.0', pageMode: state.pageMode, textarea: state.textarea, rows: state.rows, lastDetectTime: state.lastDetectTime, lastTrigger: state.lastTrigger, messages: state.messages.length, sheetStatus: state.sheetStatus, sheetError: state.sheetError, url: location.href };
    console.group('%c[DatingOps] 診断', 'color:#48bb78;font-weight:bold'); console.log(d); console.groupEnd();
    navigator.clipboard?.writeText(JSON.stringify(d, null, 2)).then(() => showNotify('診断情報コピー', 'info')).catch(() => {});
    return d;
  }

  // ============================================================
  // トリガイベント
  // ============================================================
  function bindTriggerEvents() {
    const debouncedDetect = debounce(trigger => runDetection(trigger), OBSERVER_DEBOUNCE_MS);
    document.addEventListener('click', e => {
      if (state.stopped || !e.target || e.target.closest(`#${PANEL_ID}`)) return;
      for (const sel of CONFIG.clickTriggerSelectors) {
        try { if (e.target.matches(sel) || e.target.closest(sel)) { debouncedDetect('click'); return; } } catch {}
      }
    }, true);
    document.addEventListener('change', e => {
      if (state.stopped || !e.target || e.target.tagName !== 'SELECT' || e.target.closest(`#${PANEL_ID}`)) return;
      for (const sel of CONFIG.changeTriggerSelectors) {
        try { if (e.target.matches(sel)) { debouncedDetect('change'); return; } } catch {}
      }
    }, true);
  }

  // ============================================================
  // MutationObserver
  // ============================================================
  function findObserveTarget() {
    for (const sel of CONFIG.listAreaSelectors) { const el = $(sel); if (el) return { target: el, hint: sel }; }
    const scope = findRowScope();
    if (scope.el !== document.body) return { target: scope.el, hint: `scope: ${scope.hint}` };
    return { target: document.body, hint: 'body (fallback)' };
  }

  function startObserver() {
    state.observer?.disconnect();
    const info = findObserveTarget();
    state.observeTargetHint = info.hint;
    if (info.target === document.body) logger.warn('一覧領域特定できず、bodyを監視');
    else logger.info(`監視対象: ${info.hint}`);
    const handler = debounce(() => { if (!state.stopped) runDetection('mutation'); }, OBSERVER_DEBOUNCE_MS);
    state.observer = new MutationObserver(muts => {
      if (muts.some(m => m.target.id === PANEL_ID || m.target.closest?.(`#${PANEL_ID}`))) return;
      handler();
    });
    state.observer.observe(info.target, { childList: true, subtree: true });
  }

  // ============================================================
  // URL変化検知
  // ============================================================
  function setupUrlChangeDetection() {
    state.lastUrl = location.href;
    // history hook
    const wrap = fn => function() { const r = fn.apply(this, arguments); onUrlChange('history'); return r; };
    history.pushState = wrap(history.pushState);
    history.replaceState = wrap(history.replaceState);
    window.addEventListener('popstate', () => onUrlChange('popstate'));
    // polling
    state.urlPollTimer = setInterval(() => { if (location.href !== state.lastUrl) onUrlChange('poll'); }, URL_POLL_MS);
  }

  function onUrlChange(src) {
    if (state.stopped) return;
    state.lastUrl = location.href;
    logger.info(`URL変化検知 (${src})`);
    startObserver();
    runStagedDetection();
  }

  // ============================================================
  // ヘルスチェック
  // ============================================================
  function startHealthCheck() {
    state.healthTimer = setInterval(() => {
      if (state.stopped) { clearInterval(state.healthTimer); return; }
      if (!$(`#${PANEL_ID}`)) createPanel();
      const pm = state.pageMode;
      const taOk = state.textarea.status === 'found';
      const rowsOk = pm === 'personal' || state.rows.count > 0;
      if (!taOk || (pm === 'list' && !rowsOk)) {
        logger.info('ヘルスチェック: 再検出');
        runDetection('health');
      }
    }, HEALTH_CHECK_MS);
  }

  // ============================================================
  // 初期化
  // ============================================================
  function init() {
    if (state.initialized || !location.pathname.includes('/staff/')) return;
    logger.info('初期化開始');
    state.messages = storageGet('messages', []);
    state.sheetUrl = storageGet('sheetUrl', '');
    if (state.messages.length > 0) state.sheetStatus = 'success';
    state.pageMode = detectPageMode();
    createPanel();
    bindTriggerEvents();
    startObserver();
    setupUrlChangeDetection();
    startHealthCheck();
    runStagedDetection();
    state.initialized = true;
    window.__datingOps = { state, diag: runDiagnostic, rescan: () => runDetection('manual'), insert: insertText, fetchSheet };
    logger.info('初期化完了');
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
