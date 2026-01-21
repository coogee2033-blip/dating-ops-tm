// ==UserScript==
// @name         Dating Ops - MEM
// @namespace    dating-ops-tm
// @version      1.1.0
// @description  MEM (mem44.com) サイト用返信アシストツール
// @author       Dating Ops Team
// @match        https://mem44.com/staff/*
// @match        https://*.mem44.com/staff/*
// @grant        none
// @run-at       document-idle
// @noframes
// ==/UserScript==

(function() {
  'use strict';

  // ============================================================
  // 実装メモ:
  // - 返信欄検出: findReplyTextarea() で main → iframe の順に探索
  //   優先順: message1 → .msg.wd100 → message → content
  // - 行検出: countTargetRows() で tr.rowitem, tr[id^="row"] 等を検索
  // - 返信欄未検出でも行検出は継続（早期return禁止）
  // - 除外対象: admin_memo, staff_memo, two_memo, futari_memo, memo, note
  // - 既存テキストがあれば上書き禁止
  // ============================================================

  // ============================================================
  // sites.js - サイト固有設定（埋め込み）
  // ============================================================
  const SITE_CONFIGS = {
    mem: {
      siteType: 'mem',
      siteName: 'MEM',
      panelPages: [/\/staff\//],
      
      // 返信欄セレクタ（優先度順）
      textareaSelectors: [
        'textarea[name="message1"]',      // 最優先
        'textarea.msg.wd100',             // 特有のクラス
        'textarea.msg',                   // msgクラスのみ
        'textarea[name="message"]',
        'textarea[name="body"]',
        'textarea[name="mail_body"]',
        'textarea[name="content"]',
        'textarea:not([name*="memo"]):not([name*="admin"]):not([name*="note"])', // 汎用フォールバック
      ],
      
      // 挿入禁止セレクタ（管理者メモ・二人メモ等）
      excludeSelectors: [
        'textarea[name="admin_memo"]',
        'textarea[name="staff_memo"]',
        'textarea[name="memo"]',
        'textarea[name="note"]',
        'textarea[name="two_memo"]',
        'textarea[name="admin_note"]',
        'textarea[name="futari_memo"]',
        'textarea.admin-memo',
        'textarea.staff-memo',
        'textarea#admin_memo',
        'textarea#staff_memo',
        '[data-type="admin"]',
        '[data-type="memo"]',
      ],
      
      // 行検出用セレクタ
      rowSelectors: [
        'tr.rowitem',
        'tr[id^="row"]',
        'tr[class*="row"]',
        'table.list tr',
        'table tr:has(td.chatview)',
        'table tr:has(input[type="checkbox"])',
      ],
      
      rowChatCellSelector: 'td.chatview',
      rowCheckboxSelector: 'input[type="checkbox"]',
      
      // iframe検索用（全iframe対象）
      iframeSelectors: [
        'iframe',
      ],
      
      ui: {
        title: 'MEM 返信アシスト',
        primaryColor: '#48bb78',
        accentColor: '#276749',
      },
    },
  };

  function getConfig(siteType) {
    return SITE_CONFIGS[siteType] || null;
  }

  function shouldShowPanel(siteType, pathname) {
    const config = getConfig(siteType);
    if (!config) return false;
    return config.panelPages.some(pattern => pattern.test(pathname));
  }

  // ============================================================
  // common.js - 共通コアモジュール（埋め込み）
  // ============================================================
  const PANEL_ID = 'dating-ops-panel';
  const STORAGE_PREFIX = 'datingOps_';
  const OBSERVER_DEBOUNCE_MS = 300;
  const HEALTH_CHECK_MS = 3000;
  const MAX_DETECT_RETRY = 5;
  const DETECT_RETRY_INTERVAL = 1000;

  const state = {
    siteType: 'mem',
    config: null,
    initialized: false,
    stopped: false,
    panel: null,
    observer: null,
    healthTimer: null,
    detectRetryCount: 0,
    detectTimer: null,
    messages: [],
    sheetUrl: '',
    sheetStatus: 'idle',
    sheetError: null,
    
    // 返信欄情報（拡張）
    textareaInfo: {
      status: 'unknown',
      element: null,
      selector: null,
      where: null,
      iframeSrc: null,
      reason: null,
    },
    
    // 行検出情報（新規）
    rowInfo: {
      displayCount: 0,
      checkedCount: 0,
      selectorHit: null,
    },
    
    panelPos: { x: 20, y: 20 },
    dragging: false,
    dragOffset: { x: 0, y: 0 },
  };

  const LOG_STYLES = {
    info: 'background:#48bb78;color:#fff;padding:2px 6px;border-radius:3px;',
    warn: 'background:#f6ad55;color:#000;padding:2px 6px;border-radius:3px;',
    error: 'background:#fc8181;color:#000;padding:2px 6px;border-radius:3px;',
  };

  function log(level, msg, data = null) {
    const tag = `DatingOps:${state.siteType}`;
    const style = LOG_STYLES[level] || '';
    if (data !== null) {
      console.log(`%c${tag}%c ${msg}`, style, '', data);
    } else {
      console.log(`%c${tag}%c ${msg}`, style, '');
    }
  }

  const logger = {
    info: (m, d) => log('info', m, d),
    warn: (m, d) => log('warn', m, d),
    error: (m, d) => log('error', m, d),
  };

  function $(sel, ctx = document) {
    try {
      if (!ctx || typeof ctx.querySelector !== 'function') return null;
      return ctx.querySelector(sel);
    } catch (e) { return null; }
  }

  function $$(sel, ctx = document) {
    try {
      if (!ctx || typeof ctx.querySelectorAll !== 'function') return [];
      return Array.from(ctx.querySelectorAll(sel));
    } catch (e) { return []; }
  }

  function debounce(fn, ms) {
    let t = null;
    return (...args) => {
      clearTimeout(t);
      t = setTimeout(() => fn(...args), ms);
    };
  }

  function storageGet(key, def = null) {
    try {
      const v = localStorage.getItem(STORAGE_PREFIX + key);
      return v ? JSON.parse(v) : def;
    } catch { return def; }
  }

  function storageSet(key, val) {
    try {
      localStorage.setItem(STORAGE_PREFIX + key, JSON.stringify(val));
    } catch (e) { logger.warn('localStorage保存失敗', e); }
  }

  function escapeHtml(s) {
    const d = document.createElement('div');
    d.textContent = s;
    return d.innerHTML;
  }

  // ============================================================
  // 除外判定
  // ============================================================
  function isExcluded(el) {
    if (!el || !state.config) return true;
    
    for (const sel of state.config.excludeSelectors) {
      try { if (el.matches(sel)) return true; } catch {}
    }
    
    const dataType = el.getAttribute('data-type');
    if (dataType && (dataType.includes('admin') || dataType.includes('memo'))) return true;
    
    const name = (el.name || '').toLowerCase();
    const id = (el.id || '').toLowerCase();
    const className = (el.className || '').toLowerCase();
    
    const dangerousKeywords = ['memo', 'admin', 'note', 'futari', 'two_memo', 'staff'];
    for (const kw of dangerousKeywords) {
      if (name.includes(kw) || id.includes(kw)) return true;
    }
    
    for (const kw of ['admin', 'memo', 'note']) {
      if (className.includes(kw)) return true;
    }
    
    return false;
  }

  // ============================================================
  // 返信欄検出（改良版）
  // ============================================================
  function findReplyTextarea() {
    const cfg = state.config;
    if (!cfg) {
      return { el: null, where: null, selector: null, reason: '設定が読み込まれていません' };
    }

    // (1) main document を検索
    for (const sel of cfg.textareaSelectors) {
      try {
        const els = $$(sel, document);
        for (const el of els) {
          if (!isExcluded(el) && isVisible(el)) {
            logger.info(`[main] 返信欄検出: ${sel}`, el);
            return { el, where: 'main', selector: sel, reason: null };
          }
        }
      } catch (e) {}
    }

    // (2) 全ての iframe を検索（同一オリジンのみ）
    const allIframes = $$('iframe', document);
    for (let i = 0; i < allIframes.length; i++) {
      const iframe = allIframes[i];
      try {
        const doc = iframe.contentDocument || iframe.contentWindow?.document;
        if (!doc) continue;
        
        for (const sel of cfg.textareaSelectors) {
          try {
            const els = $$(sel, doc);
            for (const el of els) {
              if (!isExcluded(el)) {
                const iframeSrc = iframe.src || iframe.name || `iframe[${i}]`;
                logger.info(`[iframe:${i}] 返信欄検出: ${sel}`, { iframe: iframeSrc, el });
                return { el, where: 'iframe', selector: sel, iframeIndex: i, iframeSrc, reason: null };
              }
            }
          } catch (e) {}
        }
        
        if (doc.body?.contentEditable === 'true') {
          const iframeSrc = iframe.src || iframe.name || `iframe[${i}]`;
          logger.info(`[iframe:${i}] contentEditable検出`, iframeSrc);
          return { el: doc.body, where: 'iframe', selector: 'body[contenteditable]', iframeIndex: i, iframeSrc, reason: null };
        }
      } catch (e) {
        if (e.name === 'SecurityError') {
          logger.warn(`iframe[${i}] クロスオリジン - スキップ`);
        }
      }
    }

    const tried = cfg.textareaSelectors.join(', ');
    return { el: null, where: null, selector: null, reason: `検索済みセレクタ: ${tried}` };
  }

  function isVisible(el) {
    if (!el) return false;
    const style = window.getComputedStyle(el);
    if (style.display === 'none' || style.visibility === 'hidden') return false;
    return true;
  }

  function updateTextareaInfo(result) {
    state.textareaInfo = {
      status: result.el ? 'found' : 'not_found',
      element: result.el,
      selector: result.selector,
      where: result.where,
      iframeSrc: result.iframeSrc || null,
      reason: result.reason,
    };
  }

  // ============================================================
  // 行検出（新規）
  // ============================================================
  function countTargetRows() {
    const cfg = state.config;
    if (!cfg || !cfg.rowSelectors) {
      state.rowInfo = { displayCount: 0, checkedCount: 0, selectorHit: null };
      return;
    }

    let displayCount = 0;
    let checkedCount = 0;
    let selectorHit = null;

    for (const sel of cfg.rowSelectors) {
      try {
        const rows = $$(sel, document);
        if (rows.length > 0) {
          displayCount = rows.length;
          selectorHit = sel;
          
          if (cfg.rowCheckboxSelector) {
            for (const row of rows) {
              const cb = $(cfg.rowCheckboxSelector, row);
              if (cb && cb.checked) checkedCount++;
            }
          }
          
          logger.info(`行検出: ${sel} → ${displayCount}件 (checked: ${checkedCount})`);
          break;
        }
      } catch (e) {}
    }

    if (displayCount === 0) {
      const fallbackRows = $$('table tr', document).filter(tr => {
        if ($(tr, 'th')) return false;
        return $('td', tr) !== null;
      });
      if (fallbackRows.length > 0) {
        displayCount = fallbackRows.length;
        selectorHit = 'table tr (fallback)';
        logger.info(`行検出(fallback): ${displayCount}件`);
      }
    }

    state.rowInfo = { displayCount, checkedCount, selectorHit };
  }

  // ============================================================
  // 統合検出
  // ============================================================
  function detectAll() {
    const taResult = findReplyTextarea();
    updateTextareaInfo(taResult);
    countTargetRows();
    return taResult.el;
  }

  function detectAllWithRetry() {
    if (state.stopped) return;
    clearTimeout(state.detectTimer);
    state.detectRetryCount = 0;
    
    function attempt() {
      if (state.stopped) return;
      
      const el = detectAll();
      
      if (!el && state.rowInfo.displayCount > 0) {
        state.detectRetryCount++;
        if (state.detectRetryCount < MAX_DETECT_RETRY) {
          logger.info(`返信欄再検出 (${state.detectRetryCount}/${MAX_DETECT_RETRY}) - 行は${state.rowInfo.displayCount}件検出済み`);
          state.detectTimer = setTimeout(attempt, DETECT_RETRY_INTERVAL);
          updateUI();
          return;
        } else {
          logger.warn('返信欄検出リトライ上限 - 行は検出済み');
        }
      }
      
      state.detectRetryCount = 0;
      updateUI();
    }
    
    attempt();
  }

  // ============================================================
  // テキスト挿入
  // ============================================================
  function insertText(text) {
    const ta = state.textareaInfo.element;
    if (!ta) {
      showNotify('返信欄が見つかりません', 'error');
      logger.error('挿入失敗: テキストエリアなし');
      return false;
    }
    if (isExcluded(ta)) {
      showNotify('この欄には挿入できません（管理者用）', 'error');
      logger.error('挿入失敗: 除外対象');
      return false;
    }
    
    const currentValue = ta.tagName === 'TEXTAREA' || ta.tagName === 'INPUT'
      ? ta.value
      : ta.textContent || ta.innerHTML;
    if (currentValue && currentValue.trim().length > 0) {
      showNotify('既にテキストが入力されています。上書きしません。', 'warn');
      logger.warn('挿入スキップ: 既存テキストあり', currentValue.substring(0, 50));
      return false;
    }
    
    try {
      if (ta.tagName === 'TEXTAREA' || ta.tagName === 'INPUT') {
        ta.value = text;
        ta.dispatchEvent(new Event('input', { bubbles: true }));
        ta.dispatchEvent(new Event('change', { bubbles: true }));
      } else if (ta.contentEditable === 'true' || ta.tagName === 'BODY') {
        ta.innerHTML = escapeHtml(text).replace(/\n/g, '<br>');
        ta.dispatchEvent(new Event('input', { bubbles: true }));
      }
      ta.focus();
      showNotify('挿入しました', 'success');
      logger.info('テキスト挿入完了');
      return true;
    } catch (e) {
      showNotify('挿入に失敗しました', 'error');
      logger.error('挿入エラー', e);
      return false;
    }
  }

  // ============================================================
  // Google Sheet
  // ============================================================
  function normalizeSheetUrl(url) {
    if (!url) return null;
    const u = url.trim();
    if (u.includes('/export?format=csv') || u.includes('output=csv') || u.includes('/gviz/tq')) return u;
    const m = u.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (m) {
      const id = m[1];
      const gidM = u.match(/gid=(\d+)/);
      const gid = gidM ? gidM[1] : '0';
      return `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${gid}`;
    }
    return u;
  }

  function parseCsv(csv) {
    return csv.split('\n').map(line => {
      line = line.trim();
      if (!line) return null;
      if (line.startsWith('"')) {
        const end = line.indexOf('",');
        if (end > 0) return line.substring(1, end).replace(/""/g, '"');
        if (line.endsWith('"')) return line.slice(1, -1).replace(/""/g, '"');
      }
      const comma = line.indexOf(',');
      return comma > 0 ? line.substring(0, comma) : line;
    }).filter(Boolean);
  }

  async function fetchSheet(url) {
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
    try {
      const res = await fetch(normalized, { method: 'GET', mode: 'cors', credentials: 'omit' });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const csv = await res.text();
      const msgs = parseCsv(csv);
      if (msgs.length === 0) throw new Error('データがありません');
      state.messages = msgs;
      state.sheetUrl = url;
      state.sheetStatus = 'success';
      storageSet('sheetUrl', url);
      storageSet('messages', msgs);
      logger.info(`Sheet読み込み成功: ${msgs.length}件`);
      showNotify(`${msgs.length}件読み込み`, 'success');
    } catch (e) {
      state.sheetStatus = 'error';
      if (e.message.includes('Failed to fetch')) {
        state.sheetError = 'CORSエラー: シートを「リンクを知っている全員」に公開してください';
      } else if (e.message.includes('403')) {
        state.sheetError = '権限エラー(403): 公開設定を確認してください';
      } else if (e.message.includes('404')) {
        state.sheetError = 'シート未発見(404): URLを確認してください';
      } else {
        state.sheetError = e.message;
      }
      logger.error('Sheet読み込み失敗', e);
    }
    updateUI();
  }

  // ============================================================
  // 通知
  // ============================================================
  function showNotify(msg, type = 'info') {
    const old = $('#dating-ops-notify');
    if (old) old.remove();
    const colors = { info: '#48bb78', success: '#48bb78', warn: '#f6ad55', error: '#fc8181' };
    const el = document.createElement('div');
    el.id = 'dating-ops-notify';
    Object.assign(el.style, {
      position: 'fixed', bottom: '20px', right: '20px', padding: '10px 16px',
      background: colors[type] || colors.info, color: '#fff', borderRadius: '6px',
      boxShadow: '0 4px 12px rgba(0,0,0,0.3)', zIndex: '2147483647',
      fontSize: '13px', fontFamily: 'system-ui, sans-serif', maxWidth: '280px',
    });
    el.textContent = msg;
    document.body.appendChild(el);
    setTimeout(() => {
      el.style.transition = 'opacity 0.3s';
      el.style.opacity = '0';
      setTimeout(() => el.remove(), 300);
    }, 3500);
  }

  // ============================================================
  // パネルUI
  // ============================================================
  function createStyles() {
    if ($('#dating-ops-styles')) return;
    const cfg = state.config;
    const primary = cfg?.ui?.primaryColor || '#48bb78';
    const accent = cfg?.ui?.accentColor || '#276749';
    const css = `
      #${PANEL_ID}{position:fixed;width:320px;max-height:85vh;background:#1e1e2e;border:1px solid ${primary};border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,0.5);z-index:2147483646;font-family:system-ui,sans-serif;font-size:13px;color:#cdd6f4;overflow:hidden}
      #${PANEL_ID} *{box-sizing:border-box}
      #${PANEL_ID} .header{display:flex;justify-content:space-between;align-items:center;padding:8px 12px;background:linear-gradient(135deg,${primary},${accent});cursor:move;user-select:none}
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
      #${PANEL_ID} .status-val{font-weight:500}
      #${PANEL_ID} .st-ok{color:#a6e3a1}
      #${PANEL_ID} .st-warn{color:#f9e2af}
      #${PANEL_ID} .st-err{color:#f38ba8}
      #${PANEL_ID} .st-info{color:#89b4fa}
      #${PANEL_ID} input,#${PANEL_ID} textarea{width:100%;padding:7px 9px;background:#313244;border:1px solid #45475a;border-radius:4px;color:#cdd6f4;font-size:12px;font-family:inherit}
      #${PANEL_ID} input:focus,#${PANEL_ID} textarea:focus{outline:none;border-color:${primary}}
      #${PANEL_ID} textarea{resize:vertical;min-height:60px;max-height:120px}
      #${PANEL_ID} .btn-row{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px}
      #${PANEL_ID} .btn{padding:6px 10px;border:none;border-radius:4px;font-size:11px;font-weight:500;cursor:pointer}
      #${PANEL_ID} .btn-pri{background:${primary};color:#fff}
      #${PANEL_ID} .btn-pri:hover{background:${accent}}
      #${PANEL_ID} .btn-sec{background:#45475a;color:#cdd6f4}
      #${PANEL_ID} .btn-sec:hover{background:#585b70}
      #${PANEL_ID} .btn-red{background:#f38ba8;color:#1e1e2e}
      #${PANEL_ID} .btn-red:hover{background:#eba0ac}
      #${PANEL_ID} .msg-list{max-height:130px;overflow-y:auto;background:#313244;border-radius:4px}
      #${PANEL_ID} .msg-item{padding:6px 8px;border-bottom:1px solid #45475a;cursor:pointer;font-size:11px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
      #${PANEL_ID} .msg-item:last-child{border-bottom:none}
      #${PANEL_ID} .msg-item:hover{background:#45475a}
      #${PANEL_ID} .error-box{background:rgba(243,139,168,0.15);border:1px solid #f38ba8;border-radius:4px;padding:8px;font-size:11px;color:#f38ba8;margin-bottom:10px;white-space:pre-wrap}
      #${PANEL_ID} .help{font-size:10px;color:#6c7086;margin-top:4px}
      #${PANEL_ID} .detail-box{background:#45475a;border-radius:4px;padding:6px;font-size:10px;color:#a6adc8;margin-top:6px;word-break:break-all}
      #${PANEL_ID} .detail-box .label{color:#6c7086}
    `;
    const style = document.createElement('style');
    style.id = 'dating-ops-styles';
    style.textContent = css;
    document.head.appendChild(style);
  }

  function createPanelHtml() {
    const cfg = state.config;
    const title = cfg?.ui?.title || 'Dating Ops';
    const taInfo = state.textareaInfo;
    const rowInfo = state.rowInfo;
    
    let taStatusText, taStatusClass;
    if (taInfo.status === 'found') {
      taStatusText = taInfo.where === 'iframe' ? `iframe内OK` : '検出OK';
      taStatusClass = 'st-ok';
    } else if (taInfo.status === 'not_found') {
      taStatusText = '未検出';
      taStatusClass = 'st-warn';
    } else if (taInfo.status === 'error') {
      taStatusText = 'エラー';
      taStatusClass = 'st-err';
    } else {
      taStatusText = '未確認';
      taStatusClass = 'st-warn';
    }
    
    let rowStatusText, rowStatusClass;
    if (rowInfo.displayCount > 0) {
      rowStatusText = `${rowInfo.displayCount}件`;
      if (rowInfo.checkedCount > 0) {
        rowStatusText += ` (✓${rowInfo.checkedCount})`;
      }
      rowStatusClass = 'st-ok';
    } else {
      rowStatusText = '0件';
      rowStatusClass = 'st-warn';
    }
    
    const sheetStatusText = { idle: '未設定', loading: '読込中...', success: `${state.messages.length}件`, error: 'エラー' }[state.sheetStatus] || '不明';
    const sheetStatusClass = { idle: 'st-warn', loading: 'st-info', success: 'st-ok', error: 'st-err' }[state.sheetStatus] || 'st-warn';

    let detailHtml = '';
    if (taInfo.selector) {
      detailHtml += `<div><span class="label">セレクタ:</span> ${escapeHtml(taInfo.selector)}</div>`;
    }
    if (taInfo.where === 'iframe' && taInfo.iframeSrc) {
      detailHtml += `<div><span class="label">iframe:</span> ${escapeHtml(taInfo.iframeSrc)}</div>`;
    }
    if (taInfo.reason) {
      detailHtml += `<div><span class="label">理由:</span> ${escapeHtml(taInfo.reason)}</div>`;
    }
    if (rowInfo.selectorHit) {
      detailHtml += `<div><span class="label">行セレクタ:</span> ${escapeHtml(rowInfo.selectorHit)}</div>`;
    }

    return `
      <div class="header">
        <span class="header-title">${escapeHtml(title)}</span>
        <button class="header-close" id="dp-close" title="閉じる">✕</button>
      </div>
      <div class="body">
        <div class="section">
          <div class="section-title">ステータス</div>
          <div class="status-box">
            <div class="status-row"><span class="status-label">サイト:</span><span class="status-val st-info">${state.siteType.toUpperCase()}</span></div>
            <div class="status-row"><span class="status-label">返信欄:</span><span class="status-val ${taStatusClass}">${taStatusText}</span></div>
            <div class="status-row"><span class="status-label">表示行:</span><span class="status-val ${rowStatusClass}">${rowStatusText}</span></div>
            <div class="status-row"><span class="status-label">Sheet:</span><span class="status-val ${sheetStatusClass}">${sheetStatusText}</span></div>
          </div>
          ${detailHtml ? `<div class="detail-box">${detailHtml}</div>` : ''}
        </div>
        ${state.sheetError ? `<div class="error-box">${escapeHtml(state.sheetError)}</div>` : ''}
        <div class="section">
          <div class="section-title">Google Sheet</div>
          <input type="text" id="dp-sheet-url" placeholder="https://docs.google.com/spreadsheets/d/..." value="${escapeHtml(state.sheetUrl || '')}">
          <div class="help">シートを「リンクを知っている全員」に公開</div>
          <div class="btn-row">
            <button class="btn btn-pri" id="dp-load-sheet">Sheet読込</button>
            <button class="btn btn-sec" id="dp-rescan">再検出</button>
          </div>
        </div>
        <div class="section">
          <div class="section-title">直接入力（1行=1メッセージ）</div>
          <textarea id="dp-direct" placeholder="メッセージを1行ずつ..."></textarea>
          <div class="btn-row"><button class="btn btn-pri" id="dp-apply-direct">適用</button></div>
        </div>
        <div class="section">
          <div class="section-title">メッセージ一覧 (${state.messages.length}件)</div>
          <div class="msg-list" id="dp-msg-list">
            ${state.messages.length > 0 ? state.messages.map((m, i) => `<div class="msg-item" data-idx="${i}" title="${escapeHtml(m)}">${escapeHtml(m.substring(0, 45))}${m.length > 45 ? '...' : ''}</div>`).join('') : '<div class="msg-item" style="color:#6c7086;">メッセージなし</div>'}
          </div>
        </div>
        <div class="btn-row">
          <button class="btn btn-red" id="dp-stop">${state.stopped ? '停止中' : '停止'}</button>
          <button class="btn btn-sec" id="dp-diag">診断</button>
        </div>
      </div>
    `;
  }

  function createPanel() {
    const old = $(`#${PANEL_ID}`);
    if (old) old.remove();
    createStyles();
    const panel = document.createElement('div');
    panel.id = PANEL_ID;
    panel.innerHTML = createPanelHtml();
    const savedPos = storageGet('panelPos');
    if (savedPos) state.panelPos = savedPos;
    panel.style.left = `${state.panelPos.x}px`;
    panel.style.top = `${state.panelPos.y}px`;
    document.body.appendChild(panel);
    state.panel = panel;
    setupPanelEvents();
    logger.info('パネル作成完了');
  }

  function setupPanelEvents() {
    const panel = state.panel;
    if (!panel) return;
    const header = panel.querySelector('.header');
    if (header) header.addEventListener('mousedown', onDragStart);
    const closeBtn = panel.querySelector('#dp-close');
    if (closeBtn) closeBtn.addEventListener('click', () => { panel.style.display = 'none'; logger.info('パネルを閉じました'); });
    const loadBtn = panel.querySelector('#dp-load-sheet');
    if (loadBtn) loadBtn.addEventListener('click', () => {
      const url = panel.querySelector('#dp-sheet-url')?.value?.trim();
      if (url) fetchSheet(url); else showNotify('URLを入力してください', 'warn');
    });
    const rescanBtn = panel.querySelector('#dp-rescan');
    if (rescanBtn) rescanBtn.addEventListener('click', () => { detectAllWithRetry(); panel.style.display = ''; showNotify('再検出開始', 'info'); });
    const applyBtn = panel.querySelector('#dp-apply-direct');
    if (applyBtn) applyBtn.addEventListener('click', () => {
      const text = panel.querySelector('#dp-direct')?.value?.trim();
      if (!text) { showNotify('テキストを入力してください', 'warn'); return; }
      const msgs = text.split('\n').filter(l => l.trim());
      if (msgs.length === 0) { showNotify('有効なメッセージがありません', 'warn'); return; }
      state.messages = msgs;
      storageSet('messages', msgs);
      showNotify(`${msgs.length}件適用`, 'success');
      updateUI();
    });
    const stopBtn = panel.querySelector('#dp-stop');
    if (stopBtn) stopBtn.addEventListener('click', () => {
      state.stopped = true;
      if (state.observer) { state.observer.disconnect(); state.observer = null; }
      if (state.healthTimer) { clearInterval(state.healthTimer); state.healthTimer = null; }
      clearTimeout(state.detectTimer);
      logger.info('完全停止');
      showNotify('監視を停止しました', 'warn');
      updateUI();
    });
    const diagBtn = panel.querySelector('#dp-diag');
    if (diagBtn) diagBtn.addEventListener('click', runDiagnostic);
    const msgList = panel.querySelector('#dp-msg-list');
    if (msgList) msgList.addEventListener('click', (e) => {
      const item = e.target.closest('.msg-item');
      if (!item) return;
      const idx = parseInt(item.dataset.idx, 10);
      if (!isNaN(idx) && state.messages[idx]) insertText(state.messages[idx]);
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
    const x = Math.max(0, Math.min(window.innerWidth - 320, e.clientX - state.dragOffset.x));
    const y = Math.max(0, Math.min(window.innerHeight - 50, e.clientY - state.dragOffset.y));
    state.panelPos = { x, y };
    state.panel.style.left = `${x}px`;
    state.panel.style.top = `${y}px`;
  }

  function onDragEnd() {
    if (!state.dragging) return;
    state.dragging = false;
    storageSet('panelPos', state.panelPos);
    if (state.panel) state.panel.style.transition = '';
  }

  function updateUI() {
    const panel = state.panel;
    if (!panel) return;
    panel.innerHTML = createPanelHtml();
    setupPanelEvents();
  }

  function runDiagnostic() {
    const diag = {
      siteType: state.siteType,
      stopped: state.stopped,
      initialized: state.initialized,
      textarea: {
        status: state.textareaInfo.status,
        selector: state.textareaInfo.selector,
        where: state.textareaInfo.where,
        iframeSrc: state.textareaInfo.iframeSrc,
        reason: state.textareaInfo.reason,
        hasElement: !!state.textareaInfo.element,
      },
      rows: state.rowInfo,
      messages: state.messages.length,
      sheetStatus: state.sheetStatus,
      sheetError: state.sheetError,
      url: location.href,
      allTextareas: $$('textarea', document).map(t => ({ name: t.name, id: t.id, class: t.className })),
    };
    console.group('%c[DatingOps] 診断情報', 'color:#48bb78;font-weight:bold');
    console.log(diag);
    console.log('設定:', state.config);
    console.groupEnd();
    navigator.clipboard?.writeText(JSON.stringify(diag, null, 2))
      .then(() => showNotify('診断情報をコピーしました', 'info'))
      .catch(() => showNotify('診断情報をコンソールに出力しました', 'info'));
    return diag;
  }

  function startObserver() {
    if (state.observer) state.observer.disconnect();
    const handler = debounce(() => {
      if (state.stopped) return;
      ensurePanel();
      detectAllWithRetry();
    }, OBSERVER_DEBOUNCE_MS);
    state.observer = new MutationObserver((muts) => {
      const isOwn = muts.some(m => m.target.id === PANEL_ID || m.target.closest?.(`#${PANEL_ID}`));
      if (isOwn) return;
      handler();
    });
    state.observer.observe(document.body, { childList: true, subtree: true });
    logger.info('DOM監視開始');
  }

  function ensurePanel() {
    if (state.stopped) return;
    const panel = $(`#${PANEL_ID}`);
    if (!panel) { logger.warn('パネル消失、再作成'); createPanel(); }
    else state.panel = panel;
  }

  function startHealthCheck() {
    if (state.healthTimer) clearInterval(state.healthTimer);
    state.healthTimer = setInterval(() => {
      if (state.stopped) { clearInterval(state.healthTimer); return; }
      ensurePanel();
    }, HEALTH_CHECK_MS);
  }

  function init() {
    if (state.initialized) { logger.warn('既に初期化済み'); return; }
    state.config = getConfig('mem');
    if (!state.config) { logger.error('設定が見つかりません'); return; }
    if (!shouldShowPanel('mem', location.pathname)) { logger.info('パネル非表示ページ', location.pathname); return; }
    logger.info('初期化開始: mem');
    state.messages = storageGet('messages', []);
    state.sheetUrl = storageGet('sheetUrl', '');
    if (state.messages.length > 0) state.sheetStatus = 'success';
    createPanel();
    detectAllWithRetry();
    startObserver();
    startHealthCheck();
    state.initialized = true;
    window.__datingOps = { state, diag: runDiagnostic, stop: () => { state.stopped = true; }, rescan: detectAllWithRetry, insert: insertText, fetchSheet };
    logger.info('初期化完了');
  }

  // ============================================================
  // 起動
  // ============================================================
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
