/**
 * common.js - 共通コアモジュール
 * 監視・検出・UI・ログ・状態管理
 * 
 * 依存: sites.js が先に読み込まれていること
 */

(function() {
  'use strict';

  // ========================================
  // 定数
  // ========================================
  const PANEL_ID = 'dating-ops-panel';
  const STORAGE_PREFIX = 'datingOps_';
  const OBSERVER_DEBOUNCE_MS = 300;
  const HEALTH_CHECK_MS = 3000;
  const MAX_DETECT_RETRY = 5;
  const DETECT_RETRY_INTERVAL = 1000;

  // ========================================
  // 状態
  // ========================================
  const state = {
    siteType: null,
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
    textareaInfo: {
      status: 'unknown',
      element: null,
      selector: null,
      inIframe: false,
      reason: null,
    },
    panelPos: { x: 20, y: 20 },
    dragging: false,
    dragOffset: { x: 0, y: 0 },
  };

  // ========================================
  // ログ
  // ========================================
  const LogLevel = { INFO: 'info', WARN: 'warn', ERROR: 'error' };
  
  const LOG_STYLES = {
    info: 'background:#4a90d9;color:#fff;padding:2px 6px;border-radius:3px;',
    warn: 'background:#f6ad55;color:#000;padding:2px 6px;border-radius:3px;',
    error: 'background:#fc8181;color:#000;padding:2px 6px;border-radius:3px;',
  };

  function log(level, msg, data = null) {
    const tag = `DatingOps:${state.siteType || '?'}`;
    const style = LOG_STYLES[level] || '';
    if (data !== null) {
      console.log(`%c${tag}%c ${msg}`, style, '', data);
    } else {
      console.log(`%c${tag}%c ${msg}`, style, '');
    }
    if (level === LogLevel.ERROR && data instanceof Error) {
      console.error(data);
    }
  }

  const logger = {
    info: (m, d) => log(LogLevel.INFO, m, d),
    warn: (m, d) => log(LogLevel.WARN, m, d),
    error: (m, d) => log(LogLevel.ERROR, m, d),
  };

  // ========================================
  // ユーティリティ
  // ========================================
  
  function $(sel, ctx = document) {
    try {
      if (!ctx || typeof ctx.querySelector !== 'function') return null;
      return ctx.querySelector(sel);
    } catch (e) {
      return null;
    }
  }

  function $$(sel, ctx = document) {
    try {
      if (!ctx || typeof ctx.querySelectorAll !== 'function') return [];
      return Array.from(ctx.querySelectorAll(sel));
    } catch (e) {
      return [];
    }
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
    } catch (e) {
      logger.warn('localStorage保存失敗', e);
    }
  }

  function escapeHtml(s) {
    const d = document.createElement('div');
    d.textContent = s;
    return d.innerHTML;
  }

  // ========================================
  // テキストエリア検出
  // ========================================

  /**
   * 要素が除外対象かチェック
   */
  function isExcluded(el) {
    if (!el || !state.config) return true;
    
    for (const sel of state.config.excludeSelectors) {
      try {
        if (el.matches(sel)) return true;
      } catch {}
    }
    
    // data-type属性チェック
    const dataType = el.getAttribute('data-type');
    if (dataType && (dataType.includes('admin') || dataType.includes('memo'))) {
      return true;
    }
    
    // name/id属性に memo/admin が含まれるか
    const name = (el.name || '').toLowerCase();
    const id = (el.id || '').toLowerCase();
    if (name.includes('memo') || name.includes('admin') || name.includes('note')) return true;
    if (id.includes('memo') || id.includes('admin') || id.includes('note')) return true;
    
    return false;
  }

  /**
   * テキストエリアを検出
   */
  function detectTextarea() {
    const cfg = state.config;
    if (!cfg) {
      setTextareaInfo('error', null, null, false, '設定が読み込まれていません');
      return null;
    }

    // 1. 通常DOMから優先度順に検索
    for (const sel of cfg.textareaSelectors) {
      const els = $$(sel, document);
      for (const el of els) {
        if (!isExcluded(el)) {
          setTextareaInfo('found', el, sel, false, null);
          logger.info(`テキストエリア検出: ${sel}`);
          return el;
        }
      }
    }

    // 2. iframe内を検索（同一オリジンのみ）
    for (const iframeSel of cfg.iframeSelectors) {
      const iframes = $$(iframeSel, document);
      for (const iframe of iframes) {
        try {
          const doc = iframe.contentDocument || iframe.contentWindow?.document;
          if (!doc) continue;
          
          for (const sel of cfg.textareaSelectors) {
            const els = $$(sel, doc);
            for (const el of els) {
              if (!isExcluded(el)) {
                setTextareaInfo('found', el, sel, true, null);
                logger.info(`iframe内テキストエリア検出: ${iframeSel} > ${sel}`);
                return el;
              }
            }
          }
          
          // contentEditable body
          if (doc.body?.contentEditable === 'true') {
            setTextareaInfo('found', doc.body, 'body[contenteditable]', true, null);
            logger.info(`iframe内contentEditable検出: ${iframeSel}`);
            return doc.body;
          }
        } catch (e) {
          if (e.name === 'SecurityError') {
            logger.warn(`iframe クロスオリジン: ${iframeSel}`);
          }
        }
      }
    }

    // 見つからなかった
    const reason = `検索セレクタ: ${cfg.textareaSelectors.join(', ')}`;
    setTextareaInfo('not_found', null, null, false, reason);
    logger.warn('テキストエリア未検出', reason);
    return null;
  }

  function setTextareaInfo(status, el, sel, inIframe, reason) {
    state.textareaInfo = { status, element: el, selector: sel, inIframe, reason };
  }

  /**
   * テキストエリア検出をリトライ付きで実行
   */
  function detectTextareaWithRetry() {
    if (state.stopped) return;
    
    clearTimeout(state.detectTimer);
    state.detectRetryCount = 0;
    
    function attempt() {
      if (state.stopped) return;
      
      const el = detectTextarea();
      if (el) {
        state.detectRetryCount = 0;
        updateUI();
        return;
      }
      
      state.detectRetryCount++;
      if (state.detectRetryCount < MAX_DETECT_RETRY) {
        logger.info(`テキストエリア再検出 (${state.detectRetryCount}/${MAX_DETECT_RETRY})`);
        state.detectTimer = setTimeout(attempt, DETECT_RETRY_INTERVAL);
      } else {
        logger.warn('テキストエリア検出リトライ上限');
        updateUI();
      }
    }
    
    attempt();
  }

  // ========================================
  // テキスト挿入
  // ========================================

  /**
   * テキストを挿入（安全チェック付き）
   */
  function insertText(text) {
    const ta = state.textareaInfo.element;
    
    if (!ta) {
      showNotify('返信欄が見つかりません', 'error');
      logger.error('挿入失敗: テキストエリアなし');
      return false;
    }
    
    // 除外チェック
    if (isExcluded(ta)) {
      showNotify('この欄には挿入できません（管理者用）', 'error');
      logger.error('挿入失敗: 除外対象');
      return false;
    }
    
    // 既存テキストチェック
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

  // ========================================
  // Google Sheet
  // ========================================

  function normalizeSheetUrl(url) {
    if (!url) return null;
    const u = url.trim();
    
    if (u.includes('/export?format=csv') || u.includes('output=csv') || u.includes('/gviz/tq')) {
      return u;
    }
    
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
    return csv.split('\n')
      .map(line => {
        line = line.trim();
        if (!line) return null;
        if (line.startsWith('"')) {
          const end = line.indexOf('",');
          if (end > 0) return line.substring(1, end).replace(/""/g, '"');
          if (line.endsWith('"')) return line.slice(1, -1).replace(/""/g, '"');
        }
        const comma = line.indexOf(',');
        return comma > 0 ? line.substring(0, comma) : line;
      })
      .filter(Boolean);
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

  // ========================================
  // 通知
  // ========================================

  function showNotify(msg, type = 'info') {
    const old = $('#dating-ops-notify');
    if (old) old.remove();
    
    const colors = { info: '#4a90d9', success: '#48bb78', warn: '#f6ad55', error: '#fc8181' };
    
    const el = document.createElement('div');
    el.id = 'dating-ops-notify';
    Object.assign(el.style, {
      position: 'fixed',
      bottom: '20px',
      right: '20px',
      padding: '10px 16px',
      background: colors[type] || colors.info,
      color: '#fff',
      borderRadius: '6px',
      boxShadow: '0 4px 12px rgba(0,0,0,0.3)',
      zIndex: '2147483647',
      fontSize: '13px',
      fontFamily: 'system-ui, sans-serif',
      maxWidth: '280px',
    });
    el.textContent = msg;
    document.body.appendChild(el);
    
    setTimeout(() => {
      el.style.transition = 'opacity 0.3s';
      el.style.opacity = '0';
      setTimeout(() => el.remove(), 300);
    }, 3500);
  }

  // ========================================
  // パネルUI
  // ========================================

  function createStyles() {
    if ($('#dating-ops-styles')) return;
    
    const cfg = state.config;
    const primary = cfg?.ui?.primaryColor || '#4a90d9';
    const accent = cfg?.ui?.accentColor || '#2c5282';

    const css = `
      #${PANEL_ID} {
        position: fixed;
        width: 300px;
        max-height: 85vh;
        background: #1e1e2e;
        border: 1px solid ${primary};
        border-radius: 8px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.5);
        z-index: 2147483646;
        font-family: system-ui, sans-serif;
        font-size: 13px;
        color: #cdd6f4;
        overflow: hidden;
      }
      #${PANEL_ID} * { box-sizing: border-box; }
      #${PANEL_ID} .header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 8px 12px;
        background: linear-gradient(135deg, ${primary}, ${accent});
        cursor: move;
        user-select: none;
      }
      #${PANEL_ID} .header-title { font-weight: 600; color: #fff; }
      #${PANEL_ID} .header-close {
        width: 22px; height: 22px;
        display: flex; align-items: center; justify-content: center;
        background: rgba(255,255,255,0.2);
        border: none; border-radius: 4px;
        color: #fff; cursor: pointer; font-size: 14px;
      }
      #${PANEL_ID} .header-close:hover { background: rgba(255,255,255,0.3); }
      #${PANEL_ID} .body { padding: 10px; overflow-y: auto; max-height: calc(85vh - 40px); }
      #${PANEL_ID} .section { margin-bottom: 10px; }
      #${PANEL_ID} .section-title {
        font-size: 11px; font-weight: 600; color: #89b4fa;
        text-transform: uppercase; margin-bottom: 6px;
      }
      #${PANEL_ID} .status-box {
        background: #313244; border-radius: 6px; padding: 8px; font-size: 12px;
      }
      #${PANEL_ID} .status-row {
        display: flex; justify-content: space-between; margin-bottom: 4px;
      }
      #${PANEL_ID} .status-row:last-child { margin-bottom: 0; }
      #${PANEL_ID} .status-label { color: #6c7086; }
      #${PANEL_ID} .status-val { font-weight: 500; }
      #${PANEL_ID} .st-ok { color: #a6e3a1; }
      #${PANEL_ID} .st-warn { color: #f9e2af; }
      #${PANEL_ID} .st-err { color: #f38ba8; }
      #${PANEL_ID} .st-info { color: #89b4fa; }
      #${PANEL_ID} input, #${PANEL_ID} textarea {
        width: 100%; padding: 7px 9px;
        background: #313244; border: 1px solid #45475a;
        border-radius: 4px; color: #cdd6f4;
        font-size: 12px; font-family: inherit;
      }
      #${PANEL_ID} input:focus, #${PANEL_ID} textarea:focus {
        outline: none; border-color: ${primary};
      }
      #${PANEL_ID} textarea { resize: vertical; min-height: 60px; max-height: 120px; }
      #${PANEL_ID} .btn-row { display: flex; gap: 6px; flex-wrap: wrap; margin-top: 8px; }
      #${PANEL_ID} .btn {
        padding: 6px 10px; border: none; border-radius: 4px;
        font-size: 11px; font-weight: 500; cursor: pointer;
      }
      #${PANEL_ID} .btn-pri { background: ${primary}; color: #fff; }
      #${PANEL_ID} .btn-pri:hover { background: ${accent}; }
      #${PANEL_ID} .btn-sec { background: #45475a; color: #cdd6f4; }
      #${PANEL_ID} .btn-sec:hover { background: #585b70; }
      #${PANEL_ID} .btn-red { background: #f38ba8; color: #1e1e2e; }
      #${PANEL_ID} .btn-red:hover { background: #eba0ac; }
      #${PANEL_ID} .msg-list {
        max-height: 130px; overflow-y: auto;
        background: #313244; border-radius: 4px;
      }
      #${PANEL_ID} .msg-item {
        padding: 6px 8px; border-bottom: 1px solid #45475a;
        cursor: pointer; font-size: 11px;
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
      }
      #${PANEL_ID} .msg-item:last-child { border-bottom: none; }
      #${PANEL_ID} .msg-item:hover { background: #45475a; }
      #${PANEL_ID} .error-box {
        background: rgba(243,139,168,0.15); border: 1px solid #f38ba8;
        border-radius: 4px; padding: 8px; font-size: 11px; color: #f38ba8;
        margin-bottom: 10px; white-space: pre-wrap;
      }
      #${PANEL_ID} .help { font-size: 10px; color: #6c7086; margin-top: 4px; }
      #${PANEL_ID} .reason-box {
        background: #45475a; border-radius: 4px; padding: 6px;
        font-size: 10px; color: #f9e2af; margin-top: 6px;
      }
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
    
    const taStatusText = {
      unknown: '未確認',
      found: taInfo.inIframe ? 'iframe内OK' : '検出OK',
      not_found: '未検出',
      error: 'エラー',
    }[taInfo.status] || '不明';
    
    const taStatusClass = {
      unknown: 'st-warn',
      found: 'st-ok',
      not_found: 'st-warn',
      error: 'st-err',
    }[taInfo.status] || 'st-warn';
    
    const sheetStatusText = {
      idle: '未設定',
      loading: '読込中...',
      success: `${state.messages.length}件`,
      error: 'エラー',
    }[state.sheetStatus] || '不明';
    
    const sheetStatusClass = {
      idle: 'st-warn',
      loading: 'st-info',
      success: 'st-ok',
      error: 'st-err',
    }[state.sheetStatus] || 'st-warn';

    return `
      <div class="header">
        <span class="header-title">${escapeHtml(title)}</span>
        <button class="header-close" id="dp-close" title="閉じる">✕</button>
      </div>
      <div class="body">
        <div class="section">
          <div class="section-title">ステータス</div>
          <div class="status-box">
            <div class="status-row">
              <span class="status-label">サイト:</span>
              <span class="status-val st-info">${state.siteType?.toUpperCase() || '-'}</span>
            </div>
            <div class="status-row">
              <span class="status-label">返信欄:</span>
              <span class="status-val ${taStatusClass}" id="dp-ta-status">${taStatusText}</span>
            </div>
            <div class="status-row">
              <span class="status-label">Sheet:</span>
              <span class="status-val ${sheetStatusClass}" id="dp-sheet-status">${sheetStatusText}</span>
            </div>
          </div>
          ${taInfo.reason ? `<div class="reason-box" id="dp-ta-reason">${escapeHtml(taInfo.reason)}</div>` : ''}
        </div>
        
        ${state.sheetError ? `<div class="error-box" id="dp-error">${escapeHtml(state.sheetError)}</div>` : ''}
        
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
          <div class="btn-row">
            <button class="btn btn-pri" id="dp-apply-direct">適用</button>
          </div>
        </div>
        
        <div class="section">
          <div class="section-title">メッセージ一覧 (<span id="dp-msg-count">${state.messages.length}</span>件)</div>
          <div class="msg-list" id="dp-msg-list">
            ${state.messages.length > 0 
              ? state.messages.map((m, i) => `<div class="msg-item" data-idx="${i}" title="${escapeHtml(m)}">${escapeHtml(m.substring(0, 45))}${m.length > 45 ? '...' : ''}</div>`).join('')
              : '<div class="msg-item" style="color:#6c7086;">メッセージなし</div>'
            }
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
    // 既存パネル削除（重複防止）
    const old = $(`#${PANEL_ID}`);
    if (old) old.remove();
    
    createStyles();
    
    const panel = document.createElement('div');
    panel.id = PANEL_ID;
    panel.innerHTML = createPanelHtml();
    
    // 位置復元
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
    
    // ドラッグ
    const header = panel.querySelector('.header');
    if (header) {
      header.addEventListener('mousedown', onDragStart);
    }
    
    // 閉じる
    const closeBtn = panel.querySelector('#dp-close');
    if (closeBtn) {
      closeBtn.addEventListener('click', () => {
        panel.style.display = 'none';
        logger.info('パネルを閉じました');
      });
    }
    
    // Sheet読込
    const loadBtn = panel.querySelector('#dp-load-sheet');
    if (loadBtn) {
      loadBtn.addEventListener('click', () => {
        const url = panel.querySelector('#dp-sheet-url')?.value?.trim();
        if (url) fetchSheet(url);
        else showNotify('URLを入力してください', 'warn');
      });
    }
    
    // 再検出
    const rescanBtn = panel.querySelector('#dp-rescan');
    if (rescanBtn) {
      rescanBtn.addEventListener('click', () => {
        detectTextareaWithRetry();
        panel.style.display = '';
        showNotify('再検出開始', 'info');
      });
    }
    
    // 直接入力
    const applyBtn = panel.querySelector('#dp-apply-direct');
    if (applyBtn) {
      applyBtn.addEventListener('click', () => {
        const text = panel.querySelector('#dp-direct')?.value?.trim();
        if (!text) {
          showNotify('テキストを入力してください', 'warn');
          return;
        }
        const msgs = text.split('\n').filter(l => l.trim());
        if (msgs.length === 0) {
          showNotify('有効なメッセージがありません', 'warn');
          return;
        }
        state.messages = msgs;
        storageSet('messages', msgs);
        showNotify(`${msgs.length}件適用`, 'success');
        updateUI();
      });
    }
    
    // 停止
    const stopBtn = panel.querySelector('#dp-stop');
    if (stopBtn) {
      stopBtn.addEventListener('click', () => {
        state.stopped = true;
        if (state.observer) {
          state.observer.disconnect();
          state.observer = null;
        }
        if (state.healthTimer) {
          clearInterval(state.healthTimer);
          state.healthTimer = null;
        }
        clearTimeout(state.detectTimer);
        logger.info('完全停止');
        showNotify('監視を停止しました', 'warn');
        updateUI();
      });
    }
    
    // 診断
    const diagBtn = panel.querySelector('#dp-diag');
    if (diagBtn) {
      diagBtn.addEventListener('click', runDiagnostic);
    }
    
    // メッセージリストクリック
    const msgList = panel.querySelector('#dp-msg-list');
    if (msgList) {
      msgList.addEventListener('click', (e) => {
        const item = e.target.closest('.msg-item');
        if (!item) return;
        const idx = parseInt(item.dataset.idx, 10);
        if (!isNaN(idx) && state.messages[idx]) {
          insertText(state.messages[idx]);
        }
      });
    }
    
    // グローバルドラッグイベント
    document.addEventListener('mousemove', onDragMove);
    document.addEventListener('mouseup', onDragEnd);
  }

  function onDragStart(e) {
    if (e.target.closest('.header-close')) return;
    state.dragging = true;
    state.dragOffset = {
      x: e.clientX - state.panelPos.x,
      y: e.clientY - state.panelPos.y,
    };
    if (state.panel) state.panel.style.transition = 'none';
  }

  function onDragMove(e) {
    if (!state.dragging || !state.panel) return;
    const x = Math.max(0, Math.min(window.innerWidth - 300, e.clientX - state.dragOffset.x));
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
    
    // 再生成してイベント再設定
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
        inIframe: state.textareaInfo.inIframe,
        reason: state.textareaInfo.reason,
        hasElement: !!state.textareaInfo.element,
      },
      messages: state.messages.length,
      sheetStatus: state.sheetStatus,
      sheetError: state.sheetError,
      url: location.href,
    };
    
    console.group('%c[DatingOps] 診断情報', 'color:#4a90d9;font-weight:bold');
    console.log(diag);
    console.log('設定:', state.config);
    console.groupEnd();
    
    navigator.clipboard?.writeText(JSON.stringify(diag, null, 2))
      .then(() => showNotify('診断情報をコピーしました', 'info'))
      .catch(() => showNotify('診断情報をコンソールに出力しました', 'info'));
    
    return diag;
  }

  // ========================================
  // MutationObserver
  // ========================================

  function startObserver() {
    if (state.observer) state.observer.disconnect();
    
    const handler = debounce(() => {
      if (state.stopped) return;
      ensurePanel();
      detectTextareaWithRetry();
    }, OBSERVER_DEBOUNCE_MS);
    
    state.observer = new MutationObserver((muts) => {
      // 自分自身の変更は無視
      const isOwn = muts.some(m => 
        m.target.id === PANEL_ID || m.target.closest?.(`#${PANEL_ID}`)
      );
      if (isOwn) return;
      handler();
    });
    
    state.observer.observe(document.body, { childList: true, subtree: true });
    logger.info('DOM監視開始');
  }

  function ensurePanel() {
    if (state.stopped) return;
    const panel = $(`#${PANEL_ID}`);
    if (!panel) {
      logger.warn('パネル消失、再作成');
      createPanel();
    } else {
      state.panel = panel;
    }
  }

  function startHealthCheck() {
    if (state.healthTimer) clearInterval(state.healthTimer);
    
    state.healthTimer = setInterval(() => {
      if (state.stopped) {
        clearInterval(state.healthTimer);
        return;
      }
      ensurePanel();
    }, HEALTH_CHECK_MS);
  }

  // ========================================
  // 初期化
  // ========================================

  function init(siteType) {
    if (state.initialized) {
      logger.warn('既に初期化済み');
      return;
    }
    
    // sites.js 確認
    if (!window.DatingOpsSites) {
      logger.error('sites.js が読み込まれていません');
      return;
    }
    
    state.siteType = siteType;
    state.config = window.DatingOpsSites.getConfig(siteType);
    
    if (!state.config) {
      logger.error(`設定が見つかりません: ${siteType}`);
      return;
    }
    
    // URL確認
    if (!window.DatingOpsSites.shouldShowPanel(siteType, location.pathname)) {
      logger.info('パネル非表示ページ', location.pathname);
      return;
    }
    
    logger.info(`初期化開始: ${siteType}`);
    
    // 保存データ復元
    state.messages = storageGet('messages', []);
    state.sheetUrl = storageGet('sheetUrl', '');
    if (state.messages.length > 0) state.sheetStatus = 'success';
    
    // パネル作成
    createPanel();
    
    // テキストエリア検出
    detectTextareaWithRetry();
    
    // 監視開始
    startObserver();
    startHealthCheck();
    
    state.initialized = true;
    
    // グローバルAPI
    window.__datingOps = {
      state,
      diag: runDiagnostic,
      stop: () => { state.stopped = true; logger.info('停止'); },
      rescan: detectTextareaWithRetry,
      insert: insertText,
      fetchSheet,
    };
    
    logger.info('初期化完了');
  }

  // グローバル公開
  window.DatingOpsCore = { init };

})();
