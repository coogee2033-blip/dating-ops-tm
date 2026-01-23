// ==UserScript==
// @name         Dating Ops - MEM
// @namespace    dating-ops-tm
// @version      2.4.5
// @description  MEM (mem44.com) 返信アシスト - 返信欄ベース行検出 + 一括挿入機能
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
 * 実装メモ (v2.4.5)
 * ============================================================
 * 
 * 【v2.4.4 主な変更】
 *   - 行検出を「返信欄（textarea）数ベース」に変更
 *   - findBatchReplyTextareas() で表示中の返信欄を検出
 *   - 一括挿入ボタン「表示分に一括挿入」を追加
 *   - rows.count は batchTextareasCount を最優先に使用
 *   - mailbox iframe より box_char/複数textarea を優先
 *   - MESSAGE1_TEXTAREA_SELECTOR を拡張（name/id含む）
 *   - href も取得し src 空/別名でも判定可能に
 *   - looksLikeRowReplyTextarea を緩和（検出漏れ防止）
 *   - 表示件数は全docから最大値を採用
 *   - docCandidates で各doc診断情報を可視化
 * 
 * 【iframe対応】
 *   - getAccessibleDocs(): main + 同一オリジンiframe(ネスト含む)
 *   - 行検出/テーブルスコアリングをdoc横断で実行
 * 
 * 【返信欄検出】
 *   - textarea[name="message1"] 最優先
 *   - memo/admin/note/futari/two_memo/staff 除外
 *   - 既存テキストがあれば上書き禁止
 * ============================================================
 */

(function() {
  'use strict';

  // ============================================================
  // 定数
  // ============================================================
  const SITE_TYPE = 'mem';
  const PANEL_ID = 'dating-ops-panel';
  const STORAGE_PREFIX = 'datingOps_mem_';
  const OBSERVER_DEBOUNCE_MS = 300;
  const HEALTH_CHECK_MS = 3000;
  const URL_POLL_MS = 500;
  const STAGED_DETECT_DELAYS = [0, 300, 800, 1500];
  const MIN_TABLE_SCORE = 5;

  // ============================================================
  // サイト設定
  // ============================================================
  const CONFIG = {
    pageModes: {
      list: [/\/staff\/index/, /\/staff\/.*list/, /\/staff\/.*inbox/],
      personal: [/\/staff\/personalbox/, /\/staff\/personal/, /\/staff\/.*detail/],
    },
    textareaSelectors: [
      'textarea[name="message1"]',
      'textarea.msg.wd100',
      'textarea.msg',
      'textarea[name="message"]',
      'textarea[name="body"]',
      'textarea[name="mail_body"]',
      'textarea[name="content"]',
    ],
    excludeKeywords: ['memo', 'admin', 'note', 'futari', 'two_memo', 'staff'],
    excludeClassKeywords: ['admin', 'memo', 'note'],
    rowSelectors: ['tr.rowitem', 'tr[id^="row"]', 'tr[class*="row"]', 'tbody > tr', 'tr'],
    fallbackRowSelector: 'td.chatview',
    countPatterns: [/表示\s*(\d+)\s*件/, /(\d+)\s*件/, /全\s*(\d+)/],
    listHeaderKeywords: ['キャラ情報', '会話', '状況', 'ふたりメモ', 'ユーザー情報', '更新', '名前', '返信', 'ステータス', 'チェック'],
    clickTriggerSelectors: [
      'input[type="submit"][value*="更新"]', 'input[value="更新"]', '.reload-btn',
      'input[type="submit"][value*="切替"]', 'input[type="submit"][value*="表示"]',
      '.pagination a', '.paging a', 'a[href*="page="]',
      'input[value*="一括"]', 'input[value*="チェック"]',
    ],
    changeTriggerSelectors: ['select[name*="box"]', 'select[name*="folder"]', 'select[name*="type"]'],
    ui: { title: 'MEM 返信アシスト', primaryColor: '#48bb78', accentColor: '#276749' },
  };

  // ============================================================
  // 状態
  // ============================================================
  const state = {
    initialized: false,
    stopped: false,
    panel: null,
    observers: [],
    healthTimer: null,
    urlPollTimer: null,
    stagedDetectIndex: 0,
    stagedDetectTimer: null,
    lastUrl: '',
    pageMode: 'other',
    textarea: { status: 'pending', element: null, selector: null, where: null, iframeSrc: null },
    rows: {
      count: 0,
      batchTextareasCount: 0,
      displayCountNum: null,
      scopeHint: null,
      usedSelector: null,
      mismatchWarning: null,
      tableScore: null,
      where: null,
      iframeSrc: null,
      docsSearched: 0,
    },
    lastDetectTime: null,
    lastTrigger: null,
    messages: [],
    sheetUrl: '',
    sheetStatus: 'idle',
    sheetError: null,
    panelPos: { x: 20, y: 20 },
    dragging: false,
    dragOffset: { x: 0, y: 0 },
    observeTargetHints: [],
    iframeLoadListeners: new WeakSet(),
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
  // アクセス可能なドキュメント一覧取得 (再帰的iframe探索)
  // ============================================================
  const debouncedReInit = debounce(() => {
    if (state.stopped) return;
    logger.info('iframe load → 再初期化');
    startObservers();
    runStagedDetection();
  }, 500);

  function getAccessibleDocs() {
    const docs = [];
    const seenDocs = new WeakSet();

    function addDoc(doc, win, where, iframeSrc) {
      if (!doc || seenDocs.has(doc)) return;
      seenDocs.add(doc);
      // href も取得（iframeSrc が空/別名でも location.href で判別可能にする）
      let href = null;
      try { href = win.location.href; } catch { /* cross-origin */ }
      docs.push({ doc, win, where, iframeSrc, href });
    }

    function scanDocForIframes(baseDoc) {
      let iframes = [];
      try { iframes = Array.from(baseDoc.querySelectorAll('iframe')); } catch { iframes = []; }
      for (const iframe of iframes) {
        if (!state.iframeLoadListeners.has(iframe)) {
          state.iframeLoadListeners.add(iframe);
          iframe.addEventListener('load', debouncedReInit, { once: true });
        }
        try {
          const doc = iframe.contentDocument;
          const win = iframe.contentWindow;
          if (doc && win) {
            const src = iframe.src || iframe.name || iframe.id || 'iframe';
            addDoc(doc, win, 'iframe', src);
            scanDocForIframes(doc);
          }
        } catch { /* cross-origin */ }
      }
    }

    addDoc(document, window, 'main', null);
    scanDocForIframes(document);
    return docs;
  }

  // ============================================================
  // ページモード判定
  // ============================================================
  function detectPageMode() {
    const path = location.pathname;
    for (const pattern of CONFIG.pageModes.list) { if (pattern.test(path)) return 'list'; }
    for (const pattern of CONFIG.pageModes.personal) { if (pattern.test(path)) return 'personal'; }
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
    for (const kw of CONFIG.excludeKeywords) { if (name.includes(kw) || id.includes(kw)) return true; }
    for (const kw of CONFIG.excludeClassKeywords) { if (cls.includes(kw)) return true; }
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
  // テーブルスコアリング
  // ============================================================
  function getAllTables(doc) { return $$('table', doc || document); }

  function scoreListTable(table) {
    if (!table) return 0;
    let score = 0;
    const tbody = $('tbody', table) || table;
    const rows = $$('tr', tbody).filter(tr => !$('th', tr));
    const rowCount = rows.length;
    score += Math.min(rowCount * 1.5, 30);
    if (rowCount < 3) score -= 20;
    const checkboxes = $$('input[type="checkbox"]', table);
    if (checkboxes.length >= 1) score += 10;
    if (checkboxes.length >= 5) score += 5;
    if ($('td.chatview', table)) score += 15;
    if ($('tr.rowitem', table)) score += 10;
    const headerText = ($$('th', table).map(th => th.textContent || '').join(' ') +
                        $$('thead td', table).map(td => td.textContent || '').join(' ')).toLowerCase();
    for (const kw of CONFIG.listHeaderKeywords) { if (headerText.includes(kw.toLowerCase())) { score += 5; break; } }
    if (table.parentElement?.tagName === 'FORM') score -= 10;
    try { const rect = table.getBoundingClientRect(); if (rect.width < 300) score -= 15; } catch {}
    return score;
  }

  function findBestListTableInDoc(doc) {
    const tables = getAllTables(doc);
    let best = null, bestScore = -Infinity;
    for (const table of tables) {
      const score = scoreListTable(table);
      if (score > bestScore) { bestScore = score; best = table; }
    }
    if (best && bestScore >= MIN_TABLE_SCORE) {
      const tbody = $('tbody', best) || best;
      return { table: best, tbody, score: bestScore, hint: `bestTable(score=${bestScore})` };
    }
    return null;
  }

  // ============================================================
  // 一括返信欄検出 (v2.4.5 大幅改修)
  // ============================================================

  /**
   * 単体送信フォーム内のtextareaかどうか判定
   * （mailbox等の送信フォーム内textareaを除外するためのフィルタ）
   */
  function isInsideSingleSendForm(ta) {
    if (!ta) return false;
    const form = ta.closest('form');
    if (!form) return false;
    // フォーム内に送信ボタンがあり、かつtextareaが1つだけ → 単体送信フォームの可能性
    const btns = form.querySelectorAll('input[type="submit"], button[type="submit"], input[value*="送信"], button');
    const textareas = form.querySelectorAll('textarea');
    return btns.length > 0 && textareas.length === 1;
  }

  /**
   * textarea の「返信欄らしさ」をスコアリング
   * スコアが高いほど返信欄として採用されやすい
   */
  function scoreReplyTextareaCandidate(ta) {
    if (!ta) return -100;
    let s = 0;
    const name = (ta.name || '').toLowerCase();
    const id = (ta.id || '').toLowerCase();
    const cls = (typeof ta.className === 'string' ? ta.className : '').toLowerCase();
    const ph = (ta.placeholder || '').toLowerCase();

    // 返信欄っぽいワードを加点
    if (name.includes('message') || name.includes('msg') || name.includes('body') || name.includes('content')) s += 5;
    if (id.includes('message') || id.includes('msg') || id.includes('body')) s += 3;
    if (cls.includes('msg') || cls.includes('message') || cls.includes('reply') || cls.includes('content')) s += 3;
    if (ph.includes('メッセージ') || ph.includes('送信') || ph.includes('返信') || ph.includes('入力')) s += 2;

    // message1 は特に優先
    if (name.includes('message1') || id.includes('message1')) s += 10;

    // memo/admin/note 系は大減点（保険）
    if (name.includes('memo') || name.includes('note') || name.includes('admin') || name.includes('staff')) s -= 15;
    if (id.includes('memo') || id.includes('note') || id.includes('admin')) s -= 15;
    if (cls.includes('memo') || cls.includes('note') || cls.includes('admin')) s -= 10;

    // サイズが大きいほど加点（返信欄は大きい想定）
    try {
      const cols = ta.cols || 0;
      const rows = ta.rows || 0;
      s += Math.min(3, Math.floor(cols / 20));
      s += Math.min(2, Math.floor(rows / 3));
    } catch {}

    return s;
  }

  /**
   * 行(tr)から最も返信欄らしい textarea を1つ選ぶ
   */
  function findReplyTextareaInRow(tr, win) {
    if (!tr) return null;
    const textareas = tr.querySelectorAll('textarea');
    if (textareas.length === 0) return null;

    let bestTa = null;
    let bestScore = -Infinity;

    for (const ta of textareas) {
      if (ta.tagName !== 'TEXTAREA') continue;
      if (isExcludedTextarea(ta)) continue;
      if (!isElementVisible(ta, win)) continue;
      if (isInsideSingleSendForm(ta)) continue;

      const score = scoreReplyTextareaCandidate(ta);
      if (score > bestScore) {
        bestScore = score;
        bestTa = ta;
      }
    }

    return bestTa;
  }

  /**
   * textarea のサンプル情報を取得（診断用）
   */
  function getTextareaSample(ta) {
    if (!ta) return null;
    const name = (ta.name || '').slice(0, 30);
    const id = (ta.id || '').slice(0, 30);
    const cls = (typeof ta.className === 'string' ? ta.className : '').slice(0, 50);
    return { name, id, class: cls };
  }

  /**
   * 指定doc内のスコープから返信欄を収集
   * v2.4.5: message1固定ではなく全textareaをスコアリングして選ぶ
   */
  function findBatchReplyTextareasInDoc(doc, win, isBoxCharDoc) {
    const resultFromTable = [];
    const seen = new Set();

    const best = findBestListTableInDoc(doc);
    const scopeTbody = best ? (best.tbody || $('tbody', best.table) || best.table) : null;
    const tableScore = best ? best.score : null;

    // (1) bestTableベース: 各行から最も返信欄らしい textarea を1つ拾う
    if (scopeTbody) {
      const rows = $$('tr', scopeTbody).filter(tr => !$('th', tr));
      let rowIndex = 0;
      for (const tr of rows) {
        const ta = findReplyTextareaInRow(tr, win);
        if (!ta) { rowIndex++; continue; }
        if (seen.has(ta)) { rowIndex++; continue; }
        seen.add(ta);
        resultFromTable.push({
          el: ta,
          selector: 'scored-textarea',
          rowIndex,
          scopeHint: best ? best.hint : 'bestTable',
          score: scoreReplyTextareaCandidate(ta),
        });
        rowIndex++;
      }
    }

    // (2) doc全体スキャン（bestTableで0件かつ isBoxCharDoc のときのみ）
    const resultFromDoc = [];
    if (resultFromTable.length === 0 && isBoxCharDoc) {
      const seen2 = new Set();
      const allTas = $$('textarea', doc.body || doc);
      let idx = 0;
      for (const ta of allTas) {
        if (ta.tagName !== 'TEXTAREA') continue;
        if (isExcludedTextarea(ta)) continue;
        if (!isElementVisible(ta, win)) continue;
        if (isInsideSingleSendForm(ta)) continue;
        if (seen2.has(ta)) continue;
        seen2.add(ta);
        resultFromDoc.push({
          el: ta,
          selector: 'doc-wide-scored',
          rowIndex: idx++,
          scopeHint: 'doc:scored',
          score: scoreReplyTextareaCandidate(ta),
        });
      }
    }

    // 採用ルール: table由来を優先、0件ならdoc全体
    let chosen = resultFromTable.length > 0 ? resultFromTable : resultFromDoc;
    let scopeHint = resultFromTable.length > 0 ? (best ? best.hint : 'bestTable') : 'doc:scored';

    // 診断用: 全textarea数とサンプル
    const allTextareas = $$('textarea', doc.body || doc);
    const allTextareasCount = allTextareas.length;
    const sampleTextareas = allTextareas.slice(0, 5).map(getTextareaSample);

    return {
      textareas: chosen,
      scopeHint,
      tableScore,
      totalCandidates: chosen.length,
      allTextareasCount,
      sampleTextareas,
    };
  }

  /**
   * 指定doc内の表示件数を抽出
   */
  function extractDisplayCountFromDoc(doc) {
    try {
      const bodyText = doc.body ? doc.body.innerText : '';
      for (const pat of CONFIG.countPatterns) {
        const m = bodyText.match(pat);
        if (m && m[1]) {
          const n = parseInt(m[1], 10);
          if (!isNaN(n) && n >= 0 && n < 10000) return n;
        }
      }
    } catch {}
    return null;
  }

  /**
   * mailbox iframe かどうか判定（listモードでは最下位扱い）
   */
  /**
   * mailbox 系の判定（単体返信用、listモードでは最下位）
   * iframeSrc だけでなく href でも判定
   */
  function isMailboxLike(srcOrHref) {
    if (!srcOrHref) return false;
    const s = String(srcOrHref).toLowerCase();
    return s.includes('/staff/mailbox') && (s.includes('method=empty') || s.includes('method=send'));
  }

  /**
   * box_char 系の判定（一覧テーブル用iframe）
   * iframeSrc だけでなく href でも判定
   */
  function isBoxCharLike(srcOrHref) {
    if (!srcOrHref) return false;
    const s = String(srcOrHref).toLowerCase();
    return s.includes('box_char') || s.includes('/staff/box_char');
  }

  /**
   * docInfo から mailbox/boxChar を判定（src と href 両方チェック）
   */
  function isMailboxDoc(docInfo) {
    return isMailboxLike(docInfo.iframeSrc) || isMailboxLike(docInfo.href);
  }
  function isBoxCharDoc(docInfo) {
    return isBoxCharLike(docInfo.iframeSrc) || isBoxCharLike(docInfo.href);
  }

  /**
   * 全アクセス可能docから一括返信欄を検出
   * 
   * v2.4.4 優先順位:
   * 1) chosenCount > 0 を優先
   * 2) isMailbox は最下位（chosenCount>=2 の非mailboxが1つでもあれば後ろ）
   * 3) isBoxChar を優先
   * 4) chosenCount が大きい方
   * 5) displayCount との差が小さい方（extractDisplayCount の最大値ロジック後の値）
   * 6) tableScore が高い方
   */
  function findBatchReplyTextareas() {
    const docs = getAccessibleDocs();
    const globalDisplayCount = extractDisplayCount(); // 最大値ロジック適用済み

    // 各docのスコアと返信欄情報を収集
    const docResults = [];
    for (const { doc, win, where, iframeSrc, href } of docs) {
      const isMailbox = isMailboxDoc({ iframeSrc, href });
      const isBoxChar = isBoxCharDoc({ iframeSrc, href });
      // isBoxCharDoc を渡して doc全体スキャンの許可判定に使う
      const res = findBatchReplyTextareasInDoc(doc, win, isBoxChar);
      const displayCountLocal = extractDisplayCountFromDoc(doc);
      docResults.push({
        doc, win, where, iframeSrc, href,
        textareas: res.textareas,
        chosenCount: res.textareas.length,
        totalCandidates: res.totalCandidates,
        allTextareasCount: res.allTextareasCount,
        sampleTextareas: res.sampleTextareas,
        scopeHint: res.scopeHint,
        tableScore: res.tableScore,
        displayCountLocal,
        isMailbox,
        isBoxChar,
      });
    }

    /**
     * ソート優先順位 (v2.4.5):
     * (1) isBoxChar && chosenCount > 0 を最優先（box_char が1件でもあれば mailbox より先）
     * (2) chosenCount > 0 を優先
     * (3) mailbox は最下位（box_char が取れていれば絶対に採用しない）
     * (4) chosenCount が大きい方を優先
     * (5) globalDisplayCount との差が小さい方を優先
     * (6) tableScore が高い方
     */
    docResults.sort((a, b) => {
      const aCount = a.chosenCount;
      const bCount = b.chosenCount;

      // (1) isBoxChar && chosenCount > 0 を最優先
      const aBoxCharWithCount = a.isBoxChar && aCount > 0;
      const bBoxCharWithCount = b.isBoxChar && bCount > 0;
      if (aBoxCharWithCount && !bBoxCharWithCount) return -1;
      if (!aBoxCharWithCount && bBoxCharWithCount) return 1;

      // (2) chosenCount > 0 を優先
      if (aCount > 0 && bCount === 0) return -1;
      if (aCount === 0 && bCount > 0) return 1;
      if (aCount === 0 && bCount === 0) return 0;

      // (3) mailbox は最下位（box_char以外でも非mailboxを優先）
      if (!a.isMailbox && b.isMailbox) return -1;
      if (a.isMailbox && !b.isMailbox) return 1;

      // (4) chosenCount が大きい方を優先
      if (aCount !== bCount) return bCount - aCount;

      // (5) globalDisplayCount との差が小さい方を優先
      const aDiff = globalDisplayCount !== null ? Math.abs(aCount - globalDisplayCount) : Infinity;
      const bDiff = globalDisplayCount !== null ? Math.abs(bCount - globalDisplayCount) : Infinity;
      if (aDiff !== bDiff) return aDiff - bDiff;

      // (6) tableScore が高い方
      const aScore = a.tableScore ?? -Infinity;
      const bScore = b.tableScore ?? -Infinity;
      return bScore - aScore;
    });

    // box_char で chosenCount > 0 があれば無条件でそれを採用
    // なければ最初に chosenCount > 0 の doc を採用
    for (const dr of docResults) {
      if (dr.chosenCount > 0) {
        // scopeHint にどのdocを選んだか情報を追加
        const hint = dr.scopeHint + (dr.isMailbox ? ' [mailbox]' : '') + (dr.isBoxChar ? ' [box_char]' : '');
        return {
          textareas: dr.textareas,
          count: dr.chosenCount,
          where: dr.where,
          iframeSrc: dr.iframeSrc,
          href: dr.href,
          scopeHint: hint,
          tableScore: dr.tableScore,
          docsSearched: docs.length,
          totalCandidates: dr.totalCandidates,
          allTextareasCount: dr.allTextareasCount,
        };
      }
    }

    return { textareas: [], count: 0, where: null, iframeSrc: null, href: null, scopeHint: 'none', tableScore: null, docsSearched: docs.length, totalCandidates: 0, allTextareasCount: 0 };
  }

  // ============================================================
  // 表示件数抽出
  // ============================================================
  /**
   * 全docsから表示件数を抽出し、最大値を返す
   * v2.4.4: mailbox=10, list=21 なら 21 を採用。0〜3は誤爆しやすいので除外。
   */
  function extractDisplayCount() {
    const docs = getAccessibleDocs();
    const candidates = [];
    for (const { doc } of docs) {
      const n = extractDisplayCountFromDoc(doc);
      // 4以上のみ候補にする（0〜3は誤爆しやすい）
      if (n !== null && n >= 4) {
        candidates.push(n);
      }
    }
    if (candidates.length === 0) {
      // 4未満でも何か取れていればそれを返す（fallback）
      for (const { doc } of docs) {
        const n = extractDisplayCountFromDoc(doc);
        if (n !== null) return n;
      }
      return null;
    }
    // 最大値を返す
    return Math.max(...candidates);
  }

  // ============================================================
  // 行検出 (v2.4.0: 返信欄ベース最優先)
  // ============================================================
  function detectRows() {
    if (state.pageMode !== 'list') {
      state.rows = {
        count: state.pageMode === 'personal' ? 1 : 0,
        batchTextareasCount: 0,
        displayCountNum: null,
        scopeHint: 'N/A (not list mode)',
        usedSelector: null,
        mismatchWarning: null,
        tableScore: null,
        where: null,
        iframeSrc: null,
        docsSearched: 0,
      };
      return;
    }

    const dispNum = extractDisplayCount();
    const batchResult = findBatchReplyTextareas();

    if (batchResult.count > 0) {
      let warn = null;
      if (dispNum !== null && dispNum > 0) {
        if (batchResult.count === 0) {
          warn = '返信欄未描画または検出失敗';
        } else if (Math.abs(batchResult.count - dispNum) >= 5) {
          warn = 'ページング/折り畳み/スクロール外の可能性';
        }
      }

      state.rows = {
        count: batchResult.count,
        batchTextareasCount: batchResult.count,
        totalCandidates: batchResult.totalCandidates,
        displayCountNum: dispNum,
        scopeHint: `batchTextareas(${batchResult.scopeHint})`,
        usedSelector: 'textarea-batch',
        mismatchWarning: warn,
        tableScore: batchResult.tableScore,
        where: batchResult.where,
        iframeSrc: batchResult.iframeSrc,
        href: batchResult.href,
        docsSearched: batchResult.docsSearched,
      };
      logger.info(`[Rows] 返信欄ベース: ${batchResult.count}件 (${batchResult.where})`);
      return;
    }

    const docs = getAccessibleDocs();
    for (const { doc, where, iframeSrc } of docs) {
      const best = findBestListTableInDoc(doc);
      if (best) {
        const rows = $$('tr', best.tbody).filter(tr => !$('th', tr));
        if (rows.length > 0) {
          let warn = '返信欄未検出 (trカウントfallback)';
          if (dispNum !== null && dispNum > 0 && Math.abs(rows.length - dispNum) >= 5) {
            warn = 'scope誤認の可能性 (trカウントfallback)';
          }
          state.rows = {
            count: rows.length,
            batchTextareasCount: 0,
            displayCountNum: dispNum,
            scopeHint: `fallback-tr(${best.hint})`,
            usedSelector: 'tbody > tr (fallback)',
            mismatchWarning: warn,
            tableScore: best.score,
            where, iframeSrc,
            docsSearched: docs.length,
          };
          logger.warn(`[Rows] fallback: tr=${rows.length}件 (返信欄は0)`);
          return;
        }
      }
    }

    state.rows = {
      count: 0,
      batchTextareasCount: 0,
      displayCountNum: dispNum,
      scopeHint: 'not found',
      usedSelector: null,
      mismatchWarning: dispNum !== null && dispNum > 0 ? '返信欄未描画または検出失敗' : null,
      tableScore: null,
      where: null,
      iframeSrc: null,
      docsSearched: docs.length,
    };
    logger.warn(`[Rows] 全${docs.length}doc探索したが返信欄0`);
  }

  // ============================================================
  // 返信欄検出 (単体用)
  // ============================================================
  function findReplyTextarea() {
    const docs = getAccessibleDocs();
    for (const { doc, win, where, iframeSrc } of docs) {
      for (const sel of CONFIG.textareaSelectors) {
        for (const el of $$(sel, doc)) {
          if (!isExcludedTextarea(el) && isElementVisible(el, win)) {
            return { element: el, selector: sel, where, iframeSrc };
          }
        }
      }
      try { if (doc.body?.contentEditable === 'true') return { element: doc.body, selector: 'body[contenteditable]', where, iframeSrc }; } catch {}
    }
    return null;
  }

  function detectTextarea() {
    const res = findReplyTextarea();
    if (res) { state.textarea = { status: 'found', element: res.element, selector: res.selector, where: res.where, iframeSrc: res.iframeSrc }; return true; }
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
    logger.info(`検出 [${trigger}] mode=${state.pageMode} rows=${state.rows.count}(${state.rows.usedSelector || '-'}) ta=${state.textarea.status}`);
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
        if (state.pageMode === 'list' && state.rows.count === 0 && state.stagedDetectIndex < STAGED_DETECT_DELAYS.length) next();
      }, delay);
    }
    next();
  }

  // ============================================================
  // テキスト挿入（単体用）
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
    } catch { showNotify('挿入失敗', 'error'); return false; }
  }

  // ============================================================
  // テキスト挿入（要素指定、一括用）
  // ============================================================
  function insertTextToElement(el, text) {
    if (!el || el.tagName !== 'TEXTAREA') return false;
    if (isExcludedTextarea(el)) return false;
    const cur = el.value || '';
    if (cur.length > 0) return false;
    try {
      el.value = text;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    } catch { return false; }
  }

  // ============================================================
  // 一括挿入 (v2.4.0 新規)
  // ============================================================
  function batchInsertVisible() {
    if (state.pageMode !== 'list') {
      showNotify('一括送信モード以外では使用できません', 'warn');
      return;
    }
    if (state.messages.length === 0) {
      showNotify('メッセージがありません', 'warn');
      return;
    }

    const batchResult = findBatchReplyTextareas();
    if (batchResult.count === 0) {
      showNotify('表示中の返信欄が見つかりません', 'warn');
      return;
    }

    let inserted = 0;
    let skipped = 0;
    let msgIndex = 0;
    let lastInsertedEl = null;

    for (const item of batchResult.textareas) {
      if (msgIndex >= state.messages.length) break;
      const el = item.el;
      const text = state.messages[msgIndex];
      if (insertTextToElement(el, text)) {
        inserted++;
        lastInsertedEl = el;
        msgIndex++;
      } else {
        skipped++;
      }
    }

    if (lastInsertedEl) {
      try { lastInsertedEl.focus(); } catch {}
    }

    const remaining = state.messages.length - msgIndex;
    showNotify(`挿入${inserted}件 / スキップ${skipped}件 / 残${remaining}件`, inserted > 0 ? 'success' : 'warn');
    logger.info(`[BatchInsert] 挿入=${inserted}, スキップ=${skipped}, 残メッセージ=${remaining}`);
  }

  // ============================================================
  // Google Sheet
  // ============================================================
  function normalizeSheetUrl(url) {
    if (!url) return null;
    const u = url.trim();
    if (u.includes('/export?format=csv') || u.includes('output=csv') || u.includes('/gviz/tq')) return u;
    const m = u.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (m) { const gidM = u.match(/gid=(\d+)/); return `https://docs.google.com/spreadsheets/d/${m[1]}/export?format=csv&gid=${gidM ? gidM[1] : '0'}`; }
    return u;
  }

  function parseCsv(csvText) {
    if (!csvText) return [];

    // 最低限のCSVパーサ（引用符対応）
    const rows = [];
    let row = [];
    let cur = '';
    let inQuotes = false;

    for (let i = 0; i < csvText.length; i++) {
      const ch = csvText[i];
      const next = csvText[i + 1];

      if (inQuotes) {
        if (ch === '"' && next === '"') { // escaped quote
          cur += '"';
          i++;
        } else if (ch === '"') {
          inQuotes = false;
        } else {
          cur += ch;
        }
        continue;
      }

      if (ch === '"') { inQuotes = true; continue; }
      if (ch === ',') { row.push(cur); cur = ''; continue; }
      if (ch === '\r') { continue; }
      if (ch === '\n') {
        row.push(cur);
        rows.push(row);
        row = [];
        cur = '';
        continue;
      }
      cur += ch;
    }

    // last cell
    row.push(cur);
    rows.push(row);

    // B列（index=1）優先。無ければA列。
    // 先頭行のみヘッダー判定（本文に「メッセージ」「内容」が含まれても落とさない）
    const out = [];
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
      const r = rows[rowIndex];
      if (!r) continue;
      const a = (r[0] ?? '').trim();
      const b = (r[1] ?? '').trim();

      // 空行スキップ
      if (!a && !b) continue;

      // 先頭行のみヘッダー判定
      if (rowIndex === 0) {
        const aIsHeader = /^(番号|no\.?|no)$/i.test(a);
        const bIsHeader = /(メッセージ|内容)/.test(b);
        if (aIsHeader && bIsHeader) continue;
      }

      const msg = (b || a).trim();
      if (!msg) continue;
      out.push(msg);
    }

    return out;
  }

  function fetchSheet(url) {
    state.sheetStatus = 'loading'; state.sheetError = null; updateUI();
    const normalized = normalizeSheetUrl(url);
    if (!normalized) { state.sheetStatus = 'error'; state.sheetError = 'URLが無効です'; updateUI(); return; }
    logger.info('Sheet読み込み開始', normalized);
    if (typeof GM_xmlhttpRequest !== 'undefined') {
      GM_xmlhttpRequest({
        method: 'GET', url: normalized, timeout: 15000,
        onload: function(res) {
          if (res.status >= 200 && res.status < 300) { const msgs = parseCsv(res.responseText); if (msgs.length === 0) { state.sheetStatus = 'error'; state.sheetError = 'データがありません'; } else { state.messages = msgs; state.sheetUrl = url; state.sheetStatus = 'success'; storageSet('sheetUrl', url); storageSet('messages', msgs); showNotify(`${msgs.length}件読み込み`, 'success'); } }
          else if (res.status === 403) { state.sheetStatus = 'error'; state.sheetError = '権限エラー(403)'; }
          else if (res.status === 404) { state.sheetStatus = 'error'; state.sheetError = 'シート未発見(404)'; }
          else { state.sheetStatus = 'error'; state.sheetError = `HTTPエラー: ${res.status}`; }
          updateUI();
        },
        onerror: function() { state.sheetStatus = 'error'; state.sheetError = 'ネットワークエラー'; updateUI(); },
        ontimeout: function() { state.sheetStatus = 'error'; state.sheetError = 'タイムアウト'; updateUI(); },
      });
    } else {
      fetch(normalized, { method: 'GET', mode: 'cors', credentials: 'omit' })
        .then(r => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.text(); })
        .then(csv => { const msgs = parseCsv(csv); if (msgs.length === 0) throw new Error('データがありません'); state.messages = msgs; state.sheetUrl = url; state.sheetStatus = 'success'; storageSet('sheetUrl', url); storageSet('messages', msgs); showNotify(`${msgs.length}件読み込み`, 'success'); })
        .catch(e => { state.sheetStatus = 'error'; state.sheetError = e.message.includes('Failed to fetch') ? 'CORSエラー' : e.message; })
        .finally(() => updateUI());
    }
  }

  // ============================================================
  // 通知
  // ============================================================
  function showNotify(msg, type) {
    const old = $('#dating-ops-notify'); if (old) old.remove();
    const colors = { info: '#48bb78', success: '#48bb78', warn: '#f6ad55', error: '#fc8181' };
    const el = document.createElement('div'); el.id = 'dating-ops-notify'; el.textContent = msg;
    Object.assign(el.style, { position: 'fixed', bottom: '20px', right: '20px', padding: '10px 16px', background: colors[type] || colors.info, color: '#fff', borderRadius: '6px', boxShadow: '0 4px 12px rgba(0,0,0,0.3)', zIndex: '2147483647', fontSize: '13px', fontFamily: 'system-ui, sans-serif', maxWidth: '300px' });
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
      #${PANEL_ID} .btn-grn{background:#a6e3a1;color:#1e1e2e}
      #${PANEL_ID} .btn-grn:hover{background:#7bc96f}
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
    const s = document.createElement('style'); s.id = 'dating-ops-styles'; s.textContent = css; document.head.appendChild(s);
  }

  function buildPanelHtml() {
    const pm = state.pageMode, r = state.rows, t = state.textarea;
    const modeText = pm === 'list' ? '一括送信' : pm === 'personal' ? '個別送信' : 'その他';
    const modeClass = pm === 'list' ? 'st-info' : pm === 'personal' ? 'st-ok' : 'st-warn';

    let rowText, rowClass;
    if (pm === 'personal') { rowText = '1 (個別送信ページ)'; rowClass = 'st-ok'; }
    else if (pm === 'list') {
      if (r.usedSelector === 'textarea-batch') {
        rowText = r.count > 0 ? `${r.count}件 (返信欄)` : '0件';
      } else {
        rowText = r.count > 0 ? `${r.count}件 (fallback)` : '0件';
      }
      rowClass = r.count > 0 ? 'st-ok' : 'st-warn';
    } else { rowText = '-'; rowClass = 'st-info'; }

    let taText, taClass;
    if (t.status === 'found') { taText = `${t.selector} (${t.where})`; taClass = 'st-ok'; } else { taText = '未検出'; taClass = 'st-warn'; }

    const sheetMap = { idle: { t: '未設定', c: 'st-warn' }, loading: { t: '読込中...', c: 'st-info' }, success: { t: `${state.messages.length}件`, c: 'st-ok' }, error: { t: 'エラー', c: 'st-err' } };
    const sh = sheetMap[state.sheetStatus] || sheetMap.idle;

    let detail = `<div><span class="label">最終検出:</span> ${formatTime(state.lastDetectTime)}</div>`;
    detail += `<div><span class="label">トリガ:</span> ${state.lastTrigger || '-'}</div>`;
    detail += `<div><span class="label">rows.where:</span> ${r.where || '-'}${r.iframeSrc ? ' (' + escapeHtml(r.iframeSrc.substring(0, 25)) + ')' : ''}</div>`;
    detail += `<div><span class="label">scopeHint:</span> ${escapeHtml(r.scopeHint || '-')}</div>`;
    detail += `<div><span class="label">usedSelector:</span> ${escapeHtml(r.usedSelector || '-')}</div>`;
    detail += `<div><span class="label">batchTextareas:</span> ${r.batchTextareasCount}</div>`;
    if (r.tableScore !== null) detail += `<div><span class="label">tableScore:</span> ${r.tableScore}</div>`;
    detail += `<div><span class="label">displayCount:</span> ${r.displayCountNum !== null ? r.displayCountNum : '-'}</div>`;
    detail += `<div><span class="label">docsSearched:</span> ${r.docsSearched}</div>`;
    if (t.iframeSrc) detail += `<div><span class="label">ta.iframe:</span> ${escapeHtml(t.iframeSrc.substring(0, 30))}</div>`;
    if (state.observeTargetHints.length > 0) detail += `<div><span class="label">監視:</span> ${escapeHtml(state.observeTargetHints.join(' | '))}</div>`;

    let warn = '';
    if (r.mismatchWarning) warn = `<div class="warn-box">⚠️ ${escapeHtml(r.mismatchWarning)}</div>`;

    const msgHtml = state.messages.length > 0
      ? state.messages.map((m, i) => `<div class="msg-item" data-idx="${i}" title="${escapeHtml(m)}">${escapeHtml(m.length > 40 ? m.substring(0, 40) + '...' : m)}</div>`).join('')
      : '<div class="msg-item" style="color:#6c7086;">メッセージなし</div>';

    const batchBtnHtml = pm === 'list' ? `<button class="btn btn-grn" id="dp-batch-apply">表示分に一括挿入</button>` : '';

    return `
      <div class="header"><span class="header-title">${escapeHtml(CONFIG.ui.title)}</span><button class="header-close" id="dp-close">✕</button></div>
      <div class="body">
        ${warn}
        <div class="section">
          <div class="section-title">ステータス</div>
          <div class="status-box">
            <div class="status-row"><span class="status-label">モード:</span><span class="status-val ${modeClass}">${modeText}</span></div>
            <div class="status-row"><span class="status-label">対象:</span><span class="status-val ${rowClass}">${rowText}</span></div>
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
          <div class="btn-row">${batchBtnHtml}</div>
        </div>
        <div class="btn-row"><button class="btn btn-red" id="dp-stop">${state.stopped ? '停止中' : '停止'}</button><button class="btn btn-sec" id="dp-diag">診断</button></div>
      </div>
    `;
  }

  function createPanel() {
    const old = $(`#${PANEL_ID}`); if (old) old.remove();
    createStyles();
    const panel = document.createElement('div'); panel.id = PANEL_ID; panel.innerHTML = buildPanelHtml();
    const pos = storageGet('panelPos', null); if (pos) state.panelPos = pos;
    panel.style.left = `${state.panelPos.x}px`; panel.style.top = `${state.panelPos.y}px`;
    document.body.appendChild(panel); state.panel = panel; bindPanelEvents();
  }

  function bindPanelEvents() {
    const p = state.panel; if (!p) return;
    p.querySelector('.header')?.addEventListener('mousedown', onDragStart);
    p.querySelector('#dp-close')?.addEventListener('click', () => p.style.display = 'none');
    p.querySelector('#dp-load-sheet')?.addEventListener('click', () => { const url = p.querySelector('#dp-sheet-url')?.value?.trim(); url ? fetchSheet(url) : showNotify('URLを入力', 'warn'); });
    p.querySelector('#dp-rescan')?.addEventListener('click', () => { showNotify('再検出', 'info'); runDetection('manual'); });
    p.querySelector('#dp-apply-direct')?.addEventListener('click', () => { const txt = p.querySelector('#dp-direct')?.value?.trim(); if (!txt) { showNotify('テキストを入力', 'warn'); return; } const msgs = txt.split('\n').map(l => l.trim()).filter(Boolean); if (msgs.length === 0) { showNotify('有効なメッセージなし', 'warn'); return; } state.messages = msgs; storageSet('messages', msgs); state.sheetStatus = 'success'; showNotify(`${msgs.length}件適用`, 'success'); updateUI(); });
    p.querySelector('#dp-stop')?.addEventListener('click', () => { state.stopped = true; stopObservers(); clearInterval(state.healthTimer); clearInterval(state.urlPollTimer); clearTimeout(state.stagedDetectTimer); showNotify('停止しました', 'warn'); updateUI(); });
    p.querySelector('#dp-diag')?.addEventListener('click', runDiagnostic);
    p.querySelector('#dp-batch-apply')?.addEventListener('click', batchInsertVisible);
    p.querySelector('#dp-msg-list')?.addEventListener('click', e => { const item = e.target.closest('.msg-item'); if (item) { const i = parseInt(item.dataset.idx, 10); if (!isNaN(i) && state.messages[i]) insertText(state.messages[i]); } });
    document.addEventListener('mousemove', onDragMove);
    document.addEventListener('mouseup', onDragEnd);
  }

  function onDragStart(e) { if (e.target.closest('.header-close')) return; state.dragging = true; state.dragOffset = { x: e.clientX - state.panelPos.x, y: e.clientY - state.panelPos.y }; if (state.panel) state.panel.style.transition = 'none'; }
  function onDragMove(e) { if (!state.dragging || !state.panel) return; state.panelPos = { x: Math.max(0, Math.min(window.innerWidth - 350, e.clientX - state.dragOffset.x)), y: Math.max(0, Math.min(window.innerHeight - 50, e.clientY - state.dragOffset.y)) }; state.panel.style.left = `${state.panelPos.x}px`; state.panel.style.top = `${state.panelPos.y}px`; }
  function onDragEnd() { if (!state.dragging) return; state.dragging = false; storageSet('panelPos', state.panelPos); if (state.panel) state.panel.style.transition = ''; }
  function updateUI() { if (state.panel) { state.panel.innerHTML = buildPanelHtml(); bindPanelEvents(); } }

  function runDiagnostic() {
    // docCandidates: 各docの診断情報を収集
    const docs = getAccessibleDocs();
    const docCandidates = [];
    for (const { doc, win, where, iframeSrc, href } of docs) {
      try {
        const isMailbox = isMailboxDoc({ iframeSrc, href });
        const isBoxChar = isBoxCharDoc({ iframeSrc, href });
        const res = findBatchReplyTextareasInDoc(doc, win, isBoxChar);
        const best = findBestListTableInDoc(doc);
        docCandidates.push({
          where,
          iframeSrc: iframeSrc || null,
          href: href || null,
          isMailbox,
          isBoxChar,
          displayCountExtracted: extractDisplayCountFromDoc(doc),
          tableScore: best ? best.score : null,
          chosenCount: res.textareas.length,
          totalCandidates: res.totalCandidates,
          allTextareasCount: res.allTextareasCount,
          sampleTextareas: res.sampleTextareas,
          scopeHint: res.scopeHint,
        });
      } catch (e) {
        docCandidates.push({ where, iframeSrc, href, error: String(e) });
      }
    }

    const d = {
      siteType: SITE_TYPE,
      version: '2.4.5',
      pageMode: state.pageMode,
      textarea: { status: state.textarea.status, selector: state.textarea.selector, where: state.textarea.where, iframeSrc: state.textarea.iframeSrc },
      rows: state.rows,
      globalDisplayCount: extractDisplayCount(),
      docCandidates,
      lastDetectTime: state.lastDetectTime,
      lastTrigger: state.lastTrigger,
      observeTargetHints: state.observeTargetHints,
      messagesCount: state.messages.length,
      sheetStatus: state.sheetStatus,
      sheetError: state.sheetError,
      url: location.href,
    };
    console.group('%c[DatingOps] 診断', 'color:#48bb78;font-weight:bold'); console.log(d); console.groupEnd();
    navigator.clipboard?.writeText(JSON.stringify(d, null, 2)).then(() => showNotify('診断情報コピー', 'info')).catch(() => {});
    return d;
  }

  // ============================================================
  // トリガイベント
  // ============================================================
  function bindTriggerEvents() {
    const debouncedDetect = debounce(trigger => runDetection(trigger), OBSERVER_DEBOUNCE_MS);
    document.addEventListener('click', e => { if (state.stopped || !e.target || e.target.closest(`#${PANEL_ID}`)) return; for (const sel of CONFIG.clickTriggerSelectors) { try { if (e.target.matches(sel) || e.target.closest(sel)) { debouncedDetect('click'); return; } } catch {} } }, true);
    document.addEventListener('change', e => { if (state.stopped || !e.target || e.target.tagName !== 'SELECT' || e.target.closest(`#${PANEL_ID}`)) return; for (const sel of CONFIG.changeTriggerSelectors) { try { if (e.target.matches(sel)) { debouncedDetect('change'); return; } } catch {} } }, true);
  }

  // ============================================================
  // MutationObserver
  // ============================================================
  function findObserveTargetInDoc(doc) {
    const best = findBestListTableInDoc(doc);
    if (best) return { target: best.tbody, hint: `bestTable(${best.score})` };
    const cv = $('td.chatview', doc); if (cv) { const t = cv.closest('table'); if (t) return { target: $('tbody', t) || t, hint: 'chatview' }; }
    const ri = $('tr.rowitem', doc); if (ri) { const t = ri.closest('table'); if (t) return { target: $('tbody', t) || t, hint: 'rowitem' }; }
    return { target: doc.body || doc, hint: 'body' };
  }

  function stopObservers() {
    for (const obs of state.observers) { try { obs.disconnect(); } catch {} }
    state.observers = []; state.observeTargetHints = [];
  }

  function startObservers() {
    stopObservers();
    const docs = getAccessibleDocs();
    const handler = debounce(() => { if (!state.stopped) runDetection('mutation'); }, OBSERVER_DEBOUNCE_MS);
    for (const { doc, where } of docs) {
      try {
        const info = findObserveTargetInDoc(doc);
        const label = where === 'main' ? `main:${info.hint}` : `iframe:${info.hint}`;
        state.observeTargetHints.push(label);
        const obs = new MutationObserver(muts => { if (where === 'main' && muts.some(m => m.target.id === PANEL_ID || m.target.closest?.(`#${PANEL_ID}`))) return; handler(); });
        obs.observe(info.target, { childList: true, subtree: true });
        state.observers.push(obs);
        logger.info(`Observer開始: ${label}`);
      } catch (e) { logger.warn(`Observer失敗: ${where}`, e); }
    }
    if (state.observers.length === 0) logger.warn('Observerが1つもアタッチできず');
  }

  // ============================================================
  // URL変化検知
  // ============================================================
  function setupUrlChangeDetection() {
    state.lastUrl = location.href;
    const wrap = fn => function() { const r = fn.apply(this, arguments); onUrlChange('history'); return r; };
    history.pushState = wrap(history.pushState);
    history.replaceState = wrap(history.replaceState);
    window.addEventListener('popstate', () => onUrlChange('popstate'));
    state.urlPollTimer = setInterval(() => { if (location.href !== state.lastUrl) onUrlChange('poll'); }, URL_POLL_MS);
  }

  function onUrlChange(src) {
    if (state.stopped) return;
    state.lastUrl = location.href;
    logger.info(`URL変化検知 (${src})`);
    startObservers();
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
      const dispNum = state.rows.displayCountNum;
      if (!taOk || (pm === 'list' && !rowsOk) || (pm === 'list' && dispNum > 0 && state.rows.count === 0)) {
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
    logger.info('初期化開始 v2.4.5');
    state.messages = storageGet('messages', []);
    state.sheetUrl = storageGet('sheetUrl', '');
    if (state.messages.length > 0) state.sheetStatus = 'success';
    state.pageMode = detectPageMode();
    createPanel();
    bindTriggerEvents();
    startObservers();
    setupUrlChangeDetection();
    startHealthCheck();
    runStagedDetection();
    state.initialized = true;
    window.__datingOps = { state, diag: runDiagnostic, rescan: () => runDetection('manual'), insert: insertText, batchInsert: batchInsertVisible, fetchSheet, getAccessibleDocs, findBatchReplyTextareas };
    logger.info('初期化完了');
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
