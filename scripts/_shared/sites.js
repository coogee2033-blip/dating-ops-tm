/**
 * sites.js - サイト固有設定（参照用）
 * OLV (olv29.com) / MEM (mem44.com) のURL・セレクタ・設定を定義
 * 
 * 実際のユーザースクリプトでは各 user.js に埋め込まれています
 */

(function() {
  'use strict';

  const SITE_CONFIGS = {
    olv: {
      siteType: 'olv',
      siteName: 'OLV',
      
      // パネル表示対象URL（正規表現）
      panelPages: [
        /\/staff\//,
      ],
      
      // 返信欄セレクタ（優先度順）
      textareaSelectors: [
        'textarea[name="message1"]',      // 最優先
        'textarea.msg.wd100',             // OLV特有のクラス
        'textarea.msg',                   // msgクラスのみ
        'textarea[name="message"]',
        'textarea[name="body"]',
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
      
      // 行検出用セレクタ（OLV特有）
      rowSelectors: [
        'tr.rowitem',
        'tr[id^="row"]',
        'tr[class*="row"]',
        'table.list tr',
        'table tr:has(td.chatview)',
        'table tr:has(input[type="checkbox"])',
      ],
      
      // 行内のチャットセル
      rowChatCellSelector: 'td.chatview',
      
      // 行内のチェックボックス
      rowCheckboxSelector: 'input[type="checkbox"]',
      
      // iframe検索用（全iframe対象）
      iframeSelectors: [
        'iframe',
      ],
      
      // UI設定
      ui: {
        title: 'OLV 返信アシスト',
        primaryColor: '#4a90d9',
        accentColor: '#2c5282',
      },
    },
    
    mem: {
      siteType: 'mem',
      siteName: 'MEM',
      
      panelPages: [
        /\/staff\//,
      ],
      
      textareaSelectors: [
        'textarea[name="message1"]',
        'textarea.msg.wd100',
        'textarea.msg',
        'textarea[name="message"]',
        'textarea[name="body"]',
        'textarea[name="mail_body"]',
        'textarea[name="content"]',
        'textarea:not([name*="memo"]):not([name*="admin"]):not([name*="note"])',
      ],
      
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

  /**
   * ホスト名からサイトタイプを判定
   */
  function detectSiteType(hostname) {
    if (!hostname) return null;
    if (hostname.includes('olv29') || hostname.includes('olv')) return 'olv';
    if (hostname.includes('mem44') || hostname.includes('mem')) return 'mem';
    return null;
  }

  /**
   * 設定を取得
   */
  function getConfig(siteType) {
    return SITE_CONFIGS[siteType] || null;
  }

  /**
   * パネル表示対象ページか判定
   */
  function shouldShowPanel(siteType, pathname) {
    const config = getConfig(siteType);
    if (!config) return false;
    return config.panelPages.some(pattern => pattern.test(pathname));
  }

  // グローバル公開
  window.DatingOpsSites = {
    SITE_CONFIGS,
    detectSiteType,
    getConfig,
    shouldShowPanel,
  };

})();
