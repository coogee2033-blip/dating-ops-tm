# Dating Ops - Tampermonkey Scripts

OLV (olv29.com) / MEM (mem44.com) サイト向けの返信アシストツールです。

## 機能

- Google Spreadsheet からメッセージテンプレートを読み込み
- 直接入力によるメッセージ管理
- ワンクリックで返信欄にメッセージ挿入
- SPA対応（ページ遷移・DOM更新でもパネル維持）
- iframe内テキストエリア対応（同一オリジン）

## 安全機能

- **送信操作は一切しません**（自動送信・擬似送信禁止）
- **管理者メモ・二人メモ等への誤挿入防止**
- **既存テキストがある場合は上書きしません**
- **textarea[name="message1"] を最優先で検出**

## ファイル構成

```
dating-ops-tm/
├── README.md
├── .gitignore
└── scripts/
    ├── _shared/
    │   ├── common.js    # 共通コアロジック（参照用）
    │   └── sites.js     # サイト固有設定（参照用）
    ├── olv/
    │   └── olv.user.js  # OLV用スクリプト（完全版・即動作）
    └── mem/
        └── mem.user.js  # MEM用スクリプト（完全版・即動作）
```

## 導入手順

### 1. Tampermonkey をインストール

- [Chrome](https://chrome.google.com/webstore/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)
- [Firefox](https://addons.mozilla.org/firefox/addon/tampermonkey/)
- [Edge](https://microsoftedge.microsoft.com/addons/detail/tampermonkey/iikmkjmpaadaobahmlepeloendndfphd)

### 2. スクリプトをインストール

#### 方法A: コピー＆ペースト（推奨）

1. Tampermonkey のダッシュボードを開く
2. 「新規スクリプトを追加」をクリック
3. 以下のファイルの内容を**全て**コピーして貼り付け：
   - OLV用: `scripts/olv/olv.user.js`
   - MEM用: `scripts/mem/mem.user.js`
4. 保存（Ctrl+S または Cmd+S）

#### 方法B: ファイルからインポート

1. このリポジトリをクローン
2. Tampermonkey ダッシュボード → ユーティリティ → ファイルからインポート
3. 該当の `.user.js` ファイルを選択

### 3. 対象サイトにアクセス

- OLV: `https://olv29.com/staff/*`
- MEM: `https://mem44.com/staff/*`

`/staff/` 配下のページでパネルが自動表示されます。

## 使い方

### パネル操作

| 操作 | 説明 |
|------|------|
| ドラッグ | ヘッダーをドラッグしてパネル移動（位置自動保存） |
| ✕ ボタン | パネルを一時非表示（再読み込みで復活） |
| 再検出 | テキストエリアを再スキャン |
| 停止 | 監視を完全停止 |
| 診断 | 状態情報をコンソール＆クリップボードに出力 |

### Google Sheet 読み込み

1. Google Spreadsheet を作成
2. A列にメッセージを1行ずつ入力
3. 「共有」→「リンクを知っている全員」に変更
4. URLをパネルに貼り付けて「Sheet読込」

### 直接入力

Sheet読み込みが失敗する場合は、直接入力欄を使用：

1. テキストエリアにメッセージを1行ずつ入力
2. 「適用」ボタンをクリック

### メッセージ挿入

メッセージ一覧から選択してクリックすると、返信欄に挿入されます。

**安全チェック**:
- 返信欄以外（管理者メモ等）には挿入されません
- 既にテキストがある場合は挿入されません

## トラブルシューティング

### パネルが表示されない

1. Tampermonkey が有効か確認
2. URL が `https://olv29.com/staff/*` または `https://mem44.com/staff/*` か確認
3. コンソール（F12）でエラーを確認
4. ページを再読み込み

### Sheet読み込みエラー

| エラー | 対処 |
|--------|------|
| CORSエラー | シートを「リンクを知っている全員」に公開 |
| 権限エラー(403) | 公開設定を再確認 |
| 未発見(404) | URL を確認 |

**回避策**: 直接入力欄を使用

### 返信欄が検出されない

1. 「再検出」ボタンをクリック
2. 「診断」で詳細を確認
3. `textarea[name="message1"]` が存在するか確認

### 挿入できない

- 「既にテキストが入力されています」→ 欄を空にしてから再試行
- 「この欄には挿入できません」→ 管理者用フィールドです

## デバッグコマンド

コンソールで使用可能：

```javascript
// 診断情報出力
window.__datingOps.diag()

// 監視停止
window.__datingOps.stop()

// 再検出
window.__datingOps.rescan()

// テキスト挿入
window.__datingOps.insert("テストメッセージ")

// Sheet読み込み
window.__datingOps.fetchSheet("https://...")

// 状態確認
console.log(window.__datingOps.state)
```

## 更新方法

1. 最新のコードを取得
2. Tampermonkey で該当スクリプトを開く
3. 新しいコードで全体を置き換え
4. 保存

## 設計メモ

### 安定性

- **MutationObserver** + デバウンス（300ms）で DOM 変化を監視
- **ヘルスチェック**（3秒間隔）でパネル存在を保証
- **リトライ制御**（最大5回）で無限ループ防止
- **null安全設計** - 全DOM操作にガード

### 誤挿入防止

- `excludeSelectors` で管理者用フィールドを除外
- name/id 属性に `memo`, `admin`, `note` を含む要素を除外
- 既存テキストがあれば挿入拒否

### iframe対応

- 同一オリジンの iframe は `contentDocument` でアクセス
- クロスオリジンは SecurityError をキャッチしてログ

## ライセンス

Private - All rights reserved
