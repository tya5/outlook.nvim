# outlook.nvim 設計ドキュメント (v1)

## 1. ゴールと非ゴール

**ゴール**
- LazyVim ユーザーが Neovim を離れずに Outlook(デスクトップ版 / Classic Outlook)のメールを閲覧できる。
- 既存の LazyVim UI(snacks.picker, snacks.win, snacks.notify, which-key)に極力乗っかり、独自UIを最小化する。
- Neovim 側は「表示と操作のディスパッチ」に徹する薄いプラグインとし、Outlook とのやりとりは外部ヘルパープロセスに閉じ込める。

**v1 の非ゴール(段階的に後回し)**
- メール送信・返信・新規作成(Object Model Guard の送信プロンプト、下書き作成フローなど検討事項が多いため v2 送り)
- カレンダー/予定表・会議招待の作成/応答
- COM イベント(NewMailEx)によるリアルタイム通知 — v1 はポーリングのみ
- New Outlook (Web/Graphベースの新UI) 対応 — Classic Outlook 前提

## 2. 全体アーキテクチャ

```
┌─────────────────────────┐        stdin/stdout        ┌──────────────────────────────┐
│  Neovim (Lua)            │   改行区切り JSON-RPC風     │  outlook-helper.ps1            │
│  lua/outlook/*.lua        │ <──────────────────────>  │  Windows PowerShell 5.1 (STA)  │
│  - job管理・request/response│                            │  New-Object -ComObject         │
│  - snacks.picker / win     │                            │      Outlook.Application       │
│  - which-key登録            │                            │  常駐・ポーリング                │
└─────────────────────────┘                            └───────────────┬──────────────┘
                                                                        │ COM
                                                                 ┌──────▼───────┐
                                                                 │ Outlook (Classic) │
                                                                 └───────────────┘
```

- Python/pywin32 は不採用。ヘルパーは **Windows PowerShell 5.1 (`powershell.exe`)** で実装し、Outlook COM オブジェクトモデルへ `New-Object -ComObject Outlook.Application` で直接アクセスする。
  - PowerShell 5.1 の既定ホストは STA アパートメントで起動するため、pywin32 で必要な `pythoncom.CoInitialize()` 相当の明示処理が不要。
  - pwsh(PowerShell 7)はデフォルトMTAのため対象外。`powershell.exe` を明示的に呼ぶ。
  - Windows 同梱コンポーネントのみで完結するため、ユーザーは Python/pip の追加インストールが不要。
- ヘルパーは **常駐プロセス**。`jobstart(cmd, {rpc = false})` で起動し、プラグインのライフサイクル中は起動しっぱなしにする(Outlook Dispatch/プロファイル読み込みのコストをリクエスト毎に払わないため)。
- 通信は **改行区切りJSON**(LSP風の `Content-Length` ヘッダは使わず、1行1JSONのシンプルな形式)。`vim.json.decode/encode` と PowerShell の `ConvertTo-Json -Compress`/`ConvertFrom-Json` で完結する。

## 3. IPC プロトコル

### 3.1 メッセージ形式

Neovim → ヘルパー(リクエスト、1行1JSON、`\n` 区切り):
```json
{"id": 1, "method": "list_messages", "params": {"folder": "inbox", "limit": 50, "unread_only": false}}
```

ヘルパー → Neovim(レスポンス):
```json
{"id": 1, "ok": true, "result": {"items": [ ... ]}}
```
```json
{"id": 1, "ok": false, "error": {"code": "OUTLOOK_NOT_RUNNING", "message": "Outlook が起動していません"}}
```

- `id` は Lua 側が発番するインクリメンタルな整数。Lua側は `id -> callback` のテーブルで応答を待つ非同期リクエスト/レスポンスにする(`on_stdout` はチャンク単位で来るので `\n` でバッファリングしてJSONをパースする)。
- ヘルパー起点のプッシュ通知(将来の新着メール通知等)は `id` なしの `{"event": "new_mail", "data": {...}}` 形式を予約しておくが、v1では未使用(ポーリングのみ)。

### 3.2 v1 で実装するメソッド

| method | 用途 |
|---|---|
| `ping` | ヘルパー起動確認・Outlook接続確認 |
| `list_folders` | Inbox配下含むフォルダ一覧取得 |
| `list_messages` | 指定フォルダのメール一覧(件名/差出人/受信日時/既読状態/EntryID) |
| `get_message` | 1件の本文取得(EntryID+StoreID指定、`Body`優先・必要なら`HTMLBody`) |
| `mark_read` / `mark_unread` | 既読状態変更 |
| `search_messages` | `Items.Restrict` によるDASL/Jetフィルタ検索(差出人・件名・日付・未読) |

各メソッドは PowerShell 側で `Items.Sort("[ReceivedTime]", $true)` → `Restrict(...)` の順で処理し、大量メールの全走査を避ける。一覧表示に必要な項目(件名/差出人/日時/既読)のみ返し、本文は `get_message` で個別取得(遅延ロード)。

### 3.3 メールの識別子

- 一覧・キャッシュのキーは `EntryID` + `StoreID` の組。Outlook 側で移動されるとEntryIDのみでは不安定なため、両方保持して `GetItemFromID(entryId, storeId)` で再取得する。

## 4. エラーハンドリング / 既知の制約

- **Outlook 未起動時**: `GetActiveObject("Outlook.Application")` 相当(PowerShellでは `[Runtime.InteropServices.Marshal]::GetActiveObject` )で先に既存インスタンスの有無を確認し、無ければ明確なエラー `OUTLOOK_NOT_RUNNING` を返す(自動起動はしない — 起動に時間がかかる上、意図しないプロファイル起動を避ける)。
- **New Outlook 判定**: COM Dispatch 自体が失敗する、またはバージョンチェックで新UIと判定できた場合は `NEW_OUTLOOK_UNSUPPORTED` を返し、Classic Outlook への切り替えを促すメッセージを表示する。
- **RPCサーバーが利用できません 等のCOMエラー**: Outlook がクラッシュ/再起動された場合、ヘルパー側で `Application`/`Namespace` ハンドルを次リクエスト時に再取得するリトライ処理を入れる。
- **文字コード**: PowerShell と Neovim 間は UTF-8 前提。`$OutputEncoding` / コンソールコードページの設定をヘルパー起動スクリプトで明示する(日本語件名・本文の文字化け対策)。
- v1 は読み取り専用のため、送信系で問題になる Object Model Guard の送信確認プロンプトは発生しない想定(`SenderEmailAddress`アクセス等、読み取り系の一部APIでもガードが発火するケースがあるとされるため、実装時に実機確認は必要)。

## 5. Neovim 側モジュール構成

```
lua/outlook/
  init.lua        -- setup(opts) のみ。実体は config.lua に委譲(folke系プラグインの慣習)
  config.lua       -- defaults/options/extend()。lazy.nvim/snacks.nvim等と同じ形
  health.lua       -- :checkhealth outlook (Windows/powershell.exe/snacks.nvim有無を確認)
  helper.lua       -- jobstart管理, request(method, params, callback), 改行区切りJSONの送受信
  picker.lua       -- snacks.picker連携(一覧・アクション), 無ければ vim.ui.select にフォールバック
  preview.lua      -- snacks.win によるメール本文プレビュー/読み取りウィンドウ
  commands.lua     -- :Outlook系ユーザーコマンド定義
  keymaps.lua      -- <leader>m 配下のキーマップ + which-key group登録
```

- **picker.lua**: `LazyVim.has("snacks.nvim")` で判定し、あれば `Snacks.picker.pick` でメール一覧(件名/差出人/日時、未読は強調表示)+ プレビュー(`preview` フィールドで本文抜粋)+ アクション(既読切替、本文を`preview.lua`のウィンドウで開く)を提供。無い環境では `vim.ui.select` + 別コマンドでの本文表示にデグレードする。
- **preview.lua**: `Snacks.win` で読み取り専用フローティングウィンドウを開き、本文を非編集バッファとして表示(`bo.modifiable=false`, `bo.filetype="mail"` 等)。
- 通知(取得失敗、Outlook未起動等)は `Snacks.notify` / 無ければ `vim.notify` にフォールバック。

## 6. キーマップ / コマンド (案)

- `<leader>m` を "mail" グループとして which-key に登録(`opts.spec` に group追加)。
  - `<leader>mm` : メール一覧を開く(picker)
  - `<leader>mu` : 未読のみ一覧
  - `<leader>mr` : カーソル下のメールの既読/未読トグル
  - `<leader>ms` : 検索(`vim.ui.input`で件名/差出人条件を受けて`search_messages`)
- ユーザーコマンド: `:OutlookOpen`, `:OutlookRefresh`, `:OutlookHelperRestart`(ヘルパープロセスの再起動、デバッグ用)。

## 7. 配布・インストール前提

- 対象環境: Windows + Outlook (Classic, デスクトップ版) + Windows PowerShell 5.1(標準搭載)。
- lazy.nvim ユーザー向けの通常プラグインスペックとして配布(LazyVim公式extrasへの登録はv1では狙わず、将来的に安定後に検討)。
- README に「LazyVimへの組み込み例」としてコピペ用の `lua/plugins/outlook.lua` スペック片を掲載する(実質的にextra相当のUXを個人配布で提供)。

## 8. 未解決・要検討事項(実装着手前に確認したい点)

1. 対象メールボックスが単一(既定プロファイル)前提か、複数アカウント/共有メールボックス対応も見るか。
2. `list_messages` のページング方式(件数指定のみ/日付レンジ/カーソルベース)をどうするか。
3. 本文表示は プレーンテキスト(`Body`)のみで十分か、HTML本文の簡易テキスト化まで要るか。
4. ヘルパーの配置場所(プラグインリポジトリ同梱の `.ps1` を `stdpath("data")` にコピーして実行 か、リポジトリ内パスを直接叩くか)と実行ポリシー(`-ExecutionPolicy Bypass` を都度指定する形でよいか)。

## 9. 段階的ロードマップ

- **v1**: 本ドキュメントのスコープ(閲覧・既読管理・検索)
- **v2**: 下書き作成・送信(`.Display()`優先でユーザー確認を挟む)、添付ファイル一覧/保存
- **v3**: カレンダー(予定一覧・会議応答)、複数アカウント対応
- **v4検討**: 新着通知(ポーリング間隔短縮 or COMイベント化の再検討)
