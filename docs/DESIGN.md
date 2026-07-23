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
| `set_flag` / `clear_flag` | フォローアップフラグの設定/解除(`MailItem.FlagStatus`。v1は`none`⇔`flagged`の二値のみ、`complete`は未対応) |
| `search_messages` | `Items.Restrict` によるDASL/Jetフィルタ検索(差出人・件名・日付・未読) |

各メソッドは PowerShell 側で `Items.Sort("[ReceivedTime]", $true)` → `Restrict(...)` の順で処理し、大量メールの全走査を避ける。一覧表示に必要な項目(件名/差出人/日時/既読/フラグ状態)のみ返し、本文は `get_message` で個別取得(遅延ロード)。

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

- **picker.lua**: `snacks.nvim` の有無で判定し、あれば `Snacks.picker.pick` でメール一覧(件名/差出人/日時、未読は強調表示)+ プレビュー(既定は件名/差出人/日時のヘッダのみ。本文は含まない — 一覧取得時に`Body`へアクセスしないための設計判断。3.2節参照)+ アクション(既読切替、本文を`preview.lua`のウィンドウで開く、`<C-l>`でpicker内プレビューに本文を読み込む)を提供。無い環境では `vim.ui.select` + 別コマンドでの本文表示にデグレードする。
  - **`<C-l>`(本文をプレビューに読み込む)**: `get_message` を呼び本文を取得して、`preview()` コールバックが直近に受け取った `ctx`(`current_preview_ctx` として保持)へ直接書き込む。あくまでユーザーがキーを押した時だけ実行される明示操作であり、カーソル移動に連動した自動読み込みは行わない(一覧を素早くスクロールしただけでOutlook COMへ`get_message`が連打されるのを避けるため)。取得した本文は `entry_id` をキーに `body_cache` へ保持し、同一メッセージの再表示では再取得しない。取得成功時は `open_message` と同じ「未読なら既読にする」処理(`fetch_and_mark_read`)を共有する。
- **preview.lua**: `Snacks.win` で読み取り専用フローティングウィンドウを開き、本文を非編集バッファとして表示(`bo.modifiable=false`, `bo.filetype="mail"` 等)。
- 通知(取得失敗、Outlook未起動等)は `Snacks.notify` / 無ければ `vim.notify` にフォールバック。

## 6. キーマップ / コマンド (実装済み)

- `<leader>m` を "mail" グループとして which-key に登録(`opts.keys ~= false` の場合、`keymaps.lua` が `which-key.add` で登録)。
  - `<leader>mm` : メール一覧を開く(picker)
  - `<leader>mu` : 未読のみ一覧
  - `<leader>ms` : 検索(`vim.ui.input`で件名/差出人条件を受けて`search_messages`)
  - 既読/未読トグル専用の既定キーマップは持たない。メッセージを開く、または `<C-l>` で本文を読み込むと自動的に既読になり(通常のメールクライアントのUXに合わせた設計判断)、`snacks.picker` 使用時のみ picker 内 `<C-r>` で明示トグルできる。`<C-l>` はpicker側の固定バインドで `opts.keys` の対象外。
  - フォローアップフラグの設定/解除も同様に、`snacks.picker` 使用時のみ picker 内 `<C-f>` でトグルできる(`none`⇔`flagged`)。専用コマンドは持たない。
  - **`<C-n>` (もっと読み込む)**: 一覧・検索結果の件数上限を50件増やして再取得し、picker内のitemsを閉じずに差し替える。詳細・制約は6.2節。
- ユーザーコマンド: `:OutlookOpen`, `:OutlookRefresh`(キャッシュ無視で再取得), `:OutlookUnread`, `:OutlookSearch`, `:OutlookHelperRestart`(ヘルパープロセスの再起動、デバッグ用)。

## 6.1 レイテンシ対策(実装済み)

PowerShellプロセス起動 + Outlook COM接続には無視できない遅延があるため、以下で体感速度を確保する。

- **常駐ヘルパー**: `helper.lua` はプロセスを一度起動したら保持し、リクエスト毎の再起動を避ける。
- **prewarm**: `opts.prewarm = true` にすると `setup()` 実行時点でヘルパー起動+Outlook接続をバックグラウンドで開始する。lazy.nvim の読み込みトリガーを `cmd`/`keys`(=ユーザー操作まで未ロード)ではなく `event = "VeryLazy"` 等にしているユーザー向けのオプション。
- **picker側の表示タイミング**: `snacks.picker`・`vim.ui.select` いずれの経路でも `picker.list()` はキャッシュ/helperから結果(またはエラー)が揃った時点で初めてUIを開く(pickerウィンドウを先に開いて後から非同期に流し込む方式ではない)。そのためキャッシュヒット時は体感即時、ミス時のみ「Outlook: 読み込み中…」を通知してから結果を待つ(キャッシュヒットではこの通知は出さない)。将来 `snacks.picker` の非同期`finder`に載せ替えて「pickerを即開いて後から流し込む」体験にする余地はあるが、実機で snacks API を確認できるまで見送っている。
- **短時間キャッシュ + 多重リクエスト抑止**: `picker.lua` が同一パラメータの結果を `cache_ttl_ms`(既定15秒)の間再利用し、同時に発火した同一リクエストは1本のOutlook COM呼び出しに集約する(`:OutlookRefresh` はキャッシュを無視して強制再取得)。

## 6.2 「もっと読み込む」(`<C-n>`) — 本格的な無限スクロールではない理由

snacks.nvimのソース(`lua/snacks/picker/core/finder.lua`, `lua/snacks/picker/core/picker.lua`)を直接確認して設計した。

- **`finder`はフィルタ(検索クエリ文字列)が変わった時に1回だけ呼ばれる**。非同期`finder`(コールバックで逐次`add(item)`)も存在するが、これは「1回のfinder呼び出しの中で結果を少しずつ流し込む」ためのものであり、スクロール位置をトリガーに追加のfinder呼び出しが走る仕組みは無い。
- **開いているpickerに項目を追加する公開APIは無い**(`picker:add_items()`のようなものは存在しない)。一方 **`picker:find({ refresh = true })`** で finder を強制的に再実行でき、**`picker:refresh()`** は「選択をクリアし、finder/matcherを再実行する」と説明されている(`refresh_row()`で既読/フラグ変更後の行更新に使用。3.2節の`ConvertTo-MessageSummary`とは無関係)。
- Outlook COM の `Items` コレクションには offset/cursor に相当するものが無く、`Restrict`/`Sort`後に`foreach`で先頭から数える以外の方法が無い(3.2節)。

以上から、「スクロールで自動的に追加取得される」形の無限スクロールは snacks.picker の設計と噛み合わない。実装した `<C-n>` は代わりに:

1. 現在の`limit`に50を足して同じ`method`/`params`で`list_messages`/`search_messages`を再実行(＝フォルダの先頭から`limit`件を数え直すだけで、差分だけを安く取ってくるわけではない)
2. 結果を、`Snacks.picker.pick({items=...})`に渡した**同じテーブルオブジェクト**の中身を丸ごと入れ替える形で反映(新しいテーブルに差し替えるとpicker内部の参照と食い違う可能性があるため)
3. `picker:refresh()`をベストエフォートで呼び、その場での再描画を試みる

`picker.lua`の`current_list_state`(module-level, 同時に開けるoutlook pickerは1つの前提)が`{method, params, items, limit, loading}`を保持し、`M.show`が`opts.method`/`opts.params`を受け取った時だけ有効化される(`:OutlookSearch`もこれに対応済み)。`state.loading`で多重リクエストを防止し、取得件数が増えなかった場合は「これ以上のメールはありません」を通知する。

**実機未確認事項**: `picker:refresh()`が実際に「同じitemsテーブルの中身が変わっていること」を検知して再描画するか(finderが「テーブル」として渡された場合、再実行時に元のテーブル参照を`ipairs`で再走査するのか、それとも初回にのみ内部へ取り込んで以降は無視するのか)は、この開発環境(snacks無し)では確認できていない。ダメだった場合の見た目上のフォールバックは「`<C-n>`を押しても件数だけ増えて表示が変わらず、picker を閉じて`:OutlookOpen`し直すと反映されている」という、今回の既読/フラグ表示更新と同種の症状になる想定。

## 7. 配布・インストール前提

- 対象環境: Windows + Outlook (Classic, デスクトップ版) + Windows PowerShell 5.1(標準搭載)。
- lazy.nvim ユーザー向けの通常プラグインスペックとして配布(LazyVim公式extrasへの登録はv1では狙わず、将来的に安定後に検討)。
- README に「LazyVimへの組み込み例」としてコピペ用の `lua/plugins/outlook.lua` スペック片を掲載する(実質的にextra相当のUXを個人配布で提供)。

## 8. 未解決・要検討事項

1. 対象メールボックスが単一(既定プロファイル)前提か、複数アカウント/共有メールボックス対応も見るか。(v1は既定プロファイル1つのみ)
2. `list_messages` のページング方式(件数指定のみ/日付レンジ/カーソルベース)をどうするか。(v1は件数上限のみ)
3. 本文表示は プレーンテキスト(`Body`)のみで十分か、HTML本文の簡易テキスト化まで要るか。(v1は`Body`のみ)
4. ~~ヘルパーの配置場所~~ → 解決済み: リポジトリ同梱の `helper/outlook-helper.ps1` をそのまま `-File` 指定で実行する(コピー不要)。`lua/outlook/helper.lua` がプラグイン自身のパスから相対的にスクリプトパスを解決する。
5. ~~`Invoke-*` の `$null` 集約~~ → 解決済み: `Invoke-GetMessage`/`Invoke-SetRead` は対象アイテムが見つからない場合 `New-HelperError -Code "ITEM_NOT_FOUND"` を返すようになり、`$null`(Outlook未接続)とは区別される。`Get-FolderByName` 経由のフォルダ未検出(`list_messages`/`search_messages`)は現状も `$null` → `OUTLOOK_NOT_RUNNING` に丸めたまま(v1が扱う3フォルダ名は固定なので実運用上は起こりにくい)。
6. Outlook側での既読状態変更(他クライアント/デバイスからの変更)を、この一覧キャッシュ(既定15秒)がどこまで許容して古いままにするかは実運用で様子見。
7. `list_folders` は実装済みだが呼び出し元(コマンド/キーマップ)が無い未使用API。`Get-FolderByName` が `inbox`/`sent`/`drafts` の既定フォルダ名しか解決できず、`list_folders` が返す任意の `path`(例 `"受信トレイ/プロジェクトA"`)からフォルダを引くには PS側にパス解決処理を追加する必要があるため、`:OutlookFolders` のようなフォルダ選択コマンドの追加は本体スコープでは見送っている。追加する場合は Get-FolderByName の拡張とセットで行うこと。
8. フラグは v1 では `none`⇔`flagged` の二値トグルのみ(`<C-f>`)。`FlagStatus = 1`(`olFlagComplete`, 完了マーク)・`FlagRequest`("Follow up"以外のカスタム文言)・`FlagDueBy`(期限日時)は未対応。必要になれば `Invoke-SetFlag` にオプション引数を足す形で拡張できる。

## 9. 段階的ロードマップ

- **v1**: 本ドキュメントのスコープ(閲覧・既読管理・検索)
- **v2**: 下書き作成・送信(`.Display()`優先でユーザー確認を挟む)、添付ファイル一覧/保存
- **v3**: カレンダー(予定一覧・会議応答)、複数アカウント対応
- **v4検討**: 新着通知(ポーリング間隔短縮 or COMイベント化の再検討)
