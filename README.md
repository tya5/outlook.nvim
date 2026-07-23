# outlook.nvim

[![CI](https://github.com/tya5/outlook.nvim/actions/workflows/ci.yml/badge.svg)](https://github.com/tya5/outlook.nvim/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Neovim(LazyVim)から Microsoft Outlook(デスクトップ版 / Classic Outlook, Windows)のメールを閲覧するための薄いプラグインです。

Outlook との連携は Outlook COM オブジェクトモデルを介して行い、Neovim 側は表示とキー操作のディスパッチに徹します。一覧・プレビューなどのUIは可能な限り [snacks.nvim](https://github.com/folke/snacks.nvim) など LazyVim 標準のコンポーネントを再利用し、独自UIは最小限に留めます。

> [!NOTE]
> **Status:** v1(閲覧専用)実装済み。送信・カレンダー等は未実装です。詳細・ロードマップは [docs/DESIGN.md](docs/DESIGN.md) を参照してください。
>
> この開発環境には Windows / Outlook が無いため、**PowerShell ヘルパー(`helper/outlook-helper.ps1`)側は実機で動作確認できていません**。Lua側(Neovimプラグイン本体)は [plenary.nvim](https://github.com/nvim-lua/plenary.nvim) によるモックテストでカバーしています([tests/README.md](tests/README.md) 参照)。実機での動作確認は歓迎します — 気づいた点は [Issues](https://github.com/tya5/outlook.nvim/issues) へどうぞ。

## 目次

- [Features (v1)](#features-v1)
- [仕組み](#仕組み)
- [動作環境](#動作環境)
- [インストール](#インストール)
- [Configuration](#configuration)
- [Commands](#commands)
- [Keymaps](#keymaps)
- [Health check](#health-check)
- [設計ドキュメント](#設計ドキュメント)
- [Contributing](#contributing)
- [ライセンス](#ライセンス)

## Features (v1)

- Inbox / 未読のみの一覧表示(`snacks.picker`、無ければ `vim.ui.select` にフォールバック。プレビューは件名/差出人/日時のヘッダのみ — 一覧取得時に本文は取得しない設計)
- `snacks.picker` 使用時、`<C-l>` でカーソル下メッセージの本文をpicker内のプレビュー欄にその場で読み込める(明示操作のみ。カーソル移動での自動読み込みはしない — 一覧を素早くスクロールしてもOutlookへ余計なCOM呼び出しは飛ばない設計)
- 全文表示(読み取り専用フローティングウィンドウ。`<CR>`で開いた時に本文を取得)
- 既読/未読の切り替え(メッセージを開く、または `<C-l>` で本文を読み込むと自動的に既読になる。picker 内では `<C-r>` で明示的にトグルも可能)
- フォローアップフラグの設定/解除(`snacks.picker` 使用時、picker内 `<C-f>` で `none`⇔`flagged` をトグル。完了マークは未対応)
- 件名・差出人での検索
- 「もっと読み込む」(`snacks.picker`使用時、picker内`<C-e>`)で一覧・検索結果を件数上限を増やして再取得(50件ずつ。Outlook COM/snacks.picker双方に増分取得の仕組みが無いため、都度フォルダの先頭から取り直す実装。詳細は[docs/DESIGN.md](docs/DESIGN.md) 6.1節)
- 一覧結果の短時間キャッシュ + 同時リクエストの1本化(Outlook COM呼び出しのレイテンシを隠す。詳細は [docs/DESIGN.md](docs/DESIGN.md) 6.1節)

v2以降で送信・添付ファイル・カレンダー・複数アカウント対応を検討しています。詳細は [docs/DESIGN.md](docs/DESIGN.md) の「段階的ロードマップ」を参照してください。

## 仕組み

```
┌─────────────────────┐   改行区切りJSON    ┌──────────────────────────┐   COM   ┌─────────┐
│ Neovim (Lua)          │ <───────────────>  │ outlook-helper.ps1         │ <────> │ Outlook  │
│ lua/outlook/*.lua      │  stdin/stdout       │ Windows PowerShell 5.1(STA) │        │ (Classic)│
└─────────────────────┘                     └──────────────────────────┘         └─────────┘
```

- Neovim プラグイン本体(Lua)は表示とキー操作のディスパッチに徹し、Outlook COM を直接触りません。
- Outlook との実際のやり取りは、プラグインが `jobstart` で起動する常駐の **Windows PowerShell** プロセス(`helper/outlook-helper.ps1`)が担当します。Python/pywin32 等の追加インストールは不要です(PowerShell は Windows 標準搭載)。
- 両者は改行区切りのJSON(1行1リクエスト/レスポンス)でやり取りします。プロトコルの詳細・エラーコード・レイテンシ対策(キャッシュ・タイムアウト等)は [docs/DESIGN.md](docs/DESIGN.md) にまとめています。

## 動作環境

- Windows
- Microsoft Outlook(デスクトップ版, Classic Outlook。"New Outlook" は COM 非対応のため未サポート)
- Windows PowerShell 5.1(Windows 標準搭載。Outlook COM オブジェクトモデルへのアクセスに使用)
- Neovim + [lazy.nvim](https://github.com/folke/lazy.nvim)(LazyVim 環境を想定)

## インストール

```lua
-- lua/plugins/outlook.lua
return {
  "tya5/outlook.nvim",
  cmd = { "OutlookOpen", "OutlookRefresh", "OutlookUnread", "OutlookSearch" },
  keys = {
    { "<leader>mm", "<cmd>OutlookOpen<cr>", desc = "Mail: open inbox" },
    { "<leader>mu", "<cmd>OutlookUnread<cr>", desc = "Mail: unread only" },
    { "<leader>ms", "<cmd>OutlookSearch<cr>", desc = "Mail: search" },
  },
  opts = {},
}
```

デフォルトキーマップ(`<leader>mm`/`<leader>mu`/`<leader>ms`)は `setup()` 実行時にプラグイン自身が登録します。上の `keys` は LazyVim の作法通り「これらのキーを押すまで読み込まない」ための遅延読み込みトリガーの宣言であり、実際のマッピングと二重にはなりません(不要なら省略しても `opts.keys` により自前登録されます)。

### 起動を `cmd`/`keys` ではなく常時ロードにしたい場合(レイテンシ最適化)

PowerShellヘルパーの起動 + Outlook COM接続には多少時間がかかります。初回オープン時の待ちをなくしたい場合は、`event = "VeryLazy"` などで起動時に読み込み、`opts.prewarm = true` でバックグラウンド接続を先に済ませておけます。

```lua
return {
  "tya5/outlook.nvim",
  event = "VeryLazy",
  opts = { prewarm = true },
}
```

which-key を使っている場合、`<leader>m` は "mail" グループとして自動登録されます(`opts.keys = true` の場合)。

## Configuration

`setup()` に渡せるオプションと既定値です(すべて省略可能)。

```lua
require("outlook").setup({
  keys = true,               -- <leader>m 配下のデフォルトキーマップを登録する
  prewarm = false,            -- setup()時点でhelperプロセスを起動しておく(上記参照)
  cache_ttl_ms = 15000,        -- 一覧/検索結果を再利用する期間(ミリ秒)
  request_timeout_ms = 30000, -- helperからの応答を待つ上限(ミリ秒)。超えるとエラー扱い
})
```

## Commands

| Command | 説明 |
|---|---|
| `:OutlookOpen` | Inbox一覧を開く |
| `:OutlookRefresh` | キャッシュを無視してInboxを再取得 |
| `:OutlookUnread` | 未読のみの一覧を開く |
| `:OutlookSearch` | 件名/差出人で検索(入力プロンプトが開く) |
| `:OutlookHelperRestart` | PowerShellヘルパーを再起動(デバッグ用) |

## Keymaps

既定では `opts.keys = true`(デフォルト)のとき、以下が登録されます。`opts.keys = false` にすれば登録されず、[インストール](#インストール)の例のように自分で `keys` を定義できます。

| Keymap | 説明 |
|---|---|
| `<leader>mm` | Inbox一覧を開く(`:OutlookOpen`) |
| `<leader>mu` | 未読のみの一覧を開く(`:OutlookUnread`) |
| `<leader>ms` | 検索(`:OutlookSearch`) |

`snacks.picker` 使用時、picker内では以下のキーも使えます(いずれも `opts.keys` の設定とは独立で、picker側の固定バインドです)。

| Keymap (picker内) | 説明 |
|---|---|
| `<CR>` | 全文を開く(自動的に既読になる) |
| `<C-l>` | カーソル下メッセージの本文をプレビュー欄に読み込む(自動的に既読になる。読み込み済みならキャッシュから即表示) |
| `<C-r>` | 既読/未読を明示的にトグル |
| `<C-f>` | フォローアップフラグを `none`⇔`flagged` でトグル |
| `<C-e>` | もっと読み込む(50件ずつ追加取得。取得済み件数が増えない=末尾に達した場合は通知) |

## Health check

```vim
:checkhealth outlook
```

Windows / `powershell.exe` の有無、Neovimバージョン、`snacks.nvim` の有無を確認します。

## 設計ドキュメント

- [docs/DESIGN.md](docs/DESIGN.md) — アーキテクチャ、IPCプロトコル、モジュール構成、ロードマップ、未解決事項
- [docs/HANDOFF.md](docs/HANDOFF.md) — v1実装後の自己レビュー指摘とその対応記録(実機未検証部分の注意点含む)
- [tests/README.md](tests/README.md) — テストの実行方法と、モック/実機確認のカバー範囲

## Contributing

- Lua のフォーマットは [StyLua](https://github.com/JohnnyMorganz/StyLua)(`stylua.toml` 準拠)を使用しています。コミット前に `stylua .` を実行してください。
- テストの実行方法は [tests/README.md](tests/README.md) を参照してください。PR前に一通り通してください。
- Issue・PR歓迎です。特にWindows + Outlook実機での動作報告・不具合報告は大歓迎です。

## ライセンス

[MIT](LICENSE)
