# outlook.nvim

Neovim(LazyVim)から Microsoft Outlook(デスクトップ版 / Classic Outlook, Windows)のメールを閲覧するための薄いプラグインです。

Outlook との連携は Outlook COM オブジェクトモデルを介して行い、Neovim 側は表示とキー操作のディスパッチに徹します。一覧・プレビューなどのUIは可能な限り [snacks.nvim](https://github.com/folke/snacks.nvim) など LazyVim 標準のコンポーネントを再利用し、独自UIは最小限に留めます。

> **Status:** v1(閲覧専用)実装済み。送信・カレンダー等は未実装です。詳細・ロードマップは [docs/DESIGN.md](docs/DESIGN.md) を参照してください。実行には Windows + Outlook(Classic) が必要なため、この開発環境では PowerShell ヘルパーの実機動作確認はできていません(Lua側は plenary.nvim によるモックテスト済み。[tests/README.md](tests/README.md) 参照)。

## Features (v1)

- Inbox / 未読のみの一覧表示(`snacks.picker`、無ければ `vim.ui.select` にフォールバック。プレビューは件名/差出人/日時のヘッダのみ — 一覧取得時に本文は取得しない設計)
- 全文表示(読み取り専用フローティングウィンドウ。`<CR>`で開いた時に本文を取得)
- 既読/未読の切り替え(メッセージを開くと自動的に既読になる。picker 内では `<C-r>` で明示的にトグルも可能)
- 件名・差出人での検索
- 一覧結果の短時間キャッシュ + 同時リクエストの1本化(Outlook COM呼び出しのレイテンシを隠す。詳細は `docs/DESIGN.md`)

## 動作環境

- Windows
- Microsoft Outlook(デスクトップ版, Classic Outlook)
- Windows PowerShell 5.1(Windows 標準搭載。Outlook COM オブジェクトモデルへのアクセスに使用し、Python 等の追加インストールは不要です)
- Neovim + [lazy.nvim](https://github.com/folke/lazy.nvim)(LazyVim 環境を想定)

## インストール

```lua
-- lua/plugins/outlook.lua
return {
  "yourname/outlook.nvim",
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

### 起動を`cmd`/`keys`ではなく常時ロードにしたい場合(レイテンシ最適化)

PowerShellヘルパーの起動 + Outlook COM接続には多少時間がかかります。初回オープン時の待ちをなくしたい場合は、`event = "VeryLazy"` などで起動時に読み込み、`opts.prewarm = true` でバックグラウンド接続を先に済ませておけます。

```lua
return {
  "yourname/outlook.nvim",
  event = "VeryLazy",
  opts = { prewarm = true },
}
```

which-key を使っている場合、`<leader>m` は "mail" グループとして自動登録されます(`opts.keys = true` の場合)。

## Configuration

```lua
require("outlook").setup({
  keys = true,         -- <leader>m 配下のデフォルトキーマップを登録する
  prewarm = false,     -- setup()時点でhelperプロセスを起動しておく(上記参照)
  cache_ttl_ms = 15000, -- 一覧/検索結果を再利用する期間(ミリ秒)
})
```

## Commands

| Command | 説明 |
|---|---|
| `:OutlookOpen` | Inbox一覧を開く |
| `:OutlookRefresh` | キャッシュを無視してInboxを再取得 |
| `:OutlookUnread` | 未読のみの一覧を開く |
| `:OutlookSearch` | 件名/差出人で検索 |
| `:OutlookHelperRestart` | PowerShellヘルパーを再起動(デバッグ用) |

## Health check

```vim
:checkhealth outlook
```

Windows / `powershell.exe` の有無、Neovimバージョン、`snacks.nvim` の有無を確認します。

## 設計ドキュメント

- [docs/DESIGN.md](docs/DESIGN.md) — アーキテクチャ、IPCプロトコル、モジュール構成、ロードマップ
- [tests/README.md](tests/README.md) — テストの実行方法と、モック/実機確認のカバー範囲

## Contributing

Lua のフォーマットは [StyLua](https://github.com/JohnnyMorganz/StyLua) (`stylua.toml` 準拠) を使用しています。PR前に `stylua .` を実行してください。テストは `tests/README.md` を参照してください。

## ライセンス

[MIT](LICENSE)
