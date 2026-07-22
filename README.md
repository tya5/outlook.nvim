# outlook.nvim

Neovim(LazyVim)から Microsoft Outlook(デスクトップ版 / Classic Outlook, Windows)のメールを閲覧するための薄いプラグインです。

Outlook との連携は Outlook COM オブジェクトモデルを介して行い、Neovim 側は表示とキー操作のディスパッチに徹します。一覧・プレビューなどのUIは可能な限り [snacks.nvim](https://github.com/folke/snacks.nvim) など LazyVim 標準のコンポーネントを再利用し、独自UIは最小限に留めます。

> **Status:** 設計段階です。実装はまだ含まれていません。詳細は [docs/DESIGN.md](docs/DESIGN.md) を参照してください。

## 動作環境

- Windows
- Microsoft Outlook(デスクトップ版, Classic Outlook)
- Windows PowerShell 5.1(Windows 標準搭載。Outlook COM オブジェクトモデルへのアクセスに使用し、Python 等の追加インストールは不要です)
- Neovim + [lazy.nvim](https://github.com/folke/lazy.nvim)(LazyVim 環境を想定)

## インストール(予定)

LazyVim ユーザー向けの `lua/plugins/outlook.lua` に置く想定のスペック例です。`cmd`/`keys` をトリガーに遅延読み込みする、LazyVim / lazy.nvim の一般的な作法に合わせています。

```lua
-- lua/plugins/outlook.lua
return {
  "yourname/outlook.nvim",
  cmd = { "OutlookOpen", "OutlookRefresh" },
  keys = {
    { "<leader>mm", "<cmd>OutlookOpen<cr>", desc = "Mail: open inbox" },
  },
  opts = {},
}
```

which-key を使っている場合、`<leader>m` を "mail" グループとして表示させるには次のように spec を追加します(LazyVim 標準の group 登録の作法)。

```lua
-- lua/plugins/outlook.lua に追記、または別ファイルで which-key の opts を拡張
{
  "folke/which-key.nvim",
  optional = true,
  opts = {
    spec = {
      { "<leader>m", group = "mail" },
    },
  },
}
```

## Configuration

```lua
require("outlook").setup({
  -- 現時点ではデフォルトのみ。オプションは実装が進み次第ここに追記します。
})
```

## Health check

```vim
:checkhealth outlook
```

Windows / `powershell.exe` の有無、Neovimバージョン、`snacks.nvim` の有無を確認します。

## 設計ドキュメント

- [docs/DESIGN.md](docs/DESIGN.md) — アーキテクチャ、IPCプロトコル、モジュール構成、ロードマップ

## Contributing

Lua のフォーマットは [StyLua](https://github.com/JohnnyMorganz/StyLua) (`stylua.toml` 準拠) を使用しています。PR前に `stylua .` を実行してください。

## ライセンス

[MIT](LICENSE)
