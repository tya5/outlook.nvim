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

```lua
-- lua/plugins/outlook.lua
return {
  "yourname/outlook.nvim",
  opts = {},
}
```

## 設計ドキュメント

- [docs/DESIGN.md](docs/DESIGN.md) — アーキテクチャ、IPCプロトコル、モジュール構成、ロードマップ

## ライセンス

[MIT](LICENSE)
