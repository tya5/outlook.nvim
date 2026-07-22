# テスト

Windows / Outlook / PowerShell が無い環境でも回せる、Lua側(Neovimプラグイン本体)のユニットテストです。[plenary.nvim](https://github.com/nvim-lua/plenary.nvim) の busted風フレームワークを使用します。

## 何をテストしているか / していないか

- `helper_spec.lua`: `lua/outlook/helper.lua` の IPC層。`vim.fn.jobstart`/`vim.fn.chansend` をスタブし、実プロセスの代わりに疑似的な stdout チャンクを直接注入して、改行区切りJSONのバッファリング・リクエストID対応・プロセス終了時のエラー処理を検証する。
- `picker_spec.lua`: `lua/outlook/picker.lua` のキャッシュ/多重リクエスト抑止ロジック。`outlook.helper` を `package.loaded` 差し替えでフェイクに置き換え、実際のIPCなしに「同一リクエストのキャッシュ再利用」「force時のキャッシュバイパス」「同時リクエストの1本化」を検証する。
- **対象外**: `helper/outlook-helper.ps1`(PowerShell/Outlook COM 部分)そのもの。実行にはWindows + Outlookが必要で、この開発環境では検証できない。プロトコル(JSON入出力の形)は`helper_spec.lua`側でカバーしている。

## 実行方法

```sh
nvim --headless --noplugin -u tests/minimal_init.lua \
  -c "PlenaryBustedDirectory tests/ {minimal_init = 'tests/minimal_init.lua'}"
```

初回実行時、`tests/minimal_init.lua` が `plenary.nvim` を `stdpath("data")` 配下に自動clone します(リポジトリには含めません)。
