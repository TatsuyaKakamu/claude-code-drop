# ClaudeDrop

Outlook のメールや、エクスプローラーのファイル・フォルダを小さな窓へドラッグすると、Claude Desktop に新しい Claude Code セッションを開く常駐アプリ。

## 動き

1. メール: 選んでいるメールを `Desktop\mail_drop` に `.msg` で保存する（クラシック Outlook のみ）
   ファイル・フォルダ: 種類を問わず、そのままのパスを使う（コピーしない）
2. `claude://code/new?q=...` を開き、入力欄にパスを入れる（送信は Enter）

## 使い方

- 起動: `start.vbs` をダブルクリック（スタートアップにショートカットを置けば自動起動）
- 終了: 窓の × で閉じる
- 設定: `config.json`（文言・保存先・最初の指示文）

## メモ

- `folder` を指定すると毎回フォルダの信頼確認が出る（Claude Desktop の仕様）ので、既定は空。
- 窓の位置は `window.json`、ログは `ClaudeDrop.log` に残る（どちらも git 管理外）。
