@../../CLAUDE.md

# coworking-system（コワーキング入退場管理）

このフォルダは独立した Git リポジトリ。冒頭の `@../../CLAUDE.md` でコココリンAI全体の運用ルールを取り込む。
本ファイルと親 CLAUDE.md に食い違いがある場合は、コード開発に関する事項を除き親の CLAUDE.md を優先する。

## 概要

- コワーキングスペースの入退場・利用管理システム。
- 実装は Google Apps Script（`gas/`）＋ HTML 画面（`admin.html` = 管理、`tablet.html` = 受付タブレット）。
- ローカルと GAS の同期は clasp（`gas/.clasp.json`）。

## 開発メモ

- GAS コード: `gas/コード.js`、マニフェスト: `gas/appsscript.json`。
- VS Code ワークスペース: `R8.1.code-workspace`。
- スタッフ向け説明: `cococorin_スタッフマニュアル.docx / .pdf`。

## 注意

- `apps/` 配下の「稼働中システム」であり、企画の文書ワークフロー（orchestrator → 専門エージェント）の対象外。コード開発として扱う。
