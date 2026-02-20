# company-ai-report

企業の公式情報（企業名・HP・任意でディープサーチ）を入力に、**AIの活用方法**と**期待効果**をレポート化し、スプレッドシート（CSV）で瞬時に出力する仕組みです。AI支援会社向けの提案・商談用レポート生成用。

## 何が入っているか

- **docs/**
  - **output-spec.md** … 生成物の仕様（スプレッドシートの列・クオリティ基準）
  - **input-spec.md** … クオリティ維持のための必須・推奨入力
  - **agent-design.md** … エージェントの役割・フロー・データ型
- **src/** … パイプライン実装（入力解決 → プロファイル → 活用案生成 → CSV出力）
- **scripts/run-report.ts** … 1企業分のレポート生成CLI
- **inputs/example_company.json** … 入力サンプル

## 前提

- Node.js が入っていること
- `tsx` で TypeScript を実行できること（`npm install` で入る）
- LLM API キー（OpenAI または Anthropic）を用意していること

## セットアップ

1. 依存のインストール

```bash
cd 個人開発/company-ai-report
npm install
```

2. 環境変数

`.env` に LLM 用のキーを設定する（`.env.example` をコピーして編集）。

```bash
cp .env.example .env
# OPENAI_API_KEY=sk-... または ANTHROPIC_API_KEY=...
```

## 使い方

### 1企業分のレポートを生成

```bash
npm run report -- --input inputs/example_company.json
```

出力は `data/out/` に保存される。

- `{企業名サニタイズ}_summary.csv` … 企業サマリ（1行）
- `{企業名サニタイズ}_utilization.csv` … 活用案・効果一覧（複数行）

### 提案シート型（正本フォーマット）で生成する

「考えられる施策」形式（共通ブロック ＋ 施策ブロック × N）のCSVを出したい場合:

```bash
npm run proposal -- --input inputs/example_company.json
```

出力: `data/out/{企業名}_proposal_sheet.csv`（列: ブロック種別, 施策名, 項目, 値, 単位, メモ）。  
スプレッドシートに貼り付け、フィルタで「共通」「施策」を分けて正本と同じレイアウトに整えられる。設計は `docs/agent-design_提案シート型.md` 参照。

### 4タブ＋前提条件・承認を一括出力（スプレッドシートにドカンと入れる）

「考えられる施策」「費用対効果など」「パッケージ」「ロードマップ」「前提条件・承認」の **5シート** を一括で作りたい場合:

```bash
npm run full:example
```

- **Google 認証あり**（`.env` に `GOOGLE_APPLICATION_CREDENTIALS=/path/to/service-account.json` を設定）: 新規スプレッドシートが作成され、5シートに一括でデータが入る。URL が表示される。
- **認証なし**: `data/out/` に 5つの CSV が出力される。手動でスプレッドシートにインポート可能。

詳細は `docs/4タブ構成_仕様.md` 参照。前提条件の承認フェーズも「前提条件・承認」シートで管理できる。

### 入力ファイルの形式

`docs/input-spec.md` を参照。必須は **企業名** と **公式情報（URL または テキスト）**。

- **URL** を渡すと、スクリプトがそのURLを取得してテキスト化する（簡易HTMLパース）。要約が既にある場合は `type: "text"` で `value` に貼る。
- **ディープサーチ** は別途実行し、結果を `deepSearchResult` に貼ると、活用案の具体性が上がる。

例（テキストで渡す場合）:

```json
{
  "companyName": "株式会社〇〇",
  "officialInfo": { "type": "text", "value": "会社概要や事業説明のテキスト..." },
  "industry": "製造業",
  "focusAreas": ["営業", "開発"],
  "deepSearchResult": "（任意）Perplexity や Grok のリサーチ結果..."
}
```

### 出力を Google Sheets で使う

1. 生成された CSV を Google ドライブにアップロードし、「スプレッドシートで開く」で開く。
2. または、Google Sheets API で直接書き込む機能を追加する（未実装。`src/pipeline.ts` の `writeReportCsv` の代わりに Sheet 用の関数を用意すれば拡張可能）。

## クオリティ維持のポイント

- **入力**: 公式情報（HPの要約 or URL）を必ず渡す。業種・注目領域があると提案が具体化する。
- **出力**: `docs/output-spec.md` の列定義と禁止事項を守るよう、プロンプトに反映済み。
- 既に「こういう表で出したい」仕様がある場合は、`docs/output-spec.md` を差し替え、`src/pipeline.ts` の列名・CSV出力部分だけ合わせれば同じパイプラインで使える。

## ライセンス

MIT（想定）。自社利用に合わせて変更して問題ありません。
