---
name: gen-report
description: 企業名とURLを入力してAI活用レポート（Google Sheets）を生成する
allowed-tools: AskUserQuestion, Bash, Write, Read
---

# AI活用レポート生成

企業のAI活用レポートをGoogle Sheetsに自動生成します。

## ステップ1: 情報収集

AskUserQuestion ツールを使って以下の情報をユーザーに聞いてください。
一度に複数の質問をまとめて出してかまいません。

必須情報:
- 会社名
- 公式サイトURL（または会社の説明文）

任意情報（わからない場合はスキップ可）:
- 業種（例: IT・SaaS, 建設, 不動産, 製造業）
- 注力部門（例: 営業, 現場, 経理）
- 補足メモ（担当者コメント等）

引数 `$ARGUMENTS` が渡されている場合はそこから会社名やURLを読み取り、
不足している情報だけを質問してください。

## ステップ2: 入力JSONの作成

このリポジトリのルートディレクトリ（`package.json` があるディレクトリ）の
`inputs/` フォルダに JSONファイルを作成します。

ファイル名は会社名をアルファベットかローマ字に変換したもの（例: `sony.json`）。

JSONフォーマット（任意フィールドは情報がある場合のみ含める）:

```json
{
  "companyName": "会社名",
  "officialInfo": {
    "type": "url",
    "value": "https://example.com"
  },
  "industry": "業種",
  "focusAreas": ["部門1", "部門2"],
  "memo": "補足メモ"
}
```

URLではなく説明文の場合は `"type": "text"` にします。

## ステップ3: レポート生成

リポジトリのルートディレクトリで以下のコマンドを実行します:

```bash
npm run full -- --input inputs/<作成したファイル名>
```

実行中は `GOOGLE_APPLICATION_CREDENTIALS` の設定有無によって出力先が変わります:
- 設定あり → Google Sheetsにスプレッドシートを生成
- 設定なし → `data/out/` にCSVを出力

## ステップ4: 結果報告

コマンドの出力から結果を取り出してユーザーに報告してください:
- Google Sheetsの場合: スプレッドシートのURLを表示
- CSVの場合: 出力されたファイルパスを表示
