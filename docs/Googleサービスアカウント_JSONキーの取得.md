# Google サービスアカウント JSON キーのダウンロード手順

スプレッドシートに一括出力するために、**サービスアカウント**の JSON キーを取得します。

---

## 1. Google Cloud コンソールを開く

1. ブラウザで [Google Cloud Console](https://console.cloud.google.com/) を開く。
2. Google アカウントでログインする。
3. 画面上部の**プロジェクト選択**で、使うプロジェクトを選ぶ（なければ「新しいプロジェクト」で作成）。

---

## 2. スプレッドシート API を有効にする

1. 左メニュー **「APIとサービス」** → **「ライブラリ」** を開く。
2. 「Google Sheets API」で検索し、**「Google Sheets API」** をクリック。
3. **「有効にする」** をクリック。

---

## 3. サービスアカウントを作る

1. 左メニュー **「APIとサービス」** → **「認証情報」** を開く。
2. **「＋ 認証情報を作成」** → **「サービス アカウント」** を選ぶ。
3. **サービス アカウント名**（例: `company-ai-report-sheets`）を入力。
4. **「作成して続行」** → ロールは「編集者」や「オーナー」など、スプレッドシートを作成・編集できる権限を選ぶ（またはスキップして後で付与）。
5. **「完了」** をクリック。

---

## 4. JSON キーをダウンロードする

1. **「認証情報」** ページに戻る。
2. **「サービス アカウント」** セクションで、今作ったサービス アカウント（メールアドレスが表示されている行）をクリック。
3. 開いた画面で上タブの **「キー」** をクリック。
4. **「鍵を追加」** → **「新しい鍵を作成」** を選ぶ。
5. **鍵のタイプ** で **「JSON」** を選び、**「作成」** をクリック。
6. JSON ファイルが自動でダウンロードされる（名前は `プロジェクト名-xxxxx.json` のような形式）。

---

## 5. プロジェクトで使う

1. ダウンロードした JSON ファイルを、**他人に共有しない場所**に置く（例: `個人開発/company-ai-report/keys/sheets-service-account.json`。`keys/` は `.gitignore` に追加しておく）。
2. `.env` に次のように書く（パスは実際のファイルの場所に合わせる）:

**相対パスで書く（おすすめ）** — プロジェクトの `keys/` に JSON を置いた場合:

```env
GOOGLE_APPLICATION_CREDENTIALS=keys/sheets-service-account.json
```

※ 実行時にカレントディレクトリが `個人開発/company-ai-report` になるので、`keys/` はその直下のフォルダです。

**絶対パスで書く場合** — 「ユーザー名」は **Google Cloud の名前ではなく、Mac のログインユーザー名**（`/Users/` の下のフォルダ名）です。ターミナルで `echo $HOME` と打つと `/Users/shotahiraga` のように出るので、その `shotahiraga` の部分がそれです。

```env
GOOGLE_APPLICATION_CREDENTIALS=/Users/shotahiraga/kaihatsu/個人開発/company-ai-report/keys/sheets-service-account.json
```

3. `npm run full:example` を実行すると、このサービス アカウントでスプレッドシートが作成される。

---

## 注意

- **JSON キーは秘密情報**です。Git にコミットしたり、公開場所に置かないでください。
- `.gitignore` に `keys/` や `*.json`（キー用）を追加することを推奨します。
