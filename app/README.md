# BEV要件織り込み管理 - 起動手順

## 起動方法

WSL2のターミナルで以下を実行:

```bash
cd /home/kazuhiko-kobayashi-dnjp/workspace/manage_swreq/app

# JIRA連携あり（トークンが有効になったら）
JIRA_URL=https://jira.geniie.net \
JIRA_USER=kazuhiko.kobayashi.j3j@jp.denso.com \
JIRA_TOKEN=<JIRAトークン> \
JIRA_PROJECT=BEV26EGSD \
node server.js

# JIRA連携なしで起動（JIRAチケット自動作成は無効になる）
node server.js
```

ブラウザで `http://localhost:3000` を開く。

## JIRA Personal Access Token の取得

現在のトークン（JIRAパスワード形式）は401エラーになっています。
Atlassian JIRAでは Personal Access Token (PAT) が推奨されています:

1. JIRA にログイン
2. 右上のアイコン → プロフィール → Personal Access Tokens
3. 新しいトークンを作成
4. 生成されたトークンを `JIRA_TOKEN` に設定
5. `JIRA_USER` は不要（PATはBearer認証のため）

server.jsのJIRA認証をBearerに変更する場合はお知らせください。

## ファイル構成

```
app/
  server.js          # Express APIサーバー
  package.json
  data/
    requirements.json  # 要件データ（929件）
  public/
    index.html         # フロントエンド
```

## データ永続化

- 編集・追加・削除は `data/requirements.json` にリアルタイム書き込み
- Excelへの書き戻しは未対応（必要なら追加可）
