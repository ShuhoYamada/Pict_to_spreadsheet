# 🛡️ セキュアセットアップガイド

このシステムはGoogleドライブとスプレッドシートを連携する写真ファイル管理システムです。

## ⚠️ セキュリティ重要事項

以下のファイルは**絶対にGitHubにアップロードしないでください**：
- `.env` - 機密情報（クライアントシークレット等）が含まれています
- `config.js` - APIキーが設定されている場合
- `node_modules/` - 依存関係フォルダ

## 🆕 最新機能 (v2.1.0)

### 🔗 写真ハイパーリンク機能
- **P区分写真**: 構成部品列に自動でハイパーリンクを設定
- **M区分写真**: 対応するP区分行の素材列にハイパーリンクを設定
- **セル内容保護**: 既存のテキストを上書きしない安全な設計
- **自動共有設定**: Google Driveファイルの共有設定を自動で有効化

## 🚀 初回セットアップ手順

### 1. 依存関係のインストール
```bash
npm install
```

### 2. 環境変数の設定
1. `.env.template`を`.env`にコピー
2. Google Cloud Consoleで取得した実際の値を設定

```bash
cp .env.template .env
# .envファイルを編集して実際の値を入力
```

### 3. Google Cloud Console設定

#### 必要なAPI
- Google Drive API（ファイル読み込み、共有設定、ハイパーリンク用）
- Google Sheets API（データ書き込み、ハイパーリンク設定用）

#### OAuth 2.0設定
- 承認済みのJavaScriptドメイン: `http://localhost:3000`
- 承認済みのリダイレクトURI: `http://localhost:3000/auth/callback`

### 4. サーバー起動
```bash
npm start
# または
python start_server.py
```

## 🔧 設定値の取得方法

### Google Cloud Console
1. [Google Cloud Console](https://console.cloud.google.com/)にアクセス
2. 新しいプロジェクトを作成
3. Google Drive API、Google Sheets APIを有効化
4. OAuth 2.0クライアントIDを作成
5. クライアントIDとシークレットを`.env`に設定

### 詳細な設定手順
詳しくは`SETUP_GUIDE.md`を参照してください。

## 🛠️ トラブルシューティング

### よくある問題
1. **認証エラー**: OAuth設定を確認
2. **API制限エラー**: Google Cloud Consoleでクォータを確認
3. **ポート競合**: PORT環境変数を変更
4. **ハイパーリンク設定エラー**: Google Drive APIのレート制限の可能性、時間をおいて再実行

## 📞 サポート

問題が発生した場合は、以下の情報を含めてお問い合わせください：
- エラーメッセージ
- ブラウザとバージョン
- 設定内容（機密情報は除く）