# Web公開メモ

## まず決めること

このアプリは Node.js のサーバーを使っています。最初の公開先は Azure App Service が扱いやすいです。

## Microsoftログインの本番設定

Azure Portal の「アプリの登録」で、次を設定します。

1. サポートされているアカウントの種類は、個人用OneDriveを使うなら「任意の組織ディレクトリ内のアカウントと個人用 Microsoft アカウント」を選びます。
2. プラットフォームは「シングルページ アプリケーション（SPA）」を選びます。
3. リダイレクトURIに次を追加します。

```text
https://公開したドメイン/auth-callback.html
```

ローカル確認も続ける場合は、これも残します。

```text
http://localhost:3000/auth-callback.html
```

## 公開環境に入れる環境変数

```text
MICROSOFT_CLIENT_ID=Azureのアプリケーション（クライアント）ID
MICROSOFT_TENANT_ID=consumers
AUTH_MODE=microsoft-only
MICROSOFT_ALLOWED_ORIGINS=https://公開したドメイン
GEMINI_API_KEY=GeminiのAPIキー
GEMINI_MODEL=gemini-3-flash-preview
```

`AUTH_MODE=microsoft-only` にすると、公開サイトではメールだけの簡易ログインを隠します。

`MICROSOFT_ALLOWED_ORIGINS` は、Microsoftログインを開始してよいURLの安全リストです。公開URLを入れてください。

## OneDrive保存

ノートは各ユーザーの OneDrive アプリフォルダに `lecture-note-data.json` として保存されます。アプリ側でDBサーバーはまだ使いません。

必要なMicrosoft Graph権限は次です。

```text
User.Read
Files.ReadWrite.AppFolder
```

## 公開前チェック

```text
npm install
npm run check
npm start
```

ブラウザで次を確認します。

1. Microsoftログインできる
2. ノート作成後にOneDrive同期できる
3. 別ブラウザまたは別端末で同じMicrosoftアカウントにログインしてノートが復元される
4. AIで整えるが動く

## 注意

今の方式は、まず公開できる最小構成です。長期運用では、トークン更新、同期衝突の扱い、AI利用回数制限、プライバシーポリシーの整備を追加すると安心です。
