# microsoft-graph-sample
Microsoft Graph SDK を使った ユーザ情報、カレンダー情報を操作するサンプル

## Feature
- Microsoft Graph SDK
- .NET Framework 4.8

## Note
- Azure AD にてアプリケーション登録、APIアクセス許可、及びクレデンシャルを取得して下さい。
  - `Calendars.ReadWrite` `User.ReadWrite.All` のアクセス許可を想定しています。

- 取得した値を `Program.cs` にて設定して下さい。
```cs
var clientId = "xxx";
var secret = "xxx";
var tenantId = "xxx";
```
