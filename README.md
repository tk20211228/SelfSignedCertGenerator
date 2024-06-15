# 自己署名証明書の作成とエクスポート

このプロジェクトは、PowerShell スクリプトを用いて自己署名証明書を作成し、エクスポートするものです。

## 主な機能

- ルート CA 証明書、CA 証明書、クライアント証明書の作成
- 証明書のエクスポート（.pfx、.p12、.cer、.pem 形式）
- 証明書情報の JSON 形式での出力
- エラーログの作成

## 使い方

1. `config.json` ファイルを編集して、証明書の詳細を設定します。設定可能な項目は以下の通りです：

```json
{
  "root": {
    "name": "test-rootCA",
    "domain": "test-root-ca.com",
    "password": "test123!!",
    "friend": "test-root-ca-friend"
  },
  "ca": {
    "name": "test-CA",
    "domain": "test-ca.com",
    "password": "test123!!",
    "friend": "test-ca-friend"
  },
  "client": {
    "name": "test-client",
    "domain": "test-client.com",
    "password": "test123!!",
    "friend": "test-client-friend"
  }
}
```

2. PowerShell を管理者権限で実行し、スクリプトを起動します。

```
.\CreateAndExportSelfSignedCert.ps1
```

3. スクリプトが完了すると、証明書とログが出力されます。出力ディレクトリはスクリプトの実行結果に表示されます。

## 注意事項

- このスクリプトは管理者権限で実行する必要があります。管理者権限で実行されていない場合、スクリプトは自動的に管理者権限で再実行されます。

```
function Test-IsAdmin {
  if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole("Administrators")) {
    Start-Process powershell.exe "-ExecutionPolicy RemoteSigned -File `"$PSCommandPath`"" -Verb RunAs
    exit
  }

```

- エラーが発生した場合、エラーログが生成され、詳細情報が出力されます。

```
  catch {
    $errorMessage = $_.Exception.Message
    $errorLocation = $_.InvocationInfo.MyCommand.Path

```
