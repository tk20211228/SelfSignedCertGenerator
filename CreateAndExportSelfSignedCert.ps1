# 現在のタイムスタンプを取得
function Get-Timestamp {
  return (Get-Date -Format "yyyyMMdd_HHmmss")
}
# ログディレクトリの新規作成
function New-LogDirectory {
  param (
    [string]$itemPath
  )
  $year = Get-Date -Format "yyyy"
  $timestamp = Get-Timestamp
  $logPath = Join-Path -Path $PSScriptRoot -ChildPath "$itemPath\$year\$timestamp"
  # ログディレクトリが存在しない場合は作成
  if (!(Test-Path -Path $logPath)) {
    $dir = New-Item -ItemType Directory -Path $logPath
    $logPath = $dir.FullName
  }
  Write-Host "New-LogDirectory: $logPath"
  return $logPath
}

# 証明書作成
# 関数の返り値に含まれないようにnullを設定 Out-Nullを設定
function New-SelfSignedCertificateWrapper {
  param (
    [hashtable]$params,
    [string]$logPath,
    [string]$type, # CA or Client
    [string]$pass,
    [string]$logFilePath,
    [String]$fileName 
  )

  $certificate = New-SelfSignedCertificate @params
  "certificate-$type : $($certificate.GetType().FullName)" | Out-File -FilePath $logFilePath -Append

  $certPassword = ConvertTo-SecureString -String $pass -Force -AsPlainText
  $exportPath = Join-Path -Path $logPath -ChildPath $type
  "exportPath: $exportPath" | Out-File -FilePath $logFilePath -Append
  New-Item -ItemType Directory -Path $exportPath | Out-Null

  # pause
  Export-PfxCertificate -Cert $certificate -FilePath "$exportPath\$($fileName)-${type}Cert.pfx" -Password $certPassword | Out-Null
  Export-PfxCertificate -Cert $certificate -FilePath "$exportPath\$($fileName)-${type}Cert.p12" -Password $certPassword | Out-Null

  Export-Certificate -Cert $certificate -FilePath "$exportPath\$($fileName)-${type}Cert.cer" | Out-Null
  $certContent = Get-Content -Path "$exportPath\$($fileName)-${type}Cert.cer" -Encoding Byte
  $pemContent = [System.Convert]::ToBase64String($certContent)
  $pemContent = "-----BEGIN CERTIFICATE-----`n" + ($pemContent -replace "(.{64})", "`$1`n") + "`n-----END CERTIFICATE-----"
  $pemContent | Out-File -FilePath "$exportPath\$($fileName)-${type}Cert.pem" -Encoding ASCII

  return $certificate
}
# 管理者権限での実行を確認
function Test-IsAdmin {
  if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole("Administrators")) {
    Start-Process powershell.exe "-ExecutionPolicy RemoteSigned -File `"$PSCommandPath`"" -Verb RunAs
    exit
  }
}
# 設定ファイルの読み込み
function Get-Configuration {
  param (
    [string]$configFilePath,
    [string]$logFilePath
  )

  if (!(Test-Path -Path $configFilePath)) {
    Write-Host "Configuration file not found." | Out-File -FilePath $logFilePath -Append
    exit
  }

  return Get-Content -Path $configFilePath -Raw | ConvertFrom-Json
}

# メイン処理
function Main {
  try {
    # スクリプトのルートディレクトリに移動
    Set-Location -Path $PSScriptRoot
  
    Test-IsAdmin
    $config = Get-Configuration -configFilePath "config.json"
    # Write-Host "configType: $($config.GetType().FullName)"
    
    $logDirectory = New-LogDirectory -itemPath "/items"
    $logFilePath = Join-Path -Path $logDirectory -ChildPath "operation_log.txt"

    $jsonFilePath = Join-Path -Path $logDirectory -ChildPath "configInfo.json"
    $config | ConvertTo-Json -Depth 100 | Set-Content -Path $jsonFilePath

    # ルートCA証明書の作成
    $rootCAName = "$($config.root.name)_$(Get-Timestamp)"
    $rootCertParams = @{
      Subject           = "CN=$rootCAName, O=YourOrganization, C=YourCountry"
      DnsName           = $config.root.subjectAlternativeName
      CertStoreLocation = "Cert:\CurrentUser\My"
      TextExtension     = @("2.5.29.19={text}CA=true")
      KeyUsage          = @("CertSign", "CRLSign", "DigitalSignature")
      KeyLength         = 4096
      KeyAlgorithm      = "RSA"
      HashAlgorithm     = "SHA256"
      NotAfter          = (Get-Date).AddYears(10)
    }
    $rootCACert = New-SelfSignedCertificateWrapper -params $rootCertParams -logPath $logDirectory -type "rootCA" -pass $config.root.password -logFilePath $logFilePath -fileName $rootCAName

    "rootCACertType: $($rootCACert.GetType().FullName)" | Out-File -FilePath $logFilePath -Append

    # CA証明書のパラメータ
    $CAName = "$($config.ca.name)_$(Get-Timestamp)"
    $caParams = @{
      Subject           = "CN=$CAName, O=YourOrganization, C=YourCountry"
      DnsName           = $config.ca.subjectAlternativeName
      KeyExportPolicy   = "Exportable"
      KeySpec           = "Signature"
      KeyLength         = 4096
      KeyAlgorithm      = "RSA"
      HashAlgorithm     = "SHA256"
      CertStoreLocation = "Cert:\CurrentUser\My"
      FriendlyName      = $config.ca.friend
      Signer            = $rootCACert
      NotAfter          = (Get-Date).AddYears(10)
    }
  
    # 証明書を作成
    # pause
    $CACert = New-SelfSignedCertificateWrapper -params $caParams -logPath $logDirectory -type "CA" -pass $config.ca.password -logFilePath $logFilePath $logFilePath -fileName $CAName
    "CACertType: $($CACert.GetType().FullName)" | Out-File -FilePath $logFilePath -Append    
    
    # pause
    # Cient証明書のパラメータ
    $clientName = "$($config.client.subject.CN)_$(Get-Timestamp)"
    # $clientSubject = @($config.client.subject) | ForEach-Object { "$($_.Key)=$($_.Value)" } -join ", "
    $clientSubject = ($config.client.subject.PSObject.Properties | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ", "

    $clientParams = @{
      # Subject           = "CN=$clientName"
      Subject           = $clientSubject
      DnsName           = $config.client.subjectAlternativeName
      KeyExportPolicy   = "Exportable"
      KeySpec           = "Signature"
      KeyLength         = 2048
      KeyAlgorithm      = "RSA"
      HashAlgorithm     = "SHA256"
      CertStoreLocation = "Cert:\CurrentUser\My"
      FriendlyName      = $config.client.friend
      Signer            = $CACert
      NotAfter          = (Get-Date).AddYears(5)
    }
    New-SelfSignedCertificateWrapper -params $clientParams -logPath $logDirectory -type "Client" -pass $config.client.password -logFilePath $logFilePath $logFilePath -fileName $clientName
    
    $csvFilePath = Join-Path -Path $logDirectory -ChildPath "upload.csv"
    "Local File Path,Password" | Out-File -FilePath $csvFilePath -Append -Encoding UTF8
    "$logDirectory\rootCA\$($config.root.name)-rootCACert.cer," | Out-File -FilePath $csvFilePath -Append -Encoding UTF8
    "$logDirectory\CA\$($config.root.name)-CACert.pfx,$($config.ca.password)" | Out-File -FilePath $csvFilePath -Append -Encoding UTF8
    "$logDirectory\Client\$($config.root.name)-ClientCert.p12,$($config.client.password)" | Out-File -FilePath $csvFilePath -Append -Encoding UTF8

    # storeに保存証明書の削除
    $rootCACertificate = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object { $_.Subject -like "*$rootCAName*" }
    Remove-Item -Path "Cert:\CurrentUser\My\$($rootCACertificate.Thumbprint)" -DeleteKey
    $caCertificate = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object { $_.Subject -Match $CAName }
    Remove-Item -Path "Cert:\CurrentUser\My\$($caCertificate.Thumbprint)" -DeleteKey
    $clientCertificate = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object { $_.Subject -Match $config.client.subject.CN }
    Remove-Item -Path "Cert:\CurrentUser\My\$($clientCertificate.Thumbprint)" -DeleteKey

    pause
  
    # ディレクトリを開く
    Invoke-Item $logDirectory
  }
  catch {
    $errorMessage = $_.Exception.Message
    $errorLocation = $_.InvocationInfo.MyCommand.Path
    $errorLine = $_.InvocationInfo.ScriptLineNumber
    $errorPosition = $_.InvocationInfo.OffsetInLine
    $errorCategory = $_.CategoryInfo.Category
    $errorType = $_.Exception.GetType().FullName
    $errorFullMessage = "Error: $errorMessage`r`nLocation: $errorLocation`r`nLine: $errorLine, Position: $errorPosition`r`nCategory: $errorCategory`r`nType: $errorType`r`n"
  
    $errorLogPath = Join-Path -Path $logDirectory -ChildPath "error_log.txt"
    $errorFullMessage | Out-File -FilePath $errorLogPath -Append
    Write-Host "An error occurred. See $errorLogPath for details." | Out-File -FilePath $logFilePath -Append
    Invoke-Item $logDirectory
  }
  exit
}
Main
