# スクリプトの実行について

PowerShellスクリプトを実行する際に、セキュリティポリシーによってブロックされることがあります。
以下に、状況に応じた3つの実行許可方法を記載します。

---

## 1. 現在のPowerShellセッションのみ実行を許可する (安全)

一時的にスクリプトの実行を許可する方法です。この設定は現在のPowerShellウィンドウを閉じるとリセットされるため、安全です。

PowerShellで以下のコマンドを実行してください。

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

このコマンドにより、現在開いているPowerShellウィンドウ内でのみ `script.ps1` のようなスクリプトが実行可能になります。PCの全体設定は変更されません。

---

## 2. 特定のスクリプトまたはフォルダを「ブロック解除」する (推奨)

ダウンロードしたファイルや、WSL (Windows Subsystem for Linux) 経由で作成されたファイルは、PowerShellによってブロックされることがあります。
この方法はそのファイルやフォルダのブロックのみを解除するため、実務的で安全な方法として推奨されます。

特定のファイルのみブロックを解除する場合：
```powershell
Unblock-File -Path .\script.ps1
```

フォルダ内のすべてのスクリプトを再帰的にブロック解除する場合：
```powershell
Get-ChildItem . -Recurse | Unblock-File
```

---

## 3. 現在のユーザーに対して永続的に実行を許可する (一般的)

現在のWindowsユーザーに限り、スクリプトの実行ポリシーを変更します。

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

この設定 (`RemoteSigned`) の意味は以下の通りです。
- **自作したスクリプト (.ps1) は実行可能**
- **インターネットからダウンロードしたスクリプト (.ps1) は、信頼された発行元によって署名されている場合のみ実行可能**
