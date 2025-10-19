# PSSpecialFolder

Windows で使われている特殊フォルダーのフルパスを取得する PowerShell モジュールです。

## 最小 PowerShell バージョン

5.1

## サポートする Windows バージョン

Windows 10 バージョン 22H2<br>
Windows 11 バージョン 23H2、24H2、25H2

## 提供する関数

### Get-SpecialFolder

Windows で使われている特殊フォルダーのフルパスを取得します。

### Get-SpecialFolderPath

指定した特殊フォルダーのフルパスを取得します。

### New-SpecialFolder

指定した特殊フォルダーを作成します。

### Show-SpecialFolder

Windows で使われている特殊フォルダーのフルパスをダイアログ上に表示します。

ダブルクリックすると、そのフォルダーを開きます。

右クリックすると、コンテキストメニューが表示されます。
フォルダーのパスをクリップボードにコピー、PowerShell やコマンドプロンプトの起動、フォルダーのプロパティの表示ができます。

この関数は WPF を利用しているので、PowerShell Core 6.x では使用できません。

## インストール

Install-Module -Name PSSpecialFolder
