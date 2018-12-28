Set-StrictMode -Version Latest

class SpecialFolder {
	[string]$Title
	[string]$Path
	
	hidden [string]$Dir
	hidden [__ComObject]$FolderItem
}

class FolderOption {
	[string]$Title
	[string]$Path
}

$shell = New-Object -ComObject Shell.Application

function newSpecialFolder {
	[OutputType([SpecialFolder])]
	param ([string]$Dir, [FolderOption]$Option = (@{}))
	
	if (!$Dir) { return }
	if ($Dir.Substring(0, 2) -eq "\\") { $Dir = "file:" + $Dir }
	elseif ($Dir.Substring(0, 6) -ne "shell:" -and $Dir.Substring(0, 5) -ne "file:") { $Dir = "file:\\\" + $Dir }
	
	try { $folder = $shell.NameSpace($Dir) }
	catch { return }
	
	if (!$folder) { return }
	$folderItem = $shell.NameSpace($Dir).Self
	
	$title =
		if ($Dir -match "^shell:(?:(?:\w|\s)+)$") { $Dir.Substring(6) }
		elseif ($Dir -match "^shell:::.*\{\w{8}-\w{4}-\w{4}-\w{4}-\w{12}\}$") {
			$clsid = $Dir.Substring($Dir.Length - 38)
			(Get-Item "Microsoft.PowerShell.Core\Registry::HKEY_CLASSES_ROOT\CLSID\$clsid").GetValue("")
		}
		else { $Dir -replace "^.+\\(.+?)$", "`$1" }
	
	$path = $folderItem.Path
	if ($path.Substring(0,2) -eq "::") { $path = "shell:" + $path }
	
	if ($Option.Title) { $title = $Option.Title }
	if ($Option.Path) { $path = $Option.Path }
	
	return [SpecialFolder]@{ Title = $title; Path = $path; Dir = $Dir; FolderItem = $folderItem }
}

function newShellCommand {
	[OutputType([SpecialFolder])]
	param ([string]$Dir)
	
	if (!$Dir) { return }
	
	$path = "Microsoft.PowerShell.Core\Registry::HKEY_CLASSES_ROOT\CLSID\$($Dir.Substring($Dir.Length - 38))"
	if (!(Test-Path $path)) { return }
	
	return [SpecialFolder]@{ Title = (Get-Item $path).GetValue(""); Path = $Dir; Dir = $Dir }
}
