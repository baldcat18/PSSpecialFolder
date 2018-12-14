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

function newSpecialFolder {
	[OutputType([SpecialFolder])]
	param ([string]$Dir, [FolderOption]$Option = (@{}))
	
}
