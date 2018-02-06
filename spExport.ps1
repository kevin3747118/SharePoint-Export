[CmdletBinding()]
Param(
#[Parameter(Mandatory=$true)][System.String]$Url = $(Read-Host -prompt "Web Url"),
[Parameter(Mandatory=$true)][System.String]$Library = $(Read-Host -prompt "Document Library")
)

#Load sp2007 module
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")

$site = new-object microsoft.sharepoint.spsite("http://win-scfojl5gm38:20834/_layouts/viewlsts.aspx")
$web = $site.OpenWeb()
$site.Dispose()

$folder = $web.GetFolder($Library)
$folder

if(!$folder.Exists){
	Write-Error "Library cannot be found !"
	$web.Dispose()
	return
}

$directory = $pwd.Path

$rootDirectory = Join-Path $pwd $folder.Name

if (Test-Path $rootDirectory) {
	Write-Error "The folder $Library in the current directory already exists, please remove it !"
	$web.Dispose()
	return
}

#progress variables
$global:counter = 0
$global:total = 0
#recursively count all files to pull
function count($folder) {
	if ($folder.Name -ne "Forms") {
		$global:total += $folder.Files.Count
		$folder.SubFolders | Foreach { count $_ }
	}
}
write "counting files, please wait..."
count $folder
write "files count $global:total"

function progress($path) {
	$global:counter++
	$percent = $global:counter / $global:total * 100
	write-progress -activity "Pulling documents from $Library" -status $path -PercentComplete $percent
}

#Write file to disk
function Save ($file, $directory) {
	$data = $file.OpenBinary()
	$path = Join-Path $directory $file.Name
	progress $path
	[System.IO.File]::WriteAllBytes($path, $data)
}

$formsDirectory = Join-Path $rootDirectory "Forms"

function Pull($folder, [string]$directory) {
	$directory = Join-Path $directory $folder.Name
	if ($directory -eq $formsDirectory) {
		return
	}
	mkdir $directory | out-null
	
	$folder.Files | Foreach { Save $_ $directory }

	$folder.Subfolders | Foreach { Pull $_ $directory }
}

Write "Copying files"
Pull $folder $directory

$web.Dispose()
