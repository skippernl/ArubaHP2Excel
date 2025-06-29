<#
.SYNOPSIS
Document-ALL parses the all the configuration from a cisco device into a Excel files.
.DESCRIPTION
Document-ALL parses the all the configuration from a cisco device into a Excel files.
.PARAMETER SrcDir
[REQUIRED] This is the path to the Cisco config/credential files
Optional switches are explained in the Cisco2Excel script  
.NOTES
Author: Xander Angenent (@XaAng70)
Idea: Drew Hjelm (@drewhjelm) (creates csv of ruleset only)
Last Modified: 2020/11/05
#>
Param
(
    [Parameter(Mandatory = $true)]
    $SrcDir,
    [switch]$SkipFilter = $false
)

Function CreateFiles ($ConfigFilesArray) {
    Foreach ( $ConfigFile in $ConfigFilesArray ) {
        $PSArgument = $BaseArgumentList
        $PSArgument += "-HPConfig '$($ConfigFile.FullName)'"
        Invoke-Expression ".\HP2Excel.ps1 $PSArgument"
    }
}

$BaseArgumentList = @()
$PSArgument = @()
if ($SkipFilter) { $BaseArgumentList += "-Skipfilter" }

#Get All *.conf Files
if (!(Test-Path $SrcDir)) {
    Write-Output "Path not found stopping script."
    exit 1
}
$GetConfigFiles = $SrcDir + "\" + "*.*"
$ConfigFiles = Get-ChildItem $GetConfigFiles | Sort-Object Name
If ($ConfigFiles) { CreateFiles $ConfigFiles }
else { 
    Write-Output "No config (*.cfg) files found."
}
