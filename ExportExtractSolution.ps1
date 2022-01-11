#Export&ExtractSolution.ps1
clear;
. "$PSScriptRoot\StartFunctions.ps1"

#Import Settings from Config file
[xml]$ConfigFile = Get-Content "$PSScriptRoot\Settings.xml"

#Parameters
$password = $ConfigFile.Settings.UserSettings.Password
$Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $ConfigFile.Settings.UserSettings.UserName, $password

#Solution Packager Location
$solutionPackager = $ConfigFile.Settings.DirectorySettings.SolutionPackager

#Location where the content will be stored and retrieved from
$fileLocation = $ConfigFile.Settings.DirectorySettings.SolutionFiles
$solutionFilesFolder = $ConfigFile.Settings.DirectorySettings.SolutionFilesSrc

#Name of Solution
$solutionZipFileName = $ConfigFile.Settings.DirectorySettings.ExportedFileName

#Map file
$mappingFile = $ConfigFile.Settings.DirectorySettings.MapFile

#Connection Object
$global:conn = Connect-CrmOnline -Credential $Cred -ServerUrl $ConfigFile.Settings.CRMSettings.FromServerURL

#Export the Solution from CRM
$exportSol = Export-CrmSolution  -conn $conn -SolutionName $ConfigFile.Settings.CRMSettings.SolutionName -SolutionFilePath $fileLocation -SolutionZipFileName $solutionZipFileName
Write-Output $exportSol

#To extract solution using Solution Packager
& "$PSScriptRoot\$solutionPackager" /action:Extract /zipfile:"$fileLocation\$solutionZipFileName" /folder:"$solutionFilesFolder"  /map:"$PSScriptRoot\$mappingFile" /allowDelete:Yes



