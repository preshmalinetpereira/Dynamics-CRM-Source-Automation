#Pack&ImportSolution.ps1
clear;
. "$PSScriptRoot\StartFunctions.ps1"
#Import Settings from Config file
[xml]$ConfigFile = Get-Content "$PSScriptRoot\Settings.xml"
#Parameters

$Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $ConfigFile.Settings.UserSettings.UserName, $password


#Solution Packager Location
$solutionPackager = $ConfigFile.Settings.DirectorySettings.SolutionPackager;

#Location where the content will be stored and retrieved from
$fileLocation = $ConfigFile.Settings.DirectorySettings.SolutionFiles;
$solutionFilesFolder = $ConfigFile.Settings.DirectorySettings.SolutionFilesSrc;

#Name of Solution
$fileNameforPack = $ConfigFile.Settings.DirectorySettings.ImportedFileName;

#Map file
$mappingFile = $ConfigFile.Settings.DirectorySettings.MapFile;

#Connection Object
$global:conn = Connect-CrmOnline -Credential $Cred -ServerUrl $ConfigFile.Settings.CRMSettings.ToServerURL

#Pack the Solution
& "$PSScriptRoot\$solutionPackager" /action:Pack /zipfile:"$fileLocation\$fileNameforPack" /folder:"$solutionFilesFolder"  /map:"D:\DynamicsCRMAutomation\$mappingFile"

#Import the Solution to CRM
$importSol = Import-CrmSolution -conn $conn -SolutionFilePath "$fileLocation\$fileNameforPack" 
Write-Output $importSol




