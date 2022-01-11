<# Make Sure to set execution policy and run as admin
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine #>
clear;
#Public CRM Functions. 
. "$PSScriptRoot\StartFunctions.ps1"

function Register-Assembly {
    <#
    .DESCRIPTION
    Input Parameters--
    AssemblyProperties : $assemblyProperties = $assemblyFile.GetName().FullName.Split(",= ".ToCharArray(), [StringSplitOptions]::RemoveEmptyEntries);
    AssemblyContent : Get-Base64 "Path"
    orgService : service of type Microsoft.Xrm.Tooling.Connector.CrmServiceClient
    
    Output Paramters--
    <Guid> 
    

    .EXAMPLE
        null
    #>

    PARAM
    (
        [parameter(Mandatory = $true)]$AssemblyProperties,
        [parameter(Mandatory = $true)]$AssemblyContent,
        [parameter(Mandatory = $true)]$orgService
    )
    Write-Host "       > Registering Assembly" -ForegroundColor Green;
    $PluginAssemblyId = New-CrmRecord -conn $orgService -EntityLogicalName pluginassembly -Fields @{"version" = $AssemblyProperties[2];
        "name" = $AssemblyProperties[0];
        "culture" = $AssemblyProperties[4];
        "publickeytoken" = $AssemblyProperties[6];
        "sourcetype" = New-CrmOptionSetValue -Value 0; #Database
        "isolationmode" = New-CrmOptionSetValue -Value 2; #Sandbox
        "content" = $AssemblyContent;
    }
    return $PluginAssemblyId;
}

function Register-Type {
    PARAM
    (
        [parameter(Mandatory = $true)]$PluginAssemblyId,
        [parameter(Mandatory = $true)]$TypeName,
        [parameter(Mandatory = $true)]$FriendlyName,
        [parameter(Mandatory = $false)]$WorkflowActivityGroupName,
        [parameter(Mandatory = $true)]$Name,
        [parameter(Mandatory = $true)]$orgService
    )
    Write-Host "       > Registering Type" -ForegroundColor Green;
    $PluginAssemblyRef = New-CrmEntityReference -EntityLogicalName pluginassembly -Id $PluginAssemblyId;
    if ($($TypeName -ne $null) -and $($FriendlyName -ne $null) -and $($Name -ne $null)) {
        [hastable]$fields = @{"pluginassemblyid" = $PluginAssemblyRef;
            "typename" = $TypeName;
            "friendlyname" = $FriendlyName;
            "name" = $Name;
        }
        if (($WorkflowActivityGroupName -ne $null) -and ($WorkflowActivityProjectName -ne $null)) {
            $fields = $fields + @{ "workflowactivitygroupname" = $WorkflowActivityGroupName;
            }
        }
        $PluginTypeId = New-CrmRecord -conn $orgService -EntityLogicalName plugintype -Fields $fields
    }
    else {
        $PluginTypeId = $null;
        Write-Host "failed! [Error : either TypeName, Friendlyname or Name is null]" -ForegroundColor Red;        
    }
    return $PluginTypeId;
}

function Register-PluginStep {
    <#
    .DESCRIPTION
    Input Parameters--
    Mode :  1 = Asynchronous 0 = Synchronous
    Rank : Execution Order {recommended 1}
    Stage : 10 = Pre-validation, 20 = Pre-operation, 40 = Post-operation
    Supported Deployment : 0 = serveronly, 1 = offlineonly, 2 = both
    Invocation Source : 0 = Parent, 1 = Child
    orgService : service of type Microsoft.Xrm.Tooling.Connector.CrmServiceClient
    
    Output Paramters--
    <Guid> 
    

    .EXAMPLE
    $pluginStep = Register-PluginStep -orgService $Service -PluginTypeId <Guid> -Mode 1 -Configuration "" -Rank 1 -Name Name -Stage 40 -SupportedDeployment 0 -MessageName "Create" -PrimaryEntityName "contact"
    #>     
    
    PARAM
    (
        [parameter(Mandatory = $true)]$PluginTypeId,
        [parameter(Mandatory = $true)]$Mode,
        [parameter(Mandatory = $true)]$Configuration,
        [parameter(Mandatory = $true)][int]$Rank,
        [parameter(Mandatory = $true)]$Name,
        [parameter(Mandatory = $true)]$Stage,
        [parameter(Mandatory = $true)]$Description,
        [parameter(Mandatory = $true)]$SupportedDeployment,
        [parameter(Mandatory = $true)]$MessageName,
        [parameter(Mandatory = $true)]$PrimaryEntityName,
        [parameter(Mandatory = $true)]$orgService
    )
    Write-Host "       > Registering Plugin Step" -ForegroundColor Green;
    $sdkmessage = Get-CrmRecordCustom -EntityLogicalName sdkmessage -FilterAttribute name -FilterValue $MessageName -orgService $orgService
    $sdkmessagefilter = Get-CrmRecordCustom -EntityLogicalName sdkmessagefilter -FilterAttribute primaryobjecttypecode -FilterValue $PrimaryEntityName -orgService $orgService
    $PluginTypeRef = New-CrmEntityReference -EntityLogicalName plugintype -Id $PluginTypeId;
    if ($($Mode -ne $null) -and $($Stage -ne $null) -and $($Name -ne $null) -and $($SupportedDeployment -ne $null)) {
        if ($Rank -eq $null) {$Rank = 1}
        $PluginStepId = New-CrmRecord -conn $orgService -EntityLogicalName sdkmessageprocessingstep -Fields @{"plugintypeid" = $PluginTypeRef;
            "mode" = New-CrmOptionSetValue -Value $Mode; 
            "configuration" = "";
            "name" = $Name; 
            "rank" = $Rank; 
            "description" = $Description;
            "stage" = New-CrmOptionSetValue -Value $Stage; 
            "supporteddeployment" = New-CrmOptionSetValue -Value $SupportedDeployment; 
            "invocationsource" = New-CrmOptionSetValue -Value 0; 
            "sdkmessageid" = $sdkmessage.ToEntityReference();
            "sdkmessagefilterid" = $sdkmessagefilter.ToEntityReference();
        }
    }
    else {
        $PluginStepId = $null;
        Write-Host "failed! [Error : either Mode, Stage, Supported Deployment Type or Name is null]" -ForegroundColor Red;        
    }

    return $PluginStepId;
}

function Register-PluginImage {
    <#
    .DESCRIPTION
    Input Parameters--
    ImageType :  1-Post Image 2-Both 0-PreImage
    Attributes : comma separated string
    messagepropertyname : For Create "Id", For Update "Target" 
    orgService : service of type Microsoft.Xrm.Tooling.Connector.CrmServiceClient
    
    Output Paramters--
    <Guid> 
    

    .EXAMPLE
    $image = Register-PluginImage -orgService $Service -PluginStepId <Guid> -EntityAlias "test" -ImageType 1 -Name "Test" -Attributes "address1_city,fullname,ownerid" -MessagePropertyName "Id" 
    #>
    
    PARAM
    (
        [parameter(Mandatory = $true)]$PluginStepId,
        [parameter(Mandatory = $true)]$EntityAlias,
        [parameter(Mandatory = $true)]$ImageType,
        [parameter(Mandatory = $true)]$Name,
        [parameter(Mandatory = $true)]$Attributes,
        [parameter(Mandatory = $true)]$MessagePropertyName,
        [parameter(Mandatory = $true)]$orgService
    )
    Write-Host "       > Registering Plugin Image" -ForegroundColor Green;
    $PluginStepRef = New-CrmEntityReference -EntityLogicalName sdkmessageprocessingstep -Id $PluginStepId;
    if ($($EntityAlias -ne $null) -and $($ImageType -ne $null) -and $($MessagePropertyName -ne $null) -and $($Name -ne $null) -and $($Attributes -ne $null)) {
        $PluginImageId = New-CrmRecord -conn $orgService -EntityLogicalName sdkmessageprocessingstepimage -Fields @{
            "entityalias" = $EntityAlias; 
            "name" = $Name; 
            "imagetype" = New-CrmOptionSetValue -Value $ImageType; 
            "attributes" = $Attributes; 
            "sdkmessageprocessingstepid" = $PluginStepRef;
            "messagepropertyname" = $MessagePropertyName;
        }
    }
    else {
        $PluginImageId = $null;
        Write-Host "failed! [Error : Either Mode, Stage, Supported Deployment Type, Attributes or Name is null]" -ForegroundColor Red;        
    }
    return $PluginImageId;
}

[xml]$config = Get-Content "$PSScriptRoot\Config.xml"
$ConnectionString = $config.Configuration.CrmConnectionString;
$Service = New-Object -TypeName Microsoft.Xrm.Tooling.Connector.CrmServiceClient -ArgumentList $ConnectionString
$assemblyConfiguration = "debug";
$d = Get-Date;
Write-Host "$d - Deploy Assemblies ($assemblyConfiguration) start" -ForegroundColor Cyan;
$rows = $config.SelectNodes("/Configuration/Assemblies/Assembly");
foreach ($row in $rows) {
    # -----Type------------
    $Type_FriendlyName = $row.Solution.PluginType | select FriendlyName;
    $Type_Name = $row.Solution.PluginType | select Name;
    $Type_TypeName = $row.Solution.PluginType | select TypeName;
    $Type_WorkflowActivityGroupName = $($row.Solution.PluginType | select WorkflowActivityGroupName).WorkflowActivityGroupName;

    # -----Plugin Step----------
    $PluginStep_Name = $row.Solution.PluginType.Step | select Name;
    $PluginStep_Description = $row.Solution.PluginType.Step | select Description;
    $PluginStep_MessageName = $row.Solution.PluginType.Step | select MessageName;
    $PluginStep_Mode = $row.Solution.PluginType.Step | select Mode;
    $PluginStep_PrimaryEntityName = $row.Solution.PluginType.Step | select PrimaryEntityName;
    $PluginStep_Rank = $row.Solution.PluginType.Step | select Rank;
    $PluginStep_SecureConfig = $row.Solution.PluginType.Step | select SecureConfiguration;
    $PluginStep_Stage = $row.Solution.PluginType.Step | select Stage;
    $PluginStep_SupportedDeployment = $row.Solution.PluginType.Step | select SupportedDeployment;

    # -----Plugin Image----------
    $PluginImage_Attributes = $row.Solution.PluginType.Step.Image | select Attributes;
    $PluginImage_EntityAlias = $row.Solution.PluginType.Step.Image | select EntityAlias;
    $PluginImage_Name = $row.Solution.PluginType.Step.Image | select Name;
    $PluginImage_MessagePropertyName = $row.Solution.PluginType.Step.Image | select MessagePropertyName;
    $PluginImage_ImageType = $row.Solution.PluginType.Step.Image | select ImageType;

    $assemblyToDeploy = $($row.Solution | select Assembly).Assembly
    $assemblyPath = $row.Path.ToString();
    $Type = $($row.Solution | select Type).Type
    $assembly = Get-ChildItem $assemblyPath -recurse -include $assemblyToDeploy;
    $filePath = $assembly.FullName.ToString();
    if ($filePath.Contains("bin") -and $filePath.ToLower().Contains($assemblyConfiguration)) {

        Write-Host "  > Deploying assembly $filePath ...";			
        $assemblyFile = [System.Reflection.Assembly]::LoadFile($filePath);
        $assemblyProperties = $assemblyFile.GetName().FullName.Split(",= ".ToCharArray(), [StringSplitOptions]::RemoveEmptyEntries);
        $assemblyShortName = $assemblyProperties[0];
        $assemblyContent = Get-Base64 $filePath;
        Write-Host "  > Searching assembly $assemblyShortName in Dynamics CRM..." -NoNewline;
        $crmAssembly = Get-Assembly -Name $assemblyShortName -orgService $Service;
        if ($crmAssembly -ne $null) {
            Write-Host "found!" -ForegroundColor Green; 
            try {
                #Unregister the assembly & Register
                Write-Host "     > Unregistering assemblies" -ForegroundColor Red;
                $crmpluginType = Get-CrmRecordCustom -orgService $Service -EntityLogicalName plugintype -FilterAttribute typename -FilterValue $Type_TypeName.TypeName
                if ($crmpluginType -ne $null -and $PluginStep_Name -ne $null) {
                    $crmpluginStep = Get-CrmRecordCustom -orgService $Service -EntityLogicalName sdkmessageprocessingstep -FilterAttribute name -FilterValue $PluginStep_Name.Name
                    if ($crmpluginStep -ne $null -and $PluginImage_Name -ne $null) {
                        $crmpluginImage = Get-CrmRecordCustom -orgService $Service -EntityLogicalName sdkmessageprocessingstepimage -FilterAttribute name -FilterValue $PluginImage_Name.Name
                        if ($crmpluginImage -ne $null) {
                            $Service.Delete($crmpluginImage.LogicalName, $crmpluginImage.Id);
                        }
                        $Service.Delete($crmpluginStep.LogicalName, $crmpluginStep.Id);
                    }
                    $Service.Delete($crmpluginType.LogicalName, $crmpluginType.Id);
                }
                $Service.Delete($crmAssembly.LogicalName, $crmAssembly.Id);
            }
            catch [Exception] {
                Write-Host "failed! [Error : $_.Exception]" -ForegroundColor Red;
                break;
            }
        }
        Write-Host "     > Registering assemblies" -ForegroundColor Green;
        try { 
            #Create Plugin Assembly Record
            $pluginAssembly = Register-Assembly -AssemblyProperties $assemblyProperties -AssemblyContent $assemblyContent -orgService $Service
        }
        catch [Exception] {
            Write-Host "        > Register Assembly failed! [Error : $_.Exception]" -ForegroundColor Red;
            break;
        }

        #Create Plugin Type Record
        try {
            if ($Type_WorkflowActivityGroupName -eq $null) {
                $pluginType = Register-Type -orgService $Service -PluginAssemblyId $pluginAssembly.Guid -TypeName $Type_TypeName.TypeName -FriendlyName $Type_FriendlyName.FriendlyName -Name $Type_Name.Name 
            }
            else {
                $pluginType = Register-Type -orgService $Service -PluginAssemblyId $pluginAssembly.Guid -TypeName $Type_TypeName.TypeName -FriendlyName $Type_FriendlyName.FriendlyName -Name $Type_Name.Name -WorkflowActivityGroupName $Type_WorkflowActivityGroupName #-WorkflowActivityProjectName $Type_WorkflowActivityProjectName
            }
        }
        catch [Exception] {
            Write-Host "        > Register Type failed! [Error : $_.Exception]" -ForegroundColor Red;
            break;
        }

        if ($Type -eq "Plugin") {
            #Create Plugin Step Record
            try {
                $pluginStep = Register-PluginStep -orgService $Service -PluginTypeId $pluginType.Guid -Description $PluginStep_Description.Description -Mode $PluginStep_Mode.Mode -Configuration $PluginStep_SecureConfig.SecureConfiguration -Rank $PluginStep_Rank.Rank -Name $PluginStep_Name.Name -Stage $PluginStep_Stage.Stage -SupportedDeployment $PluginStep_SupportedDeployment.SupportedDeployment -MessageName $PluginStep_MessageName.MessageName -PrimaryEntityName $PluginStep_PrimaryEntityName.PrimaryEntityName
            }
            catch [Exception] {
                Write-Host "        > Register Plugin Step failed! [Error : $_.Exception]" -ForegroundColor Red;
                break;
            }

            #Create Plugin Image Record    
            try {
                $image = Register-PluginImage -orgService $Service -PluginStepId $pluginStep.Guid -EntityAlias $PluginImage_EntityAlias.EntityAlias -ImageType $PluginImage_ImageType.ImageType -Name $PluginImage_Name.Name -Attributes $PluginImage_Attributes.Attributes -MessagePropertyName $PluginImage_MessagePropertyName.MessagePropertyName
            }
            catch [Exception] {
                Write-Host "        > Register Plugin Image failed! [Error : $_.Exception]" -ForegroundColor Red;
                break;
            }
        }
    }
}

$d = Get-Date;
Write-Host "$d - Deploy Assemblies ($assemblyConfiguration) done" -ForegroundColor Cyan; 
break;  


