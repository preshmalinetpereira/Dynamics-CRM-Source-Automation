function Add-Crm-Sdk {
    # Load SDK assemblies
    if($(Get-Module -Name Microsoft.Xrm.Data.PowerShell) -eq $null){
    Import-Module "$PSScriptRoot\Microsoft.Xrm.Data.PowerShell\Microsoft.Xrm.Data.PowerShell.psm1";
    }
    Add-Type -Path "$PSScriptRoot\Microsoft.Xrm.Sdk.dll";
    Add-Type -Path "$PSScriptRoot\Microsoft.Crm.Sdk.Proxy.dll";
    Add-Type -Path "$PSScriptRoot\Microsoft.IdentityModel.Clients.ActiveDirectory.dll";

    try {
        Add-Type -Path "$PSScriptRoot\Microsoft.Xrm.Tooling.Connector.dll";
    }
    catch [Exception] {
        Write-Host "failed! [Error : $_.Exception]" -ForegroundColor Red;
        break;
    }
}

function Get-Configuration() {
    $configFilePath = "$PSScriptRoot\Config.xml";
    $content = Get-Content $configFilePath;
    return [xml]$content;
}

function Get-Assembly {
    <#
    .DESCRIPTION
    Parameters--
    Name : Assembly Name
    orgService : service of type Microsoft.Xrm.Tooling.Connector.CrmServiceClient

    .EXAMPLE
    Get-Assembly -Name Test.TestPlugin -orgService $service
    #>
    PARAM
    (
        [parameter(Mandatory = $true)]$Name,
        [parameter(Mandatory = $true)]$orgService
    )

    $query = New-Object -TypeName Microsoft.Xrm.Sdk.Query.QueryExpression -ArgumentList "pluginassembly";
    $query.Criteria.AddCondition("name", [Microsoft.Xrm.Sdk.Query.ConditionOperator]::Equal, $Name);
    $query.ColumnSet = New-Object -TypeName Microsoft.Xrm.Sdk.Query.ColumnSet -ArgumentList $true;
    $results = $orgService.RetrieveMultiple($query);
    $records = $results.Entities;

    if ($records.Count -eq 1) {
        return $records[0];
    }
    return $null;
}

function Get-Base64 {
    <#
    .DESCRIPTION
    Parameters--
    Path : Path to dll file

    .EXAMPLE
    Get-Base64 -Path "D:/Path"
    #>
    PARAM
    (
        [parameter(Mandatory = $true)]$path
    )
    
    $content = [System.IO.File]::ReadAllBytes($path);
    $content64 = [System.Convert]::ToBase64String($content);
    return $content64;
}

function Get-CrmRecordCustom {
    <#
    .DESCRIPTION
    Input Parameters--
    EntityLogicalName : Path to dll file
    FilterAttribute : Attribute name to filter on
    FilterValue : Value for filter
    orgService : service of type Microsoft.Xrm.Tooling.Connector.CrmServiceClient
    
    Output Paramters--
    <Entity> $records
    

    .EXAMPLE
    Get-CrmRecordCustom -orgService $Service -EntityLogicalName plugintype -FilterAttribute pluginassemblyid -FilterValue "78521358525045032505324"
    #>
    PARAM
    (
        [parameter(Mandatory = $true)]$EntityLogicalName,
        [parameter(Mandatory = $true)]$FilterAttribute,
        [parameter(Mandatory = $true)]$FilterValue,
        [parameter(Mandatory = $true)]$orgService
    )

    $query = New-Object -TypeName Microsoft.Xrm.Sdk.Query.QueryExpression -ArgumentList $EntityLogicalName;
    $query.Criteria.AddCondition($FilterAttribute, [Microsoft.Xrm.Sdk.Query.ConditionOperator]::Equal, $FilterValue);
    $query.ColumnSet = New-Object -TypeName Microsoft.Xrm.Sdk.Query.ColumnSet -ArgumentList $true;
    $query.TopCount = 1;
    $results = $orgService.RetrieveMultiple($query);
    $records = $results.Entities;
    return $records;
}

Add-Crm-Sdk;