<#
    .SYNOPSIS


    .DESCRIPTION


    .PARAMETER RegistryPath
    The Path in the registry where the key to query for will be located.

    .PARAMETER RegistryKey
    The key in the registry that is queried against.

    .PARAMETER RootOUPath
    The path that devices are placed in, but before the OU that's in the registry.  
    Example: the Domain is some.company.org with an OU called 'Devices' under it.  
    The RootOUPath parameter would then be "Devices".

    .PARAMETER RegistryOU
    The name of the property in the RegistryKey that's being queried against.  
    Example: if the registry Key property is "OrganizationalUnit" then the parameter would be that.  
    Default is OU

    .PARAMETER AssetFamily

    .INPUTS
    Accepts inputs from pipeline

    .OUTPUTS
    None.

    .EXAMPLE


    .LINK

#>


param(
    [parameter(ValueFromPipelineByPropertyName,
    HelpMessage='Path to registry where the key will be created. Default is HKLM:\SOFTWARE')]
    [string]$RegistryPath = 'HKLM:\SOFTWARE',
    [parameter(ValueFromPipelineByPropertyName,
    HelpMessage='Registry where values will be written. Default is CustomInv')]
    [string]$RegistryKey = 'CustomInv',
    [parameter(ValueFromPipelineByPropertyName,
    HelpMessage='Root OU that devices are in, not including AD Root. Default is Devices')]
    [string]$RootOUPath = 'Devices',
    [parameter(ValueFromPipelineByPropertyName,
    HelpMessage='Registry Key Property where OU value is written. Default is OU')]
    [string]$RegistryOU = 'OU',
    [parameter(ValueFromPipelineByPropertyName,
    HelpMessage='Registry Key Property where Asset Family value is written. Default is AssetFamily')]
    [string]$AssetFamily = 'AssetFamily',
    [parameter(HelpMessage='If set, returns value and does not move device.')]
    [switch]$WhatIf,
    [parameter(HelpMessage='Task Sequence Variable to use if running under a task sequence.')]
    [string]$TSVariable
)

$Registry = "$RegistryPath\$RegistryKey"

## Test Registry Location
if((Test-Path -Path $Registry) -ne $true){
    if((Test-Path -Path "$RegistryPath") -ne $true){
        Write-Error "Registry Path does not exist."
        Exit 1
    } else {
        Write-Error "Registry Key does not exist."
        Exit 1
    }
}

## Get Reg Values
Get-ItemProperty -Path $Registry | ForEach-Object {
    if($_ -match $RegistryOU){
        $OU = (Get-ItemProperty -Path $Registry).$RegistryOU
    }
    if($_ -match $AssetFamily){
        $DeviceType = (Get-ItemProperty -Path $Registry).$AssetFamily
    }
}

## Get Domain and Start OU Path
$Root = [ADSI]"LDAP://RootDSE"
$DestinationOU = $Root.Get("rootDomainNamingContext")



## Test if Parameter has / in it and split if needed to create multiple OU entries in $DestinationOU
if($RootOUPath -match '/'){
    ($RootOUPath).Split('/') | ForEach-Object{
        $DestinationOU = "OU=$_,$DestinationOU"
    }
} else {
    $DestinationOU = "OU=$RootOUPath,$DestinationOU"
}

## Add Registry OU entry at end
$DestinationOU = "OU=$OU,$DestinationOU"

## Add final OU for device
Switch($DeviceType){
    'Mobile'{$DestinationOU = "OU=Laptops,$DestinationOU"}
    'Desktop'{$DestinationOU = "OU=Desktops,$DestinationOU"}
}

if($WhatIf -eq $true){
    Return $DestinationOU
} else {
    Try{
        $TSEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment
        $TSEnv.Value("$TSVariable") = $DestinationOU
    }Catch{
        ## Inspired by https://social.technet.microsoft.com/forums/scriptcenter/en-US/37ab13a4-4ddb-460a-8a6a-0eac5887e0c0/using-adsi-and-ldap-to-move-to-an-ou
        $SystemInfo = New-Object -ComObject "ADSystemInfo"
        $ComputerDN = $SystemInfo.GetType().InvokeMember("ComputerName", "GetProperty", $Null, $SystemInfo, $Null)
        $Computer = [ADSI]"LDAP://$ComputerDN"
        $Computer.psbase.MoveTo($([ADSI]"LDAP://$DestinationOU")) 
    }
}
