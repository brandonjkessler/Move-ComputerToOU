#Requires -RunAsAdministrator
<#
    .SYNOPSIS


    .DESCRIPTION


    .PARAMETER RegistryPath

    .PARAMETER RegistryKey

    .PARAMETER RootOUPath

    .PARAMETER RegistryOU

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
    [parameter(HelpMessage='Registry Key Property where Asset Family value is written. Default is AssetFamily')]
    [string]$WhatIf
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
$DestinationOU = (Get-ADDomain).DistinguishedName

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
    Get-ADComputer -Identity $env:COMPUTERNAME | Move-ADObject -TargetPath $DestinationOU
}
