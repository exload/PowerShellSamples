<# 
.DESCRIPTION 
    Скрипт создания из MSI файла Application.

.PARAMETER path
    Путь к MSI файлу

.PARAMETER CMSite
    Имя сайта CCM, в контексте которого создаётся приложение

.PARAMETER Description
    Опциональный параметр для описания приложения.

.EXAMPLE
    .\new-app.ps1 -path "\\server\Far\x64\v3.0 build 5577\Far30b5577.x64.20200327.msi" -Description 'TEST!!!!'
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory=$True)]
    [string]$path,

    [Parameter(Mandatory=$false)]
    [string]$Description,

    [Parameter(Mandatory=$false)]
    [string]$CMSite = '%имя сайта SCCM%:'
)
function Get-MSIinfo{
    <#
    .SYNOPSIS
        Функция извлекает характеристики из MSI файла

    .PARAMETER Path
        Путь к MSI файлу

    .INPUTS
        Путь к MSI файлу

    .OUTPUTS
        Возвращает массив со значениями

    .EXAMPLE
        Get-MSIinfo -path "\\server\Far\x64\v3.0 build 5577\Far30b5577.x64.20200327.msi"
    #>
    [CmdletBinding()]
    param (
    # Путь к MSI файлу
    [Parameter(Mandatory=$true)]
    [string]
    $Path
    )
    
    process {
        $com_object = New-Object -com WindowsInstaller.Installer
        try {
            $FilePath = Get-ChildItem $path
            $FilePath
            $FilePath.FullName
            $database = $com_object.GetType().InvokeMember(
                "OpenDatabase",
                "InvokeMethod",
                $Null,
                $com_object,
                @($FilePath.FullName, 0)
            )

            $query = "SELECT * FROM Property"
            $View = $database.GetType().InvokeMember(
                "OpenView",
                "InvokeMethod",
                $Null,
                $database,
                ($query)
            )

            $View.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $View, $Null)

            $record = $View.GetType().InvokeMember(
                "Fetch",
                "InvokeMethod",
                $Null,
                $View,
                $Null
            )

            $msi_props = @{}
            while ($null -ne $record) {
                $msi_props[$record.GetType().InvokeMember("StringData", "GetProperty", $Null, $record, 1)] = $record.GetType().InvokeMember("StringData", "GetProperty", $Null, $record, 2)
                $record = $View.GetType().InvokeMember(
                    "Fetch",
                    "InvokeMethod",
                    $Null,
                    $View,
                    $Null
                )
            }
            $msi_props
        }
        catch{
            throw "Failed to get MSI file properties the error was: {0}." -f $_
        }
    }
}
Import-Module 'C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1' -ErrorAction Stop

$current_location  = Get-Location

$app_info = @{}
$app_info = Get-MSIinfo $path

$ProductCode = $app_info.ProductCode
$ProductVersion = $app_info.ProductVersion
$ProductName = $app_info.ProductName
$Manufacturer = $app_info.Manufacturer

$app_name = $ProductName + ' ' + $ProductVersion
$Deployment_Type_Name = 'Installation ' + $app_name

$msi_file = Get-ChildItem $path
$install_command = 'msiexec /i ' + $msi_file.Name + ' /q REBOOT=REALLYSUPPRESS'
$uninstall_command = 'msiexec /x ' + $ProductCode + ' /q'

Set-Location $CMSite

$new_app = New-CMApplication -Name $app_name -LocalizedName $ProductName -Publisher $Manufacturer -SoftwareVersion $ProductVersion -Description $Description

$new_app | Add-CMMsiDeploymentType -ContentLocation $msi_file.FullName -InstallCommand $install_command `
    -InstallationBehaviorType InstallForSystem -LogonRequirementType WhetherOrNotUserLoggedOn -ProductCode $ProductCode -RebootBehavior BasedOnExitCode `
    -UninstallCommand $uninstall_command -UserInteractionMode Hidden -DeploymentTypeName $Deployment_Type_Name | Out-Null

set-location $current_location

Write-Output "Подготовлено к развёртыванию приложение: $app_name"