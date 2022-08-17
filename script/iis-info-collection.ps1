function Create-Output {
    param (
        $WebsiteObj,
        $WebsiteAppPoolObj,
        $WebAppObj,
        $WebAppAppPoolObj

    )

    if ($WebAppObj -ne $null) {
        $WebAppName = $WebAppObj.path.Trim('/')
        $WebAppPhysicalPath = $WebAppObj.PhysicalPath
    }

    If ($WebAppAppPoolObj -ne $null) {
        $WebAppAppPoolName = $WebAppAppPoolObj.Name -join';'
        $WebAppAppPoolRuntimeVer = $WebappAppPoolObj.ManagedRuntimeVersion -join ';'
    }


    $OutputObj = [PSCustomObject]@{
                    Dns_HostName                          = $ComputerInfo.CsDNSHostName
                    Os_Name                               = $ComputerInfo.OsName
                    IIS_Version                           = $IISInfo.'Version String'
                    Website_Name                          = $WebsiteObj.Name
                    Website_Id                            = $WebsiteObj.Id -join ';'
                    Website_State                         = $WebsiteObj.State -join ';'
                    Website_PhysicalPath                  = $WebsiteObj.PhysicalPath
                    Website_Bindings                      = $WebsiteObj.Bindings.Collection -join ';'
                    Website_Attributes                    = ($WebsiteObj.Attributes | ForEach-Object { $_.name + "=" + $_.value }) -join ';'
                    Website_AppPool_Name                  = $WebsiteAppPoolObj.Name -join';'
                    Website_AppPool_State                 = $WebsiteAppPoolObj.State -join ';'
                    Website_AppPool_ManagedRuntimeVersion = $WebsiteAppPoolObj.ManagedRuntimeVersion -join ';'
                    Website_AppPool_ManagedPipelineMode   = $WebsiteAppPoolObj.ManagedPipelineMode -join ';'
                    Website_AppPool_StartMode             = $WebsiteAppPoolObj.StartMode -join ';'
                    WebApp_Name                           = $WebAppName
                    WebApp_PhysicalPath                   = $WebAppPhysicalPath
                    WebApp_AppPool_Name                   = $WebAppAppPoolName
                    WebApp_AppPool_ManagedRuntimeVersion  = $WebAppAppPoolRuntimeVer
    }

    $CsvOutput = ($OutputObj | Select @{N = 'Host Name'; E = {$_.Dns_HostName}},
                        @{N = 'OS Name'; E = {$_.Os_Name}},
                        @{N = 'IIS Version'; E = {$_.IIS_Version}},
                        @{N = 'Website ID'; E = {$_.Website_Id}},
                        @{N = 'Website Name'; E = {$_.Website_Name}},
                        @{N = 'Website State'; E = {$_.Website_State}},
                        @{N = 'Website Physical Path'; E = {$_.Website_PhysicalPath}},
                        @{N = 'Website Binding'; E = {$_.Website_Bindings}},
                        @{N = 'Website Application Pool'; E = {$_.Website_AppPool_Name}},
                        @{N = 'Website Application Pool State'; E = {$_.Website_AppPool_State}},
                        @{N = 'Website Application Pool CLR Version'; E = {$_.Website_AppPool_ManagedRuntimeVersion}},
                        @{N = 'Web App Name'; E = {$_.WebApp_Name}},
                        @{N = 'Web App Physical Path'; E = {$_.WebApp_PhysicalPath}},
                        @{N = 'Web App Application Pool'; E = {$_.WebApp_AppPool_Name}},
                        @{N = 'Web App Application Pool State'; E = {$_.WebApp_AppPool_State}}
    )

    $OutputObj | Select @{N = 'Website ID'; E = {$_.Website_Id}},
                        @{N = 'Website Name'; E = {$_.Website_Name}},
                        @{N = 'Website State'; E = {$_.Website_State}},
                        @{N = 'Website Physical Path'; E = {$_.Website_PhysicalPath}},
                        @{N = 'Website Binding'; E = {$_.Website_Bindings}},
                        @{N = 'Application Pool'; E = {$_.Website_AppPool_Name}},
                        @{N = 'Pool State'; E = {$_.Website_AppPool_State}},
                        @{N = 'Pool CLR Version'; E = {$_.Website_AppPool_ManagedRuntimeVersion}},
                        @{N = 'Web App Name'; E = {$_.WebApp_Name}},
                        @{N = 'Web App Physical Path'; E = {$_.WebApp_PhysicalPath}},
                        @{N = 'Web App Application Pool'; E = {$_.WebApp_AppPool_Name}},
                        @{N = 'Web App Application Pool State'; E = {$_.WebApp_AppPool_State}} | Format-Table -AutoSize

    Write-Csv $CsvOutput
}

function Write-Csv() {
    param(
        $CsvOutput
    )

    $OutputPath = "c:\" + $ComputerInfo.CsDNSHostName.ToLower() + "_output.csv"
    $CsvOutput | Export-Csv $OutputPath -NoTypeInformation -Append
}

$psVersion = "PowerShell Version: " + $PSVersionTable.PSVersion
Write-Output ""
Write-Output $psVersion

$policy=Get-ExecutionPolicy -Scope CurrentUser
Write-Output "Execution Policy for current user: $policy"

Write-Output ""
Write-Output "------------------------------"
Write-Output "| Basic Computer Information |"
Write-Output "------------------------------"

$ComputerInfo = Get-ComputerInfo

$ComputerInfo | Select @{N = 'Computer Name'; E = {$_.CsName}},
                          @{N = 'DNS Host Name'; E={$_.CsDNSHostName}},
                          @{N = 'Domain'; E = {$_.CsDomain}},
                          @{N = 'OS Name'; E = {$_.OsName}},
                          @{N = 'OS Version'; E = {$_.OsVersion}},
                          @{N = 'OS Build Number'; E = {$_.OsBuildNumber}},
                          @{N = 'OS Architecture'; E = {$_.OsArchitecture}},
                          @{N = 'Windows Product Name'; E = {$_.WindowsProductName}} | Format-Table

Write-Output ""
Write-Output "-------------------------------"
Write-Output "| .NET Framework Installation |"
Write-Output "-------------------------------"

Get-ChildItem ‘HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP’ -Recurse |
Get-ItemProperty -Name version -EA 0 |
Where { $_.PSChildName -Match ‘^(?!S)\p{L}’} |
Select @{N = '.NET Framework'; E = {$_.PSChildName}}, version | Format-Table


Write-Output ""
Write-Output "-------------------"
Write-Output "| IIS Information |"
Write-Output "-------------------"


#Import-Module ServerManager
if ((Get-WindowsFeature Web-Server).Installed) {
    Import-Module WebAdministration
    Import-Module IISAdministration

    #Get-ItemProperty -Path registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\InetStp\ | Select-Object | Format-Table

    $IISInfo = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\InetStp' |
                Select @{N = "Product Version"; E = {(Get-ItemProperty ($_.InstallPath + "\w3wp.exe")).VersionInfo.ProductVersion}},
                       @{N = 'Version String'; E = {$_.VersionString}},
                       @{N = 'Install Path'; E = {$_.InstallPath}})

    $IISInfo | Format-Table

    Write-Output ""
    Write-Output "------------------------"
    Write-Output "| IIS Application Pool |"
    Write-Output "------------------------"

    Get-IISAppPool | Format-Table

    Write-Output ""
    Write-Output "-------------------------------------------"
    Write-Output "| Website and Web Application Information |"
    Write-Output "-------------------------------------------"

    $Websites = Get-Website

    foreach ($Website in $Websites) {

        $WebApps = Get-WebApplication -Site $Website.Name
        $WebsiteAppPool = Get-IISAppPool -Name $Website.ApplicationPool

        if ($WebApps -eq $null) {
            Create-Output $Website $WebsiteAppPool $null $null
        } else {
            foreach ($WebApp in $WebApps) {
                $WebAppAppPool = Get-IISAppPool -Name $WebApp.applicationPool
                Create-Output $Website $WebsiteAppPool $WebApp $WebAppAppPool
            }
        }

     }
} else {
    Write-Output "IIS is not installed on this server."
    Write-Output ""

    $Output = ($ComputerInfo | Select @{N = 'Host Name'; E = {$_.CsDNSHostName}},
                            @{N = 'OS Name'; E = {$_.OsName}},
                            @{N = 'IIS Version'; E = {"Not Installed"}},
                            @{N = 'Website ID'; E = {}},
                            @{N = 'Website Name'; E = {}},
                            @{N = 'Website State'; E = {}},
                            @{N = 'Website Physical Path'; E = {}},
                            @{N = 'Website Binding'; E = {}},
                            @{N = 'Website Application Pool'; E = {}},
                            @{N = 'Website Pool State'; E = {}},
                            @{N = 'Website Pool CLR Version'; E = {}},
                            @{N = 'Web App Name'; E = {}},
                            @{N = 'Web App Physical Path'; E = {}},
                            @{N = 'Web App Application Pool'; E = {}},
                            @{N = 'Web App Application Pool State'; E = {}})
    Write-Csv $Output
}
