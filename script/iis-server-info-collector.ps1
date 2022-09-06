function Map-OutputObject {
    param (
        [parameter(mandatory = $true)][object] $WebsiteObj,
        [parameter(mandatory = $true)][object] $WebsiteAppPoolObj,
        [object] $WebAppObj = $null,
        [object] $WebAppAppPoolObj = $null

    )

    if ($WebAppObj -ne $null) {
        $WebAppName = $WebAppObj.path.Trim('/')
        $WebAppPhysicalPath = $WebAppObj.PhysicalPath
    }

    If ($WebAppAppPoolObj -ne $null) {
        $WebAppAppPoolName = $WebAppAppPoolObj.Name -join';'
        $WebAppAppPoolState = $WebAppAppPoolObj.State -join ';'
        $WebAppAppPoolRuntimeVer = $WebappAppPoolObj.ManagedRuntimeVersion -join ';'
    }

    $OutputObject = [PSCustomObject]@{
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
                    WebApp_AppPool_State                  = $WebAppAppPoolState
                    WebApp_AppPool_ManagedRuntimeVersion  = $WebAppAppPoolRuntimeVer
    }

    return $OutputObject
}

function Create-CsvOutputObject {
    param (
        [parameter(mandatory = $true)][object] $OutputObject
    )

    $CsvOutput = ($OutputObject |
      Select @{N = 'Host Name';                      E = {$_.Dns_HostName}},
             @{N = 'OS Name';                        E = {$_.Os_Name}},
             @{N = 'IIS Version';                    E = {$_.IIS_Version}},
             @{N = 'Website ID';                     E = {$_.Website_Id}},
             @{N = 'Website Name';                   E = {$_.Website_Name}},
             @{N = 'Website State';                  E = {$_.Website_State}},
             @{N = 'Website Physical Path';          E = {$_.Website_PhysicalPath}},
             @{N = 'Website Binding';                E = {$_.Website_Bindings}},
             @{N = 'Website Application Pool';       E = {$_.Website_AppPool_Name}},
             @{N = 'Website Application Pool State'; E = {$_.Website_AppPool_State}},
             @{N = 'Website CLR Version';            E = {$_.Website_AppPool_ManagedRuntimeVersion}},
             @{N = 'Web App Name';                   E = {$_.WebApp_Name}},
             @{N = 'Web App Physical Path';          E = {$_.WebApp_PhysicalPath}},
             @{N = 'Web App Application Pool';       E = {$_.WebApp_AppPool_Name}},
             @{N = 'Web App Application Pool State'; E = {$_.WebApp_AppPool_State}},
             @{N = 'Web App CLR Version';            E = {$_.WebApp_AppPool_ManagedRuntimeVersion}}
    )

    return $CsvOutput
}

function Create-ConsoleOutputObject {
    param (
        [parameter(mandatory = $true)][object] $OutputObject
    )

    $ConsoleOutput = ($OutputObject |
      Select @{N = 'Website ID';               E = {$_.Website_Id}},
             @{N = 'Website Name';             E = {$_.Website_Name}},
             @{N = 'Website State';            E = {$_.Website_State}},
             @{N = 'Website Physical Path';    E = {$_.Website_PhysicalPath}},
             @{N = 'Website Binding';          E = {$_.Website_Bindings}},
             @{N = 'Application Pool';         E = {$_.Website_AppPool_Name}},
             @{N = 'CLR Version';              E = {$_.Website_AppPool_ManagedRuntimeVersion}},
             @{N = 'Web App Name';             E = {$_.WebApp_Name}},
             @{N = 'Web App Physical Path';    E = {$_.WebApp_PhysicalPath}},
             @{N = 'Web App Application Pool'; E = {$_.WebApp_AppPool_Name}},
             @{N = 'Web App CLR Version';      E = {$_.WebApp_AppPool_ManagedRuntimeVersion}}
    )

    return $ConsoleOutput
}

function Write-Csv() {
    param(
        [parameter(mandatory = $true)][object] $CsvOutput
    )

    $OutputPath = "c:\iis-info\" + $ComputerInfo.CsDNSHostName.ToLower() + "_output.csv"
    $CsvOutput | Export-Csv $OutputPath -NoTypeInformation
}

function Load-Module {
    param (
        [parameter(Mandatory = $true)][string] $Name
    )

    $RetVal = $true

    if (!(Get-Module -Name $Name)) {
        $RetVal = Get-Module -ListAvailable | where { $_.Name -eq $Name }

        if ($RetVal) {
            try {
                Import-Module $Name -ErrorAction SilentlyContinue
                $RetVal = $true
            } catch {
                Write-Output "Failed to import $Name module"
                $RetVal = $false
            }
        } else {
            $RetVal = $false
        }
    } else {
        Import-Module $Name
    }

    return $RetVal
}

function Is-IISInstalled {
    $IsWebServerInstalled = $false

    # For Windows Server 2008, we have to import ServerManager module to use Get-WindowsFeature command
    if ($ComputerInfo.OsName.Contains("Windows Server 2008")) {
        $ret = Load-Module "ServerManager"

        if ($ret) {
            $IsWebServerInstalled = (Get-WindowsFeature Web-Server).Installed
        } else {
            $res = Get-WmiObject -class Win32_ServerFeature -Namespace "ROOT\CIMV2" -Filter "Name LIKE 'Web Server%'"
            if ($res) {
                $IsWebServerInstalled = $true
            }
        }
    } else {
        $IsWebServerInstalled = (Get-WindowsFeature Web-Server).Installed
    }

    return $IsWebServerInstalled
}

function Get-IISApplicationPool {
    param (
        [parameter(Mandatory = $false)][string] $PoolName
    )

    $ret = Load-Module "IISAdministration"
    if ($ret) {
        if ($PoolName) {
            return Get-IISAppPool -Name $PoolName | Select Name, State, ManagedRuntimeVersion, StartMode, ManagedPipelineMode
        } else {
            return (Get-IISAppPool | Select Name, State, ManagedRuntimeVersion, StartMode, ManagedPipelineMode)
        }
    } else {
        if ($PoolName) {
            $AppPool = Get-WmiObject -class ApplicationPool -Namespace "ROOT\WebAdministration" -Filter "Name='$PoolName'"
            return Map-IISAppPool $AppPool
        }
        else {
            $AppPools = Get-WmiObject -class ApplicationPool -Namespace "ROOT\WebAdministration"
            $IISAppPools = @()
            foreach($AppPool in $AppPools) {
                $IISAppPools += Map-IISAppPool $AppPool
            }

            return $IISAppPools
       }
    }
}

function Map-IISAppPool {

    param(
        [parameter(mandatory = $true)][object] $AppPool
    )

    #https://docs.microsoft.com/en-us/dotnet/api/microsoft.web.administration.startmode?view=iis-dotnet
    $StartMode = @{
        0 = "OnDemand"
        1 = "AlwaysRunning"
    }

    #https://docs.microsoft.com/en-us/dotnet/api/microsoft.web.administration.managedpipelinemode?view=iis-dotnet
    $ManagedPipelineMode = @{
        0 = "Integrated"
        1 = "Classic"
    }

    #https://docs.microsoft.com/en-us/iis/wmi-provider/applicationpool-getstate-method
    $PoolState = @{
        0 = "Starting"
        1 = "Started"
        2 = "Stopping"
        3 = "Stopped"
        4 = "Unknown"
    }

    return [PSCustomObject]@{
            Name                  = $AppPool.Name
            State                 = $PoolState[[int]$AppPool.GetState().ReturnValue]
            StartMode             = $StartMode[$AppPool.StartMode]
            ManagedPipelineMode   = $ManagedPipelineMode[$AppPool.ManagedPipelineMode]
            ManagedRuntimeVersion = $appPool.ManagedRuntimeVersion
    }
}

#######################################################################################
# ENTRY POINT
#######################################################################################
$OutputPath = "c:\iis-info"

if (!(Test-Path $OutputPath)) {
    New-Item $OutputPath -ItemType Directory
}

Write-Output ""
Write-Output "------------------------------"
Write-Output "| Basic Computer Information |"
Write-Output "------------------------------"

if (!(Get-Command "Get-ComputerInfo" -ErrorAction SilentlyContinue)) {
    $Win32ComSystem = Get-WmiObject -class Win32_ComputerSystem -Namespace "ROOT\CIMV2"
    $Win32OSSystem = Get-WmiObject -class Win32_OperatingSystem -Namespace "ROOT\CIMV2"

    $ComputerInfo = [PSCustomObject]@{
        CsName           = $Win32ComSystem.Name
        CsDNSHostName    = $Win32ComSystem.DNSHostName
        CsDomain         = $Win32ComSystem.Domain
        OsName           = $Win32OSSystem.Caption
        OsVersion        = $Win32OSSystem.Version
        OsBuildNumber    = $Win32OSSystem.BuildNumber
        OsArchitecture   = $Win32OSSystem.OSArchitecture
    }
} else {
    $ComputerInfo = Get-ComputerInfo | Select CsName, CsDNSHostName, CsDomain, OsName, OsVersion, OsBuildNumber, OsArchitecture
}

$ComputerInfo | Select @{N = 'Computer Name';   E = {$_.CsName}},
                       @{N = 'DNS Host Name';   E = {$_.CsDNSHostName}},
                       @{N = 'Domain';          E = {$_.CsDomain}},
                       @{N = 'OS Name';         E = {$_.OsName}},
                       @{N = 'OS Version';      E = {$_.OsVersion}},
                       @{N = 'OS Build Number'; E = {$_.OsBuildNumber}},
                       @{N = 'OS Architecture'; E = {$_.OsArchitecture}}| Format-Table * -AutoSize -Wrap

Write-Output ""
Write-Output "-------------------------------"
Write-Output "| .NET Framework Installation |"
Write-Output "-------------------------------"

Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -Recurse |
  Get-ItemProperty -Name version -EA 0 |
  Where { $_.PSChildName -Match '^(?!S)\p{L}'} |
    Select @{N = '.NET Framework'; E = {$_.PSChildName}}, version |
Format-Table * -AutoSize


Write-Output ""
Write-Output "-------------------"
Write-Output "| IIS Information |"
Write-Output "-------------------"

if (Is-IISInstalled) {
    $IISInfo = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\InetStp' |
                Select @{N = 'Product';  E = {$_.SetupString}},
                       @{N = "Product Version"; E = {(Get-ItemProperty ($_.InstallPath + "\w3wp.exe")).VersionInfo.ProductVersion}},
                       @{N = 'Version String';  E = {$_.VersionString}},
                       @{N = 'Install Path';    E = {$_.InstallPath}}
    )

    $IISInfo | Format-Table

    Write-Output ""
    Write-Output "------------------------"
    Write-Output "| IIS Application Pool |"
    Write-Output "------------------------"
    Get-IISApplicationPool | Select @{N = "Name";             E = {$_.Name}},
                                    @{N = "Runtime Version";  E = {$_.ManagedRuntimeVersion}},
                                    @{N = "Status";           E = {$_.State}},
                                    @{N = "Start Mode";       E = {$_.StartMode}},
                                    @{N = "Pipeline Mode";    E = {$_.ManagedPipelineMode}} | Format-Table * -AutoSize

    Write-Output ""
    Write-Output "-------------------------------------------"
    Write-Output "| Website and Web Application Information |"
    Write-Output "-------------------------------------------"
    $ret = Load-Module "WebAdministration"
    if (!$ret) {
        Write-Output "Failed to import WebAdministration module."
    }

    $Websites = Get-Website
    $CsvOutputList = @()
    $ConsoleOutputList = @()

    foreach ($Website in $Websites) {

        $WebApps = Get-WebApplication -Site $Website.Name
        $WebsiteAppPool = Get-IISApplicationPool $Website.ApplicationPool

        $OutputObj = Map-OutputObject $Website $WebsiteAppPool $null $null
        $CsvOutputList += Create-CsvOutputObject $OutputObj
        $ConsoleOutputList += Create-ConsoleOutputObject $OutputObj

        if ($WebApps -ne $null) {
            foreach ($WebApp in $WebApps) {
                $WebAppAppPool = Get-IISApplicationPool $WebApp.applicationPool
                $OutputObj = Map-OutputObject $Website $WebsiteAppPool $WebApp $WebAppAppPool
                $CsvOutputList += Create-CsvOutputObject $OutputObj
                $ConsoleOutputList += Create-ConsoleOutputObject $OutputObj
            }
        }

    }

    $ConsoleOutputList | Format-Table * -AutoSize -Wrap
    Write-Csv $CsvOutputList
} else {
    Write-Output "IIS is not installed on this server."
    Write-Output ""

    $Output = ($ComputerInfo |
      Select @{N = 'Host Name';                      E = {$_.CsDNSHostName}},
             @{N = 'OS Name';                        E = {$_.OsName}},
             @{N = 'IIS Version';                    E = {"Not Installed"}},
             @{N = 'Website ID';                     E = {}},
             @{N = 'Website Name';                   E = {}},
             @{N = 'Website State';                  E = {}},
             @{N = 'Website Physical Path';          E = {}},
             @{N = 'Website Binding';                E = {}},
             @{N = 'Website Application Pool';       E = {}},
             @{N = 'Website Pool State';             E = {}},
             @{N = 'Website CLR Version';            E = {}},
             @{N = 'Web App Name';                   E = {}},
             @{N = 'Web App Physical Path';          E = {}},
             @{N = 'Web App Application Pool';       E = {}},
             @{N = 'Web App Application Pool State'; E = {}},
             @{N = 'Web App CLR Version';            E = {}})

    Write-Csv $Output
}
