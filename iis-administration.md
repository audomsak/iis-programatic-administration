# IIS Programatic Administration

## Summary

Following table is result from this [decision path](decision-path.png).

| Windows version            | Default PowerShell version | Windows Management Framework (WMF) 5.1 | IIS version | Script type                         | API, Interface             | Remoting                 |
|----------------------------|--------------------|----------------------------------------|-------------|-------------------------------------|----------------------------|--------------------------|
| Windows Server 2022        | 5.1                | Installed by default                   | 10          | PS Script / VBScript (no PS Module) | PS Module, WMI, AppCmd.exe | PS Remoting / WMI Remoting |
| Windows Server 2019        | 5.1                | Installed by default                   | 10          | PS Script / VBScript (no PS Module) | PS Module, WMI, AppCmd.exe | PS Remoting / WMI Remoting |
| Windows Server 2016        | 5.1                | Installed by default                   | 10          | PS Script / VBScript (no PS Module) | PS Module, WMI, AppCmd.exe | PS Remoting / WMI Remoting |
| Windows Server 2012 R2     | 4.0                | Needs to be installed                   | 8.5         | PS Script / VBScript (no PS Module) | PS Module, WMI, AppCmd.exe | PS Remoting / WMI Remoting |
| Windows Server 2012        | 3.0                | Needs to be installed                   | 8           | PS Script / VBScript (no PS Module) | PS Module, WMI, AppCmd.exe | PS Remoting / WMI Remoting |
| Windows Server 2008 R2 SP1 | 2.0                | Needs to be installed                   | 7.5         | PS Script / VBScript (no PS Module) | PS Module, WMI, AppCmd.exe | PS Remoting / WMI Remoting |
| Windows Server 2008        | N/A                | N/A                                    | 7.0         | VBScript                            | WMI, AppCmd.exe            | WMI Remoting             |
| Windows Server 2003        | N/A                | N/A                                    | 6.0         | VBScript                            | WMI                        | WMI Remoting             |

***Note***

1. **IISAdministration** PowerShell module is available only in IIS 10.
2. PowerShell (old) module for IIS is **WebAdministration**.
3. **Get-ComputerInfo** PowerShell module requires WMF 5.1 to be installed.
4. **WebAdministration** WMI Namespace requires **IIS Management Scripts and Tools** Windows feature enabled.
5. WMI Namespaces for IIS 7.0 and later is **Root\WebAdministration**.
6. WMI Namespaces for IIS 6.0 is : **Root\MicrosoftIISv2**.
7. **AppCmd**.exe is the single command line tool for managing IIS7 and above.

## WMI Remoting

### System Requirements

#### Windows Firewall

For Windows Server 2003 [see this document](wmi-firewall-config.pdf)

**Applicable for:** Windows Server 2008, 2012, 2016

Following apps need to be allowed to communicate through firewall:

* Windows Management Instrumentation (WMI)

#### Windows Defender Firewall

**Applicable for:** Windows Server 2019, Windows Server 2022

Following apps need to be allowed to communicate through firewall:

* Windows Management Instrumentation (WMI)

#### Windows Services

Following services need to be started on target/remote computer:

* Windows Management Instrumentation

#### User Account

A local or domain user account in the Administrtor group is required.

## PowerShell (PS) Remoting

### System Requirements

## References

* [IIS Official Documentation](https://docs.microsoft.com/en-us/iis/)

* [Internet Information Services (IIS) releases](https://docs.microsoft.com/en-us/lifecycle/products/internet-information-services-iis)

* [IIS Previous Versions Documentation](https://docs.microsoft.com/en-us/previous-versions/iis/)

* [IIS WMI Provider Architecture](https://docs.microsoft.com/en-us/previous-versions/iis/6.0-sdk/ms525673(v=vs.90))

* [IIS Administration Technologies](https://docs.microsoft.com/en-us/previous-versions/iis/6.0-sdk/ms525806(v=vs.90))

* [Using WMI to Configure IIS](https://docs.microsoft.com/en-us/previous-versions/iis/6.0-sdk/ms525309(v=vs.90))

* [An Introduction to Windows PowerShell and IIS](https://docs.microsoft.com/en-us/iis/manage/powershell/an-introduction-to-windows-powershell-and-iis)

* [PowerShell and IIS: 20 practical examples](https://octopus.com/blog/iis-powershell)

* [Enabling IIS Remote Management Using PowerShell](https://mcpmag.com/articles/2014/10/21/enabling-iis-remote-management.aspx)
