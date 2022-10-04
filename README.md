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

5. WMI Namespaces for IIS 7.0 and later is **\ROOT\WebAdministration**.

6. WMI Namespaces for IIS 6.0 is **\ROOT\MicrosoftIISv2**.

7. **AppCmd.exe** is the single command line tool for managing IIS 7 onward.

## WMI Remoting

### System Requirements

#### Windows Firewall

For Windows Server 2003, [see this document](wmi-firewall-config.pdf)

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

### Running the script

[IIS Server Info Collector](script/iis-server-info-collector.vbs) and [Runner](script/iis-server-info-collector-runner.vbs) VBScripts are provided.

* Following these steps to query IIS information on local computer (localhost):

  1. Open Command Prompt (CMD)

  2. Type `cscript iis-server-info-collector-runner.vbs` then press **Enter**

  3. The script will produce some output to console and generate CSV file in **C:\iis-info** folder with filename format as `{DNS Hostname}_output.csv`

* Following these steps to query IIS information on remote computer(s):

  1. Create a text file contains a list of remote computer hostname(s) like this:

     ```txt
      server1
      server2
      server3.example.com
     ```

  2. Open Command Prompt (CMD)

  3. Type `cscript iis-server-info-collector-runner.vbs {path to text file from step 1}` then press **Enter**

  4. You will be asked whether you're running the script on a computer lives in the same domain as remote computer(s) or not.
     * If Yes, enter **y**. The script will use current logged in account for remote authentication.
     * If No, you have to enter username and password for login to remote computer(s). This usename should be able to login to all remote computer(s) (e.g. domain user) listed in the text file in the step 1.

  The script will produce some output to console and generate CSV file for each remote computer in **C:\iis-info** folder with filename format as `{DNS Hostname}_output.csv`. In case of there are more than 1 remote computers, the script will also genereate a `iis-info-all.csv` file that all `*_output.csv` content are merged.

## PowerShell (PS) Remoting

### System Requirements

[Remoting Requirements](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_remote_requirements?view=powershell-5.1)

#### Windows Firewall

**Applicable for:** Windows Server 2008, 2012, 2016

Following apps need to be allowed to communicate through firewall:

* Windows Remote Management

#### Windows Defender Firewall

**Applicable for:** Windows Server 2019, Windows Server 2022

Following apps need to be allowed to communicate through firewall:

* Windows Remote Management

#### Windows Services

Following services need to be started on target/remote computer:

* Windows Remote Management (WS-Management)

#### User Account

A local or domain user account in the Administrtor group is required.

### Running the script

[IIS Server Info Collector](script/iis-server-info-collector.ps1) and [Runner](script/iis-server-info-collector-runner.ps1) PowerShell scripts are provided. The steps to run the script are same as in the [WMI Remoting](#wmi-remoting) section except the script name - use `.ps1` script instead of `.vbs`.



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

* [Enable PSRemoting with Group Policy](https://www.serveracademy.com/enable-psremoting-with-group-policy/)

* [Setting up WMI-Access Through Active Directory & Group Policy](https://support.infrasightlabs.com/help-pages/setting-up-wmi-access-through-ad-gpo/)

* [Set Up a Group Policy to Allow WMI on Your Domain](https://cherwellsupport.com/webhelp/en/5.0/3295.htm)
