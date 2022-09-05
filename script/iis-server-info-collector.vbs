Option Explicit

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

Const HKEY_LOCAL_MACHINE = &H80000002
Const CIMV2_NAMESPACE = "\ROOT\CIMV2"
Const WEB_ADMINISTRATION_NAMESPACE = "\ROOT\WebAdministration"
Const MICROSOFT_IISV2_NAMESPACE = "\ROOT\MicrosoftIISv2"

'SWbemSecurity Impersonation Level
'https://docs.microsoft.com/en-us/windows/win32/api/wbemdisp/ne-wbemdisp-wbemimpersonationlevelenum
Const wbemImpersonationLevelImpersonate = 3

'SWbemSecurity Authentication Level
'https://docs.microsoft.com/en-us/windows/win32/api/wbemdisp/ne-wbemdisp-wbemauthenticationlevelenum
Const wbemAuthenticationLevelPktPrivacy = 6

Dim strComputer
Dim strUser
Dim strPassword
Dim bolInSameDomain

bolInSameDomain = False

If (WScript.Arguments.Count > 0) Then
    strComputer = WScript.Arguments.Item(0)
    strUser = WScript.Arguments.Item(1)
    strPassword = WScript.Arguments.Item(2)

    If (StrComp(WScript.Arguments.Item(3), "y", 1) = 0) Then
        bolInSameDomain = True
    End If
Else
    strComputer = "localhost"
End If

' ------------------ Collect Basic Computer Information ----------------------------

WScript.StdOut.WriteLine "------------------------------"
WScript.StdOut.WriteLine "| Basic Computer Information |"
WScript.StdOut.WriteLine "------------------------------"

Dim objComputerInfo : Set objComputerInfo = CollectBasicComputerInfo()

' ------------------ Collect .NET Framework Information ----------------------------

WScript.StdOut.WriteLine "-------------------------------"
WScript.StdOut.WriteLine "| .NET Framework Installation |"
WScript.StdOut.WriteLine "-------------------------------"

CollectDotNetInfo()

' ------------------ Collect IIS Server Information --------------------------------

WScript.StdOut.WriteLine "-------------------"
WScript.StdOut.WriteLine "| IIS Information |"
WScript.StdOut.WriteLine "-------------------"

If (IsIISInstalled(objComputerInfo.OSName)) Then
    Dim strIISProduct

    strIISProduct = CollectIISServerInfo()

    ' ------------------ Collect IIS Application Pool Information ----------------------

    WScript.StdOut.WriteLine "------------------------"
    WScript.StdOut.WriteLine "| IIS Application Pool |"
    WScript.StdOut.WriteLine "------------------------"

    CollectIISApplicationPoolInfo(strIISProduct)

    ' ------------------ Collect Website and Web Application Information ---------------

    WScript.StdOut.WriteLine "-------------------------------------------"
    WScript.StdOut.WriteLine "| Website and Web Application Information |"
    WScript.StdOut.WriteLine "-------------------------------------------"

    Dim arrWebsiteAndAppList : Set arrWebsiteAndAppList = CollectWebSiteAndWebApplicationInfo(strIISProduct)
    WriteCsvOutputFile True, strIISProduct, arrWebsiteAndAppList
Else
    WScript.StdOut.WriteLine "IIS is not installed on this server."
    WriteCsvOutputFile False, "", Nothing
End If

' ------------------ Write CSV output file -----------------------------------------


'
' ----------------------------------------------------------------------------------
' Functions
' ----------------------------------------------------------------------------------
Function WriteCsvOutputFile(bolIsIISInstalled, strIISVersion, arrWebsiteList)
    Const ForWriting = 2

    Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objCsvFile : Set objCsvFile = objFSO.CreateTextFile("C:\iis-info\" & LCase(objComputerInfo.DNSHostName) & "_output.csv", ForWriting, True, False)
    Dim obj
    Dim strCsvOut

    strCsvOut = "Host Name,OS Name,IIS Version,Website ID,Website Name,Website State,Website Physical Path,Website Binding," & _
                "Website Application Pool,Website Application Pool State,Website CLR Version,Web App Name," & _
                "Web App Physical Path,Web App Application Pool,Web App Application Pool State,Web App CLR Version"

    objCsvFile.WriteLine strCsvOut

    if (bolIsIISInstalled = False) Then
        strCsvOut = objComputerInfo.DNSHostName & "," & objComputerInfo.OSName & ",Not Installed,,,,,,,,,,,,,"
        objCsvFile.WriteLine strCsvOut

    ElseIf (bolIsIISInstalled And (IsNull(arrWebsiteList) Or arrWebsiteList.Count = 0)) Then
        Dim strIISValue
        strIISValue = strIISProduct & " is installed but there isn't any website"
        strCsvOut = objComputerInfo.DNSHostName & "," & objComputerInfo.OSName & "," & strIISValue & ",,,,,,,,,,,,,"
        objCsvFile.WriteLine strCsvOut
    Else
        Dim objWebsite
        For Each objWebsite In arrWebsiteList
            strCsvOut = ""
            strCsvOut = objComputerInfo.DNSHostName & "," & _
                        objComputerInfo.OSName & "," & _
                        strIISVersion & "," & _
                        objWebsite.ID & "," & _
                        objWebsite.Name & "," & _
                        objWebsite.State & "," & _
                        objWebsite.PhysicalPath & "," & _
                        objWebsite.Binding & "," & _
                        objWebsite.ApplicationPool.Name & "," & _
                        objWebsite.ApplicationPool.State & "," & _
                        objWebsite.ApplicationPool.ManagedRuntimeVersion

            If (objWebsite.WebApplication Is Nothing) Then
                strCsvOut = strCsvOut & ",,,,,"
            Else
                strCsvOut = strCsvOut & "," & _
                            objWebsite.WebApplication.Name & "," & _
                            objWebsite.WebApplication.PhysicalPath & "," & _
                            objWebsite.WebApplication.ApplicationPool.Name & "," & _
                            objWebsite.WebApplication.ApplicationPool.State & "," & _
                            objWebsite.WebApplication.ApplicationPool.ManagedRuntimeVersion
            End If

            objCsvFile.WriteLine strCsvOut
        Next
    End If
End Function

Function CollectWebSiteAndWebApplicationInfo(strIISProduct)
    Dim arrWebsiteList : Set arrWebsiteList = CreateObject("System.Collections.ArrayList")

    Dim strQuery
    Dim objItem, colItems
    Dim objWebsite

    If (Instr(strIISProduct, "6.0")) Then
        ' Server State. Ref: https://docs.microsoft.com/en-us/previous-versions/iis/6.0-sdk/ms524905(v=vs.90)
        Dim dicServerState : Set dicServerState = CreateObject("Scripting.Dictionary")
        dicServerState.Add 1, "Starting"
        dicServerState.Add 2, "Started"
        dicServerState.Add 3, "Stopping"
        dicServerState.Add 4, "Stopped"
        dicServerState.Add 5, "Pausing"
        dicServerState.Add 6, "Paused"
        dicServerState.Add 7, "Continuing"

        Dim objSWbemServices: Set objSWbemServices = GetSWbemServices(MICROSOFT_IISV2_NAMESPACE)

        strQuery = "SELECT Name, ServerState FROM IISWebServer"
        Set colItems = objSWbemServices.ExecQuery(strQuery)

        For Each objItem In colItems
            Set objWebsite = New WebSiteInfo
            Dim objWebServerSetting : Set objWebServerSetting = objSWbemServices.Get("IIsWebServerSetting='" & objItem.Name & "'")
            Dim objWebVirtualDirSetting : Set objWebVirtualDirSetting = objSWbemServices.Get("IIsWebVirtualDirSetting='" & objItem.Name & "/ROOT'")
            Dim objSvrBinding
            Dim objScriptMap
            Dim strSvrBinding
            Dim strClrVersion

            For Each objSvrBinding In objWebServerSetting.ServerBindings
                strSvrBinding = objSvrBinding.Hostname & ":" & _
                                objSvrBinding.IP & ":" & _
                                objSvrBinding.Port & ";"
            Next

            For Each objScriptMap In objWebVirtualDirSetting.Properties_("ScriptMaps").Value
                Dim strExtensions
                Dim strScriptProcessor

                strExtensions = objScriptMap.Extensions
                strScriptProcessor = objScriptMap.ScriptProcessor

                If (StrComp(strExtensions, ".aspx", 1) = 0) Then
                    If (Instr(strScriptProcessor, "v1.1")) Then
                        strClrVersion = "v1.1"
                    ElseIf (Instr(strScriptProcessor, "v2.0") Or Instr(strScriptProcessor, "v3.0") Or Instr(strScriptProcessor, "v3.5")) Then
                        strClrVersion = "v2.0"
                    ElseIf (Instr(strScriptProcessor, "v4.0") Or Instr(strScriptProcessor, "v4.5")) Then
                        strClrVersion = "v4.0"
                    End If
                    Exit For
                End If
            Next

            With objWebsite
                .ID = objItem.Name
                .Name = objWebServerSetting.ServerComment
                .State = dicServerState(objItem.ServerState)
                .Binding = Left(strSvrBinding, Len(strSvrBinding) - 1)
                .PhysicalPath = objWebVirtualDirSetting.Properties_("Path").Value
                Set .ApplicationPool = GetApplicationPool(strIISProduct, "W3SVC/AppPools/" & objWebVirtualDirSetting.AppPoolId)(0)
                ' Hack!! In IIS6, .NET Runtime is set at website level but the ApplicationPool.ManagedRuntimeVersion
                ' is used for both display and export to CSV so just set its value here
                .ApplicationPool.ManagedRuntimeVersion = strClrVersion
            End With

            'Note. There is no concept of web application under website in IIS6
            'so no need to query for web application

            arrWebsiteList.Add objWebsite
        Next
    Else
        Dim dicSiteState : Set dicSiteState = CreateObject("Scripting.Dictionary")
        dicSiteState.Add 0, "Starting"
        dicSiteState.Add 1, "Started"
        dicSiteState.Add 2, "Stopping"
        dicSiteState.Add 3, "Stopped"
        dicSiteState.Add 4, "Unknown"

        strQuery = "SELECT * FROM Site"
        Set colItems = ExecuteWMIQuery(WEB_ADMINISTRATION_NAMESPACE, strQuery)
        For Each objItem In colItems
            Set objWebsite = New WebsiteInfo
            Dim strBinding
            Dim objBinding

            strBinding = ""
            For Each objBinding In objItem.Bindings
                strBinding = strBinding & objBinding.Protocol & " " & objBinding.BindingInformation & ";"
            Next

            With objWebsite
                .ID = objItem.ID
                .Name = objItem.Name
                .State = dicSiteState(objItem.GetState)
                .PhysicalPath = GetPhysicalPath(objItem.Name, "/", "/", strIISProduct)
                .Binding = Left(strBinding, Len(strBinding) - 1)
                Set .ApplicationPool = GetApplicationPoolBySiteNameAndPath(objItem.Name, "/", strIISProduct)
            End With

            arrWebsiteList.Add objWebsite

            Dim colApps : Set colApps = objItem.Associators_("SiteContainsApplication")
            Dim objApp

            For Each objApp In colApps
                If (StrComp(objApp.Path, "/", 1) <> 0) Then 'Ignore the Website itself
                    Dim objParentWebsite : Set objParentWebsite = New WebsiteInfo
                    Dim objWebApplication : Set objWebApplication = New WebApplicationInfo


                    With objParentWebsite
                        .ID = objWebsite.ID
                        .Name = objWebsite.Name
                        .State = objWebsite.State
                        .PhysicalPath = objWebsite.PhysicalPath
                        .Binding = objWebsite.Binding
                        Set .ApplicationPool = objWebsite.ApplicationPool
                    End With

                    With objWebApplication
                        .Name = Mid(objApp.Path, 2)
                        .PhysicalPath = GetPhysicalPath(objWebsite.Name, "/", objApp.Path, strIISProduct)
                        Set .ApplicationPool =  GetApplicationPoolBySiteNameAndPath(objItem.Name, objApp.Path, strIISProduct)
                    End With

                    Set objParentWebsite.WebApplication = objWebApplication
                    arrWebsiteList.Add objParentWebsite
                End If
            Next
        Next
    End If

    '  Prepare and Display Output
    Dim arrWebsiteIdList : Set arrWebsiteIdList = CreateObject("System.Collections.ArrayList")
    Dim arrWebsiteNameList : Set arrWebsiteNameList = CreateObject("System.Collections.ArrayList")
    Dim arrWebsiteStateList : Set arrWebsiteStateList = CreateObject("System.Collections.ArrayList")
    Dim arrWebsitePhysicalPathList : Set arrWebsitePhysicalPathList = CreateObject("System.Collections.ArrayList")
    Dim arrWebsiteBindingList : Set arrWebsiteBindingList = CreateObject("System.Collections.ArrayList")
    Dim arrAppPoolList : Set arrAppPoolList = CreateObject("System.Collections.ArrayList")
    Dim arrAppStateList : Set arrAppStateList = CreateObject("System.Collections.ArrayList")
    Dim arrCLRVersionList : Set arrCLRVersionList = CreateObject("System.Collections.ArrayList")
    Dim arrWebAppNameList : Set arrWebAppNameList = CreateObject("System.Collections.ArrayList")
    Dim arrWebAppPhysicalPathList : Set arrWebAppPhysicalPathList = CreateObject("System.Collections.ArrayList")
    Dim arrWebAppAppPoolList : Set arrWebAppAppPoolList = CreateObject("System.Collections.ArrayList")
    Dim arrWebAppAppPoolStateList : Set arrWebAppAppPoolStateList = CreateObject("System.Collections.ArrayList")
    Dim arrWebAppCLRVersionList : Set arrWebAppCLRVersionList = CreateObject("System.Collections.ArrayList")

    Dim dicOutput : Set dicOutput = CreateObject("Scripting.Dictionary")

    For Each objWebsite In arrWebsiteList
        arrWebsiteIdList.Add objWebsite.ID
        arrWebsiteNameList.Add objWebsite.Name
        arrWebsiteStateList.Add objWebsite.State
        arrWebsitePhysicalPathList.Add objWebsite.PhysicalPath
        arrWebsiteBindingList.Add objWebsite.Binding
        arrAppPoolList.Add objWebsite.ApplicationPool.Name
        arrAppStateList.Add objWebsite.ApplicationPool.State
        arrCLRVersionList.Add objWebsite.ApplicationPool.ManagedRuntimeVersion

        If (objWebsite.WebApplication Is Nothing) Then
            arrWebAppNameList.Add ""
            arrWebAppPhysicalPathList.Add ""
            arrWebAppAppPoolList.Add ""
            arrWebAppAppPoolStateList.Add ""
            arrWebAppCLRVersionList.Add ""
        Else
            arrWebAppNameList.Add objWebsite.WebApplication.Name
            arrWebAppPhysicalPathList.Add objWebsite.WebApplication.PhysicalPath
            arrWebAppAppPoolList.Add objWebsite.WebApplication.ApplicationPool.Name
            arrWebAppAppPoolStateList.Add objWebsite.WebApplication.ApplicationPool.State
            arrWebAppCLRVersionList.Add objWebsite.WebApplication.ApplicationPool.ManagedRuntimeVersion
        End If
    Next

    dicOutput.Add "Website ID", arrWebsiteIdList
    dicOutput.Add "Website Name", arrWebsiteNameList
    dicOutput.Add "Website State", arrWebsiteStateList
    dicOutput.Add "Website Physical Path", arrWebsitePhysicalPathList
    dicOutput.Add "Website Binding", arrWebsiteBindingList
    dicOutput.Add "Application Pool", arrAppPoolList
    dicOutput.Add "CLR Version", arrCLRVersionList
    dicOutput.Add "Web App Name", arrWebAppNameList
    dicOutput.Add "Web App Physical Path", arrWebAppPhysicalPathList
    dicOutput.Add "Web App Application Pool", arrWebAppAppPoolList
    dicOutput.Add "Web App CLR Version", arrWebAppCLRVersionList

    WriteMultipleOutputTable(dicOutput)
    Set CollectWebSiteAndWebApplicationInfo = arrWebsiteList
End Function

Function GetApplicationPoolBySiteNameAndPath(strSiteName, strPath, strIISProduct)
    Dim strQuery
    Dim objItem, colItems
    Dim strAppPoolName

    strQuery = "SELECT ApplicationPool FROM Application WHERE SiteName='" & strSiteName & "' AND Path='" & strPath & "'"
    Set colItems = ExecuteWMIQuery(WEB_ADMINISTRATION_NAMESPACE, strQuery)

    For Each objItem In colItems
        strAppPoolName = objItem.ApplicationPool
    Next

    Set GetApplicationPoolBySiteNameAndPath = GetApplicationPool(strIISProduct, strAppPoolName)(0)
End function

Function GetPhysicalPath(strSiteName, strPath, strAppPath, strIISProduct)
    Dim strQuery, strRet
    Dim objItem, colItems

    If (Instr(strIISProduct, "6.0")) Then
        strQuery = "SELECT * FROM IIsWebVirtualDirSetting WHERE SiteName='" & strSiteName & "'"
        Set colItems = ExecuteWMIQuery(MICROSOFT_IISV2_NAMESPACE, strQuery)
    Else
        strQuery = "SELECT PhysicalPath FROM VirtualDirectory WHERE SiteName='" & _
                   strSiteName & "' AND ApplicationPath='" & strAppPath & "' AND Path='" &_
                   strPath & "'"
        Set colItems = ExecuteWMIQuery(WEB_ADMINISTRATION_NAMESPACE, strQuery)
        For Each objItem In colItems
            strRet = objItem.PhysicalPath
        Next
    End If

    GetPhysicalPath = strRet
End Function

Function CollectIISApplicationPoolInfo(strIISProduct)
    Dim arrAppPoolList : Set arrAppPoolList = GetApplicationPool(strIISProduct, Null)

    '  Prepare and Display Output
    Dim arrAppPoolNameList : Set arrAppPoolNameList = CreateObject("System.Collections.ArrayList")
    Dim arrAppPoolStateList : Set arrAppPoolStateList = CreateObject("System.Collections.ArrayList")
    Dim arrAppPoolStartModeList : Set arrAppPoolStartModeList = CreateObject("System.Collections.ArrayList")
    Dim arrAppPoolPipelineModeList : Set arrAppPoolPipelineModeList = CreateObject("System.Collections.ArrayList")
    Dim arrAppPoolRuntimeVersionList : Set arrAppPoolRuntimeVersionList = CreateObject("System.Collections.ArrayList")

    Dim dicOutput : Set dicOutput = CreateObject("Scripting.Dictionary")
    Dim objAppPool

    For Each objAppPool In arrAppPoolList
        arrAppPoolNameList.Add objAppPool.Name
        arrAppPoolStateList.Add objAppPool.State

        ' Application Pool's StartMode is available in IIS 7.5 onward
        If (Instr(strIISProduct, "6.0") Or Instr(strIISProduct, "7.0")) Then
            arrAppPoolStartModeList.Add "N/A"
        Else
            arrAppPoolStartModeList.Add objAppPool.StartMode
        End If

        ' Application Pool's ManagedPipelineMode and ManagedRuntimeVersion is available in IIS 7 onward
        ' (Managed) RuntimeVersion in IIS 6 is set at Web Site level
        If (Instr(strIISProduct, "6.0")) Then
            arrAppPoolPipelineModeList.Add "N/A"
            arrAppPoolRuntimeVersionList.Add "N/A"
        Else
            arrAppPoolPipelineModeList.Add objAppPool.ManagedPipelineMode
            arrAppPoolRuntimeVersionList.Add objAppPool.ManagedRuntimeVersion
        End If
    Next

    dicOutput.Add "Name", arrAppPoolNameList
    dicOutput.Add "Runtime Version", arrAppPoolRuntimeVersionList
    dicOutput.Add "Status", arrAppPoolStateList
    dicOutput.Add "Start Mode", arrAppPoolStartModeList
    dicOutput.Add "Pipeline Mode", arrAppPoolPipelineModeList

    WriteMultipleOutputTable(dicOutput)
End Function

Function GetApplicationPool(strIISProduct, strPoolName)
    Dim dicStartMode : Set dicStartMode = CreateObject("Scripting.Dictionary")
    Dim dicPoolState : Set dicPoolState = CreateObject("Scripting.Dictionary")
    Dim dicIIS6PoolState : Set dicIIS6PoolState = CreateObject("Scripting.Dictionary")
    Dim dicManagedPipelineMode : Set dicManagedPipelineMode = CreateObject("Scripting.Dictionary")

    dicStartMode.Add 0, "OnDemand"
    dicStartMode.Add 1, "AlwaysRunning"

    dicPoolState.Add 0, "Starting"
    dicPoolState.Add 1, "Started"
    dicPoolState.Add 2, "Stopping"
    dicPoolState.Add 3, "Stopped"
    dicPoolState.Add 4, "Unknown"

    ' The document here doesn't tell values mapping. These mapping came up from testing and observed
    ' https://docs.microsoft.com/en-us/previous-versions/iis/6.0-sdk/ms525967(v=vs.90)
    dicIIS6PoolState.Add 2, "Running"
    dicIIS6PoolState.Add 4, "Stopped"

    dicManagedPipelineMode.Add 0, "Integrated"
    dicManagedPipelineMode.Add 1, "Classic"

    Dim strQuery
    Dim colItems
    Dim objItem
    Dim objAppPool

    Dim arrAppPoolList : Set arrAppPoolList = CreateObject("System.Collections.ArrayList")

    If (Instr(strIISProduct, "6.0")) Then
        strQuery = "SELECT * FROM IIsApplicationPoolSetting"

        If (Not IsNull(strPoolName)) Then
            strQuery = strQuery & " WHERE NAME='" & strPoolName & "'"
        End If

        Set colItems = ExecuteWMIQuery(MICROSOFT_IISV2_NAMESPACE, strQuery)

        For Each objItem In colItems
            Set objAppPool = New ApplicationPoolInfo
            Dim arrPoolName

            arrPoolName = Split(objItem.Name, "/")
            With objAppPool
                .Name = arrPoolName(Ubound(arrPoolName))
                .State = dicIIS6PoolState.Item(objItem.AppPoolState)
            End With

            arrAppPoolList.Add objAppPool
        Next
    Else
        strQuery = "SELECT * FROM ApplicationPool"

        If (Not IsNull(strPoolName)) Then
            strQuery = strQuery & " WHERE NAME='" & strPoolName & "'"
        End If

        Set colItems = ExecuteWMIQuery(WEB_ADMINISTRATION_NAMESPACE, strQuery)

        For Each objItem In colItems
            Set objAppPool = New ApplicationPoolInfo

            On Error Resume Next
            With objAppPool
                .Name = objItem.Name
                .State = dicPoolState.Item(objItem.GetState())
                .StartMode = dicStartMode.Item(objItem.StartMode) ' Available in IIS 7.5 onward
                .ManagedPipelineMode = dicManagedPipelineMode.Item(objItem.ManagedPipelineMode)
                .ManagedRuntimeVersion = objItem.ManagedRuntimeVersion
            End With
            Err.Clear

            arrAppPoolList.Add objAppPool
        Next
    End If
    Set GetApplicationPool = arrAppPoolList
End Function

Function CollectBasicComputerInfo()
    Dim objComputerInfo : Set objComputerInfo = New ComputerInfo
    Dim strQuery
    Dim colItems
    Dim objItem

    strQuery = "SELECT * FROM Win32_ComputerSystem"
    Set colItems = ExecuteWMIQuery(CIMV2_NAMESPACE, strQuery)

    For Each objItem In colItems
        objComputerInfo.Name = objItem.Name
        objComputerInfo.DNSHostName = objItem.DNSHostName
        objComputerInfo.Domain = objItem.Domain

        Exit For
    Next

    strQuery = "SELECT * FROM Win32_OperatingSystem"
    Set colItems = ExecuteWMIQuery(CIMV2_NAMESPACE, strQuery)


    For Each objItem In colItems
        objComputerInfo.OSName = objItem.Caption
        objComputerInfo.OSVersion = objItem.Version
        objComputerInfo.OSBuildNumber = objItem.BuildNumber

        ' The WMI object in Windows Server 2003 doesn't have OSArchitecture attribute
        If (InStr(objItem.Caption, "2003")) Then
            objComputerInfo.OSArchitecture = "N/A"
        Else
            objComputerInfo.OSArchitecture = objItem.OSArchitecture
        End If

        Exit For
    Next

    Dim objOutput : Set objOutput = CreateObject("Scripting.Dictionary")
    objOutput.Add "Computer Name", objComputerInfo.Name
    objOutput.Add "DNS Host Name", objComputerInfo.DNSHostName
    objOutput.Add "Domain", objComputerInfo.Domain
    objOutput.Add "OS Name", objComputerInfo.OSName
    objOutput.Add "OS Version", objComputerInfo.OSVersion
    objOutput.Add "OS Build Number", objComputerInfo.OSBuildNumber
    objOutput.Add "OS Architecture", objComputerInfo.OSArchitecture

    WriteSingleOutputTable(objOutput)

    Set CollectBasicComputerInfo = objComputerInfo
End Function

Function CollectIISServerInfo()
    Dim strKeyPath
    Dim strfileVersion
    Dim strProduct
    Dim strInstallPath
    Dim strVersion
    Dim objReg
    Dim objOutput

    strKeyPath = "SOFTWARE\Microsoft\InetStp"
    Set objReg = GetStdRegProvObject()
    objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, "SetupString", strProduct
    objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, "InstallPath", strInstallPath
    objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, "VersionString", strVersion
    strFileVersion = CreateObject("Scripting.FileSystemObject").GetFileVersion(strInstallPath & "\w3wp.exe")

    Set objOutput = CreateObject("Scripting.Dictionary")
    objOutput.Add "Product", strProduct
    objOutput.Add "Product Version", strFileVersion
    objOutput.Add "Version String", strVersion
    objOutput.Add "Install Path", strInstallPath

    WriteSingleOutputTable(objOutput)
    CollectIISServerInfo = strProduct
End Function

Function CollectDotNetInfo()
    Dim strKeyPath

    strKeyPath = "SOFTWARE\Microsoft\NET Framework Setup\NDP"
    Dim objReg : Set objReg = GetStdRegProvObject()
    Dim colDotNetVersionItems : Set colDotNetVersionItems = CreateObject("Scripting.Dictionary")
    Dim arrDotNetVersionList : Set arrDotNetVersionList = CreateObject("System.Collections.ArrayList")

    Dim objRegExp : Set objRegExp = New RegExp
    With objRegExp
        .Pattern = "^[vCFW]"
        .IgnoreCase = False
        .Global = False
    End With

    GetAllDotNetVersion objReg, objRegExp, strKeyPath, arrDotNetVersionList

    ' Prepare and Display Output
    Dim dicOutput : Set dicOutput = CreateObject("Scripting.Dictionary")
    Dim arrNameList : Set arrNameList = CreateObject("System.Collections.ArrayList")
    Dim arrVersionList : Set arrVersionList = CreateObject("System.Collections.ArrayList")
    Dim objDotNet

    For Each objDotNet In arrDotNetVersionList
        arrNameList.Add objDotNet.Name
        arrVersionList.Add objDotNet.Version
    Next

    dicOutput.Add ".NET Framework", arrNameList
    dicOutput.Add "Version", arrVersionList

    WriteMultipleOutputTable(dicOutput)
End Function

Function GetAllDotNetVersion(objReg, objRegExp, strKeyPath, arrDotNetVersionList)
    Dim arrSubKeys
    Dim strSubKey, strFullPath, strValue

    objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys

    If (Not IsNull(arrSubKeys)) Then
    For Each strSubKey In arrSubKeys
        strFullPath = strKeyPath & "\" & strSubkey

        If (objRegExp.Test(strSubkey)) Then
            objReg.GetStringValue HKEY_LOCAL_MACHINE, strFullPath, "Version", strValue

            If Not IsNull(strValue) Then
                Dim objDotNet : Set objDotNet = New DotNetFrameworkInfo
                objDotNet.Name = strSubKey
                objDotNet.Version = strValue

                arrDotNetVersionList.Add objDotNet
            End If
        End If
        GetAllDotNetVersion objReg, objRegExp, strFullPath, arrDotNetVersionList
    Next
    End If
End Function

Function IsIISInstalled(strOSName)
    Dim bolRet
    Dim strQuery
    Dim strKeyPath
    Dim strValue

    bolRet = False

    ' Win32_ServerFeature class is not available in ROOT\CIMV2 namespace in Windows Server 2003
    ' Need to check from Registry instead
    If (InStr(strOSName, "2003") > 0) Then
        strKeyPath = "SOFTWARE\Microsoft\InetStp"
        Dim objReg : Set objReg = GetStdRegProvObject()
        objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, "InstallPath", strValue

        If (Not IsNull(strValue)) Then
            bolRet = True
        End If
    Else
        strQuery = "SELECT * FROM Win32_ServerFeature WHERE Name LIKE 'Web Server%'"
        Dim colItems : Set colItems = ExecuteWMIQuery(CIMV2_NAMESPACE, strQuery)

        ' It's weird that an error occurred when use colItems.Count
        ' Microsoft VBScript runtime error: Object doesn't support this property or method: 'colItems.Count'
        ' While the Doc says 'Count' is there (https://docs.microsoft.com/en-us/windows/win32/wmisdk/swbemobjectset)
        ' So this is bad code but no choice though

        Dim objItem
        For Each objItem in colItems
            bolRet = True
        Next
    End If

    IsIISInstalled = bolRet
End Function

Function ExecuteWMIQuery(strNamespace, strQueryStatement)
    Set ExecuteWMIQuery = GetSWbemServices(strNamespace).ExecQuery(strQueryStatement, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
End Function

Function GetStdRegProvObject()
    SET GetStdRegProvObject = GetSWbemServices("\ROOT\DEFAULT").Get("StdRegProv")
End Function

Function GetSWbemServices(strNamespace)
    Dim objSWbemServices

    If (StrComp(strComputer, "localhost", 1) = 0 Or bolInSameDomain) Then
        Set objSWbemServices = GetObject("winmgmts:{impersonationLevel=impersonate,authenticationLevel=pktPrivacy}!\\" & strComputer & strNamespace)
    Else
        'Remote connection example https://docs.microsoft.com/en-us/windows/win32/wmisdk/swbemlocator-connectserver#examples

        Dim objSWbemLocator : Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
        Set objSWbemServices = objSWbemLocator.ConnectServer(strComputer, strNamespace, strUser, strPassword)

        objSWbemServices.Security_.ImpersonationLevel = wbemImpersonationLevelImpersonate
        objSWbemServices.Security_.AuthenticationLevel = wbemAuthenticationLevelPktPrivacy
    End If

    Set GetSWbemServices = objSWbemServices
End Function

Function WriteSingleOutputTable(objOutput)
    Dim strColumnHeader, strColumnHeaderLine, strColumnValue, strOut
    Dim iColumnHeaderLength, iColumnValueLength, iColumnHeaderWidth, iColumnValueWidth

    strColumnHeader = ""
    strColumnHeaderLine = ""
    strColumnValue = ""

    Dim strKey
    Dim iSpace
    iSpace = 2

    For Each strKey In objOutput.Keys
        iColumnHeaderLength = Len(strKey)
        iColumnValueLength = Len(objOutput.Item(strKey))
        iColumnHeaderWidth = 0
        iColumnValueWidth = 0

        If (iColumnValueLength > iColumnHeaderLength) Then
            iColumnHeaderWidth = (iColumnValueLength - iColumnHeaderLength) + iSpace
            iColumnValueWidth = iSpace
        ElseIf (iColumnValueLength = iColumnHeaderLength) Then
            iColumnHeaderWidth = iSpace
            iColumnValueWidth = iSpace
        Else
            iColumnHeaderWidth = iSpace
            iColumnValueWidth = (iColumnHeaderLength - iColumnValueLength) + iSpace
        End If

        strColumnHeader = strColumnHeader & LeftJustified(strKey, iColumnHeaderWidth)
        strColumnHeaderLine = strColumnHeaderLine & LeftJustified(GenerateDashLine(Len(strKey)), iColumnHeaderWidth)
        strColumnValue = strColumnValue & LeftJustified(objOutput.Item(strKey), iColumnValueWidth)
    Next

    strOut = vbCrLf & strColumnHeader & vbCrLf & strColumnHeaderLine & vbCrLf & strColumnValue & vbCrLf
    WScript.StdOut.WriteLine strOut
End Function

Function WriteMultipleOutputTable(dicOutput)
    Dim strColumnHeader
    Dim strColumnHeaderLine
    Dim strColumnValue
    Dim strColHeaderKey

    Dim iColumnHeaderLength
    Dim iColumnValueLength
    Dim iColumnHeaderWidth
    Dim iColumnValueWidth
    Dim iMaxColumnValueLength
    Dim iMaxValueCount
    Dim iSpace

    iSpace = 2

    strColumnHeader = ""
    strColumnHeaderLine = ""
    strColumnValue = ""

    For Each strColHeaderKey In dicOutput.Keys
        iColumnHeaderLength = Len(strColHeaderKey)
        iMaxColumnValueLength = GetMaxValueLength(dicOutput.Item(strColHeaderKey))
        iMaxValueCount = dicOutput.Item(strColHeaderKey).Count
        iColumnHeaderWidth = 0

        If (iMaxColumnValueLength > iColumnHeaderLength) Then
            iColumnHeaderWidth = (iMaxColumnValueLength - iColumnHeaderLength) + iSpace
        ElseIf (iMaxColumnValueLength = iColumnHeaderLength) Then
            iColumnHeaderWidth = iSpace
        Else
            iColumnHeaderWidth = iSpace
        End If

        strColumnHeader = strColumnHeader & LeftJustified(strColHeaderKey, iColumnHeaderWidth)
        strColumnHeaderLine = strColumnHeaderLine & LeftJustified(GenerateDashLine(Len(strColHeaderKey)), iColumnHeaderWidth)
    Next

    Dim strOut
    Dim iIndex

    For iIndex = 0 To (iMaxValueCount - 1) Step 1
        For Each strColHeaderKey In dicOutput.Keys
            Dim strValue

            iMaxColumnValueLength = GetMaxValueLength(dicOutput.Item(strColHeaderKey))
            if (iMaxColumnValueLength < Len(strColHeaderKey)) Then
                iMaxColumnValueLength = Len(strColHeaderKey)
            End If

            strValue = dicOutput.Item(strColHeaderKey)(iIndex)
            iColumnValueLength = Len(strValue)
            iColumnValueWidth = iMaxColumnValueLength - iColumnValueLength + iSpace
            strColumnValue = strColumnValue & LeftJustified(strValue, iColumnValueWidth)
        Next
        strColumnValue = strColumnValue & vbCrLf
    Next

    strOut = vbCrLf & strColumnHeader & vbCrLf & strColumnHeaderLine & vbCrLf & strColumnValue & vbCrLf
    WScript.StdOut.WriteLine strOut
End Function

Function GetMaxValueLength(colValues)
    Dim strValue
    Dim iMaxValueLength
    Dim iValueLength

    For Each strValue In colValues
        iValueLength = Len(strValue)
        If (iValueLength > iMaxValueLength) Then
            iMaxValueLength = iValueLength
        End If
    Next

    GetMaxValueLength = iMaxValueLength
End Function

Function LeftJustified(ColumnValue, ColumnWidth)
    LeftJustified = ColumnValue & Space(ColumnWidth)
End Function

Function RightJustified(ColumnValue, ColumnWidth)
   RightJustified = Space(ColumnWidth - Len(ColumnValue)) & ColumnValue
End Function

Function GenerateDashLine(length)
    Dim strRet
    Dim i

    strRet = ""

    For i = 1 To length Step 1
        strRet = strRet & "-"
    Next

    GenerateDashLine = strRet
End Function

' ----------------------------------------------------------------------------------
' Classes
' ----------------------------------------------------------------------------------
Class ApplicationPoolInfo
    Private strName
    Private strState
    Private strStartMode
    Private strManagedPipelineMode
    Private strManagedRuntimeVersion

    Public Property Get Name()
        Name = strName
    End Property

    Public Property Let Name(strNameParam)
        strName = strNameParam
    End Property

    Public Property Get State()
        State = strState
    End Property

    Public Property Let State(strStateParam)
        strState = strStateParam
    End Property

    Public Property Get StartMode()
        StartMode = strStartMode
    End Property

    Public Property Let StartMode(strStartModeParam)
        strStartMode = strStartModeParam
    End Property

    Public Property Get ManagedPipelineMode()
        ManagedPipelineMode = strManagedPipelineMode
    End Property

    Public Property Let ManagedPipelineMode(strManagedPipelineModeParam)
        strManagedPipelineMode = strManagedPipelineModeParam
    End Property

    Public Property Get ManagedRuntimeVersion()
        ManagedRuntimeVersion = strManagedRuntimeVersion
    End Property

    Public Property Let ManagedRuntimeVersion(strManagedRuntimeVersionParam)
        strManagedRuntimeVersion = strManagedRuntimeVersionParam
    End Property
End Class

Class DotNetFrameworkInfo
    Private strName
    Private strVersion

    Public Property Get Name()
        Name = strName
    End Property

    Public Property Let Name(strNameParam)
        strName = strNameParam
    End Property

    Public Property Get Version()
        Version = strVersion
    End Property

    Public Property Let Version(strVersionParam)
        strVersion = strVersionParam
    End Property
End Class

Class ComputerInfo
    Private strName
    Private strDNSHostName
    Private strDomain
    Private strOSName
    Private strOSVersion
    Private strOSBuildNumber
    Private strOSArchitecture

    Public Property Get Name()
        Name = strName
    End Property

    Public Property Let Name(strNameParam)
        strName = strNameParam
    End Property

    Public Property Get DNSHostName()
        DNSHostName = strDNSHostName
    End Property

    Public Property Let DNSHostName(strDNSHostNameParam)
        strDNSHostName = strDNSHostNameParam
    End Property

    Public Property Get Domain()
        Domain = strDomain
    End Property

    Public Property Let Domain(strDomainParam)
        strDomain = strDomainParam
    End Property

    Public Property Get OSName()
        OSName = strOSName
    End Property

    Public Property Let OSName(strOSNameParam)
        strOSName = strOSNameParam
    End Property

    Public Property Get OSVersion()
        OSVersion = strOSVersion
    End Property

    Public Property Let OSVersion(strOSVersionParam)
        strOSVersion = strOSVersionParam
    End Property

    Public Property Get OSBuildNumber()
        OSBuildNumber = strOSBuildNumber
    End Property

    Public Property Let OSBuildNumber(strOSBuildNumberParam)
        strOSBuildNumber = strOSBuildNumberParam
    End Property

    Public Property Get OSArchitecture()
        OSArchitecture = strOSArchitecture
    End Property

    Public Property Let OSArchitecture(strOSArchitectureParam)
        strOSArchitecture = strOSArchitectureParam
    End Property
End Class

Class WebSiteInfo
    Private strId
    Private strName
    Private strState
    Private strPhysicalPath
    Private strBinding
    Private objAppPool
    Private objWebApp

    Private Sub Class_Initialize()
        Set ApplicationPool = Nothing
        Set WebApplication = Nothing
    End Sub

    Public Property Get ID()
        ID = strId
    End Property

    Public Property Let ID(strIdParam)
        strId = strIdParam
    End Property

    Public Property Get Name()
        Name = strName
    End Property

    Public Property Let Name(strNameParam)
        strName = strNameParam
    End Property

    Public Property Get State()
        State = strState
    End Property

    Public Property Let State(strStateParam)
        strState = strStateParam
    End Property

    Public Property Get PhysicalPath()
        PhysicalPath = strPhysicalPath
    End Property

    Public Property Let PhysicalPath(strPhysicalPathParam)
        strPhysicalPath = strPhysicalPathParam
    End Property

    Public Property Get Binding()
        Binding = strBinding
    End Property

    Public Property Let Binding(strBindingParam)
        strBinding = strBindingParam
    End Property

    Public Property Get ApplicationPool()
        Set ApplicationPool = objAppPool
    End Property

    Public Property Set ApplicationPool(objAppPoolParam)
        Set objAppPool = objAppPoolParam
    End Property

    Public Property Get WebApplication()
        Set WebApplication = objWebApp
    End Property

    Public Property Set WebApplication(objWebAppParam)
        Set objWebApp = objWebAppParam
    End Property
End Class

Class WebApplicationInfo
    Private strName
    Private strPhysicalPath
    Private objAppPool

    Private Sub Class_Initialize()
        Set ApplicationPool = Nothing
    End Sub

    Public Property Get Name()
        Name = strName
    End Property

    Public Property Let Name(strNameParam)
        strName = strNameParam
    End Property

    Public Property Get PhysicalPath()
        PhysicalPath = strPhysicalPath
    End Property

    Public Property Let PhysicalPath(strPhysicalPathParam)
        strPhysicalPath = strPhysicalPathParam
    End Property

     Public Property Get ApplicationPool()
        Set ApplicationPool = objAppPool
    End Property

    Public Property Set ApplicationPool(objAppPoolParam)
        Set objAppPool = objAppPoolParam
    End Property
End Class
