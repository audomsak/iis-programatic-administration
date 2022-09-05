Option Explicit

Const OUTPUT_FOLDER = "c:\iis-info"
Const ForReading = 1

Dim strComputer
Dim strUser
Dim strPassword
Dim strInSameDomain
Dim strScriptPath
Dim objShell
Dim objShellExec
Dim objFso

Set objShell = CreateObject("WScript.Shell")
strScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & "iis-server-info-collector.vbs"

Set objFso = CreateObject("Scripting.FileSystemObject")
If (Not objFso.FolderExists(OUTPUT_FOLDER)) Then
    objFso.CreateFolder(OUTPUT_FOLDER)
End If

If (WScript.Arguments.Count = 0) Then
    'Dim wshNetwork : Set wshNetwork = WScript.CreateObject("WScript.Network")
    'strComputer = wshNetwork.ComputerName
    strComputer = "localhost"

    ExecuteScript strComputer, "cscript /nologo " & strScriptPath & " localhost x x y"

    WScript.StdOut.Write ""
    WScript.StdOut.Write "The script was run successfully."
Else
    WScript.StdOut.Write "Are you running the script on the machine that joined the same domain as target servers? (y/n): "

    If (StrComp(WScript.StdIn.ReadLine, "y", 1) = 0) Then
        strInSameDomain = "y"
        strUser = "whatever"
        strPassword = "whatever"
    Else
        WScript.StdOut.WriteLine ""
        WScript.StdOut.WriteLine "Running the script on the machine lives in different domain of target servers" & _
                                 "requires username and password for login to remote server." & vbCrLf & _
                                 "Please note that, the user must be in Administrator group." & vbCrLf
        strInSameDomain = "n"
        WScript.StdOut.Write "Please enter your user name: "
        strUser = WScript.StdIn.ReadLine

        ' Ideally, ScriptPW.Password should be used but for some reason it might not be avaialble in all machines
        ' so simply just use StdIn here

        'Set objPassword = CreateObject("ScriptPW.Password")
        'strPassword = objPassword.GetPassword()

        WScript.StdOut.Write "Please enter your password: "
        strPassword = WScript.StdIn.ReadLine

        WScript.StdOut.WriteLine ""
    End If

    Dim filename
    filename = WScript.Arguments.Item(0)

    Dim objFile : Set objFile = objFso.OpenTextFile(filename, ForReading)

    Do Until objFile.AtEndOfStream
        strComputer = objFile.ReadLine

        Dim strCommand
        strCommand = "cscript /nologo " & strScriptPath & " " & _
                     strComputer & " " & _
                     strUser & " " & _
                     strPassword & " " & _
                     strInSameDomain
        ExecuteScript strComputer, strCommand
    Loop

    objFile.Close

    MergeCsvFile()

    WScript.StdOut.WriteLine ""
    WScript.StdOut.Write "The script was run successfully."
End If

Set objShell = Nothing
Set objFso = Nothing

Function ExecuteScript(strComputer, strCommand)
    Set objShellExec = objShell.Exec(strCommand)

    WScript.StdOut.WriteLine "Running the script against " & strComputer & " remote computer..." & vbCrLf

    Do Until objShellExec.StdOut.AtEndOfStream
        WScript.StdOut.WriteLine objShellExec.StdOut.ReadLine
    Loop

    Do Until objShellExec.StdErr.AtEndOfStream
        WScript.StdOut.WriteLine objShellExec.StdErr.ReadLine
    Loop
End Function

Function MergeCsvFile()
    Dim strOutputTxtFilePath
    Dim strOutputCsvFilePath

    strOutputCsvFilePath = OUTPUT_FOLDER & "\iis-info-all.csv"
    strOutputTxtFilePath = OUTPUT_FOLDER & "\iis-info-all.txt"

    ' Delete existing txt and csv files
    If (objFso.FileExists(strOutputCsvFilePath)) Then
        objFso.DeleteFile strOutputCsvFilePath
    End If

    Dim objOutputFolder : Set objOutputFolder = objFso.GetFolder(OUTPUT_FOLDER)
    Dim colFiles : Set colFiles = objOutputFolder.Files
    Dim objFileItem
    Dim strText

    If (Not IsNull(colFiles) And colFiles.Count > 1) Then
        Dim objOutTxtFile : Set objOutTxtFile = objFso.CreateTextFile(strOutputTxtFilePath, True, False)

        objOutTxtFile.WriteLine "Host Name,OS Name,IIS Version,Website ID,Website Name,Website State,Website Physical Path,Website Binding," & _
                "Website Application Pool,Website Application Pool State,Website CLR Version,Web App Name," & _
                "Web App Physical Path,Web App Application Pool,Web App Application Pool State,Web App CLR Version"

        For Each objFileItem In colFiles
            If (objFso.GetExtensionName(objFileItem) = "csv") Then
                Dim objCsvFile : Set objCsvFile = objFso.OpenTextFile(objFileItem.Path, ForReading, False, 0)

                ' Discard headers
                objCsvFile.SkipLine

                Do Until objCsvFile.AtEndOfStream
                    strText = objCsvFile.ReadLine
                   objOutTxtFile.WriteLine strText
                Loop
                objCsvFile.Close
            End If
        Next
        objOutTxtFile.Close

        'Rename txt to csv file
        Set objOutTxtFile = objFso.GetFile(strOutputTxtFilePath)
        objOutTxtFile.Name = Replace(objOutTxtFile.Name, ".txt", ".csv")
    End If
End Function
