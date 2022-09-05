Option Explicit

Dim strComputer
Dim strUser
Dim strPassword
Dim strInSameDomain
Dim strScriptPath
Dim objShell
Dim objShellExec

Set objShell = CreateObject("WScript.Shell")
strScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & "iis-server-info-collector.vbs"

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

    Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
    Dim objFile : Set objFile = objFso.OpenTextFile(filename)

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

    WScript.StdOut.WriteLine ""
    WScript.StdOut.Write "The script was run successfully."
End If

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
