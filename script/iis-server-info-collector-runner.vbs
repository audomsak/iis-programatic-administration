Option Explicit

Dim strComputer
Dim strUser
Dim strPassword
Dim strInSameDomain
Dim strScriptPath
Dim objShell

Set objShell = Wscript.CreateObject("WScript.Shell")
strScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & "iis-server-info-collector.vbs"

If (WScript.Arguments.Count = 0) Then
    'Dim wshNetwork : Set wshNetwork = WScript.CreateObject("WScript.Network")
    'strComputer = wshNetwork.ComputerName
    strComputer = "localhost"

    Wscript.StdOut.Write "Running the script against local computer..." & vbCrLf

    
    objShell.Run "cscript " & strScriptPath & " localhost whatever whatever y", 2, True
    
    Wscript.StdOut.Write "The script was run successfully."
Else
    Wscript.StdOut.Write "Are you running the script on the machine that joined the same domain as target servers? (y/n): "
    
    If (StrComp(Wscript.StdIn.ReadLine, "y", 1) = 0) Then
        strInSameDomain = "y"
        strUser = "whatever"
        strPassword = "whatever"
    Else
        Wscript.StdOut.WriteLine ""
        Wscript.StdOut.WriteLine "Running the script on the machine lives in different domain of target servers" & _
                                 "requires username and password for login to remote server." & vbCrLf & _
                                 "Please note that, the user must be in Administrator group." & vbCrLf
        strInSameDomain = "n"
        Wscript.StdOut.Write "Please enter your user name: "
        strUser = Wscript.StdIn.ReadLine 

        ' Ideally, ScriptPW.Password should be used but it might not be avaialble on all machine
        ' so simply just use StdIn
        'Set objPassword = CreateObject("ScriptPW.Password")
       
        Wscript.StdOut.Write "Please enter your password: "
        'strPassword = objPassword.GetPassword()
        
        strPassword = Wscript.StdIn.ReadLine

        Wscript.StdOut.WriteLine ""
    End If

    Dim filename
    filename = WScript.Arguments.Item(0)
    
    Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
    Dim objFile : Set objFile = objFso.OpenTextFile(filename)

    Do Until objFile.AtEndOfStream
        strComputer = objFile.ReadLine
        Wscript.StdOut.Write "Running the script against " & strComputer & " remote computer..." & vbCrLf

        objShell.Run "cscript " & strScriptPath & " " & _
                     strComputer & " " & _
                     strUser & " " & _
                     strPassword & " " & _
                     strInSameDomain, 2, True
    Loop

    objFile.Close

    Wscript.StdOut.WriteLine ""
    Wscript.StdOut.Write "The script was run successfully."
End If
