$PsVersion = "PowerShell Version: " + $PSVersionTable.PSVersion
Write-Output ""
Write-Output $PsVersion

$Policy=Get-ExecutionPolicy
Write-Output "Execution Policy for current session: $Policy"
Write-Output ""

$Computer
$OutputPath = "c:\iis-info"

$ExecutingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$ScriptPath = Join-Path $ExecutingScriptDirectory "iis-server-info-collector.ps1"

if (Test-Path $OutputPath) {
    Remove-Item $OutputPath -Force -Recurse
}

New-Item $OutputPath -ItemType Directory | Out-Null

if ($args -eq $null -or $args.count -eq 0) {
    Write-Output "Running the script against localhost computer..."
    Write-Output ""

    & $ScriptPath

} else {
    $IsInSameDomain = Read-Host -Prompt "Are you running the script on the machine that joined the same domain as remote servers? (y/n)"
    $RequireCredential

    if ($IsInSameDomain -eq "y") {
        $RequireCredential = $false
    } else {
        $RequireCredential = $true
        Write-Output ""
        Write-Output "Running the script on the machine lives in different domain of remote servers"
        Write-Output "requires username and password for login to remote servers."
        Write-Output ""
        Write-Output "Please note that, the user must be in Administrator group. Press Enter to continue..."
        Read-Host

        $Credential = Get-Credential -Message "Credential for access to all remote computers. Usually the domain user account in Administrator group."
    }

    foreach ($Computer in Get-Content $args[0]) {
        Write-Output ""
        Write-Output "Running the script against $computer computer..."
        Write-Output ""
       
        if ($RequireCredential) {
            $Session = New-PSSession $Computer -Credential $Credential
        } else {
            $Session = New-PSSession $Computer
        }

        try {
            Invoke-Command -Session $Session -FilePath $ScriptPath
            Copy-Item $OutputPath -Destination C:\ -Recurse -Force -FromSession $Session
            Disconnect-PSSession -Session $Session
            Remove-PSSession -Session $Session 
        } catch {
            Write-Host $_.ScriptStackTrace
        }
    }

    if (Get-ChildItem $OutputPath) {
        Get-ChildItem $OutputPath\*.csv |
        ForEach-Object { Import-Csv $_ } |
        Export-Csv $OutputPath\iis-info-all.csv -NoTypeInformation
    }
}

Write-Output ""
Write-Output "The script was run successfully."
