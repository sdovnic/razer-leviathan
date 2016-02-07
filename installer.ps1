if ($PSVersionTable.PSVersion.Major -lt 3) {
    [string] $PSScriptRoot = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
}
if ($PSVersionTable.PSVersion.Major -lt 3) {
    [string] $PSCommandPath = $MyInvocation.MyCommand.Definition
}

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process -FilePath "powershell" -WindowStyle Hidden -WorkingDirectory $PSScriptRoot -Verb runAs `
                  -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File $PSCommandPath $args"
    return
}

function Show-Balloon {
    param(
        [parameter(Mandatory=$true)] [string] $TipTitle,
        [parameter(Mandatory=$true)] [string] $TipText,
        [parameter(Mandatory=$false)] [ValidateSet("Info", "Error", "Warning")] [string] $TipIcon,
        [string] $Icon
    )
    process {
        [Void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        $FormsNotifyIcon = New-Object -TypeName System.Windows.Forms.NotifyIcon
        if (-not $Icon) { $Icon = (Join-Path -Path $PSHOME -ChildPath "powershell.exe"); }
        $DrawingIcon = [System.Drawing.Icon]::ExtractAssociatedIcon($Icon)
        $FormsNotifyIcon.Icon = $DrawingIcon
        if (-not $TipIcon) { $TipIcon = "Info"; }
        $FormsNotifyIcon.BalloonTipIcon = $TipIcon;
        $FormsNotifyIcon.BalloonTipTitle = $TipTitle
        $FormsNotifyIcon.BalloonTipText = $TipText
        $FormsNotifyIcon.Visible = $True
        $FormsNotifyIcon.ShowBalloonTip(500)
        Start-Sleep -Milliseconds 500
        $FormsNotifyIcon.Dispose()
    }
}

[string] $TaskInterval = "PT29M"
[string] $TaskName = "razer-leviathan"
[string] $AudioFile = (Join-Path -Path $PSScriptRoot -ChildPath "play.wav")
[string] $TaskCommand = (Join-Path -Path $PSScriptRoot -ChildPath "play.vbs")
[string] $TaskDescription = "Verhindert das automatische abschalten der Razer Leviathan Soundbar. Ben$([char]0x00F6)tigt die Dateien $TaskCommand und $AudioFile."
[string] $TaskFile = (Join-Path -Path $PSScriptRoot -ChildPath "$TaskName.xml")
[string] $TaskTemplate = @'
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
	<RegistrationInfo>
		<Date>2015-08-16T03:36:29</Date>
		<Author>Rally Vincent</Author>
		<Description></Description>
		<URI></URI>
	</RegistrationInfo>
	<Triggers>
        <TimeTrigger>
            <Repetition>
                <Interval></Interval>
                <StopAtDurationEnd>false</StopAtDurationEnd>
            </Repetition>
            <StartBoundary>2016-02-05T03:05:15</StartBoundary>
            <Enabled>true</Enabled>
        </TimeTrigger>
	</Triggers>
	<Settings>
		<MultipleInstancesPolicy>Parallel</MultipleInstancesPolicy>
		<DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
		<StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
		<AllowHardTerminate>true</AllowHardTerminate>
		<StartWhenAvailable>false</StartWhenAvailable>
		<RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
		<IdleSettings>
			<StopOnIdleEnd>true</StopOnIdleEnd>
			<RestartOnIdle>false</RestartOnIdle>
		</IdleSettings>
		<AllowStartOnDemand>false</AllowStartOnDemand>
		<Enabled>true</Enabled>
		<Hidden>false</Hidden>
		<RunOnlyIfIdle>false</RunOnlyIfIdle>
		<WakeToRun>false</WakeToRun>
		<ExecutionTimeLimit>PT72H</ExecutionTimeLimit>
		<Priority>7</Priority>
	</Settings>
	<Actions Context="Author">
		<Exec>
			<Command></Command>
		</Exec>
	</Actions>
</Task>
'@

if ($args.Length -gt 0) {
    if ($args[0].Contains("remove")) {
        if (Get-Command -Name Get-ScheduledTask -ErrorAction SilentlyContinue) {
            if (Get-ScheduledTask -TaskName $TaskName -TaskPath "\" -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName $TaskName -TaskPath "\" -Confirm:$False
            }
        } else {
            Write-Warning -Message "Get-ScheduledTask not supported, using Schtasks."
            $Query = schtasks /Query /TN "\$TaskName" | Out-String
            if ($Query.Contains($TaskName)) {
                [array] $ArgumentList = @("/Delete", "/TN `"\$TaskName`"", "/F")
                Start-Process -FilePath "schtasks" -ArgumentList $ArgumentList -WindowStyle Hidden
            }
        }
        Show-Balloon -TipTitle "Razer Leviathan" -TipText "Razer Leviathan Event entfernt." -TipIcon Info
        $Result = [System.Windows.Forms.MessageBox]::Show(
            "Razer Leviathan Event entfernt.", "Razer Leviathan", 0,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
} else {
    if (Get-Command -Name Get-ScheduledTask -ErrorAction SilentlyContinue) {
        if (Get-ScheduledTask -TaskName $TaskName -TaskPath "\" -ErrorAction SilentlyContinue) {
            Write-Warning -Message "Task already exist!"
        } else {
            $TaskTemplate = $TaskTemplate -replace "<Description>(.*)</Description>", "<Description>$TaskDescription</Description>"
            $TaskTemplate = $TaskTemplate -replace "<URI>(.*)</URI>", "<URI>\$TaskName</URI>"
            $TaskTemplate = $TaskTemplate -replace "<Interval>(.*)</Interval>", "<Interval>$TaskInterval</Interval>"
            $TaskTemplate = $TaskTemplate -replace "<Command>(.*)</Command>", "<Command>$TaskCommand</Command>"
            Set-Content -Path $TaskFile -Value $TaskTemplate
            Start-Process -FilePath "schtasks" -ArgumentList ("/Create", "/TN `"\$TaskName`"", "/XML `"$PSScriptRoot\$TaskName.xml`"") -WindowStyle Hidden -Wait
            Remove-Item -Path $TaskFile
        }
    } else {
        Write-Warning -Message "Get-ScheduledTask not supported, using Schtasks."
        $Query = schtasks /Query /TN "\$TaskName" | Out-String
        if ($Query.Contains($TaskName)) {
            Write-Warning -Message "Task already exist!"
        } else {
            $TaskTemplate = $TaskTemplate -replace "<Description>(.*)</Description>", "<Description>$TaskDescription</Description>"
            $TaskTemplate = $TaskTemplate -replace "<URI>(.*)</URI>", "<URI>\$TaskName</URI>"
            $TaskTemplate = $TaskTemplate -replace "<Interval>(.*)</Interval>", "<Interval>$TaskInterval</Interval>"
            $TaskTemplate = $TaskTemplate -replace "<Command>(.*)</Command>", "<Command>$TaskCommand</Command>"
            Set-Content -Path $TaskFile -Value $TaskTemplate
            Start-Process -FilePath "schtasks" -ArgumentList ("/Create", "/TN `"\$TaskName`"", "/XML `"$PSScriptRoot\$TaskName.xml`"") -WindowStyle Hidden -Wait
            Remove-Item -Path $TaskFile
        }
    }
    Show-Balloon -TipTitle "Razer Leviathan" -TipText "Razer Leviathan Event installiert." -TipIcon Info
    $Result = [System.Windows.Forms.MessageBox]::Show(
        "Razer Leviathan Event installiert.", "Razer Leviathan", 0,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}