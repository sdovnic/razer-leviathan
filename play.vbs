Set objShell = CreateObject("WScript.Shell")

strCommand = "powershell.exe -NoLogo -Command (New-Object Media.SoundPlayer 'C:\Portable\razer-leviathan\play.wav').PlaySync();"

objShell.Run strCommand, 0
