<#Sources
https://www.advancedinstaller.com/per-machine-msix.html
https://techcommunity.microsoft.com/t5/outlook-blog/the-new-outlook-for-windows-for-organization-admins/bc-p/3937574
https://learn.microsoft.com/en-us/deployoffice/outlook/troubleshoot/troubleshoot-deployment-new-outlook -This is where the link for the new outlook setup.exe is from.
https://learn.microsoft.com/en-us/microsoftteams/new-teams-bulk-install-client -This is where the link for the new teams MSIX is located. There are also exe versions as well as various MSIX versions. Such as x86 and ARM64
#>

<#
Remove previous versions of teams and outlook from all users

$TeamsAppx = Get-AppxPackage -name "*team*" -AllUsers

Foreach ($Appx in $TeamsAppx.packagefullname){
Remove-AppxPackage -Package $Appx -AllUsers
}

$OlkAppx = Get-AppxPackage -name "*outlook*" -AllUsers

Foreach ($Appx in $OlkAppx.packagefullname){
Remove-AppxPackage -Package $Appx -AllUsers
}

#>

Start-Transcript "$($env:SystemRoot)\Logs\Software\NewTeamsOLKServerInstall.log" -Append

#check side loading apps is enabled
$AppModelKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock"

If ($AppModelKey.AllowAllTrustedApps -ne 1) {
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" -Name "AllowAllTrustedApps" -value 1
}

#MSTeams-X64.msix is... you guessed it, referencing the MSIX for New Teams.
#Download found here: https://go.microsoft.com/fwlink/?linkid=2196106
Add-AppProvisionedPackage -Online -PackagePath "$($PSScriptRoot)\MSTeams-X64.msix" -SkipLicense

#Setup.exe is calling the setup executable for new outlook
#Download found here: https://go.microsoft.com/fwlink/?linkid=2207851
Start-Process "$($PSScriptRoot)\Setup.exe" -ArgumentList "--provision true --quiet --start-*" -wait

<#Dropped this for -wait param in start-process
#Ensure New outlook is installed before continuing.
Do {
Write-Host "Waiting on New Outlook to be installed."
Write-Host "If you see this message for too long you either didn't run as admin or you don't have permission to the <Drive>:\Program Files\WindowsApps Folder."
sleep 30
}
until (Test-Path -Path "$($env:ProgramFiles)\WindowsApps\*OutlookForWindows*")
#>

#Copy New outlook source files to new destination with less problems >.<
Copy-Item -Path "$($env:ProgramFiles)\WindowsApps\*OutlookForWindows*" -Destination $env:ProgramData -Recurse

#Creating Shortcut for all users
#https://medium.com/@dbilanoski/how-to-tuesdays-shortcuts-with-powershell-how-to-make-customize-and-point-them-to-places-1ee528af2763
$NewOutlookPath = gci -path "$($env:ProgramData)\*OutlookForWindows*\"

$ShortcutTarget= "$($NewOutlookPath.FullName)\Olk.exe"
$ShortcutFile = "$($env:PUBLIC)\Desktop\New Outlook.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.TargetPath = $ShortcutTarget
$Shortcut.IconLocation = "$($NewOutlookPath.FullName)\Outlook.ico"
$Shortcut.Save()

Stop-Transcript