# Set the execution policy to allow running the script
Set-ExecutionPolicy Bypass -Scope Process -Force

# Define the path to teamsbootstrapper.exe in the same directory as the script
$teamsBootstrapperPath = ".\teamsbootstrapper.exe"
$logFilePath = ".\install_log.txt"

# Start transcript logging
Start-Transcript -Path $logFilePath -Append

# Log starting information
Write-Output "Starting Microsoft Teams installation."

# Check if the teamsbootstrapper.exe exists
if (Test-Path $teamsBootstrapperPath) {
    # Log execution of installation command
    Write-Output "Executing Teams installation command."

    # Execute the installation command
    Start-Process -FilePath $teamsBootstrapperPath -ArgumentList "-p" -Wait

    # Log successful installation
    Write-Output "Microsoft Teams installation completed."
} else {
    # Log error if teamsbootstrapper.exe not found
    Write-Output "Error: teamsbootstrapper.exe not found at the specified path."
}

# Log creation of registry keys
Write-Output "Creating registry keys."

# Create the registry key and DWORD value for Teams
New-Item -Path 'HKLM:\SOFTWARE\Microsoft\' -Name 'Teams' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Teams' -Name 'disableAutoUpdate' -Value 1 -PropertyType DWord -Force | Out-Null

# Create the registry key and MultiString value for Citrix
New-Item -Path 'HKLM:\SOFTWARE\WOW6432Node\Citrix\' -Name 'WebSocketService' -Force | New-ItemProperty -Name 'ProcessWhitelist' -Value 'msedgewebview2.exe' -PropertyType MultiString -Force | Out-Null

# Log completion of registry key creation
Write-Output "Registry keys created successfully."

# Reset the execution policy to its original state
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Undefined -Force

# Log completion of the script
Write-Output "Script execution completed."

# Log created registry keys
Write-Output "Created Registry Keys:"
Get-Item 'HKLM:\SOFTWARE\Microsoft\Teams' | Format-Table -AutoSize
Get-Item 'HKLM:\SOFTWARE\WOW6432Node\Citrix\WebSocketService' | Format-Table -AutoSize

Start-Sleep -Seconds 10

#Installs Teams 2.0 Plugin in C:\Program Files (x86)\Microsoft\TeamsMeetingAddin
#New Microsoft Teams for Virtualized Desktop Infrastructure (VDI) - Microsoft Teams | Microsoft Learn
$teamsMsiPath = (Get-ChildItem -Path 'C:\Program Files\WindowsApps' -Filter 'MSTeams*').FullName
$teamsMsiFullPath = "$teamsMsiPath\MicrosoftTeamsMeetingAddinInstaller.msi"
$installDir = "C:\Program Files (x86)\Microsoft\TeamsMeetingAddin"
$arguments = "/i `"$teamsMsiFullPath`" Reboot=ReallySuppress ALLUSERS=1 TARGETDIR=`"$installDir`" /qn"
Start-Process "msiexec.exe" -ArgumentList $arguments -Wait -NoNewWindow

Start-Sleep -Seconds 10

#Add Registry Keys for loading the Add-in
New-Item -Path "HKLM:\Software\Microsoft\Office\Outlook\Addins" -Name "TeamsAddin.FastConnect" -Force -ErrorAction Ignore
New-ItemProperty -Path "HKLM:\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" -Type "DWord" -Name "LoadBehavior" -Value 3 -force
New-ItemProperty -Path "HKLM:\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" -Type "String" -Name "Description" -Value "Microsoft Teams Meeting Add-in for Microsoft Office" -force
New-ItemProperty -Path "HKLM:\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" -Type "String" -Name "FriendlyName" -Value "Microsoft Teams Meeting Add-in for Microsoft Office" -force


# Stop transcript logging
Stop-Transcript


