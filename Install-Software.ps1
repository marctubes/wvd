# Software install Script
#
# Applications to install:
# Client Certificate
# Teams (for WVD)
# FSLogix
# Chrome
# Adobe DC Reader
# Egnyte Desktop Client
# Paloalto Terminal Server Agent
# Argus Software
# Visual Studio Code
# Smart View
# FileZilla Client
# Add Sign Out Button
# Sysprep Fix
# Time Zone Redirection



#region Set logging 
$logFile = "c:\Install\" + (get-date -format 'yyyyMMdd') + '_softwareinstall.log'
function Write-Log {
    Param($message)
    Write-Output "$(get-date -format 'yyyyMMdd HH:mm:ss') $message" | Out-File -Encoding utf8 $logFile -Append
}
#endregion

#region Set Version
$version = "1.0.0"
write-log "Script Version $version"

#region Client Certificate

try {
    Import-Certificate -ErrorAction Stop -FilePath C:\Install\Software\ClientCert\certnew.p7b -CertStoreLocation Cert:\LocalMachine\Root
    if ((Get-ChildItem -Path Cert:\LocalMachine\Root | Where-Object { $_.Thumbprint -eq "B00449232F3E994BDDEC50BCA4B976184FD39D3D" }) -and
        (Get-ChildItem -Path Cert:\LocalMachine\Root | Where-Object { $_.Thumbprint -eq "E65E8B44A8D3FD9CC22E3B99694334387BBB7295" })) {
        Write-log "Client Certificate installed"
    }
    else {
        write-log "Error locating the Client Certificate "
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error installing the Client Certificate: $ErrorMessage"
}

#endregion

#region Teams

# Add VDI Registry Key
# Using a variable for path did not work in testing..
$Name = "IsWVDEnvironment"
$value = "1"
# Add Registry Path
if (!(Test-Path "HKLM:\SOFTWARE\Microsoft\Teams")) {
    try {
        New-Item -ErrorAction Stop -Path "HKLM:\SOFTWARE\Microsoft\Teams" -Force 
    }
    catch {
        $ErrorMessage = $_.Exception.message
        write-log "Error adding teams registry PATH: $ErrorMessage"
    }
}
# Add VDI Registry Value
try {
    New-ItemProperty -ErrorAction Stop -Path "HKLM:\SOFTWARE\Microsoft\Teams" -Name $name -Value $value -PropertyType DWORD -Force
    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Teams") {
        Write-Log "Teams VDI registry key is added"
    }
    else {
        write-log "Error locating the Teams registry key"
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error adding teams registry KEY: $ErrorMessage"
}
# Install the Visual Studio C++ service
try {
    invoke-webrequest -ErrorAction Stop -uri 'https://aka.ms/vs/16/release/vc_redist.x64.exe' -OutFile 'c:\Install\VC_redist.x64.exe'
    Start-Process -filepath c:\Install\VC_redist.x64.exe   -Wait -ErrorAction Stop -ArgumentList '/quiet /log c:\Install\VC_redist.log'
    if (Test-Path "c:\windows\system32\vcruntime140_1.dll") {
        Write-Log "VC_Redist has been installed"
    }
    else {
        write-log "Error locating VC_Redist service "
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error installing VC_redist: $ErrorMessage"
}
# Install the WebSocket service
try {
    invoke-webrequest -ErrorAction Stop -uri 'https://query.prod.cms.rt.microsoft.com/cms/api/am/binary/RE4AQBt' -OutFile 'c:\Install\MsRdcWebRTCSvc_x64.msi'
    Start-Process -filepath msiexec.exe -Wait -ErrorAction Stop -ArgumentList '/i c:\Install\MsRdcWebRTCSvc_x64.msi /quiet'
    if (Test-Path "C:\Program Files\Remote Desktop WebRTC Redirector\MsRdcWebRTCSvc.exe") {
        Write-Log "WebSocket has been installed"
    }
    else {
        write-log "Error locating the WebSocket Service "
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error installing WebSocket: $ErrorMessage"
}
# Install Teams
try {
    invoke-webrequest -ErrorAction Stop -uri 'https://teams.microsoft.com/downloads/desktopurl?env=production&plat=windows&arch=x64&managedInstaller=true&download=true' -OutFile 'c:\Install\Teams_windows_x64.msi'
    Start-Process -filepath msiexec.exe -Wait -ErrorAction Stop -ArgumentList '/i', 'c:\Install\Teams_windows_x64.msi', '/l*v c:\Install\teams.log', 'ALLUSER=1', 'ALLUSERS=1'
    if (Test-Path "C:\Program Files (x86)\Teams Installer\Teams.exe") {
        Write-Log "Teams has been installed"
    }
    else {
        write-log "Error locating the Teams executable"
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error installing Teams: $ErrorMessage"
}
#endregion

#region FSLogix

try {
    invoke-webrequest -ErrorAction Stop -uri 'https://aka.ms/fslogix_download' -OutFile 'c:\Install\fslogix.zip'
    Expand-Archive 'c:\Install\fslogix.zip' 'c:\Install\fslogix'
    Start-Process -ErrorAction Stop -Wait -filepath 'c:\Install\fslogix\x64\Release\FSLogixAppsSetup.exe' -ArgumentList '/quiet'
    if (Test-Path "C:\Program Files\FSLogix\Apps\frxsvc.exe") {
        Write-Log "FSLogix has been installed"
    }
    else {
        write-log "Error locating FSLogix service "
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error installing FSLogix: $ErrorMessage"
}

#endregion

#region Chrome
try {
    Start-Process -filepath msiexec.exe -Wait -ErrorAction Stop -ArgumentList '/i', 'c:\Install\Software\Chrome\GoogleChromeStandaloneEnterprise64.msi', '/quiet'
    if (Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe") {
        Write-Log "Chrome has been installed"
    }
    else {
        write-log "Error locating the Chrome executable "
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error installing FSLogix: $ErrorMessage"
}

#endregion

#region AdobeReader
# Install ref https://silentinstallhq.com/adobe-reader-dc-silent-install-how-to-guide/
try {
    Start-Process -Wait -filepath 'C:\Install\Software\AdobeReader\AcroRdrDC2001320074_en_US.exe' -ArgumentList '/sAll', '/rs', '/msi', 'EULA_ACCEPT=YES'
    if (Test-Path "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe") {
        Write-Log "Adobe Reader has been installed"
    }
    else {
        write-log "Error locating Adobe Reader service "
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error installing Adobe Reader: $ErrorMessage"
}

#endregion

#region Egnyte Desktop Client

# Install Source
# https://helpdesk.egnyte.com/hc/en-us/articles/205237150-Desktop-App-Installers

try {
    Start-Process -filepath msiexec.exe -Wait -ErrorAction Stop -ArgumentList '/i', 'c:\Install\Software\EgnyteDesktop\EgnyteDesktopApp_3.10.1_42.msi', 'ED_SILENT=1', '/passive'
    if (Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe") {
        Write-Log "Egnyte Desktop has been installed"
    }
    else {
        write-log "Error locating the Egnyte Desktop executable "
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error installing Egnyte : $ErrorMessage"
}

#endregion

#region Paloalto Terminal Server Agent

try {
    Start-Process -filepath msiexec.exe -Wait -ErrorAction Stop -ArgumentList '/i', 'c:\Install\Software\Paloalto\TaInstall64.x64-9.1.1-8.msi', '/quiet'
    if (Test-Path "C:\Program Files\Palo Alto Networks\Terminal Server Agent\TaController.exe") {
        Write-Log "Paloalto Terminal Server Agent has been installed"
    }
    else {
        write-log "Error locating the Paloalto Terminal Server Agent executable"
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error installing Paloalto Terminal Server Agent: $ErrorMessage"
}

#endregion


#region Argus Software

try {
    Start-Process -filepath msiexec.exe -Wait -ErrorAction Stop -ArgumentList '/i', 'c:\Install\software\Argus\AEMinimal.msi', '/q', 'INSTALLDIR="C:\Program Files (x86)\ARGUS Software"', 'PSETUPTYPE=4', 'ADDDEFAULT="MainProduct,AEClient,feature.configutility,feature.exceladdin"'
    Copy-Item -ErrorAction Stop 'c:\Install\software\Argus\ConnectionSettings.xml' -destination "C:\ProgramData\ARGUS Software\ARGUS Enterprise 12.1\Settings\" -Force
    New-Item -ErrorAction Stop -ItemType Directory -Name 'Cleaned' -Path "C:\ProgramData\ARGUS Software\ARGUS Enterprise 12.1\Settings\" 
    if (Test-Path "C:\Program Files (x86)\ARGUS Software\ARGUS Enterprise 12.1\Client\ARGUS Software, Inc") {
        Write-log "Argus has been installed"
    }
    else {
        write-log "Error locating the Argus File"
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error installing Argus: $ErrorMessage"
}

#endregion


#region Visual Studio Code Software
try {
    Start-Process -Wait -filepath 'C:\install\Software\VSCode\VSCodeSetup-x64-1.53.2' -ArgumentList '/VERYSILENT', '/NORESTART', '/MERGETASKS=!runcode'
    if (Test-Path "C:\Program Files\Microsoft VS Code\Code.exe") {
        Write-Log "VIsual Studio Code has been Installed"
    }
    else {
        write-log "Error locating Visual Studio Code Executable "
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error installing Visual Studio Code: $ErrorMessage"
}


#endregion

#region SmartView
# Add Smart Forms Registry Key to set IE Settings
    try {
        Start-Process -filepath c:\windows\system32\reg.exe   -Wait -ErrorAction Stop -ArgumentList 'import C:\install\Software\SmartView\SmartView 64 bit\SmartViewRegistrySettings.reg'
    }
    catch {
        $ErrorMessage = $_.Exception.message
        write-log "Error adding teams registry PATH: $ErrorMessage"
    }
#Install SmartView for Office 64 bit only
try {
    Start-Process -Wait -filepath 'C:\install\Software\SmartView\SmartView 64 bit\smartview.exe' -ArgumentList '/s /Office=64'
    if (Test-Path "C:\Oracle\SmartView\bin\SVLaunch.exe") {
        Write-Log "SmartView Plug in Installed"
    }
    else {
        write-log "Error Installing Smart View "
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error installing SmartView: $ErrorMessage"
}
#endregion

#region Filezilla Client
try {
    Start-Process -Wait -filepath 'C:\install\software\Filezilla\FileZilla_3.52.2_win64-setup.exe' -ArgumentList '/S'
    if (Test-Path "C:\Program Files\FileZilla FTP Client\filezilla.exe") {
        Write-Log "FileZilla client Installed"
    }
    else {
        write-log "Error Installing FilaZilla "
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error installing FileZilla: $ErrorMessage"
}
#endregion


#region Sign Out Button

try {
    Copy-Item -ErrorAction Stop 'C:\Install\software\SignOut\Sign Out.lnk' -destination "C:\Users\Public\Desktop" -Force
    if (Test-Path "C:\Users\Public\Desktop\sign out.lnk") {
        Write-log "Sign Out Shortcut Updated"
    }
    else {
        write-log "Error finding sign out shortcut"
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error with sign out link: $ErrorMessage"
}

#endregion

#region Sysprep Fix
# Fix for first login delays due to Windows Module Installer
try {
    ((Get-Content -path C:\DeprovisioningScript.ps1 -Raw) -replace 'Sysprep.exe /oobe /generalize /quiet /quit', 'Sysprep.exe /oobe /generalize /quit /mode:vm' ) | Set-Content -Path C:\DeprovisioningScript.ps1
    write-log "Sysprep Mode:VM fix applied"
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error updating script: $ErrorMessage"
}
#endregion

#region Time Zone Redirection

$Name = "fEnableTimeZoneRedirection"
$value = "1"
# Add Registry value
try {
    New-ItemProperty -ErrorAction Stop -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" -Name $name -Value $value -PropertyType DWORD -Force
    if ((Get-ItemProperty "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services").PSObject.Properties.Name -contains $name) {
        Write-log "Added time zone redirection registry key"
    }
    else {
        write-log "Error locating the Teams registry key"
    }
}
catch {
    $ErrorMessage = $_.Exception.message
    write-log "Error adding teams registry KEY: $ErrorMessage"
}

#endregion

