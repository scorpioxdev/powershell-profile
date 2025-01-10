################
#              #
# Setup Script #
#              #
################

# Ensure the script can run with elevated privileges
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    
    Write-Warning "Please run this script as an Administrator."
    break

}

# Function to test internet connection

function Test-InternetConnection {
    try {
        Test-Connection -ComputerName www.google.com -Count 1 -ErrorAction Stop
        return $true
    }
    catch {
        Write-Warning "Internet connection is required but not available. Please check your connection."
        return $false
    }
}


# Function to install Nerd Fonts

function Install-NerdFonts {

    param (
        [string]$FontName,
        [string]$FontDisplayName,
        [string]$Version
    )

    $nerdFonts = @(
        @{FontName = "FiraCode"; FontDisplayName = "Fira Code NF"; Version = "3.3.0"},
        @{FontName = "JetBrainsMono"; FontDisplayName = "JetBrains Mono NF"; Version = "3.3.0"},
        @{FontName = "Meslo"; FontDisplayName = "Meslo Nerd Font"; Version = "3.3.0"},
        @{FontName = "CascadiaCode"; FontDisplayName = "Cascadia Code NF"; Version = "3.3.0"}
    )

    $nerdFonts | ForEach-Object {
        Write-Host "$($_.FontName) - $($_.FontDisplayName) - $($_.Version)"
    }

    $selectedFont = Read-Host "Select a font to install"

    $font = $nerdFonts | Where-Object { $_.FontName -eq $selectedFont }

    if ($font) {
        $fontName = $font.FontName
        $fontDisplayName = $font.FontDisplayName
        $version = $font.Version
    }
    else {
        Write-Warning "Invalid selection. Please select a valid font."
        break
    }

    try {

        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

        $fontFamily = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name

        if ($fontFamily -notcontains "${FontDisplayName}") {
            
            $fontZipUrl = "https://github.com/ryanoasis/nerd-fonts/releases/download/v${Version}/${FontName}.zip"

            $zipFilePath = "$env:TEMP\$FontName.zip"

            $extractPath = "$env:TEMP\$FontName"

            $webClient = New-Object System.Net.WebClient

            $webClient.DownloadFileAsync((New-Object System.Uri($fontZipUrl)), $zipFilePath)

            while ($webClient.IsBusy) {
                Start-Sleep -Seconds 3
            }

            Expand-Archive -Path $zipFilePath -DestinationPath $extractPath -Force

            $destination = (New-Object -ComObject Shell.Application).Namespace(0x14)

            Get-ChildItem -Path $extractPath -Recurse -Filter "*.ttf" | ForEach-Object {
                if (-not(Test-Path "C:\Windows\Fonts\$($_.Name)")) {
                    $destination.CopyHere($_.FullName, 0x10)
                }
            }

            Remove-Item -Path $extractPath -Recurse -Force
            Remove-Item -Path $zipFilePath -Force
        } else {
            Write-Host "Font ${FontDisplayName} is already installed."
        }
    }
    catch {
        Write-Host "An error occurred while installing the font. Font ${FontDisplayName} could not be installed. Error: $_"
    }
}

# Check for internet connectivity before proceeding

if (-not (Test-InternetConnection)) {
    break
}

# Profile creation or update

if (!(Test-Path -Path $PROFILE -PathType Leaf)) {
    try {

        # Detect Version of PowerShell & Create Profile directories if they do not exist.

        $profilePath = ""
        
        if ($PSVersionTable.PSEdition -eq "Core") {
            $profilePath = "$env:userprofile\Documents\Powershell"
        }
        elseif ($PSVersionTable.PSEdition -eq "Desktop") {
            $profilePath = "$env:userprofile\Documents\WindowsPowerShell"
        }

        if (!(Test-Path -Path $profilePath)) {
            New-Item -Path $profilePath -ItemType "directory"
        }

        Invoke-RestMethod https://github.com/scorpioxdev/powershell-profile/raw/main/Microsoft.PowerShell_profile.ps1 -OutFile $PROFILE
        
        Write-Host "The profile @ [$PROFILE] has been created."
        Write-Host "If you want to make any personal changes or customizations, please do so at [$profilePath\Profile.ps1] as there is an updater in the installed profile which uses the hash to update the profile and will lead to loss of changes"
    }
    catch {
        Write-Error "Failed to create or update the profile. Error: $_"
    }
}
else {
    try {
        
        Get-Item -Path $PROFILE | Move-Item -Destination "oldprofile.ps1" -Force

        Invoke-RestMethod https://github.com/scorpioxdev/powershell-profile/raw/main/Microsoft.PowerShell_profile.ps1 -OutFile $PROFILE
        
        Write-Host "The profile @ [$PROFILE] has been created and old profile removed."
        
        Write-Host "Please back up any persistent components of your old profile to [$HOME\Documents\PowerShell\Profile.ps1] as there is an updater in the installed profile which uses the hash to update the profile and will lead to loss of changes"
    }
    catch {
        Write-Error "Failed to backup and update the profile. Error: $_"
    }
}

# OMP Install

try {
    winget install -e --accept-source-agreements --accept-package-agreements JanDeDobbeleer.OhMyPosh
}
catch {
    Write-Error "Failed to install Oh My Posh. Error: $_"
}

# Final check and message to the user

if ((Test-Path -Path $PROFILE) -and (winget list --name "OhMyPosh" -e) -and ($fontFamilies -contains "CaskaydiaCove NF")) {
    
    Write-Host "Setup completed successfully. Please restart your PowerShell session to apply changes."

} else {
    Write-Warning "Setup completed with errors. Please check the error messages above."
}

# Choco install

try {
    Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
}
catch {
    Write-Error "Failed to install Chocolatey. Error: $_"
}

# Terminal Icons Install

try {
    Install-Module -Name Terminal-Icons -Repository PSGallery -Force
}
catch {
    Write-Error "Failed to install Terminal Icons module. Error: $_"
}

# zoxide Install

try {
    winget install -e --id ajeetdsouza.zoxide
    Write-Host "zoxide installed successfully."
}
catch {
    Write-Error "Failed to install zoxide. Error: $_"
}