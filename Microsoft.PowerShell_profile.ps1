#############################
#     PowerShell Profile    #
#     ------------------    #
#        Version 1.0        #
#############################

$debug = $false

if ($debug) {
    Write-Host "##########################" -ForegroundColor Red
    Write-Host "#                        #" -ForegroundColor Red
    Write-Host "#   Debug mode enabled   #" -ForegroundColor Red
    Write-Host "#                        #" -ForegroundColor Red
    Write-Host "#  ONLY FOR DEVELOPMENT  #" -ForegroundColor Red
    Write-Host "#                        #" -ForegroundColor Red
    Write-Host "##########################" -ForegroundColor Red

    Write-Host "####################################" -ForegroundColor Red
    Write-Host "#                                  #" -ForegroundColor Red
    Write-Host "#     IF YOU ARE NOT DEVELOPING    #" -ForegroundColor Red
    Write-Host "#                                  #" -ForegroundColor Red
    Write-Host "#    JUST RUN \'Update-Profile\'   #" -ForegroundColor Red
    Write-Host "#                                  #" -ForegroundColor Red
    Write-Host "#       to discard all changes     #" -ForegroundColor Red
    Write-Host "#                                  #" -ForegroundColor Red
    Write-Host "# and update to the latest profile #" -ForegroundColor Red
    Write-Host "#                                  #" -ForegroundColor Red
    Write-Host "#              version             #" -ForegroundColor Red
    Write-Host "#                                  #" -ForegroundColor Red
    Write-Host "####################################" -ForegroundColor Red
}

#######################
#                     #
# Telemetry & Initial #
#        SETUP        #
#                     #
#######################

# opt-out of telemetry before doing anything, only if powershell is run as admin

if ([bool]([System.Security.Principal.WindowsIdentity]::GetCurrent()).IsSystem) {
    [System.Environment]::SetEnvironmentVariable('POWERSHELL_TELEMETRY_OPTOUT', 'true', [System.EnvironmentVariableTarget]::Machine)
}

# Initial setup - check network connectivity with 3 second timeout

$global:canConnectToGitHub = Test-Connection github.com -Count 3 -TimeoutSeconds 3


########################################
#                                      #
# Import Modules and External Profiles #
#                                      #
########################################

# Ensure that all required modules are installed before importing

$requiredModules = @(
    'posh-git',
    'Terminal-Icons',
    'PSReadLine',
    'PSScriptAnalyzer',
    'BurntToast'
)

$requiredModules | ForEach-Object {
    if (-not(Get-Module -Name $_ -ListAvailable)) {
        if ($canConnectToGitHub) {
            Install-Module -Name $_ -Force -Scope CurrentUser -AllowClobber -SkipPublisherCheck
        } else {
            Write-Host "Unable to connect to GitHub. Skipping module installation for $_" -ForegroundColor Yellow
        }
    }
}

# Import all required modules

$requiredModules | ForEach-Object {
    Import-Module -Name $_ -ErrorAction SilentlyContinue
}

# Import chocolatey profile if it exists

$ChocolateyProfile = "$env:ChocolateyInstall\helpers\chocolateyProfile.psm1"

if (Test-Path ($ChocolateyProfile)) {
    Import-Module "$ChocolateyProfile" -ErrorAction SilentlyContinue
}

#############################
#                           #
#           Check           #
#            for            #
#   Profile and PowerShell  #
#           Updates         #
#                           #
#############################

# Check for profile updates

function Update-Profile {
    try {
        $url = "https://raw.githubusercontent.com/scorpioxdev/powershell-profile/main/Microsoft.PowerShell_profile.ps1"

        $oldhash = Get-FileHash $PROFILE

        Invoke-RestMethod $url -OutFile "$env:temp/Microsoft.PowerShell_profile.ps1"

        $newhash = Get-FileHash "$env:temp/Microsoft.PowerShell_profile.ps1"

        if ($newhash.Hash -ne $oldhash.Hash) {
            Copy-Item -Path "$env:temp/Microsoft.PowerShell_profile.ps1" -Destination $PROFILE -Force

            Write-Host "Profile has been updated successfully. Please restart your shell to reflect changes." -ForegroundColor Magenta
        } else {
            Write-Host "Profile is already up to date" -ForegroundColor Green
        }
    } catch {
        Write-Error "An error occurred while updating the profile. Please try again later."
    }
    finally {
        Remove-Item "$env:temp/Microsoft.PowerShell_profile.ps1" -ErrorAction SilentlyContinue -Force
    }
}

# Skip in debug mode

if (-not $debug) {
    Update-Profile
} else {
    Write-Warning "Skipping profile update in debug mode"
}

# Check for PowerShell updates

function Update-PowerShell {
    try {
        Write-Host "Checking for PowerShell updates..." -ForegroundColor Cyan

        $updateNeeded = $false
        $currentVersion = $PSVersionTable.PSVersion.ToString()

        $gitHubApiUrl = "https://api.github.com/repos/PowerShell/PowerShell/releases/latest"

        $latestReleaseInfo = Invoke-RestMethod -Uri $gitHubApiUrl
        $latestVersion = $latestReleaseInfo.tag_name.Trim('v')

        if ($currentVersion -lt $latestVersion) {
            $updateNeeded = $true
        }

        if ($updateNeeded) {
            Write-Host "Updating PowerShell to version $latestVersion..." -ForegroundColor Yellow

            Start-Process powershell.exe -ArgumentList "-NoProfile -Command winget upgrade Microsoft.PowerShell --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow

            Write-Host "PowerShell has been updated successfully. Please restart your shell to reflect changes." -ForegroundColor Magenta
        } else {
            Write-Host "PowerShell is already up to date" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "An error occurred while updating PowerShell. Error message: $_"
    }
}

# Skip in debug mode

if (-not $debug) {
    Update-PowerShell
} else {
    Write-Warning "Skipping PowerShell update in debug mode"
}

###############
#             #
# Clear Cache #
#             #
###############

function Clear-Cache {
    
    # add clear cache logic here
    Write-Host "Clearing cache..." -ForegroundColor Cyan

    # Clear Windows Prefetch
    Write-Host "Clearing Windows Prefetch..." -ForegroundColor Yellow
    Remove-Item -Path "$env:SystemRoot\Prefetch\*" -Force -ErrorAction SilentlyContinue

    # Clear Windows Temp
    Write-Host "Clearing Windows Temp..." -ForegroundColor Yellow
    Remove-Item -Path "$env:SystemRoot\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue

    # Clear User Temp
    Write-Host "Clearing User Temp..." -ForegroundColor Yellow
    Remove-Item -Path "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue

    # Clear Internet Explorer Cache
    Write-Host "Clearing Internet Explorer Cache..." -ForegroundColor Yellow
    Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Windows\INetCache\*" -Recurse -Force -ErrorAction SilentlyContinue

    Write-Host "Cache clearing completed." -ForegroundColor Green

}

########################
#                      #
#      Admin Check     # 
# Prompt Customization #
#          and         #
#    Other Utilities   #
#                      #
########################

# Check if the shell is running as admin

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# Prompt customization

function prompt {
    if ($isAdmin) {
        "[" + (Get-Location) + "] # "
    } else {
        "[" + (Get-Location) + "] $ "
    }
}

$adminSuffix = if ($isAdmin) { " (Admin)" } else { "" }

$Host.UI.RawUI.WindowTitle = "PowerShell {0}$adminSuffix" -f $PSVersionTable.PSVersion.ToString()

# Quick access to edit profile

function Edit-Profile {
    code $PROFILE.CurrentUserAllHosts
}

function Update-Profile {
    & $profile
}

function ReloadProf {
    & $profile
}

# Other utilities

function touch($file) {
    "Creating file $file" | Out-File $file -Encoding ascii
}

function Get-PubIP {
    $pubIP = Invoke-RestMethod -Uri "https://api.ipify.org?format=json"
    $pubIP.ip
    Write-Host "Public IP: $pubIP.ip" -ForegroundColor Cyan
}

function admin {
    if ($args.Count -gt 0) {
        $argList = $args -join ' '
        Start-Process wt -Verb runAs -ArgumentList "pwsh.exe -NoExit -Command $argList"
    } else {
        Start-Process wt -Verb runAs
    }
}

function uptime {
    try {

        # check powershell version
        if ($PSVersionTable.PSVersion.Major -eq 5) {
            $lastBoot = (Get-WmiObject win32_operatingsystem).LastBootUpTime
            $bootTime = [System.Management.ManagementDateTimeConverter]::ToDateTime($lastBoot)
        } else {
            $lastBootStr = net statistics workstation | Select-String "since" | ForEach-Object { $_.ToString().Replace('Statistics since ', '') }
            # check date format
            if ($lastBootStr -match '^\d{2}/\d{2}/\d{4}') {
                $dateFormat = 'dd/MM/yyyy'
            } elseif ($lastBootStr -match '^\d{2}-\d{2}-\d{4}') {
                $dateFormat = 'dd-MM-yyyy'
            } elseif ($lastBootStr -match '^\d{4}/\d{2}/\d{2}') {
                $dateFormat = 'yyyy/MM/dd'
            } elseif ($lastBootStr -match '^\d{4}-\d{2}-\d{2}') {
                $dateFormat = 'yyyy-MM-dd'
            } elseif ($lastBootStr -match '^\d{2}\.\d{2}\.\d{4}') {
                $dateFormat = 'dd.MM.yyyy'
            }
            
            # check time format
            if ($lastBootStr -match '\bAM\b' -or $lastBootStr -match '\bPM\b') {
                $timeFormat = 'h:mm:ss tt'
            } else {
                $timeFormat = 'HH:mm:ss'
            }

            $bootTime = [System.DateTime]::ParseExact($lastBootStr, "$dateFormat $timeFormat", [System.Globalization.CultureInfo]::InvariantCulture)
        }

        # Format the start time
        ### $formattedBootTime = $bootTime.ToString("dddd, MMMM dd, yyyy HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
        $formattedBootTime = $bootTime.ToString("dddd, MMMM dd, yyyy HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture) + " [$lastBootStr]"
        Write-Host "System started on: $formattedBootTime" -ForegroundColor DarkGray

        # calculate uptime
        $uptime = (Get-Date) - $bootTime

        # Uptime in days, hours, minutes, and seconds
        $days = $uptime.Days
        $hours = $uptime.Hours
        $minutes = $uptime.Minutes
        $seconds = $uptime.Seconds

        # Uptime output
        Write-Host ("Uptime: {0} days, {1} hours, {2} minutes, {3} seconds" -f $days, $hours, $minutes, $seconds) -ForegroundColor Blue
        

    } catch {
        Write-Error "An error occurred while retrieving system uptime."
    }
}

# Enhance listing

function la { 
    Get-ChildItem -Path . -Force | Format-Table -AutoSize 
}

function ll { 
    Get-ChildItem -Path . -Force -Hidden | Format-Table -AutoSize 
}

# Quick Access to System Information
function sysinfo { Get-ComputerInfo }

# Networking Utilities
function flushdns {
	Clear-DnsClientCache
	Write-Host "DNS has been flushed"
}

######################
#                    #
# Enhance PowerShell #
#     Experience     #
#                    #
######################

# PSReadLine configurations

$PSReadLineOptions = @{
    EditMode = 'Windows'
    HistoryNoDuplicates = $true
    HistorySearchCursorMovesToEnd = $true
    PredictionSource = 'History'
    PredictionViewStyle = 'ListView'
    BellStyle = 'None'
    Colors = @{
        Command = '#32CD32' # LimeGreen
        Parameter = '#98FB98' # PaleGreen
        Operator = '#DDA0DD' # Plum
        Variable = '#4682B4' # SteelBlue
        String = '#9370DB' # MediumPurple
        Number = '#5F9EA0' # CadetBlue
        Type = '#ADFF2F' # GreenYellow
        Comment = '#D3D3D3' # LightGray
        Keyword = '#8A2BE2' # BlueViolet
        Error = '#FF4500' # OrangeRed
    }
}

Set-PSReadLineOption @PSReadLineOptions

# Custom key handlers

Set-PSReadLineKeyHandler -Key UpArrow -Function HistorySearchBackward
Set-PSReadLineKeyHandler -Key DownArrow -Function HistorySearchForward
Set-PSReadLineKeyHandler -Key Tab -Function MenuComplete
Set-PSReadLineKeyHandler -Chord 'Ctrl+d' -Function DeleteChar
Set-PSReadLineKeyHandler -Chord 'Ctrl+w' -Function BackwardDeleteWord
Set-PSReadLineKeyHandler -Chord 'Alt+d' -Function DeleteWord
Set-PSReadLineKeyHandler -Chord 'Ctrl+LeftArrow' -Function BackwardWord
Set-PSReadLineKeyHandler -Chord 'Ctrl+RightArrow' -Function ForwardWord
Set-PSReadLineKeyHandler -Chord 'Ctrl+z' -Function Undo
Set-PSReadLineKeyHandler -Chord 'Ctrl+y' -Function Redo

# Custom functions for PSReadLine

Set-PSReadLineOption -AddToHistoryHandler {
    param($line)
    
    $sensitive = @('password', 'secret', 'token' , 'apikey', 'auth', 'access', 'private', 'confidential', 'secure', 'key', 'pass', 'credential', 'authkey', 'connectionstring')

    $hasSensitive = $sensitive | Where-Object { $line -match $_ }

    return ($null -eq $hasSensitive)
}

# Improve prediction settings

Set-PSReadLineOption -PredictionSource HistoryAndPlugin
Set-PSReadLineOption -MaximumHistoryCount 10000

# Custom completion for common commands

$scriptblock = {
    param($wordToComplete, $commandAst, $cursorPosition)
    
    $customCompletions = @{
        'git' = @('add', 'commit', 'push', 'pull', 'clone', 'status', 'branch', 'checkout', 'merge', 'rebase', 'reset', 'log', 'diff', 'tag', 'stash', 'fetch', 'remote', 'config', 'init', 'help')

        'choco' = @('install', 'uninstall', 'upgrade', 'list', 'search', 'outdated', 'pin', 'unpin', 'pack', 'push', 'sources', 'feature', 'config', 'apikey', 'source', 'help')
        
        'winget' = @('install', 'uninstall', 'show', 'search', 'upgrade', 'source', 'settings', 'validate', 'export', 'import', 'hash', 'list', 'features', 'help')
        
        'scoop' = @('install', 'uninstall', 'update', 'status', 'list', 'search', 'bucket', 'cache', 'checkup', 'cleanup', 'config', 'export', 'help')
        
        'npm' = @('install', 'uninstall', 'update', 'list', 'search', 'init', 'publish', 'run', 'test', 'start', 'stop', 'help')
        
        'yarn' = @('install', 'uninstall', 'update', 'list', 'search', 'init', 'publish', 'run', 'test', 'start', 'stop', 'help')
    }

    $command = $commandAst.CommandElements[0].Value

    if ($customCompletions.ContainsKey($command)) {
        $customCompletions[$command] | Where-Object { 
            $_ -like "$wordToComplete*" 
        } | ForEach-Object { 
            [System.Management.Automation.CompletionResult]::new($_, $_, 'ParameterValue', $_) 
        }
    }
}

Register-ArgumentCompleter -Native -CommandName '*' -ScriptBlock $scriptblock

##################
#                #
# Prompt Theming #
#                #
##################

# oh-my-posh theme

function Get-Theme {
    if (Test-Path $PROFILE.CurrentUserAllHosts -PathType Leaf) {
        $existingTheme = Select-String -Raw -Path $PROFILE.
        
        CurrentUserAllHosts -Pattern "oh-my-posh init pwsh --config"

        if ($null -ne $existingTheme) {
            Invoke-Expression $existingTheme
            return
        }

        oh-my-posh init shell pwsh --config https://raw.githubusercontent.com/JanDeDobbeleer/oh-my-posh/main/themes/cobalt2.omp.json | Invoke-Expression
    } else {
        oh-my-posh init pwsh --config https://raw.githubusercontent.com/JanDeDobbeleer/oh-my-posh/main/themes/amro.omp.json | Invoke-Expression
    }
}

###############
#             #
# Final Setup #
#             #
###############

# Set the theme

Get-Theme

# Install and Setup zoxide

if (Get-Command zoxide -ErrorAction SilentlyContinue) {
    Invoke-Expression (& {
        (zoxide init --cmd cd powershell | Out-String)
    })
} else {
    Write-Host "zoxide is not installed. Installing via winget..." -ForegroundColor Yellow
    try {
        winget install -e --id ajeetdsouza.zoxide

        Write-Host "zoxide has been installed successfully. Setting up..." -ForegroundColor Green

        Invoke-Expression (& {
            (zoxide init --cmd cd powershell | Out-String)
        })
    }
    catch {
        Write-Error "Failed to install zoxide. Error message: $_"
    }
}

Set-Alias -Name z -Value __zoxide_z -Option AllScope -Scope Global -Force
Set-Alias -Name zi -Value __zoxide_zi -Option AllScope -Scope Global -Force
