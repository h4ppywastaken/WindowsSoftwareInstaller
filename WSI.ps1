<#
.SYNOPSIS
    Windows Software Installer

.DESCRIPTION
    Powershell Script to install a list of software on the local computer.
    Supported Methods:
     - Direct Download Url
     - Sourceforge ("latest" link)
     - Github
     - Chocolatey
     - Manual User Download (just opens the link from the config in firefox browser)
    See WSISW.config for examples.

    This script will automatically install chocolatey.
    If this script is not run as admin it will automatically elevate privileges.

.PARAMETER inputFile
    WSI config file

.PARAMETER folder
    Folder to save downloaded installers (default: <ScriptPath>\Installers)

.PARAMETER pretend
    Pretend Mode

.PARAMETER autoContinue
    Automatically confirm download and installation of all items

.PARAMETER noPromptOnError
    Ignore errors and continue

.EXAMPLE
    .\WSI.ps1
#>

param
(
    [string] $inputFile = ".\WSISW.config",
    [string] $folder = $(Split-Path -Parent -Path $MyInvocation.MyCommand.Path) + "\Installers",
    [switch] $pretend,
    [switch] $autoContinue,
    [switch] $noPromptOnError
)

# Self-elevate the script if required
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator'))
{
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000)
    {
        Write-Warning "Not running as administrator - automatically elevating privileges..."
        $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
        Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
    }
    else
    {
        Write-Error "Cannot elevate privilegs. Please run the script as admin!"
    }
    Exit
}

if([string]::IsNullOrEmpty($inputFile))
{
    throw "InputFile parameter not set";
}

$files = Get-Content $inputFile -ErrorAction Stop

Write-Host ""
Write-Host "###############################################################################################################" -ForegroundColor Red
Write-Host "# PRETEND MODE ENABLED !!! DOWNLOAD/INSTALLATION/CONFIRMATION WILL BE SKIPPED !!!" -ForegroundColor Red
Write-Host "###############################################################################################################" -ForegroundColor Red
Write-Host ""

Write-Host "###############################################################################################################" -ForegroundColor Green
Write-Host "# PREPARATION SECTION" -ForegroundColor Green
Write-Host "###############################################################################################################" -ForegroundColor Green
Write-Host ""
if(-not $(Test-Path($folder)))
{
    Write-Host "###############################################################################################################"
    Write-Host "$folder does not exist - creating directory..." -ForegroundColor Yellow -BackgroundColor Black
    if(-not $pretend)
    {
        New-Item $folder -ItemType Directory > $null
    }
    Write-Host "###############################################################################################################"
    Write-Host ""
}
Write-Host "###############################################################################################################"
Write-Host "Installing chocolatey..." -ForegroundColor Yellow -BackgroundColor Black
if(-not $pretend)
{
    Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    chocolatey feature enable -n useRememberedArgumentsForUpgrades
}

Write-Host "###############################################################################################################"
Write-Host "Upgrading all chocolatey packages..." -ForegroundColor Yellow -BackgroundColor Black
if(-not $pretend)
{
    chocolatey upgrade all --acceptlicense -y -r
}
Write-Host "###############################################################################################################"
Write-Host ""

Write-Host "###############################################################################################################" -ForegroundColor Green
Write-Host "# INSTALLATION SECTION" -ForegroundColor Green
Write-Host "###############################################################################################################" -ForegroundColor Green

foreach($file in $files)
{
    $repeat = $true
    while($repeat)
    {
        try
        {
            if($file[0] -eq "#" -or $file -eq "")
            {
                Continue
            }
            Write-Host ""
            Write-Host "###############################################################################################################"
            if(-not $pretend -and -not $autoContinue)
            {
                $skip = ''
                Write-Host "Next Item: $($file)`r`nPress ENTER to continue, CTRL+C to quit or type 's' to skip this item: " -NoNewline -ForegroundColor Yellow -BackgroundColor Black
                $skip = Read-Host
                if($skip -eq 's')
                {
                    Continue
                }
            }
            $file = $file.Split(';')
            # Get forwarded url from download link
            
            $url = $file[1]
            $origurl = $url

            if($url.Contains('$manual:'))
            {
                Write-Host "Opening $($file[0]) for manual downloads..." -ForegroundColor Yellow -BackgroundColor Black
                Start-Process "firefox" $url.Replace('$Manual:','')
                Continue
            }

            if($url.Contains('$choco:'))
            {
                Write-Host "Installing $($file[0]) from chocolatey..." -ForegroundColor Yellow -BackgroundColor Black
                if(-not $pretend)
                {
                    if(-not $autoContinue)
                    {
                        $skip = ''
                        Write-Host "Press ENTER to continue, CTRL+C to quit or type 's' to skip installing $($file[0]): " -NoNewline -ForegroundColor Yellow -BackgroundColor Black
                        $skip = Read-Host
                        if($skip -eq 's')
                        {
                            Continue
                        }
                    }
                    $chococmdstring = "chocolatey install " +  $url.Replace('$choco:','') + " --acceptlicense -y -r $($file[2])"
                    Invoke-Expression -Command $chococmdstring
                }
                Continue
            }
            
            Write-Host "Getting filename for $($file[0])..." -ForegroundColor Yellow -BackgroundColor Black

            if($url.Contains("sourceforge"))
            {
                $url = Invoke-WebRequest -Method Get -Uri $url -UseBasicParsing
                $url = ($url.Content | Select-String -Pattern '<meta http-equiv="refresh".*(https:\/\/.*);').matches.groups[1].Value
            }   

            if($url.Contains('$version'))
            {
                $url = $url.Substring(0, $url.IndexOf('$version'))
                $url = $url.Substring(0, $url.LastIndexOf('/'))
            }

            $tmphead = Invoke-WebRequest -Method Head -Uri $url -UseBasicParsing
            $url = $tmphead.BaseResponse.RequestMessage.RequestUri.AbsoluteUri
            if(-not $url)
            {
                $url = $tmphead.BaseResponse.ResponseUri.AbsoluteUri
            }

            if($url.Contains("github"))
            {
                $browser_download_urls = ((Invoke-WebRequest $origurl -UseBasicParsing) | ConvertFrom-Json).assets | ForEach-Object {$_.browser_download_url}
                $url = $browser_download_urls | Where-Object {$_ -match 'x64.*.exe$'} | Select-Object -Last 1
                if(-not $url)
                {
                    $url = $browser_download_urls | Where-Object {$_ -match 'x86_64.*.exe$'} | Select-Object -Last 1
                }
                if(-not $url)
                {
                    $url = $browser_download_urls | Where-Object {$_ -match '.exe$'} | Select-Object -Last 1
                }
            }

            if($origurl.Contains('$version'))
            {
                $content = (Invoke-WebRequest $url -UseBasicParsing).Content
                $fullversionlist = ($content | Select-String -Pattern '<a href=".*?((\d+\.)+\d+)([a-z])*.*">.*<\/a>' -AllMatches).Matches
                
                #get main version
                $version = $fullversionlist | ForEach-Object {$_.Groups[1].Value} |
                        Where-Object { $_ -as [version] } |
                        Sort-Object { [version] $_ } | Select-Object -Last 1

                #add subversion a,b,c,d...
                if($fullversionlist[0].Groups[$fullversionlist[0].Groups.Count-1] -match "[a-z]")
                {
                    $version = $fullversionlist | Foreach-Object {$($_.Groups[1].Value + $_.Groups[3].Value)} |
                            Where-Object {$_ -match "$($version)[a-z]{0,1}" } |
                            Sort-Object | Select-Object -Last 1
                }

                $url = $origurl.Replace('$version', $version)
            }
            
            #Extract filename from URL
            $filename = [uri]::UnescapeDataString([System.IO.Path]::GetFileName($url))
            if($filename.IndexOf("?") -ne -1)
            {
                $filename = $filename.Substring(0, $filename.IndexOf("?"))
            }

            if(-not $filename.EndsWith(".exe") -and -not $filename.EndsWith(".zip") -and -not $filename.EndsWith(".msi") -and -not $filename.EndsWith(".ts3_plugin"))
            {
                $filename += ".exe"
            }
            
            Write-Host "Filename: $($filename)" -ForegroundColor Yellow -BackgroundColor Black

            $path = [IO.Path]::Combine($folder,$filename)

            if(Test-Path($path))
            {
                Write-Host "File already downloaded - Skipping download!" -ForegroundColor Yellow -BackgroundColor Black
            }
            else
            {
                Write-Host "Downloading $($file[0]) from $url..." -ForegroundColor Yellow -BackgroundColor Black

                if([System.IO.File]::Exists($path) -eq $False )
                {
                    if(-not $pretend)
                    {
                        try
                        {
                            $wc = New-Object System.Net.WebClient
                            $wc.DownloadFile($url, $path)
                        }
                        catch
                        {
                            Invoke-WebRequest $url -Outfile $path -UseBasicParsing
                        }
                        
                    }
                }
            }

            Write-Host "Installing $($file[0])..." -ForegroundColor Yellow -BackgroundColor Black
            if(-not $pretend)
            {
                if($filename.EndsWith(".zip"))
                {
                    Write-Host "Extracting .zip archive..." -ForegroundColor Yellow -BackgroundColor Black
                    $folderpath = $path.Replace(".zip","")
                    Expand-Archive -LiteralPath $path -DestinationPath $folderpath
                    $path = (Get-ChildItem -Path $folderpath -Recurse | Where-Object {$_.Name -match "^.*.exe$"}).FullName | Select-Object -First 1
                }

                if(-not $autoContinue)
                {
                    $skip = ''
                    Write-Host "Press ENTER to continue, CTRL+C to quit or type 's' to skip installing $($file[0]): " -NoNewline -ForegroundColor Yellow -BackgroundColor Black
                    $skip = Read-Host
                    if($skip -eq 's')
                    {
                        Continue
                    }
                }
                Write-Host "Executing $($path)..." -ForegroundColor Yellow -BackgroundColor Black
                if(-not $file[2])
                {
                    Start-Process $path -Wait
                }
                else
                {
                    Start-Process $path -ArgumentList $file[2] -Wait
                }
                
            }

            $repeat = $false
        }
        catch
        {
            if(!($noPromptOnError))
            {
                Write-Error "There was an unexpected Error. Do you want to retry? (y/N)" -NoNewline
                if((Read-Host).ToLower() -eq "n")
                {
                    $repeat = $false
                }
            }
            else
            {
                $repeat = $false
            }
        }
    }
}

Write-Host ""
Write-Host "###############################################################################################################" -ForegroundColor Green
Write-Host "# FINALIZATION SECTION" -ForegroundColor Green
Write-Host "###############################################################################################################" -ForegroundColor Green
Write-Host ""

Write-Host "###############################################################################################################"
Write-Host "Upgrading all chocolatey packages to finish up..." -ForegroundColor Yellow -BackgroundColor Black
if(-not $pretend)
{
    chocolatey upgrade all --acceptlicense -y -r
}

Write-Host "###############################################################################################################"
Write-Host "Setting up Windows Scheduled Task to automatically upgrade chocolatey packages on system startup..." -ForegroundColor Yellow -BackgroundColor Black
if(-not $pretend)
{
    $create = 'y'
    Write-Host "Do you want to create a scheduled task that automatically updates all choco packages at system startup? (recommended)`r`nYou can define a list of exceptions in the next step... (Y/n): " -NoNewline -ForegroundColor Yellow -BackgroundColor Black
    $create = Read-Host
    if($create.ToLower() -eq 'y')
    {
        Write-Host "List of currently installed packages:"
        chocolatey list --local
        $upgradeexceptions = ''
        Write-Host "Type comma seperated list of packages that you want to exclude from the auto-upgrade task`r`ne.g. geforce-experience, hwinfo: " -NoNewline -ForegroundColor Yellow -BackgroundColor Black
        $upgradeexceptions = Read-Host
        # Get choco.exe path
        $chocoCmd = Get-Command -Name 'choco' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Select-Object -ExpandProperty Source

        # Settings for the scheduled task
        $exceptionargument = ''
        if($upgradeexceptions)
        {
            $exceptionargument = "--except `"'$($upgradeexceptions)'`""
        }
        $taskAction = New-ScheduledTaskAction -Execute $chocoCmd -Argument "--acceptlicense upgrade all -y $($exceptionargument)"
        $taskTrigger = New-ScheduledTaskTrigger -AtStartup
        $taskUserPrincipal = New-ScheduledTaskPrincipal -UserId 'SYSTEM'
        $taskSettings = New-ScheduledTaskSettingsSet -Compatibility Win8

        # Set up the task, and register it
        $task = New-ScheduledTask -Action $taskAction -Principal $taskUserPrincipal -Trigger $taskTrigger -Settings $taskSettings
        Register-ScheduledTask -TaskName 'Run a Choco Upgrade All at Startup' -InputObject $task -Force
    }
    
}
Write-Host "###############################################################################################################"
Write-Host "Finished! Don't forget to clean up your autostart programs!" -ForegroundColor Green
Read-Host "Press ENTER to exit"