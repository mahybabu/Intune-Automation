
<#PSScriptInfo

.VERSION 1.0

.GUID 2c339518-bee7-4cda-86a2-ab624ade94b6

.AUTHOR Aaron Parker

.COMPANYNAME stealthpuppy

.COPYRIGHT 2019, Aaron Parker. All rights reserved.

.TAGS Intune Microsoft-Intune VcRedist

.LICENSEURI https://github.com/aaronparker/Intune-Automation/blob/master/LICENSE

.PROJECTURI https://github.com/aaronparker/Intune-Automation

.ICONURI https://github.com/aaronparker/Intune-Automation/blob/master/img/IntuneVcRedist.png

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES
First version

.PRIVATEDATA
#>
#Requires -Module Microsoft.Graph.Intune
#Requires -Module VcRedist
<#
.DESCRIPTION 
 Downloads the Visual C++ Redistributables, wraps them with the Intune Win32 App Packaging Tool and deploys them to Microsoft Intune. 
#>
[CmdletBinding(SupportsShouldProcess)]
Param (
    [Parameter()]
    [string] $Path = (Resolve-Path -Path $PWD)
)

#region Functions
Function Get-IntuneWin32PackagingTool {
    [CmdletBinding()]
    Param (
        [Parameter()]
        [string] $DownloadPath = $pwd,

        [Parameter()]
        [string] $ExtractPath = (Join-Path $pwd "IntuneWin32PackagingTool")
    )

    # Get the latest Intune PowerShell SDK
    $latestRelease = (Invoke-Webrequest -uri https://api.github.com/repos/Microsoft/Intune-Win32-App-Packaging-Tool/releases -UseBasicParsing | ConvertFrom-Json)[0]

    # Return the latest version tag
    $latestVersion = $latestRelease.tag_name
    Write-Verbose -Message "Latest release is $latestVersion."

    # Output paths
    $releaseZip = Join-Path $ExtractPath "Intune-Win32-App-Packaging-Tool-$latestVersion.zip"
    If (!(Test-Path -Path $ExtractPath)) { New-Item -Path $ExtractPath -ItemType Directory | Out-Null }

    # Download and extract the latest release
    try {
        If (!(Test-Path -path $releaseZip)) {
            Write-Verbose -Message "Downloading $($latestRelease.zipball_url) to $releaseZip."
            Invoke-WebRequest -Uri $latestRelease.zipball_url -OutFile $releaseZip
            # https://github.com/microsoft/Intune-Win32-App-Packaging-Tool/releases/tag/v1.4
            # https://github.com/microsoft/Intune-Win32-App-Packaging-Tool/archive/v1.4.zip
        }
    }
    catch {
        Throw $_
        Break
    }
    finally {
        Write-Verbose -Message "Extracting $releaseZip."
        Expand-Archive -LiteralPath $releaseZip -DestinationPath $ExtractPath -Force
    }

    # Return the Intune Win32 Packaging Tool to the pipeline
    $exe = Get-ChildItem -Path $ExtractPath -Include "*.exe" -Recurse | Where-Object { $_.Name -eq "IntuneWinAppUtil.exe" } | `
        Sort-Object -Property CreationTime -Descending | Select-Object -First 1
    Write-Output $exe.FullName
}
#endregion

#region Test environment
Try {
    Get-IntuneDeviceManagement -ErrorAction SilentlyContinue | Out-Null
}
Catch {
    Throw "Failed to find MSGraph connection. Please sign in first with Connect-MSGraph."
    Break
}
#endregion

#region Package
# Create $Path if it does not exist
If (Test-Path -Path $Path) {
    Write-Verbose -Message "Path exists: $Path"
}
Else {
    Write-Verbose -Message "Creating path: $Path"
    New-Item -Path $Path -ItemType Directory
}

# Download the VcRedists
$VcRedists = Get-VcList
New-Item -Path "$Path\VcRedists" -ItemType Directory | Out-Null
Save-VcRedist -VcList $VcRedists -Path "$Path\VcRedists"

# Download the Intune Win32 App Packaging Tool
New-Item -Path "$Path\IntuneWin32PackagingTool" -ItemType Directory | Out-Null
$IntuneWin32PackagingTool = Get-IntuneWin32PackagingTool -DownloadPath "$env:Temp" -ExtractPath "$Path\IntuneWin32PackagingTool"

# Package the VcRedists
New-Item -Path "$Path\Packages" -ItemType Directory | Out-Null
ForEach ($vcRedist in $VcRedists) {

    # Paths
    $folder = Join-Path (Join-Path (Join-Path $(Resolve-Path -Path "$Path\VcRedists") $vcRedist.Release) $vcRedist.Architecture) $vcRedist.ShortName
    $filename = Join-Path $folder $(Split-Path -Path $vcRedist.Download -Leaf)
    $output = Join-Path (Join-Path (Join-Path $(Resolve-Path -Path "$Path\Packages") $vcRedist.Release) $vcRedist.Architecture) $vcRedist.ShortName
    New-Item -Path $output -ItemType Directory | Out-Null

    # Package
    If (Test-Path -Path $filename) {
        Write-Verbose "Package: [$($vcRedist.Architecture)]$($vcRedist.Name)"
        . $IntuneWin32PackagingTool -c $folder -s $filename -o $output
        $intunePackage = Get-ChildItem -Path $output -Include "*.intunewin" -Recurse | Sort-Object -Property CreationTime -Descending | Select-Object -First 1

        # Create the request body
        $requestBody = @{
            "@odata.type"                     = "#microsoft.graph.win32LobApp"
            "displayName"                     = "Microsoft $($vcRedist.Name) $($vcRedist.Architecture)"
            "description"                     = "Microsoft $($vcRedist.Name) $($vcRedist.Architecture)"
            "publisher"                       = "Microsoft"
            "isFeatured"                      = "false"
            "privacyInformationUrl"           = "https://privacy.microsoft.com/en-us/privacystatement/"
            "informationUrl"                  = $vcRedist.Url
            "owner"                           = "Microsoft"
            "developer"                       = "Microsoft"
            "notes"                           = "Microsoft $($vcRedist.Name) $($vcRedist.Architecture)"
            "isAssigned"                      = "false"
            "committedContentVersion"         = "1"
            "fileName"                        = $intunePackage.FullName
            "installCommandLine"              = "$(Split-Path -Path $vcRedist.Download -Leaf) $($vcRedist.SilentInstall)"
            "uninstallCommandLine"            = "msiexec /X$($vcRedist.ProductCode) /qn-"
            "applicableArchitectures"         = If ($vcRedist.Architecture -eq "x86") { "x86, x64" } Else { "x64" }
            "minimumSupportedOperatingSystem" = @{
                # "@odata.type" = "#microsoft.graph.windowsMinimumOperatingSystem"
                "v8_0"     = "false"
                "v8_1"     = "false"
                "v10_0"    = "false"
                "v10_1607" = "true"
                "v10_1703" = "false"
                "v10_1709" = "false"
                "v10_1803" = "false"
            }
            "minimumFreeDiskSpaceInMB"        = 200
            "minimumMemoryInMB"               = 4
            "minimumNumberOfProcessors"       = 1
            "minimumCpuSpeedInMHz"            = 1000
            "detectionRules"                  = @( @{
                    "@odata.type"            = "#microsoft.graph.win32LobAppProductCodeDetection"
                    "productCode"            = $vcRedist.ProductCode
                    "productVersionOperator" = "notConfigured"
                    "productVersion"         = $Null
                })
            
            "installExperience"               = @{
                # "@odata.type"  = "#microsoft.graph.win32LobAppInstallExperience"
                "runAsAccount" = "system"
            }
            
            # "setupFilePath"                   = $(Split-Path -Path $vcRedist.Download -Leaf)
            "setupFilePath"                   = "vc_redistx86.exe"
            # "largeIcon"                       = @{ }
            # "returnCodes"                     = @(@{ })
            # "msiInformation"                  = @{ }
            # "requirementRules"                = @(@{ })
            # "uploadState"                     = 11
            # "publishingState"                 = "processing"
            # "dependentAppCount"               = 1
            # "size"                            = 4
        }

        # Create the application
        $app = Invoke-MSGraphRequest -HttpMethod POST -Url "deviceAppManagement/mobileApps" -Content $requestBody -Headers @{"Accept" = "application/json" }
        New-LobApp -filePath $intunePackage.FullName -mobileApp $app
    }
}
#endregion
