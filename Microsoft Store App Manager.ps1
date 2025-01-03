#Requires -Version 5.1
#Requires -Modules Microsoft.PowerShell.Utility

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Yönetici hakları kontrolü
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    [System.Windows.Forms.MessageBox]::Show("Bu uygulama yönetici hakları gerektirir. Lütfen PowerShell'i yönetici olarak çalıştırın.", "Hata", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web
[System.Windows.Forms.Application]::EnableVisualStyles()

# DPI farkındalığını ayarla
Add-Type @"
    using System;
    using System.Runtime.InteropServices;

    public class DPI {
        [DllImport("user32.dll")]
        public static extern int SetProcessDPIAware();

        public static void Enable() {
            SetProcessDPIAware();
        }
    }
"@

[DPI]::Enable()

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Microsoft Store App Manager 1.3"
$form.Size = New-Object System.Drawing.Size(800, 700)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.MinimumSize = New-Object System.Drawing.Size(780, 500)

# Package selection group
$packageGroup = New-Object System.Windows.Forms.GroupBox
$packageGroup.Location = New-Object System.Drawing.Point(10, 10)
$packageGroup.Size = New-Object System.Drawing.Size(760, 150)
$packageGroup.Text = "Choose Package"
$packageGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($packageGroup)

# Package combobox
$packageCombo = New-Object System.Windows.Forms.ComboBox
$packageCombo.Location = New-Object System.Drawing.Point(10, 20)
$packageCombo.Size = New-Object System.Drawing.Size(740, 25)
$packageCombo.DropDownStyle = "DropDown"
$packageCombo.AutoCompleteMode = "SuggestAppend"
$packageCombo.AutoCompleteSource = "ListItems"
$packageCombo.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$packageGroup.Controls.Add($packageCombo)

# Product ID section
$productIdLabel = New-Object System.Windows.Forms.Label
$productIdLabel.Location = New-Object System.Drawing.Point(10, 50)
$productIdLabel.Size = New-Object System.Drawing.Size(740, 20)
$productIdLabel.Text = "Paste Microsoft Store URL or Product ID:"
$productIdLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$packageGroup.Controls.Add($productIdLabel)

$productIdTextBox = New-Object System.Windows.Forms.TextBox
$productIdTextBox.Location = New-Object System.Drawing.Point(10, 70)
$productIdTextBox.Size = New-Object System.Drawing.Size(740, 25)
$productIdTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$packageGroup.Controls.Add($productIdTextBox)

# Fetch button
$fetchButton = New-Object System.Windows.Forms.Button
$fetchButton.Location = New-Object System.Drawing.Point(10, 110)
$fetchButton.Size = New-Object System.Drawing.Size(740, 30)
$fetchButton.Text = "Fetch Package Information"
$fetchButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$packageGroup.Controls.Add($fetchButton)

# Architecture group
$archGroup = New-Object System.Windows.Forms.GroupBox
$archGroup.Location = New-Object System.Drawing.Point(10, 170)
$archGroup.Size = New-Object System.Drawing.Size(370, 60)
$archGroup.Text = "Architecture"
$archGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($archGroup)

# Architecture radio buttons
$archX64 = New-Object System.Windows.Forms.RadioButton
$archX64.Location = New-Object System.Drawing.Point(10, 30)
$archX64.Size = New-Object System.Drawing.Size(80, 20)
$archX64.Text = "x64"
$archX64.Checked = $true
$archGroup.Controls.Add($archX64)

$archX86 = New-Object System.Windows.Forms.RadioButton
$archX86.Location = New-Object System.Drawing.Point(100, 30)
$archX86.Size = New-Object System.Drawing.Size(80, 20)
$archX86.Text = "x86"
$archGroup.Controls.Add($archX86)

$archArm64 = New-Object System.Windows.Forms.RadioButton
$archArm64.Location = New-Object System.Drawing.Point(190, 30)
$archArm64.Size = New-Object System.Drawing.Size(80, 20)
$archArm64.Text = "ARM64"
$archGroup.Controls.Add($archArm64)

$archArm = New-Object System.Windows.Forms.RadioButton
$archArm.Location = New-Object System.Drawing.Point(280, 30)
$archArm.Size = New-Object System.Drawing.Size(80, 20)
$archArm.Text = "ARM"
$archGroup.Controls.Add($archArm)

# Ring group
$ringGroup = New-Object System.Windows.Forms.GroupBox
$ringGroup.Location = New-Object System.Drawing.Point(400, 170)
$ringGroup.Size = New-Object System.Drawing.Size(370, 60)
$ringGroup.Text = "Ring"
$ringGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($ringGroup)

# Ring combobox
$ringCombo = New-Object System.Windows.Forms.ComboBox
$ringCombo.Location = New-Object System.Drawing.Point(10, 20)
$ringCombo.Size = New-Object System.Drawing.Size(350, 25)
$ringCombo.DropDownStyle = "DropDownList"
$ringGroup.Controls.Add($ringCombo)

# ListView for packages
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10, 240)
$listView.Size = New-Object System.Drawing.Size(760, 200)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.CheckBoxes = $true
$listView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$listView.Columns.Add("Package", 300)
$listView.Columns.Add("Version", 100)
$listView.Columns.Add("Size", 100)
$listView.Columns.Add("Status", 240)
$form.Controls.Add($listView)

# Select/Deselect buttons
$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Location = New-Object System.Drawing.Point(10, 450)
$selectAllButton.Size = New-Object System.Drawing.Size(375, 30)
$selectAllButton.Text = "Select All"
$selectAllButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($selectAllButton)

$deselectAllButton = New-Object System.Windows.Forms.Button
$deselectAllButton.Location = New-Object System.Drawing.Point(395, 450)
$deselectAllButton.Size = New-Object System.Drawing.Size(375, 30)
$deselectAllButton.Text = "Deselect All"
$deselectAllButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($deselectAllButton)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 500)
$progressBar.Size = New-Object System.Drawing.Size(760, 20)
$progressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($progressBar)

# Status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 520)
$statusLabel.Size = New-Object System.Drawing.Size(760, 20)
$statusLabel.Text = "Ready"
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$statusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($statusLabel)



# Download button
$downloadButton = New-Object System.Windows.Forms.Button
$downloadButton.Location = New-Object System.Drawing.Point(10, 550)
$downloadButton.Size = New-Object System.Drawing.Size(760, 30)
$downloadButton.Text = "Download Selected"
$downloadButton.Enabled = $false
$downloadButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($downloadButton)

# Reset button
$resetButton = New-Object System.Windows.Forms.Button
$resetButton.Location = New-Object System.Drawing.Point(10, 590)
$resetButton.Size = New-Object System.Drawing.Size(760, 30)
$resetButton.Text = "Reset"
$resetButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($resetButton)

# Form resize event handler
$form.Add_Resize({
    $margin = 10
    $clientWidth = $form.ClientSize.Width
    
    # Adjust control widths
    $packageGroup.Width = $clientWidth - (2 * $margin)
    $packageCombo.Width = $packageGroup.Width - (2 * $margin)
    $productIdLabel.Width = $packageCombo.Width
    $productIdTextBox.Width = $packageCombo.Width
    $fetchButton.Width = $packageCombo.Width
    
    # Adjust Ring group position
    $ringGroup.Left = $clientWidth - $ringGroup.Width - $margin
    
    # Adjust ListView width
    $listView.Width = $clientWidth - (2 * $margin)
    
    # Adjust button widths
    $buttonWidth = ($clientWidth - (3 * $margin)) / 2
    $selectAllButton.Width = $buttonWidth
    $deselectAllButton.Width = $buttonWidth
    $deselectAllButton.Left = $selectAllButton.Right + $margin
    
    $progressBar.Width = $clientWidth - (2 * $margin)
    $statusLabel.Width = $progressBar.Width
    $downloadButton.Width = $progressBar.Width
    $resetButton.Width = $progressBar.Width
})


# Data definitions
$PackageList = ConvertFrom-Csv @'
Identity, Family
SpotifyAB.SpotifyMusic, SpotifyAB.SpotifyMusic_zpdnekdrzrea0
5319275A.WhatsAppDesktop, 5319275A.WhatsAppDesktop_cv1g1gvanyjgm
PythonSoftwareFoundation.Python.3.12, PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0
Microsoft.Copilot, Microsoft.Copilot_8wekyb3d8bbwe
BlenderFoundation.Blender, BlenderFoundation.Blender_ppwjx1n5r4v9t
Clipchamp.Clipchamp, Clipchamp.Clipchamp_yxz26nhyzhsrt
Microsoft.AV1VideoExtension, Microsoft.AV1VideoExtension_8wekyb3d8bbwe
Microsoft.BingNews, Microsoft.BingNews_8wekyb3d8bbwe
Microsoft.BingTranslator, Microsoft.BingTranslator_8wekyb3d8bbwe
Microsoft.BingWeather, Microsoft.BingWeather_8wekyb3d8bbwe
Microsoft.Cortana, Microsoft.549981C3F5F10_8wekyb3d8bbwe
Microsoft.DesktopAppInstaller, Microsoft.DesktopAppInstaller_8wekyb3d8bbwe
Microsoft.DirectXRuntime, Microsoft.DirectXRuntime_8wekyb3d8bbwe
Microsoft.GamingApp, Microsoft.GamingApp_8wekyb3d8bbwe
Microsoft.GamingServices, Microsoft.GamingServices_8wekyb3d8bbwe
Microsoft.GetHelp, Microsoft.GetHelp_8wekyb3d8bbwe
Microsoft.Getstarted, Microsoft.Getstarted_8wekyb3d8bbwe
Microsoft.HEIFImageExtension, Microsoft.HEIFImageExtension_8wekyb3d8bbwe
Microsoft.HEVCVideoExtension, Microsoft.HEVCVideoExtension_8wekyb3d8bbwe
Microsoft.Microsoft3DViewer, Microsoft.Microsoft3DViewer_8wekyb3d8bbwe
Microsoft.MicrosoftFamily, MicrosoftCorporationII.MicrosoftFamily_8wekyb3d8bbwe
Microsoft.MicrosoftHoloLens, Microsoft.MicrosoftHoloLens_8wekyb3d8bbwe
Microsoft.MicrosoftJournal, Microsoft.MicrosoftJournal_8wekyb3d8bbwe
Microsoft.MicrosoftOfficeHub, Microsoft.MicrosoftOfficeHub_8wekyb3d8bbwe
Microsoft.MicrosoftOneDrive, Microsoft.MicrosoftSkyDrive_8wekyb3d8bbwe
Microsoft.MicrosoftPCManager, Microsoft.MicrosoftPCManager_8wekyb3d8bbwe
Microsoft.MicrosoftSolitaireCollection, Microsoft.MicrosoftSolitaireCollection_8wekyb3d8bbwe
Microsoft.MicrosoftStickyNotes, Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe
Microsoft.MinecraftEducation, Microsoft.MinecraftEducationEdition_8wekyb3d8bbwe
Microsoft.MixedReality.Portal, Microsoft.MixedReality.Portal_8wekyb3d8bbwe
Microsoft.MPEG2VideoExtension, Microsoft.MPEG2VideoExtension_8wekyb3d8bbwe
Microsoft.MSPaint, Microsoft.MSPaint_8wekyb3d8bbwe
Microsoft.Office.Excel, Microsoft.Office.Excel_8wekyb3d8bbwe
Microsoft.Office.OneNote, Microsoft.Office.OneNote_8wekyb3d8bbwe
Microsoft.Office.PowerPoint, Microsoft.Office.PowerPoint_8wekyb3d8bbwe
Microsoft.Office.Word, Microsoft.Office.Word_8wekyb3d8bbwe
Microsoft.OutlookForWindows, Microsoft.OutlookForWindows_8wekyb3d8bbwe
Microsoft.Paint, Microsoft.Paint_8wekyb3d8bbwe
Microsoft.People, Microsoft.People_8wekyb3d8bbwe
Microsoft.PhotosLegacy, Microsoft.PhotosLegacy_8wekyb3d8bbwe
Microsoft.PowerAutomateDesktop, Microsoft.PowerAutomateDesktop_8wekyb3d8bbwe
Microsoft.PowerShell, Microsoft.PowerShell_8wekyb3d8bbwe
Microsoft.QuickAssist, MicrosoftCorporationII.QuickAssist_8wekyb3d8bbwe
Microsoft.RawImageExtension, Microsoft.RawImageExtension_8wekyb3d8bbwe
Microsoft.RemoteDesktop, Microsoft.RemoteDesktop_8wekyb3d8bbwe
Microsoft.ScreenSketch, Microsoft.ScreenSketch_8wekyb3d8bbwe
Microsoft.Services.Store.Engagement, Microsoft.Services.Store.Engagement_8wekyb3d8bbwe
Microsoft.SkypeApp, Microsoft.SkypeApp_kzf8qxf38zg5c
Microsoft.StorePurchaseApp, Microsoft.StorePurchaseApp_8wekyb3d8bbwe
Microsoft.SysinternalsSuite, Microsoft.SysinternalsSuite_8wekyb3d8bbwe
Microsoft.Todos, Microsoft.Todos_8wekyb3d8bbwe
Microsoft.VP9VideoExtensions, Microsoft.VP9VideoExtensions_8wekyb3d8bbwe
Microsoft.WebMediaExtensions, Microsoft.WebMediaExtensions_8wekyb3d8bbwe
Microsoft.WebpImageExtension, Microsoft.WebpImageExtension_8wekyb3d8bbwe
Microsoft.Whiteboard, Microsoft.Whiteboard_8wekyb3d8bbwe
Microsoft.WinDbg, Microsoft.WinDbg_8wekyb3d8bbwe
Microsoft.WindowsAlarms, Microsoft.WindowsAlarms_8wekyb3d8bbwe
Microsoft.WindowsCalculator, Microsoft.WindowsCalculator_8wekyb3d8bbwe
Microsoft.WindowsCamera, Microsoft.WindowsCamera_8wekyb3d8bbwe
MicrosoftWindows.Client.WebExperience, MicrosoftWindows.Client.WebExperience_cw5n1h2txyewy
Microsoft.WindowsCommunicationsApps, microsoft.windowscommunicationsapps_8wekyb3d8bbwe
Microsoft.WindowsConfigurationDesigner, Microsoft.WindowsConfigurationDesigner_8wekyb3d8bbwe
MicrosoftWindows.CrossDevice, MicrosoftWindows.CrossDevice_cw5n1h2txyewy
Microsoft.WindowsDefenderApplicationGuard, Microsoft.WindowsDefenderApplicationGuard_8wekyb3d8bbwe
Microsoft.Windows.DevHome, Microsoft.Windows.DevHome_8wekyb3d8bbwe
Microsoft.WindowsFeedbackHub, Microsoft.WindowsFeedbackHub_8wekyb3d8bbwe
Microsoft.WindowsHDRCalibration, MicrosoftCorporationII.WindowsHDRCalibration_8wekyb3d8bbwe
Microsoft.WindowsMaps, Microsoft.WindowsMaps_8wekyb3d8bbwe
Microsoft.WindowsNotepad, Microsoft.WindowsNotepad_8wekyb3d8bbwe
Microsoft.Windows.Photos, Microsoft.Windows.Photos_8wekyb3d8bbwe
Microsoft.WindowsScan, Microsoft.WindowsScan_8wekyb3d8bbwe
Microsoft.WindowsSoundRecorder, Microsoft.WindowsSoundRecorder_8wekyb3d8bbwe
Microsoft.WindowsStore, Microsoft.WindowsStore_8wekyb3d8bbwe
Microsoft.WindowsSubsystemForAndroid, MicrosoftCorporationII.WindowsSubsystemForAndroid_8wekyb3d8bbwe
Microsoft.WindowsSubsystemforLinux, MicrosoftCorporationII.WindowsSubsystemforLinux_8wekyb3d8bbwe
Microsoft.WindowsTerminal, Microsoft.WindowsTerminal_8wekyb3d8bbwe
Microsoft.XboxApp, Microsoft.XboxApp_8wekyb3d8bbwe
Microsoft.XboxDevices, Microsoft.XboxDevices_8wekyb3d8bbwe
Microsoft.XboxGameOverlay, Microsoft.XboxGameOverlay_8wekyb3d8bbwe
Microsoft.XboxGamingOverlay, Microsoft.XboxGamingOverlay_8wekyb3d8bbwe
Microsoft.XboxIdentityProvider, Microsoft.XboxIdentityProvider_8wekyb3d8bbwe
Microsoft.XboxSpeechToTextOverlay, Microsoft.XboxSpeechToTextOverlay_8wekyb3d8bbwe
Microsoft.Xbox.TCUI, Microsoft.Xbox.TCUI_8wekyb3d8bbwe
Microsoft.YourPhone, Microsoft.YourPhone_8wekyb3d8bbwe
Microsoft.ZuneMusic, Microsoft.ZuneMusic_8wekyb3d8bbwe
Microsoft.ZuneVideo, Microsoft.ZuneVideo_8wekyb3d8bbwe
Amazon.AmazonAppstore, Amazon.comServicesLLC.AmazonAppstore_bvztej1py64t8
AMD.AMDRadeonSoftware, AdvancedMicroDevicesInc-2.AMDRadeonSoftware_0a9344xs7nr4m
Canonical.Ubuntu, CanonicalGroupLimited.Ubuntu_79rhkp1fndgsc
Canonical.Ubuntu18.04, CanonicalGroupLimited.Ubuntu18.04onWindows_79rhkp1fndgsc
Canonical.Ubuntu20.04LTS, CanonicalGroupLimited.Ubuntu20.04LTS_79rhkp1fndgsc
Canonical.Ubuntu22.04LTS, CanonicalGroupLimited.Ubuntu22.04LTS_79rhkp1fndgsc
Debian.DebianGNULinux, TheDebianProject.DebianGNULinux_76v4gfsz19hv4
Intel.IntelGraphicsExperience, AppUp.IntelGraphicsExperience_8j3eq9eme6ctt
NVIDIA.NVIDIAControlPanel, NVIDIACorp.NVIDIAControlPanel_56jybvy8sckqj
'@

$RingList = ConvertFrom-Csv @'
Name, Value
Fast, WIF
Slow, WIS
Preview, RP
Retail, Retail
'@

# Utility Functions
function Initialize-WebRequestConfig {
    [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
    [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $global:ProgressPreference = 'SilentlyContinue'
}

function Get-PackageSize {
    param ([string]$url)
    try {
        $response = Invoke-WebRequest -UseBasicParsing -Method 'HEAD' -Uri $url
        return [math]::Round([long]$response.Headers['Content-Length'] / 1MB, 2)
    }
    catch {
        return 0
    }
}

function Save-Package {
    param (
        [string]$url,
        [string]$filePath
    )
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($url, $filePath)
}

function Reset-Form {
    $packageCombo.SelectedIndex = -1
    $productIdTextBox.Clear()
    $listView.Items.Clear()
    $progressBar.Value = 0
    $statusLabel.Text = "Ready"
    $downloadButton.Enabled = $false
    $archX64.Checked = $true
    $ringCombo.SelectedIndex = 3
}


# Initialize data
foreach ($item in $PackageList.Identity) {
    [void]$packageCombo.Items.Add($item)
}

foreach ($item in $RingList.Name) {
    [void]$ringCombo.Items.Add($item)
}
$ringCombo.SelectedIndex = 3

# Store response data
$script:lastResponse = $null

# Event Handlers
$packageCombo.Add_SelectedIndexChanged({
    if ($packageCombo.SelectedIndex -ne -1) {
        $productIdTextBox.Enabled = $false
        $productIdTextBox.Clear()
    }
    else {
        $productIdTextBox.Enabled = $true
    }
})

$productIdTextBox.Add_TextChanged({
    if ($productIdTextBox.Text.Length -gt 0) {
        $packageCombo.Enabled = $false
        $packageCombo.SelectedIndex = -1
    }
    else {
        $packageCombo.Enabled = $true
    }
})

$selectAllButton.Add_Click({
    foreach ($item in $listView.Items) {
        $item.Checked = $true
    }
})

$deselectAllButton.Add_Click({
    foreach ($item in $listView.Items) {
        $item.Checked = $false
    }
})

$resetButton.Add_Click({
    Reset-Form
})

$fetchButton.Add_Click({
    try {
        if ([string]::IsNullOrEmpty($packageCombo.Text) -and [string]::IsNullOrEmpty($productIdTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter either a Store URL or select a package.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $statusLabel.Text = "Fetching package information..."
        $progressBar.Value = 0
        $downloadButton.Enabled = $false
        
        Initialize-WebRequestConfig

        # Get selected architecture
        $selectedArch = if ($archX64.Checked) { "x64" }
                       elseif ($archX86.Checked) { "x86" }
                       elseif ($archArm64.Checked) { "ARM64" }
                       else { "ARM" }

        $wchttp = New-Object System.Net.WebClient
        $URI = "https://store.rg-adguard.net/api/GetFiles"
        
        if (-not [string]::IsNullOrEmpty($productIdTextBox.Text)) {
            $productId = $productIdTextBox.Text.Trim()
            $storeUrl = "https://apps.microsoft.com/store/detail/$productId"
            $myParameters = "type=url&url=$([System.Web.HttpUtility]::UrlEncode($storeUrl))"
        }
        else {
            $selectedPackage = $PackageList | Where-Object { $_.Identity -eq $packageCombo.Text }
            if ($selectedPackage) {
                $myParameters = "type=PackageFamilyName&url=$($selectedPackage.Family)&ring=$($RingList[$ringCombo.SelectedIndex].Value)&lang=en-US"
            }
            else {
                throw "Invalid package selection"
            }
        }

        $wchttp.Headers[[System.Net.HttpRequestHeader]::ContentType] = "application/x-www-form-urlencoded"
        $HtmlResult = $wchttp.UploadString($URI, $myParameters)
        
        # Store the response for later use
        $script:lastResponse = [regex]::Matches($HtmlResult, '<a[^>]*href="([^"]*)"[^>]*>([^<]*)</a>')
        
        if ($script:lastResponse.Count -eq 0) {
            throw "No packages found for the specified criteria"
        }

        $progressBar.Value = 50
        $statusLabel.Text = "Processing package information..."

        # Process and display packages
        $listView.Items.Clear()
        foreach ($link in $script:lastResponse) {
            $url = $link.Groups[1].Value
            $fileName = $link.Groups[2].Value

            if ($fileName -notmatch 'BlockMap' -and 
                $fileName -notmatch '\.eappx' -and 
                $fileName -notmatch '\.emsix' -and
                ($fileName -match $selectedArch -or $fileName -match '_neutral_')) {
                
                $size = Get-PackageSize -url $url
                
                $item = New-Object System.Windows.Forms.ListViewItem($fileName)
                $item.SubItems.Add($fileName.Split('_')[1])
                $item.SubItems.Add("$size MB")
                $item.SubItems.Add("Ready to download")
                $listView.Items.Add($item)
            }
        }

        $progressBar.Value = 100
        $statusLabel.Text = "Package information fetched successfully"
        $downloadButton.Enabled = $true
    }
    catch {
        $statusLabel.Text = "Error: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})









$downloadButton.Add_Click({
    try {
        $selectedItems = $listView.Items | Where-Object { $_.Checked }
        if ($selectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one package to download.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $downloadPath = Join-Path $PSScriptRoot "downloads"
        if (-not (Test-Path $downloadPath)) {
            New-Item -ItemType Directory -Path $downloadPath | Out-Null
        }

        # Determine folder name based on selection method
        $folderName = if (-not [string]::IsNullOrEmpty($productIdTextBox.Text)) {
            # Extract product ID from URL or use direct input
            if ($productIdTextBox.Text -match '\/([^\/\?]+)(?:\?|$)') {
                $matches[1]
            } else {
                $productIdTextBox.Text
            }
        } else {
            $packageCombo.Text
        }

        # Create package directory
        $packagePath = Join-Path $downloadPath $folderName
        if (-not (Test-Path $packagePath)) {
            New-Item -ItemType Directory -Path $packagePath | Out-Null
        }

        foreach ($item in $selectedItems) {
            $fileName = $item.Text
            $filePath = Join-Path $packagePath $fileName

            $matchingLink = $script:lastResponse | Where-Object { $_.Groups[2].Value -eq $fileName }
            if (-not $matchingLink) {
                $item.SubItems[3].Text = "Error: URL not found"
                continue
            }

            $url = $matchingLink.Groups[1].Value
            
            # Get file size
            $request = [System.Net.WebRequest]::Create($url)
            $totalBytes = $request.GetResponse().ContentLength
            $request.GetResponse().Close()
            
            # Create WebClient
            $webClient = New-Object System.Net.WebClient
            $totalMB = [Math]::Round($totalBytes / 1MB, 2)
            
            # Download with progress
            $downloadedBytes = 0
            $buffer = New-Object byte[] 8192
            $stream = $webClient.OpenRead($url)
            $fileStream = [System.IO.File]::Create($filePath)
            
            while (($read = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $fileStream.Write($buffer, 0, $read)
                $downloadedBytes += $read
                $downloaded = [Math]::Round($downloadedBytes / 1MB, 2)
                $percent = [Math]::Round(($downloadedBytes / $totalBytes) * 100)
                
                $progressBar.Value = $percent
                $item.SubItems[3].Text = "$downloaded MB / $totalMB MB ($percent%)"
                $statusLabel.Text = "Downloading $fileName : $downloaded MB / $totalMB MB"
                [System.Windows.Forms.Application]::DoEvents()
            }
            
            $fileStream.Close()
            $stream.Close()
            $webClient.Dispose()
            
            $item.SubItems[3].Text = "Completed ($totalMB MB)"
        }
        
        [System.Windows.Forms.MessageBox]::Show("All downloads completed!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $statusLabel.Text = "All downloads completed"
        $progressBar.Value = 100
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})





# Show form
$form.ShowDialog()