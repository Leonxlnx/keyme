$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$root = Split-Path -Parent $PSScriptRoot
$sourceRoot = Split-Path -Parent $root
$configDir = Join-Path $env:APPDATA "Keyme"
$configPath = Join-Path $configDir "config.json"
$startupShortcut = Join-Path ([Environment]::GetFolderPath("Startup")) "Keyme.lnk"
$launcher = Join-Path $PSScriptRoot "launch-hidden.vbs"
$profiles = @(
    [pscustomobject]@{ Id = "holy-panda"; Label = "Holy Panda - tactile thock" },
    [pscustomobject]@{ Id = "oil-king"; Label = "Oil King - deep linear" },
    [pscustomobject]@{ Id = "topre"; Label = "Topre - soft dome" },
    [pscustomobject]@{ Id = "box-jade"; Label = "Box Jade - crisp click" },
    [pscustomobject]@{ Id = "silent-tactile"; Label = "Silent Tactile - muted" },
    [pscustomobject]@{ Id = "ink-black"; Label = "Ink Black - low thock" },
    [pscustomobject]@{ Id = "nk-cream"; Label = "NK Cream - smooth pop" },
    [pscustomobject]@{ Id = "buckling-spring"; Label = "Buckling Spring - vintage" },
    [pscustomobject]@{ Id = "mx-black"; Label = "MX Black - classic linear" },
    [pscustomobject]@{ Id = "alps-blue"; Label = "Alps Blue - bright click" },
    [pscustomobject]@{ Id = "ceramic"; Label = "Ceramic - clean clack" },
    [pscustomobject]@{ Id = "terminal"; Label = "Terminal - retro board" },
    [pscustomobject]@{ Id = "alpaca"; Label = "Alpaca - soft pop" },
    [pscustomobject]@{ Id = "typewriter"; Label = "Typewriter - sharp strike" }
)

function Ensure-Config {
    if (-not (Test-Path $configPath)) {
        New-Item -ItemType Directory -Path $configDir -Force | Out-Null
        $defaultConfig = Join-Path $root "config\default.json"
        if (-not (Test-Path $defaultConfig)) {
            $defaultConfig = Join-Path $sourceRoot "config\default.json"
        }
        Copy-Item $defaultConfig $configPath -Force
    }
}

function Read-KeymeConfig {
    Ensure-Config
    Get-Content $configPath -Raw | ConvertFrom-Json
}

function Write-KeymeConfig {
    param([string]$Profile, [int]$Volume, [bool]$Autostart)

    New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    [pscustomobject]@{
        profile = $Profile
        volume = $Volume
        autostart = $Autostart
    } | ConvertTo-Json | Set-Content $configPath -Encoding UTF8
}

function Get-KeymeRunning {
    [bool](Get-Process -Name "keyme" -ErrorAction SilentlyContinue)
}

function Set-KeymeAutostart {
    param([bool]$Enabled)

    if ($Enabled) {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($startupShortcut)
        $shortcut.TargetPath = "$env:WINDIR\System32\wscript.exe"
        $shortcut.Arguments = "`"$launcher`""
        $shortcut.WorkingDirectory = $root
        $shortcut.Description = "Start Keyme at login"
        $shortcut.Save()
    }
    else {
        Remove-Item $startupShortcut -ErrorAction SilentlyContinue
    }
}

function Start-Keyme {
    & (Join-Path $PSScriptRoot "run.ps1")
}

function Stop-Keyme {
    & (Join-Path $PSScriptRoot "stop.ps1")
}

$config = Read-KeymeConfig

$form = New-Object Windows.Forms.Form
$form.Text = "Keyme"
$form.Width = 520
$form.Height = 500
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = [Drawing.Color]::FromArgb(244, 239, 229)
$form.ForeColor = [Drawing.Color]::FromArgb(30, 34, 34)

$card = New-Object Windows.Forms.Panel
$card.Location = New-Object Drawing.Point(24, 24)
$card.Size = New-Object Drawing.Size(456, 400)
$card.BackColor = [Drawing.Color]::FromArgb(255, 252, 245)
$form.Controls.Add($card)

$title = New-Object Windows.Forms.Label
$title.Text = "Keyme"
$title.Font = New-Object Drawing.Font("Segoe UI Variable Display", 28, [Drawing.FontStyle]::Bold)
$title.Location = New-Object Drawing.Point(28, 24)
$title.Size = New-Object Drawing.Size(260, 50)
$card.Controls.Add($title)

$subtitle = New-Object Windows.Forms.Label
$subtitle.Text = "Keyboard sounds for Windows"
$subtitle.Font = New-Object Drawing.Font("Segoe UI", 10)
$subtitle.ForeColor = [Drawing.Color]::FromArgb(91, 101, 97)
$subtitle.Location = New-Object Drawing.Point(31, 73)
$subtitle.Size = New-Object Drawing.Size(300, 24)
$card.Controls.Add($subtitle)

$status = New-Object Windows.Forms.Label
$status.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
$status.Location = New-Object Drawing.Point(31, 116)
$status.Size = New-Object Drawing.Size(390, 24)
$card.Controls.Add($status)

$profileLabel = New-Object Windows.Forms.Label
$profileLabel.Text = "Sound"
$profileLabel.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
$profileLabel.Location = New-Object Drawing.Point(31, 158)
$profileLabel.Size = New-Object Drawing.Size(160, 22)
$card.Controls.Add($profileLabel)

$profileBox = New-Object Windows.Forms.ComboBox
$profileBox.DropDownStyle = "DropDownList"
$profileBox.Location = New-Object Drawing.Point(31, 184)
$profileBox.Size = New-Object Drawing.Size(390, 28)
$profileBox.DisplayMember = "Label"
$profileBox.ValueMember = "Id"
[void]$profileBox.Items.AddRange($profiles)
$selectedProfile = $profiles | Where-Object { $_.Id -eq $config.profile } | Select-Object -First 1
if (-not $selectedProfile) { $selectedProfile = $profiles[0] }
$profileBox.SelectedItem = $selectedProfile
$card.Controls.Add($profileBox)

$volumeLabel = New-Object Windows.Forms.Label
$volumeLabel.Text = "Volume: $($config.volume)%"
$volumeLabel.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
$volumeLabel.Location = New-Object Drawing.Point(31, 236)
$volumeLabel.Size = New-Object Drawing.Size(160, 22)
$card.Controls.Add($volumeLabel)

$volumeSlider = New-Object Windows.Forms.TrackBar
$volumeSlider.Minimum = 0
$volumeSlider.Maximum = 100
$volumeSlider.TickFrequency = 10
$volumeSlider.Value = [Math]::Max(0, [Math]::Min(100, [int]$config.volume))
$volumeSlider.Location = New-Object Drawing.Point(26, 262)
$volumeSlider.Size = New-Object Drawing.Size(402, 45)
$volumeSlider.Add_ValueChanged({ $volumeLabel.Text = "Volume: $($volumeSlider.Value)%" })
$card.Controls.Add($volumeSlider)

$autostart = New-Object Windows.Forms.CheckBox
$autostart.Text = "Start automatically with Windows"
$autostart.Checked = [bool]$config.autostart
$autostart.Location = New-Object Drawing.Point(31, 315)
$autostart.Size = New-Object Drawing.Size(360, 26)
$autostart.ForeColor = [Drawing.Color]::FromArgb(30, 34, 34)
$card.Controls.Add($autostart)

$saveButton = New-Object Windows.Forms.Button
$saveButton.Text = "Apply"
$saveButton.Location = New-Object Drawing.Point(31, 354)
$saveButton.Size = New-Object Drawing.Size(90, 34)
$card.Controls.Add($saveButton)

$startButton = New-Object Windows.Forms.Button
$startButton.Text = "Restart sound"
$startButton.Location = New-Object Drawing.Point(135, 354)
$startButton.Size = New-Object Drawing.Size(120, 34)
$card.Controls.Add($startButton)

$stopButton = New-Object Windows.Forms.Button
$stopButton.Text = "Stop"
$stopButton.Location = New-Object Drawing.Point(269, 354)
$stopButton.Size = New-Object Drawing.Size(70, 34)
$card.Controls.Add($stopButton)

$openConfigButton = New-Object Windows.Forms.Button
$openConfigButton.Text = "File"
$openConfigButton.Location = New-Object Drawing.Point(353, 354)
$openConfigButton.Size = New-Object Drawing.Size(72, 36)
$card.Controls.Add($openConfigButton)

$privacy = New-Object Windows.Forms.Label
$privacy.Text = "Local only. No telemetry. No typed text is stored."
$privacy.Font = New-Object Drawing.Font("Segoe UI", 8.5)
$privacy.ForeColor = [Drawing.Color]::FromArgb(91, 101, 97)
$privacy.Location = New-Object Drawing.Point(31, 430)
$privacy.Size = New-Object Drawing.Size(420, 24)
$form.Controls.Add($privacy)

function Refresh-Status {
    if (Get-KeymeRunning) {
        $status.Text = "Status: running"
        $status.ForeColor = [Drawing.Color]::FromArgb(80, 220, 150)
    }
    else {
        $status.Text = "Status: stopped"
        $status.ForeColor = [Drawing.Color]::FromArgb(255, 190, 90)
    }
}

function Save-CurrentSettings {
    Write-KeymeConfig -Profile $profileBox.SelectedItem.Id -Volume $volumeSlider.Value -Autostart $autostart.Checked
    Set-KeymeAutostart -Enabled $autostart.Checked
}

$saveButton.Add_Click({
    Save-CurrentSettings
    Refresh-Status
})

$startButton.Add_Click({
    Save-CurrentSettings
    Stop-Keyme
    Start-Sleep -Milliseconds 250
    Start-Keyme
    Start-Sleep -Milliseconds 350
    Refresh-Status
})

$stopButton.Add_Click({
    Stop-Keyme
    Start-Sleep -Milliseconds 250
    Refresh-Status
})

$openConfigButton.Add_Click({
    Ensure-Config
    Start-Process notepad.exe $configPath
})

Refresh-Status
[void]$form.ShowDialog()
