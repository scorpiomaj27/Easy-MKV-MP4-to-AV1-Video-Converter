# Easy AV1 Converter
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- Settings ----------
$DefaultBitrate = '5M'

# Resolve script directory robustly and store in script scope
$script:ScriptDir = $PSScriptRoot
if (-not $script:ScriptDir) { $script:ScriptDir = Split-Path -Path $PSCommandPath -Parent }
if (-not $script:ScriptDir) { $script:ScriptDir = (Get-Location).Path }

# Resolve this script path for persistence operations
$script:ThisScriptPath = $PSCommandPath
if (-not $script:ThisScriptPath -and $MyInvocation -and $MyInvocation.MyCommand) {
    $script:ThisScriptPath = $MyInvocation.MyCommand.Path
}
if (-not $script:ThisScriptPath) {
    $script:ThisScriptPath = Join-Path -Path $script:ScriptDir -ChildPath 'EasyAV1Converter.ps1'
}

# Optional: explicit ffmpeg path (falls back to PATH if not found)
$script:FfmpegPath = ''

# Adaptive calibration (auto-updated by script)
$script:CalibrationFactor = 1.0
$script:ConversionHistory = @()

# Recent folder paths (auto-updated by script)
$script:RecentPaths = @()

# Track if we've already warned about competing processes
$script:HasWarnedAboutCompetingProcesses = $false

# ---------- UI ----------
function New-Form {
    $form               = New-Object System.Windows.Forms.Form
    $form.Text          = 'Easy AV1 Converter'
    $form.StartPosition = 'CenterScreen'
    $form.Size          = New-Object System.Drawing.Size(900, 640)
    $form.MaximizeBox   = $true

    # Folder dropdown label
    $lblFolder            = New-Object System.Windows.Forms.Label
    $lblFolder.Text       = "Folder:"
    $lblFolder.AutoSize   = $true
    $lblFolder.Location   = New-Object System.Drawing.Point(10, 10)
    $form.Controls.Add($lblFolder)
    
    # Folder dropdown
    $script:cmbFolder               = New-Object System.Windows.Forms.ComboBox
    $script:cmbFolder.DropDownStyle = 'DropDown'
    $script:cmbFolder.Location      = New-Object System.Drawing.Point(60, 8)
    $script:cmbFolder.Size          = New-Object System.Drawing.Size(480, 28)
    $script:cmbFolder.Text          = $script:ScriptDir
    $form.Controls.Add($script:cmbFolder)
    
    # Browse folder button
    $btnBrowseFolder           = New-Object System.Windows.Forms.Button
    $btnBrowseFolder.Text      = '...'
    $btnBrowseFolder.Location  = New-Object System.Drawing.Point(550, 8)
    $btnBrowseFolder.Size      = New-Object System.Drawing.Size(30, 22)
    $form.Controls.Add($btnBrowseFolder)

    # FFmpeg path label
    $script:lblFfmpeg           = New-Object System.Windows.Forms.Label
    $script:lblFfmpeg.Text      = 'FFmpeg: resolving...'
    $script:lblFfmpeg.AutoSize  = $true
    $script:lblFfmpeg.Location  = New-Object System.Drawing.Point(10, 35)
    $script:lblFfmpeg.MaximumSize= New-Object System.Drawing.Size(470, 0)
    $script:lblFfmpeg.AutoEllipsis = $true
    $form.Controls.Add($script:lblFfmpeg)
    
    # Browse FFmpeg button
    $btnBrowseFfmpeg           = New-Object System.Windows.Forms.Button
    $btnBrowseFfmpeg.Text      = 'Browse...'
    $btnBrowseFfmpeg.Location  = New-Object System.Drawing.Point(490, 33)
    $btnBrowseFfmpeg.Size      = New-Object System.Drawing.Size(80, 22)
    $form.Controls.Add($btnBrowseFfmpeg)

    # Current file label
    $script:lblCurrent           = New-Object System.Windows.Forms.Label
    $script:lblCurrent.Text      = 'Current: (idle)'
    $script:lblCurrent.AutoSize  = $true
    $script:lblCurrent.Location  = New-Object System.Drawing.Point(10, 60)
    $script:lblCurrent.MaximumSize= New-Object System.Drawing.Size(860, 0)
    $script:lblCurrent.AutoEllipsis = $true
    $form.Controls.Add($script:lblCurrent)

    # File list (ListView)
    $script:lst                  = New-Object System.Windows.Forms.ListView
    $script:lst.View             = 'Details'
    $script:lst.FullRowSelect    = $true
    $script:lst.MultiSelect      = $true
    $script:lst.Location         = New-Object System.Drawing.Point(10, 85)
    $script:lst.Size             = New-Object System.Drawing.Size(540, 365)
    [void]$script:lst.Columns.Add('Name', 280)
    [void]$script:lst.Columns.Add('Size', 80)
    [void]$script:lst.Columns.Add('Codec', 70)
    [void]$script:lst.Columns.Add('Est. Time', 90)
    $form.Controls.Add($script:lst)

    # Refresh
    $btnRefresh           = New-Object System.Windows.Forms.Button
    $btnRefresh.Text      = 'Refresh'
    $btnRefresh.Location  = New-Object System.Drawing.Point(10, 460)
    $btnRefresh.Size      = New-Object System.Drawing.Size(80, 28)
    $form.Controls.Add($btnRefresh)

    # View all video files
    $script:chkAll              = New-Object System.Windows.Forms.CheckBox
    $script:chkAll.Text         = 'View all video files'
    $script:chkAll.AutoSize     = $true
    $script:chkAll.Checked      = $false
    $script:chkAll.Location     = New-Object System.Drawing.Point(100, 463)
    $form.Controls.Add($script:chkAll)

    # Dark mode
    $script:chkDark               = New-Object System.Windows.Forms.CheckBox
    $script:chkDark.Text          = 'Dark mode'
    $script:chkDark.AutoSize      = $true
    $script:chkDark.Checked       = $false
    $script:chkDark.Location      = New-Object System.Drawing.Point(230, 463)
    $form.Controls.Add($script:chkDark)

    # Bitrate (fixed options)
    $lblBr                = New-Object System.Windows.Forms.Label
    $lblBr.Text           = 'Bitrate'
    $lblBr.AutoSize       = $true
    $lblBr.Location       = New-Object System.Drawing.Point(590, 35)
    $form.Controls.Add($lblBr)

    $script:cmbBr               = New-Object System.Windows.Forms.ComboBox
    $script:cmbBr.DropDownStyle = 'DropDownList'
    $script:cmbBr.Location      = New-Object System.Drawing.Point(590, 55)
    $script:cmbBr.Size          = New-Object System.Drawing.Size(140, 28)
    [void]$script:cmbBr.Items.AddRange(@('2.5M','5M','7.5M','10M'))
    $script:cmbBr.SelectedItem = '5M'
    $form.Controls.Add($script:cmbBr)

    # Rename .old checkbox
    $script:chkRename            = New-Object System.Windows.Forms.CheckBox
    $script:chkRename.Text       = 'Rename original to .old after convert'
    $script:chkRename.Checked    = $true
    $script:chkRename.AutoSize   = $true
    $script:chkRename.Location   = New-Object System.Drawing.Point(590, 100)
    $form.Controls.Add($script:chkRename)

    # Move to _Old checkbox
    $script:chkMoveOld           = New-Object System.Windows.Forms.CheckBox
    $script:chkMoveOld.Text      = 'Move files to Old Folder After Conversion'
    $script:chkMoveOld.Checked   = $true
    $script:chkMoveOld.AutoSize  = $true
    $script:chkMoveOld.Location  = New-Object System.Drawing.Point(590, 122)
    $form.Controls.Add($script:chkMoveOld)

    # Convert button
    $script:btnConvert           = New-Object System.Windows.Forms.Button
    $script:btnConvert.Text      = 'Convert to AV1'
    $script:btnConvert.Location  = New-Object System.Drawing.Point(590, 165)
    $script:btnConvert.Size      = New-Object System.Drawing.Size(140, 34)
    $form.Controls.Add($script:btnConvert)

    # Test Speed button
    $script:btnTestSpeed         = New-Object System.Windows.Forms.Button
    $script:btnTestSpeed.Text    = 'Test Speed'
    $script:btnTestSpeed.Location = New-Object System.Drawing.Point(590, 205)
    $script:btnTestSpeed.Size    = New-Object System.Drawing.Size(90, 28)
    $form.Controls.Add($script:btnTestSpeed)
    
    # Adaptive Test button
    $script:btnAdaptiveTest      = New-Object System.Windows.Forms.Button
    $script:btnAdaptiveTest.Text = 'Adaptive Test'
    $script:btnAdaptiveTest.Location = New-Object System.Drawing.Point(690, 205)
    $script:btnAdaptiveTest.Size = New-Object System.Drawing.Size(90, 28)
    $form.Controls.Add($script:btnAdaptiveTest)

    # Cancel button
    $script:btnCancel            = New-Object System.Windows.Forms.Button
    $script:btnCancel.Text       = 'Cancel'
    $script:btnCancel.Location   = New-Object System.Drawing.Point(740, 165)
    $script:btnCancel.Size       = New-Object System.Drawing.Size(120, 34)
    $script:btnCancel.Enabled    = $false
    $form.Controls.Add($script:btnCancel)

    # Log output
    $script:txtLog               = New-Object System.Windows.Forms.TextBox
    $script:txtLog.Location      = New-Object System.Drawing.Point(10, 500)
    $script:txtLog.Size          = New-Object System.Drawing.Size(860, 100)
    $script:txtLog.Multiline     = $true
    $script:txtLog.ScrollBars    = 'Vertical'
    $script:txtLog.ReadOnly      = $true
    $form.Controls.Add($script:txtLog)

    # ---------- Helpers ----------
    $script:CodecCache = @{}

    function script:Append-Log($text) {
        $script:txtLog.AppendText($text + "`r`n")
        $script:txtLog.SelectionStart = $script:txtLog.Text.Length
        $script:txtLog.ScrollToCaret()
    }

    function script:Apply-Theme {
        param([bool]$Dark)
        $bg = if ($Dark) { [System.Drawing.Color]::FromArgb(30,30,30) } else { [System.Drawing.SystemColors]::Window }
        $fg = if ($Dark) { [System.Drawing.Color]::WhiteSmoke } else { [System.Drawing.SystemColors]::WindowText }
        $form.BackColor = $bg
        foreach ($ctrl in $form.Controls) {
            try {
                if ($ctrl -is [System.Windows.Forms.TextBox]) { $ctrl.BackColor = $bg; $ctrl.ForeColor = $fg }
                elseif ($ctrl -is [System.Windows.Forms.ListView]) { $ctrl.BackColor = $bg; $ctrl.ForeColor = $fg }
                elseif ($ctrl -is [System.Windows.Forms.ComboBox]) { $ctrl.BackColor = $bg; $ctrl.ForeColor = $fg }
                elseif ($ctrl -is [System.Windows.Forms.Label]) { $ctrl.ForeColor = $fg }
                elseif ($ctrl -is [System.Windows.Forms.Button]) { $ctrl.ForeColor = $fg }
                elseif ($ctrl -is [System.Windows.Forms.CheckBox]) { $ctrl.ForeColor = $fg }
            } catch {}
        }
    }
    
    function script:Load-Calibration {
        # Calibration is already loaded from the script variables
        if ($script:ConversionHistory.Count -gt 0) {
            Append-Log ("Loaded calibration: factor={0:N2}, history={1} conversions" -f $script:CalibrationFactor, $script:ConversionHistory.Count)
        }
    }
    
    function script:Save-Calibration {
        try {
            if (-not (Test-Path -LiteralPath $script:ThisScriptPath)) { return }
            $content = Get-Content -LiteralPath $script:ThisScriptPath -Raw
            
            # Update CalibrationFactor
            $pattern = '(?m)^\$script:CalibrationFactor\s*=\s*[\d.]+'
            $replacement = ('$script:CalibrationFactor = {0:N2}' -f $script:CalibrationFactor)
            $content = [Regex]::Replace($content, $pattern, $replacement)
            
            # Update ConversionHistory (serialize to PowerShell array syntax)
            $historyLines = @()
            foreach ($h in $script:ConversionHistory) {
                $historyLines += "@{{File='{0}';Estimated={1:N1};Actual={2:N1};Ratio={3:N3};Date='{4}'}}" -f `
                    $h.File.Replace("'","''"), $h.Estimated, $h.Actual, $h.Ratio, $h.Date
            }
            $historyArray = if ($historyLines.Count -gt 0) { "@(" + ($historyLines -join ",") + ")" } else { "@()" }
            
            # Use a more robust pattern that handles nested parentheses and braces
            $pattern = '(?ms)^\$script:ConversionHistory\s*=\s*@\(.*?\)(?=\s*$)'
            $replacement = ('$script:ConversionHistory = {0}' -f $historyArray)
            $content = [Regex]::Replace($content, $pattern, $replacement)
            
            # Update RecentPaths
            $pathsArray = if ($script:RecentPaths.Count -gt 0) { "@('" + ($script:RecentPaths -join "','") + "')" } else { "@()" }
            $pattern = '(?m)^\$script:RecentPaths\s*=\s*@\(.*?\)'
            $replacement = ('$script:RecentPaths = {0}' -f $pathsArray)
            $content = [Regex]::Replace($content, $pattern, $replacement)
            
            Set-Content -LiteralPath $script:ThisScriptPath -Value $content -Encoding UTF8 -NoNewline
        } catch {
            Append-Log ("Failed to save calibration: {0}" -f $_.Exception.Message)
        }
    }
    
    function script:Update-Calibration {
        param([double]$EstimatedSeconds, [double]$ActualSeconds, [string]$FileName)
        
        # Add to history (keep last 10)
        $script:ConversionHistory += [PSCustomObject]@{
            File = $FileName
            Estimated = $EstimatedSeconds
            Actual = $ActualSeconds
            Ratio = $ActualSeconds / $EstimatedSeconds
            Date = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }
        
        if ($script:ConversionHistory.Count > 10) {
            $script:ConversionHistory = $script:ConversionHistory[-10..-1]
        }
        
        # Calculate new calibration factor (average of last 10 ratios)
        $avgRatio = ($script:ConversionHistory | Measure-Object -Property Ratio -Average).Average
        $script:CalibrationFactor = $avgRatio
        
        Append-Log ("Calibration updated: factor={0:N2} (based on {1} conversions)" -f $script:CalibrationFactor, $script:ConversionHistory.Count)
        
        Save-Calibration
    }

    function script:Persist-FfmpegPath {
        param([Parameter(Mandatory=$true)][string]$Path)
        try {
            if (-not (Test-Path -LiteralPath $script:ThisScriptPath)) { return }
            $content = Get-Content -LiteralPath $script:ThisScriptPath -Raw
            $pattern = '(?m)^\s*\$script:FfmpegPath\s*=\s*''[^'']*'''
            $replacement = ('$script:FfmpegPath = ''{0}''' -f $Path)
            if ($content -match $pattern) {
                $newContent = [Regex]::Replace($content, $pattern, $replacement)
            } else {
                $newContent = $content + ("`r`n$replacement`r`n")
            }
            if ($newContent -ne $content) {
                Set-Content -LiteralPath $script:ThisScriptPath -Value $newContent -Encoding UTF8
                Append-Log ("Saved ffmpeg path for next run: {0}" -f $Path)
            }
        } catch {
            Append-Log ("Failed to persist ffmpeg path: {0}" -f $_.Exception.Message)
        }
    }

    function script:Add-RecentPath {
        param([Parameter(Mandatory=$true)][string]$Path)
        
        # Remove if already exists
        $script:RecentPaths = @($script:RecentPaths | Where-Object { $_ -ne $Path })
        
        # Add to front
        $script:RecentPaths = @($Path) + $script:RecentPaths
        
        # Keep only last 3
        if ($script:RecentPaths.Count > 3) {
            $script:RecentPaths = $script:RecentPaths[0..2]
        }
        
        # Update dropdown
        Update-FolderDropdown
        
        # Save to file
        Save-Calibration
    }
    
    function script:Update-FolderDropdown {
        $script:cmbFolder.Items.Clear()
        
        # Add current path if not in recent paths
        $currentPath = $script:cmbFolder.Text
        if ($currentPath -and $script:RecentPaths -notcontains $currentPath) {
            [void]$script:cmbFolder.Items.Add($currentPath)
        }
        
        # Add recent paths
        foreach ($path in $script:RecentPaths) {
            if ($path -and (Test-Path -Path $path)) {
                [void]$script:cmbFolder.Items.Add($path)
            }
        }
    }
    
    function script:Change-WorkingFolder {
        param([string]$NewPath)
        
        if (-not $NewPath -or -not (Test-Path -Path $NewPath)) {
            [System.Windows.Forms.MessageBox]::Show('Invalid folder path.', 'Error', 'OK', 'Error') | Out-Null
            return
        }
        
        $script:ScriptDir = $NewPath
        $script:cmbFolder.Text = $NewPath
        Add-RecentPath -Path $NewPath
        
        # Refresh file list
        Load-Files -All:$script:chkAll.Checked
    }

    function script:Update-FfmpegLabel {
        try {
            $path = $null
            if ($script:FfmpegInUse) { $path = $script:FfmpegInUse }
            elseif ($script:FfmpegPath -and (Test-Path -LiteralPath $script:FfmpegPath)) { $path = $script:FfmpegPath }
            else {
                try { $cmd = Get-Command ffmpeg -ErrorAction Stop } catch {}
                if ($cmd -and $cmd.Path) { $path = $cmd.Path }
            }
            if (-not $path) { $path = '(not found)' }
            $script:lblFfmpeg.Text = "FFmpeg: $path"
        } catch { $script:lblFfmpeg.Text = 'FFmpeg: (error resolving)' }
    }

    function script:Get-VideoCodec([string]$FullPath) {
        if ($script:CodecCache.ContainsKey($FullPath)) { return $script:CodecCache[$FullPath] }
        $codec = $null
        $ffprobe = $null
        try { $ffprobe = (Get-Command ffprobe -ErrorAction Stop).Path } catch {}
        if (-not $ffprobe -and $script:FfmpegPath) {
            try { $ffdir = [System.IO.Path]::GetDirectoryName($script:FfmpegPath); $cand = Join-Path -Path $ffdir -ChildPath 'ffprobe.exe'; if (Test-Path -LiteralPath $cand) { $ffprobe = $cand } } catch {}
        }
        if ($ffprobe) {
            try {
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName  = $ffprobe
                $psi.Arguments = ('-v error -select_streams v:0 -show_entries stream=codec_name -of default=nw=1:nk=1 "{0}"' -f $FullPath)
                $psi.RedirectStandardOutput = $true
                $psi.RedirectStandardError  = $true
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $true
                $p = New-Object System.Diagnostics.Process
                $p.StartInfo = $psi
                [void]$p.Start()
                $out = $p.StandardOutput.ReadToEnd().Trim()
                $p.WaitForExit()
                if ($out) { $codec = $out }
            } catch {}
        }
        if (-not $codec) {
            try {
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = (if ($script:FfmpegPath -and (Test-Path -LiteralPath $script:FfmpegPath)) { $script:FfmpegPath } else { 'ffmpeg' })
                $psi.Arguments = ('-hide_banner -i "{0}"' -f $FullPath)
                $psi.RedirectStandardError = $true
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $true
                $p = New-Object System.Diagnostics.Process
                $p.StartInfo = $psi
                [void]$p.Start()
                $stderr = $p.StandardError.ReadToEnd()
                $p.WaitForExit()
                $m = [Regex]::Match($stderr, 'Video:\s*([^,\s]+)')
                if ($m.Success) { $codec = $m.Groups[1].Value }
            } catch {}
        }
        if (-not $codec) { $codec = 'unknown' }
        $script:CodecCache[$FullPath] = $codec.ToLower()
        return $script:CodecCache[$FullPath]
    }

    function script:Get-VideoFrameCount([string]$FullPath) {
        $ffprobe = $null
        try { $ffprobe = (Get-Command ffprobe -ErrorAction Stop).Path } catch {}
        if (-not $ffprobe -and $script:FfmpegPath) {
            try { $ffdir = [System.IO.Path]::GetDirectoryName($script:FfmpegPath); $cand = Join-Path -Path $ffdir -ChildPath 'ffprobe.exe'; if (Test-Path -LiteralPath $cand) { $ffprobe = $cand } } catch {}
        }
        if ($ffprobe) {
            try {
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName  = $ffprobe
                $psi.Arguments = ('-v error -select_streams v:0 -count_packets -show_entries stream=nb_read_packets -of csv=p=0 "{0}"' -f $FullPath)
                $psi.RedirectStandardOutput = $true
                $psi.RedirectStandardError  = $true
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $true
                $p = New-Object System.Diagnostics.Process
                $p.StartInfo = $psi
                [void]$p.Start()
                $out = $p.StandardOutput.ReadToEnd().Trim()
                $p.WaitForExit()
                if ($out -and $out -match '^\d+$') { return [int]$out }
            } catch {}
        }
        return 0
    }

    function script:Get-Av1Encoder {
        if ($script:Av1Encoder) { return $script:Av1Encoder }
        $ff = if ($script:FfmpegPath -and (Test-Path -LiteralPath $script:FfmpegPath)) { $script:FfmpegPath } else { 'ffmpeg' }
        try {
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $ff
            $psi.Arguments = '-encoders'
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError  = $true
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            $p = New-Object System.Diagnostics.Process
            $p.StartInfo = $psi
            [void]$p.Start()
            $txt = $p.StandardOutput.ReadToEnd() + "`n" + $p.StandardError.ReadToEnd()
            $p.WaitForExit()
            # Look for av1_nvenc anywhere in the output (simpler and more reliable)
            if ($txt -match 'av1_nvenc') { $script:Av1Encoder = 'av1_nvenc' }
            elseif ($txt -match 'libaom-av1') { $script:Av1Encoder = 'libaom-av1' }
            elseif ($txt -match 'librav1e') { $script:Av1Encoder = 'librav1e' }
        } catch {}
        if (-not $script:Av1Encoder) { $script:Av1Encoder = 'libaom-av1' }
        return $script:Av1Encoder
    }

    function script:Run-FFmpeg-AV1($inputPath, $outputPath, $bitrate) {
        $ffmpeg = $null
        if ($script:FfmpegPath -and (Test-Path -LiteralPath $script:FfmpegPath)) {
            $ffmpeg = $script:FfmpegPath
            $script:FfmpegInUse = $script:FfmpegPath
        } else {
            $cmd = $null
            try { $cmd = Get-Command ffmpeg -ErrorAction Stop } catch {}
            if ($cmd) {
                $ffmpeg = 'ffmpeg'
                if ($cmd.Path) { $script:FfmpegInUse = $cmd.Path }
            } else {
                $resp = [System.Windows.Forms.MessageBox]::Show('ffmpeg.exe not found. Would you like to browse for it?', 'FFmpeg', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Exclamation)
                if ($resp -eq [System.Windows.Forms.DialogResult]::Yes) {
                    $ofd = New-Object System.Windows.Forms.OpenFileDialog
                    $ofd.Title = 'Locate ffmpeg.exe'
                    $ofd.Filter = 'ffmpeg.exe|ffmpeg.exe|Executables|*.exe'
                    $ofd.InitialDirectory = $script:ScriptDir
                    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                        $script:FfmpegPath = $ofd.FileName
                        Persist-FfmpegPath -Path $script:FfmpegPath
                        $ffmpeg = $script:FfmpegPath
                        $script:FfmpegInUse = $script:FfmpegPath
                    }
                }
                if (-not $ffmpeg) { [System.Windows.Forms.MessageBox]::Show('Cannot proceed without ffmpeg.exe. Conversion aborted.', 'FFmpeg', 'OK', 'Exclamation') | Out-Null; return 1 }
            }
        }
        Update-FfmpegLabel
        $enc = Get-Av1Encoder
        Append-Log ("========================================")
        Append-Log ("Using AV1 encoder: {0}" -f $enc)
        if ($enc -eq 'av1_nvenc') {
            Append-Log ("NVENC detected - Using GPU hardware acceleration!")
        } else {
            Append-Log ("WARNING: Using CPU encoder - this will be SLOW!")
            Append-Log ("For faster encoding, install FFmpeg with NVENC support.")
        }
        Append-Log ("========================================")
        
        $ffArgs = @('-y')
        $ffArgs += @('-i', ('"{0}"' -f $inputPath))
        if ($enc -eq 'av1_nvenc') {
            # NVENC settings matching your working batch script
            $ffArgs += @('-c:v','av1_nvenc','-b:v',$bitrate,'-maxrate',$bitrate,'-bufsize','20M','-preset','fast')
        } elseif ($enc -eq 'libaom-av1') {
            $ffArgs += @('-c:v','libaom-av1','-b:v',$bitrate,'-cpu-used','8','-row-mt','1','-threads','0','-pix_fmt','yuv420p')
        } elseif ($enc -eq 'librav1e') {
            $ffArgs += @('-c:v','librav1e','-b:v',$bitrate,'-speed','10','-pix_fmt','yuv420p')
        } else {
            $ffArgs += @('-c:v','libaom-av1','-b:v',$bitrate,'-cpu-used','8','-row-mt','1','-threads','0','-pix_fmt','yuv420p')
        }
        $ffArgs += @('-c:a','copy', ('"{0}"' -f $outputPath))

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $ffmpeg
        $psi.Arguments = ($ffArgs -join ' ')
        $psi.RedirectStandardError = $true
        $psi.RedirectStandardOutput = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true

        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $psi
        $script:CurrentProc = $proc
        [void]$proc.Start()
        while (-not $proc.HasExited) {
            if ($script:CancelRequested) { try { $proc.Kill() } catch {}; break }
            $line = $proc.StandardError.ReadLine()
            if ($null -ne $line) { Append-Log $line }
            Start-Sleep -Milliseconds 30
            [System.Windows.Forms.Application]::DoEvents()
        }
        while (-not $proc.StandardError.EndOfStream) { Append-Log ($proc.StandardError.ReadLine()) }
        $proc.WaitForExit()
        return $proc.ExitCode
    }

    function script:Process-Queue-AV1 {
        param([string[]]$Names)
        $script:IsRunning = $true
        $script:btnConvert.Enabled = $false
        $script:btnCancel.Enabled  = $true
        
        # Track conversion times
        $conversionResults = @()

        while ($Names.Count -gt 0) {
            if ($script:CancelRequested) { Append-Log 'Cancellation requested. Stopping.'; break }
            $name = $Names[0]
            if ($Names.Count -gt 1) { $Names = $Names[1..($Names.Count-1)] } else { $Names = @() }

            $ext = [System.IO.Path]::GetExtension($name)
            if (@('.mkv','.mp4') -notcontains $ext.ToLower()) { Append-Log ("Skipping non-video: {0}" -f $name); continue }

            $in = Join-Path -Path $script:ScriptDir -ChildPath $name
            $dir  = [System.IO.Path]::GetDirectoryName($in)
            $base = [System.IO.Path]::GetFileNameWithoutExtension($in)
            $out  = Join-Path -Path $dir -ChildPath ("{0}-AV1{1}" -f $base, $ext)

            # Track for UI and cancel cleanup
            $script:CurrentName = $name
            $script:CurrentOutPath = $out
            $script:lblCurrent.Text = ("Current: {0}" -f $name)

            # If output file already exists, rename the source file to .orig extension
            if (Test-Path $out) {
                Append-Log ("Output file already exists: {0}" -f [System.IO.Path]::GetFileName($out))
                try {
                    $baseNameOnly = [System.IO.Path]::GetFileNameWithoutExtension($in)
                    $origName = "{0}.orig{1}" -f $baseNameOnly, $ext
                    $origPath = Join-Path -Path $dir -ChildPath $origName
                    
                    # If .orig already exists, add a number
                    $counter = 1
                    while (Test-Path $origPath) {
                        $origName = "{0}.orig{1}{2}" -f $baseNameOnly, $counter, $ext
                        $origPath = Join-Path -Path $dir -ChildPath $origName
                        $counter++
                    }
                    
                    Rename-Item -LiteralPath $in -NewName $origName -ErrorAction Stop
                    Append-Log ("Renamed source to: {0}" -f $origName)
                    
                    # Update input path to the renamed file
                    $in = $origPath
                } catch {
                    Append-Log ("Failed to rename source file: {0}" -f $_.Exception.Message)
                    Append-Log ("Skipping {0} - cannot proceed with existing output file." -f $name)
                    continue
                }
            }

            Append-Log ("Converting {0} -> {1} at {2}" -f $name, [System.IO.Path]::GetFileName($out), $script:cmbBr.SelectedItem)
            
            # Track conversion time for adaptive calibration
            $conversionStart = Get-Date
            $code = Run-FFmpeg-AV1 -inputPath $in -outputPath $out -bitrate $script:cmbBr.SelectedItem
            $conversionEnd = Get-Date
            $actualSeconds = ($conversionEnd - $conversionStart).TotalSeconds
            
            if ($code -eq 0) {
                Append-Log ("Done: {0} (took {1:N1} minutes)" -f $name, ($actualSeconds / 60))
                
                # Add to results
                $conversionResults += [PSCustomObject]@{
                    FileName = $name
                    ActualMinutes = [Math]::Round($actualSeconds / 60, 1)
                    ActualSeconds = $actualSeconds
                }
                
                # Update adaptive calibration if we have an estimate for this file
                foreach ($item in $script:lst.Items) {
                    if ($item.Tag -eq $name -and $item.SubItems[3].Text -match '([\d.]+)m') {
                        $estimatedMinutes = [double]$matches[1]
                        $estimatedSeconds = $estimatedMinutes * 60
                        Update-Calibration -EstimatedSeconds $estimatedSeconds -ActualSeconds $actualSeconds -FileName $name
                        break
                    }
                }
                
                if ($script:chkRename.Checked) {
                    try {
                        $oldPath = [System.IO.Path]::ChangeExtension($in, '.old')
                        Rename-Item -LiteralPath $in -NewName ([System.IO.Path]::GetFileName($oldPath)) -ErrorAction Stop
                        Append-Log ("Renamed original to {0}" -f ([System.IO.Path]::GetFileName($oldPath)))
                    } catch { Append-Log ("Failed to rename original: {0}" -f $_.Exception.Message) }
                }
                if ($script:chkMoveOld.Checked) {
                    try {
                        $oldDir = Join-Path -Path $dir -ChildPath '_Old'
                        if (-not (Test-Path -LiteralPath $oldDir)) { [void](New-Item -ItemType Directory -Path $oldDir -Force) }
                        $toMove = if ($script:chkRename.Checked) { [System.IO.Path]::ChangeExtension($in, '.old') } else { $in }
                        if (Test-Path -LiteralPath $toMove) {
                            $dest = Join-Path -Path $oldDir -ChildPath ([System.IO.Path]::GetFileName($toMove))
                            Move-Item -LiteralPath $toMove -Destination $dest -Force
                            Append-Log ("Moved file to {0}" -f $dest)
                        }
                    } catch { Append-Log ("Failed to move file to _Old: {0}" -f $_.Exception.Message) }
                }
            } else {
                Append-Log ("ffmpeg exited with code {0} for {1}" -f $code, $name)
            }

            Append-Log ('-'*60)
            $script:CurrentName = $null
            $script:CurrentOutPath = $null
            $script:lblCurrent.Text = 'Current: (idle)'
            $script:PendingNames = $Names
        }

        $script:IsRunning = $false
        if ($script:CancelRequested) {
            [System.Windows.Forms.MessageBox]::Show('Conversion cancelled.', 'FFmpeg', 'OK', 'Exclamation') | Out-Null
        } else {
            # Build completion message with actual times
            if ($conversionResults.Count -gt 0) {
                $totalMinutes = ($conversionResults | Measure-Object -Property ActualMinutes -Sum).Sum
                $msg = "CONVERSIONS COMPLETE!`r`n`r`n"
                $msg += "Files converted: {0}`r`n" -f $conversionResults.Count
                $msg += "Total time: {0:N1} minutes ({1:N1} hours)`r`n`r`n" -f $totalMinutes, ($totalMinutes / 60)
                $msg += "Per-file times:`r`n"
                $msg += ("-" * 50) + "`r`n"
                foreach ($result in $conversionResults) {
                    $msg += "{0}`r`n  {1:N1} minutes`r`n" -f $result.FileName, $result.ActualMinutes
                }
                [System.Windows.Forms.MessageBox]::Show($msg, 'Conversions Complete', 'OK', 'Exclamation') | Out-Null
            } else {
                [System.Windows.Forms.MessageBox]::Show('Conversions complete.', 'FFmpeg', 'OK', 'Exclamation') | Out-Null
            }
            $script:PendingNames = @()
        }

        $script:btnConvert.Enabled = $true
        $script:btnCancel.Enabled  = $false
    }

    function script:Check-CompetingProcesses {
        $competingProcesses = @()
        
        # Check for UniFab
        $unifab = Get-Process -Name "unifab64" -ErrorAction SilentlyContinue
        if ($unifab) { 
            $count = @($unifab).Count
            $competingProcesses += if ($count -gt 1) { "UniFab ($count instances)" } else { "UniFab (unifab64.exe)" }
        }
        
        # Check for Folding@Home
        $fahClient = Get-Process -Name "FAHClient" -ErrorAction SilentlyContinue
        if ($fahClient) { 
            $count = @($fahClient).Count
            $competingProcesses += if ($count -gt 1) { "Folding@Home Client ($count instances)" } else { "Folding@Home Client (FAHClient.exe)" }
        }
        
        $fahCore = Get-Process | Where-Object { $_.ProcessName -like "FAHCore_*" }
        if ($fahCore) { 
            $count = @($fahCore).Count
            $competingProcesses += if ($count -gt 1) { "Folding@Home Core ($count instances)" } else { "Folding@Home Core (FAHCore_*.exe)" }
        }
        
        # Check for DaVinci Resolve
        $resolve = Get-Process -Name "Resolve" -ErrorAction SilentlyContinue
        if ($resolve) { 
            $count = @($resolve).Count
            $competingProcesses += if ($count -gt 1) { "DaVinci Resolve ($count instances)" } else { "DaVinci Resolve" }
        }
        
        # Check for Invoke Community Edition - check multiple possible process names
        $invoke = @()
        $invoke += Get-Process -Name "Invoke*" -ErrorAction SilentlyContinue
        $invoke += Get-Process | Where-Object { $_.ProcessName -like "*Invoke*" -and $_.ProcessName -notlike "Invoke-*" }
        $invoke = $invoke | Select-Object -Unique
        if ($invoke) { 
            $count = @($invoke).Count
            $competingProcesses += if ($count -gt 1) { "Invoke Community Edition ($count instances)" } else { "Invoke Community Edition" }
        }
        
        return $competingProcesses
    }

    function script:Test-EncodingSpeed {
        param([string[]]$Names)
        
        if ($Names.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Select at least one file to test.', 'Test Speed', 'OK', 'Exclamation') | Out-Null
            return
        }
        
        # Check for competing processes (only warn once per session)
        if (-not $script:HasWarnedAboutCompetingProcesses) {
            $competing = Check-CompetingProcesses
            if ($competing.Count -gt 0) {
                $msg = "WARNING: The following applications may interfere with encoding performance:`r`n`r`n"
                $msg += ($competing | ForEach-Object { "  - $_" }) -join "`r`n"
                $msg += "`r`n`r`nIt is recommended to close these applications before testing.`r`n`r`nContinue anyway?"
                $result = [System.Windows.Forms.MessageBox]::Show($msg, 'Competing Processes Detected', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $script:HasWarnedAboutCompetingProcesses = $true
                if ($result -eq [System.Windows.Forms.DialogResult]::No) {
                    return
                }
            }
        }
        
        $script:btnTestSpeed.Enabled = $false
        $script:btnConvert.Enabled = $false
        [System.Windows.Forms.Application]::DoEvents()
        
        Append-Log ("========================================")
        Append-Log ("SPEED TEST - Encoding 10 seconds per file")
        Append-Log ("========================================")
        
        $enc = Get-Av1Encoder
        Append-Log ("Detected encoder: {0}" -f $enc)
        
        # Show warning if using CPU encoder
        if ($enc -ne 'av1_nvenc') {
            Append-Log ("WARNING: Not using NVENC! This will be VERY SLOW!")
            Append-Log ("Your command line may be using a different FFmpeg with NVENC support.")
            Append-Log ("Current FFmpeg: {0}" -f $script:lblFfmpeg.Text)
        }
        Append-Log ("")
        
        $results = @()
        $totalEstimatedSeconds = 0
        $fileCount = 0
        
        foreach ($name in $Names) {
            $fileCount++
            $in = Join-Path -Path $script:ScriptDir -ChildPath $name
            if (-not (Test-Path $in)) { continue }
            
            Append-Log ("Testing {0}/{1}: {2}" -f $fileCount, $Names.Count, $name)
            
            # Get total frame count
            $totalFrames = Get-VideoFrameCount -FullPath $in
            if ($totalFrames -eq 0) {
                Append-Log ("  Could not determine frame count, skipping")
                continue
            }
            Append-Log ("  Total frames: {0:N0}" -f $totalFrames)
            
            # Create temp output
            $tempOut = [System.IO.Path]::GetTempFileName() + ".mkv"
            
            try {
                # Build ffmpeg command for 2 second test
                $ffmpeg = if ($script:FfmpegPath -and (Test-Path -LiteralPath $script:FfmpegPath)) { $script:FfmpegPath } else { 'ffmpeg' }
                $bitrate = $script:cmbBr.SelectedItem
                
                # Encode from start for accuracy, limit to 10 seconds
                $ffArgs = @('-y', '-i', ('"{0}"' -f $in), '-t', '10')
                
                if ($enc -eq 'av1_nvenc') {
                    # NVENC settings matching your working batch script
                    $ffArgs += @('-c:v','av1_nvenc','-b:v',$bitrate,'-maxrate',$bitrate,'-bufsize','20M','-preset','fast')
                } elseif ($enc -eq 'libaom-av1') {
                    $ffArgs += @('-c:v','libaom-av1','-b:v',$bitrate,'-cpu-used','8','-row-mt','1','-threads','0','-pix_fmt','yuv420p')
                } elseif ($enc -eq 'librav1e') {
                    $ffArgs += @('-c:v','librav1e','-b:v',$bitrate,'-speed','10','-pix_fmt','yuv420p')
                } else {
                    $ffArgs += @('-c:v','libaom-av1','-b:v',$bitrate,'-cpu-used','8','-row-mt','1','-threads','0','-pix_fmt','yuv420p')
                }
                $ffArgs += @('-an', ('"{0}"' -f $tempOut))
                
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = $ffmpeg
                $psi.Arguments = ($ffArgs -join ' ')
                $psi.RedirectStandardError = $true
                $psi.RedirectStandardOutput = $true
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $true
                
                $proc = New-Object System.Diagnostics.Process
                $proc.StartInfo = $psi
                
                $startTime = Get-Date
                [void]$proc.Start()
                
                $lastLine = ""
                while (-not $proc.HasExited) {
                    $line = $proc.StandardError.ReadLine()
                    if ($null -ne $line) { $lastLine = $line }
                    [System.Windows.Forms.Application]::DoEvents()
                }
                while (-not $proc.StandardError.EndOfStream) { 
                    $line = $proc.StandardError.ReadLine()
                    if ($null -ne $line) { $lastLine = $line }
                }
                $proc.WaitForExit()
                $endTime = Get-Date
                $elapsed = ($endTime - $startTime).TotalSeconds
                
                # Parse FPS from output
                $fps = 0
                if ($lastLine -match 'fps=\s*([0-9.]+)') {
                    $fps = [double]$matches[1]
                } elseif ($lastLine -match 'speed=\s*([0-9.]+)x') {
                    $speed = [double]$matches[1]
                    # Estimate FPS: speed multiplier * typical framerate (assume 24fps)
                    $fps = 24 * $speed
                }
                
                # Fallback: calculate from elapsed time (10 seconds of video at ~24fps = ~240 frames)
                if ($fps -eq 0 -and $elapsed -gt 0) {
                    $fps = 240 / $elapsed
                }
                
                if ($fps -gt 0) {
                    $estimatedSeconds = $totalFrames / $fps
                    
                    # Apply adaptive calibration if enabled
                    if ($script:UseAdaptive -and $script:CalibrationFactor -ne 1.0) {
                        $estimatedSeconds = $estimatedSeconds * $script:CalibrationFactor
                        Append-Log ("  Raw speed: {0:N1} fps" -f $fps)
                        Append-Log ("  Calibration factor: {0:N2}" -f $script:CalibrationFactor)
                    } else {
                        Append-Log ("  Encoding speed: {0:N1} fps" -f $fps)
                    }
                    
                    $estimatedMinutes = [Math]::Round($estimatedSeconds / 60, 1)
                    $totalEstimatedSeconds += $estimatedSeconds
                    
                    Append-Log ("  Estimated time: {0:N1} minutes" -f $estimatedMinutes)
                    
                    # Update the ListView item with estimated time
                    foreach ($item in $script:lst.Items) {
                        if ($item.Tag -eq $name) {
                            $timeDisplay = if ($estimatedMinutes -ge 60) {
                                "{0:N1}h" -f ($estimatedMinutes / 60)
                            } else {
                                "{0:N1}m" -f $estimatedMinutes
                            }
                            $item.SubItems[3].Text = $timeDisplay
                            break
                        }
                    }
                    
                    $results += [PSCustomObject]@{
                        File = $name
                        TotalFrames = $totalFrames
                        FPS = $fps
                        EstimatedSeconds = $estimatedSeconds
                        EstimatedMinutes = $estimatedMinutes
                    }
                } else {
                    Append-Log ("  Could not determine encoding speed")
                    # Update ListView to show error
                    foreach ($item in $script:lst.Items) {
                        if ($item.Tag -eq $name) {
                            $item.SubItems[3].Text = "Error"
                            break
                        }
                    }
                }
                
            } catch {
                Append-Log ("  Error during test: {0}" -f $_.Exception.Message)
            } finally {
                if (Test-Path $tempOut) {
                    try { Remove-Item -LiteralPath $tempOut -Force } catch {}
                }
            }
            
            Append-Log ("")
        }
        
        if ($results.Count -gt 0) {
            Append-Log ("========================================")
            Append-Log ("SUMMARY")
            Append-Log ("========================================")
            Append-Log ("Total files: {0}" -f $results.Count)
            Append-Log ("Total estimated time: {0:N1} minutes ({1:N1} hours)" -f ($totalEstimatedSeconds / 60), ($totalEstimatedSeconds / 3600))
            Append-Log ("")
            Append-Log ("Per-file breakdown:")
            foreach ($r in $results) {
                Append-Log ("  {0}: {1:N1} min" -f $r.File, $r.EstimatedMinutes)
            }
            Append-Log ("========================================")
            
            # Show results in message box
            $msgBox = "SPEED TEST RESULTS`r`n`r`n"
            $msgBox += "Encoder: {0}`r`n" -f $enc
            $msgBox += "Files tested: {0}`r`n`r`n" -f $results.Count
            $msgBox += "ESTIMATED CONVERSION TIMES:`r`n"
            $msgBox += ("-" * 50) + "`r`n"
            foreach ($r in $results) {
                $msgBox += "{0}`r`n  {1:N1} minutes ({2:N0} frames @ {3:N1} fps)`r`n`r`n" -f $r.File, $r.EstimatedMinutes, $r.TotalFrames, $r.FPS
            }
            $msgBox += ("-" * 50) + "`r`n"
            $msgBox += "TOTAL TIME: {0:N1} minutes ({1:N1} hours)" -f ($totalEstimatedSeconds / 60), ($totalEstimatedSeconds / 3600)
            
            [System.Windows.Forms.MessageBox]::Show($msgBox, 'Speed Test Complete', 'OK', 'Exclamation') | Out-Null
        } else {
            [System.Windows.Forms.MessageBox]::Show('Speed test completed but no results were obtained.`r`n`r`nCheck the log for details.', 'Speed Test', 'OK', 'Exclamation') | Out-Null
        }
        
        $script:btnTestSpeed.Enabled = $true
        $script:btnConvert.Enabled = $true
    }

    # ---------- Events ----------
    $btnBrowseFolder.Add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = 'Select folder containing video files'
        $fbd.SelectedPath = $script:cmbFolder.Text
        
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Change-WorkingFolder -NewPath $fbd.SelectedPath
        }
    })
    
    $script:cmbFolder.Add_TextChanged({
        $newPath = $script:cmbFolder.Text
        if ($newPath -and (Test-Path -Path $newPath) -and $newPath -ne $script:ScriptDir) {
            Change-WorkingFolder -NewPath $newPath
        }
    })
    
    $script:cmbFolder.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $newPath = $script:cmbFolder.Text
            if ($newPath -and (Test-Path -Path $newPath)) {
                Change-WorkingFolder -NewPath $newPath
            } else {
                [System.Windows.Forms.MessageBox]::Show('Folder does not exist.', 'Error', 'OK', 'Error') | Out-Null
            }
        }
    })
    
    $btnBrowseFfmpeg.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Title = 'Select FFmpeg.exe (with NVENC support)'
        $ofd.Filter = 'ffmpeg.exe|ffmpeg.exe|Executables|*.exe'
        if ($script:FfmpegPath) {
            $ofd.InitialDirectory = [System.IO.Path]::GetDirectoryName($script:FfmpegPath)
        } else {
            $ofd.InitialDirectory = $script:ScriptDir
        }
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:FfmpegPath = $ofd.FileName
            $script:Av1Encoder = $null  # Reset encoder detection
            Persist-FfmpegPath -Path $script:FfmpegPath
            Update-FfmpegLabel
            
            # Test for NVENC support
            $enc = Get-Av1Encoder
            if ($enc -eq 'av1_nvenc') {
                Append-Log ("SUCCESS: FFmpeg with NVENC support detected!")
                [System.Windows.Forms.MessageBox]::Show("NVENC support confirmed!`r`n`r`nThis FFmpeg will provide fast hardware-accelerated encoding.", 'Success', 'OK', 'Exclamation') | Out-Null
            } else {
                Append-Log ("WARNING: Selected FFmpeg does not have NVENC support. Encoder: {0}" -f $enc)
                [System.Windows.Forms.MessageBox]::Show("WARNING: This FFmpeg does not have NVENC support.`r`n`r`nEncoder detected: {0}`r`n`r`nPlease select a different FFmpeg build with NVENC." -f $enc, 'No NVENC Support', 'OK', 'Exclamation') | Out-Null
            }
        }
    })
    
    $btnRefresh.Add_Click({ Load-Files -All:$script:chkAll.Checked })
    $script:chkAll.Add_CheckedChanged({ Load-Files -All:$script:chkAll.Checked })
    $script:chkDark.Add_CheckedChanged({ Apply-Theme -Dark:$script:chkDark.Checked })
    
    # Initialize folder dropdown
    Update-FolderDropdown

    $script:lst.add_ItemSelectionChanged({
        $script:LastSelectedNames = @()
        foreach ($it in $script:lst.SelectedItems) { $script:LastSelectedNames += [string]$it.Tag }
    })
    
    $script:btnTestSpeed.Add_Click({
        $selectedNames = @()
        if ($script:lst.SelectedItems.Count -gt 0) {
            foreach ($it in $script:lst.SelectedItems) { $selectedNames += [string]$it.Tag }
        } elseif ($script:LastSelectedNames -and $script:LastSelectedNames.Count -gt 0) {
            $selectedNames = @($script:LastSelectedNames)
        }
        if ($selectedNames.Count -eq 0) { 
            [System.Windows.Forms.MessageBox]::Show('Select at least one file to test.', 'Test Speed', 'OK', 'Exclamation') | Out-Null
            return 
        }
        Test-EncodingSpeed -Names $selectedNames
    })
    
    $script:btnAdaptiveTest.Add_Click({
        $selectedNames = @()
        if ($script:lst.SelectedItems.Count -gt 0) {
            foreach ($it in $script:lst.SelectedItems) { $selectedNames += [string]$it.Tag }
        } elseif ($script:LastSelectedNames -and $script:LastSelectedNames.Count -gt 0) {
            $selectedNames = @($script:LastSelectedNames)
        }
        if ($selectedNames.Count -eq 0) { 
            [System.Windows.Forms.MessageBox]::Show('Select at least one file to test.', 'Adaptive Test', 'OK', 'Exclamation') | Out-Null
            return 
        }
        
        # Run test with adaptive calibration
        $script:UseAdaptive = $true
        Test-EncodingSpeed -Names $selectedNames
        $script:UseAdaptive = $false
    })

    $script:btnConvert.Add_Click({
        $script:CancelRequested = $false
        $script:btnConvert.Enabled = $false
        $script:btnCancel.Enabled  = $true
        [System.Windows.Forms.Application]::DoEvents()
        $null = $script:lst.Focus()

        $selectedNames = @()
        if ($script:lst.SelectedItems.Count -gt 0) {
            foreach ($it in $script:lst.SelectedItems) { $selectedNames += [string]$it.Tag }
        } elseif ($script:LastSelectedNames -and $script:LastSelectedNames.Count -gt 0) {
            $selectedNames = @($script:LastSelectedNames)
        }
        if ($selectedNames.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Select at least one file.', 'FFmpeg', 'OK', 'Exclamation') | Out-Null; $script:btnConvert.Enabled=$true; $script:btnCancel.Enabled=$false; return }

        $script:PendingNames = @($selectedNames)
        Process-Queue-AV1 -Names $script:PendingNames
    })

    $script:btnCancel.Add_Click({
        $script:CancelRequested = $true
        if ($script:CurrentProc -and -not $script:CurrentProc.HasExited) { try { $script:CurrentProc.Kill() } catch {} }
        if ($script:CurrentOutPath -and (Test-Path -LiteralPath $script:CurrentOutPath)) {
            try { Remove-Item -LiteralPath $script:CurrentOutPath -Force; Append-Log ("Deleted partial output: {0}" -f $script:CurrentOutPath) } catch { Append-Log ("Failed to delete partial output: {0}" -f $_.Exception.Message) }
        }
        Append-Log 'Cancel requested by user.'
    })

    # ---------- Initial Load ----------
    function script:Load-Files {
        param([bool]$All = $false)
        $script:lst.Items.Clear()
        $patterns = if ($All) { @('*.mkv','*.mp4','*.avi','*.mov','*.m4v','*.webm','*.ts','*.m2ts','*.wmv','*.flv','*.mpg','*.mpeg','*.vob') } else { @('*.mkv') }
        $files = @()
        foreach ($pat in $patterns) { $files += Get-ChildItem -LiteralPath $script:ScriptDir -Filter $pat -File -ErrorAction SilentlyContinue }
        $files = $files | Sort-Object Name -Unique
        foreach ($f in $files) {
            $sizeMB = [Math]::Round(($f.Length/1MB),2)
            $codec = Get-VideoCodec -FullPath $f.FullName
            $item = New-Object System.Windows.Forms.ListViewItem($f.Name)
            [void]$item.SubItems.Add(("{0} MB" -f $sizeMB))
            [void]$item.SubItems.Add($codec)
            [void]$item.SubItems.Add("")  # Est. Time column (empty initially)
            $item.Tag = $f.Name
            [void]$script:lst.Items.Add($item)
        }
        if ($script:lst.Items.Count -eq 0) { 
            $msg = if ($All) { 'No video files found.' } else { 'No MKV files found.' }
            Append-Log $msg
        }
    }

    # Init flags
    $script:IsRunning = $false
    $script:PendingNames = @()

    # Initial label/theme/load
    Update-FfmpegLabel
    Load-Calibration
    Load-Files -All:$script:chkAll.Checked
    Apply-Theme -Dark:$script:chkDark.Checked
    return $form
}

# Run GUI
$form = New-Form
[void]$form.ShowDialog()

