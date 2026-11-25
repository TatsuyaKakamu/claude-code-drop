# ClaudeDrop: drop an Outlook mail, or files / folders from Explorer, on this
# small window to open a new Claude Code session in Claude Desktop
# (claude://code/new deep link).
# ASCII only. Japanese strings live in config.json (UTF-8).
# Run with: powershell.exe -STA -File ClaudeDrop.ps1  (see start.vbs)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$cfgText = [IO.File]::ReadAllText((Join-Path $here 'config.json'), [Text.Encoding]::UTF8)
$cfg = $cfgText | ConvertFrom-Json
$saveDir = [Environment]::ExpandEnvironmentVariables($cfg.saveDir)
$workDir = [Environment]::ExpandEnvironmentVariables($cfg.folder)
$logPath = Join-Path $here 'ClaudeDrop.log'

function Write-Log([string]$msg) {
    $line = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + ' ' + $msg + "`r`n"
    [IO.File]::AppendAllText($logPath, $line, [Text.Encoding]::UTF8)
}

function Get-SafeName([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return 'mail' }
    $t = $s -replace '[\\/:*?"<>|\r\n\t]', '_'
    $t = $t.Trim()
    if ($t.Length -gt 40) { $t = $t.Substring(0, 40) }
    if ($t.Length -eq 0) { return 'mail' }
    return $t
}

# Save the mails currently selected in Outlook (the ones being dragged) as .msg
function Save-SelectedMails {
    if (-not (Test-Path $saveDir)) { New-Item -ItemType Directory -Path $saveDir | Out-Null }
    $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    $sel = $ol.ActiveExplorer().Selection
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $paths = New-Object System.Collections.ArrayList
    for ($i = 1; $i -le $sel.Count; $i++) {
        $item = $sel.Item($i)
        $name = '{0}_{1}_{2}.msg' -f $stamp, $i, (Get-SafeName $item.Subject)
        $p = Join-Path $saveDir $name
        $item.SaveAs($p, 9)   # 9 = olMSGUnicode (3 = ANSI olMSG garbles Japanese and breaks parsers)
        [void]$paths.Add($p)
    }
    return ,$paths.ToArray()
}

function Open-ClaudeSession([string[]]$paths) {
    $lines = ($paths | ForEach-Object { '- ' + $_ }) -join "`n"
    $q = $cfg.prompt.Replace('{files}', $lines)
    # A folder passed by link always triggers the trust dialog (by design),
    # so it is only added when config.json sets "folder".
    $url = 'claude://code/new?q=' + [Uri]::EscapeDataString($q)
    if (-not [string]::IsNullOrEmpty($workDir)) {
        $url += '&folder=' + [Uri]::EscapeDataString($workDir)
    }
    Start-Process $url
}

# ---- window ----
$fullSize = 132   # normal mode: square window
$miniSize = 48    # mini mode: a small square that still takes drops

# window.json (not in git): position, mini mode and transparency (0-50 %).
# Older files have only x and y; the other keys then fall back to defaults.
$posPath = Join-Path $here 'window.json'
$pos = $null
$script:mini = $false
$script:transparency = 0
try {
    if (Test-Path $posPath) { $pos = [IO.File]::ReadAllText($posPath) | ConvertFrom-Json }
} catch { $pos = $null }
if ($null -ne $pos) {
    if ($pos.mini -eq $true) { $script:mini = $true }
    try {
        if ($null -ne $pos.transparency) {
            $script:transparency = [Math]::Max(0, [Math]::Min(50, [int]$pos.transparency))
        }
    } catch { $script:transparency = 0 }
}

$form = New-Object Windows.Forms.Form
$form.Text = $cfg.title
$form.FormBorderStyle = 'None'
$form.TopMost = $true
$form.ShowInTaskbar = $false
$form.BackColor = [Drawing.Color]::FromArgb(251, 251, 251)
$form.MinimumSize = New-Object Drawing.Size(1, 1)
if ($script:mini) { $sz = $miniSize } else { $sz = $fullSize }
$form.ClientSize = New-Object Drawing.Size($sz, $sz)
$form.StartPosition = 'Manual'
$wa = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$form.Location = New-Object Drawing.Point(($wa.Right - $form.Width - 16), ($wa.Bottom - $form.Height - 24))
$form.AllowDrop = $true
$form.Opacity = (100 - $script:transparency) / 100.0

try {
    if ($null -ne $pos -and $null -ne $pos.x -and $null -ne $pos.y) {
        $pt = New-Object Drawing.Point([int]$pos.x, [int]$pos.y)
        $rect = New-Object Drawing.Rectangle($pt, $form.Size)
        $onScreen = $false
        foreach ($scr in [Windows.Forms.Screen]::AllScreens) {
            if ($scr.WorkingArea.IntersectsWith($rect)) { $onScreen = $true }
        }
        if ($onScreen) {
            # Pull it fully inside the screen, so a mini window is never left
            # with only a few pixels showing.
            $area = [Windows.Forms.Screen]::FromRectangle($rect).WorkingArea
            $x = [Math]::Max($area.Left, [Math]::Min($pt.X, $area.Right - $form.Width))
            $y = [Math]::Max($area.Top, [Math]::Min($pt.Y, $area.Bottom - $form.Height))
            $form.Location = New-Object Drawing.Point($x, $y)
        }
    }
} catch { }

$savePos = {
    try {
        $o = [ordered]@{
            x = $form.Location.X
            y = $form.Location.Y
            mini = [bool]$script:mini
            transparency = [int]$script:transparency
        }
        [IO.File]::WriteAllText($posPath, ($o | ConvertTo-Json -Compress))
    } catch { }
}
$form.Add_ResizeEnd($savePos)      # fires after the window is moved
$form.Add_FormClosing($savePos)

Add-Type -Namespace ClaudeDrop -Name Win32 -MemberDefinition @'
[DllImport("user32.dll")]
public static extern bool SetWindowPos(IntPtr hWnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
[DllImport("user32.dll")]
public static extern bool ReleaseCapture();
[DllImport("user32.dll")]
public static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr w, IntPtr l);
[DllImport("dwmapi.dll")]
public static extern int DwmSetWindowAttribute(IntPtr hWnd, int attr, ref int value, int size);
'@

# ---- look ----
# The whole window is one PictureBox (double-buffered) painted by hand:
# a header with the title and three buttons (settings, mini, close), and a
# dashed drop zone. In mini mode only a small drop zone with an icon is drawn.
function New-Rgb([int]$r, [int]$g, [int]$b) { [Drawing.Color]::FromArgb($r, $g, $b) }
$styles = @{
    idle  = @{ fill = (New-Rgb 244 245 247); line = (New-Rgb 190 194 201); text = (New-Rgb 70 74 82) }
    hover = @{ fill = (New-Rgb 252 238 231); line = (New-Rgb 217 119 87);  text = (New-Rgb 176 80 50) }
    busy  = @{ fill = (New-Rgb 244 245 247); line = (New-Rgb 190 194 201); text = (New-Rgb 70 74 82) }
    done  = @{ fill = (New-Rgb 231 245 236); line = (New-Rgb 52 168 83);   text = (New-Rgb 24 110 50) }
    error = @{ fill = (New-Rgb 252 233 231); line = (New-Rgb 217 48 37);   text = (New-Rgb 170 30 25) }
}
$fontTitle = New-Object Drawing.Font('Yu Gothic UI', 9)
$fontBody  = New-Object Drawing.Font('Yu Gothic UI', 9.5, [Drawing.FontStyle]::Bold)
$script:state = 'idle'
$script:text = $cfg.label
$script:hot = $null   # header button under the mouse: 'close', 'min', 'gear' or $null

function Set-State([string]$st, [string]$txt) {
    $script:state = $st
    $script:text = $txt
    $canvas.Invalidate()
}

# Header buttons, right to left: close, mini, settings.
$buttonNames = @('close', 'min', 'gear')
function Get-ButtonRect([string]$name) {
    $i = [array]::IndexOf($buttonNames, $name)
    New-Object Drawing.Rectangle(($canvas.Width - 27 - 21 * $i), 5, 21, 20)
}
function Get-HotButton([Drawing.Point]$pt) {
    if ($script:mini) { return $null }
    foreach ($n in $buttonNames) {
        if ((Get-ButtonRect $n).Contains($pt)) { return $n }
    }
    return $null
}

function New-RoundPath([Drawing.RectangleF]$r, [float]$rad) {
    $p = New-Object Drawing.Drawing2D.GraphicsPath
    $d = $rad * 2
    $p.AddArc($r.X, $r.Y, $d, $d, 180, 90)
    $p.AddArc($r.Right - $d, $r.Y, $d, $d, 270, 90)
    $p.AddArc($r.Right - $d, $r.Bottom - $d, $d, $d, 0, 90)
    $p.AddArc($r.X, $r.Bottom - $d, $d, $d, 90, 90)
    $p.CloseFigure()
    return $p
}

function Draw-Zone($g, [Drawing.RectangleF]$zone, $st) {
    $path = New-RoundPath $zone 8
    $g.FillPath((New-Object Drawing.SolidBrush($st.fill)), $path)
    $pen = New-Object Drawing.Pen($st.line, 1.5)
    if ($script:state -ne 'done' -and $script:state -ne 'error') { $pen.DashPattern = [single[]](4, 3) }
    $g.DrawPath($pen, $path)
}

$canvas = New-Object Windows.Forms.PictureBox
$canvas.Dock = 'Fill'
$canvas.AllowDrop = $true
$form.Controls.Add($canvas)

$canvas.Add_Paint({
    param($s, $e)
    $g = $e.Graphics
    $g.SmoothingMode = 'AntiAlias'
    $st = $styles[$script:state]

    if ($script:mini) {
        # mini: drop zone with an icon (arrow, check when done, X on error)
        Draw-Zone $g (New-Object Drawing.RectangleF(4, 4, ($canvas.Width - 9), ($canvas.Height - 9))) $st
        $ip = New-Object Drawing.Pen($st.text, 2)
        $ip.StartCap = 'Round'; $ip.EndCap = 'Round'
        $cx = $canvas.Width / 2; $cy = $canvas.Height / 2
        if ($script:state -eq 'done') {
            $g.DrawLines($ip, [Drawing.PointF[]]@(
                (New-Object Drawing.PointF(($cx - 7), ($cy + 0))),
                (New-Object Drawing.PointF(($cx - 2), ($cy + 5))),
                (New-Object Drawing.PointF(($cx + 7), ($cy - 5)))))
        } elseif ($script:state -eq 'error') {
            $g.DrawLine($ip, ($cx - 6), ($cy - 6), ($cx + 6), ($cy + 6))
            $g.DrawLine($ip, ($cx + 6), ($cy - 6), ($cx - 6), ($cy + 6))
        } elseif ($script:state -eq 'busy') {
            $dot = New-Object Drawing.SolidBrush($st.text)
            foreach ($dx in -7, 0, 7) { $g.FillEllipse($dot, [float]($cx + $dx - 2), [float]($cy - 2), [float]4, [float]4) }
        } else {
            $g.DrawLine($ip, $cx, ($cy - 8), $cx, ($cy + 3))
            $g.DrawLine($ip, ($cx - 4), ($cy - 1), $cx, ($cy + 3))
            $g.DrawLine($ip, ($cx + 4), ($cy - 1), $cx, ($cy + 3))
            $g.DrawLine($ip, ($cx - 7), ($cy + 8), ($cx + 7), ($cy + 8))
        }
        return
    }

    # header: title + settings / mini / close buttons
    $titleRect = New-Object Drawing.Rectangle(10, 4, ($canvas.Width - 10 - 27 - 21 * 2 - 2), 22)
    [Windows.Forms.TextRenderer]::DrawText($g, $cfg.title, $fontTitle, $titleRect, (New-Rgb 95 99 104), [Windows.Forms.TextFormatFlags]'Left, VerticalCenter, EndEllipsis')
    foreach ($n in $buttonNames) {
        $br = Get-ButtonRect $n
        $fg = New-Rgb 120 124 130
        if ($script:hot -eq $n) {
            $hotPath = New-RoundPath (New-Object Drawing.RectangleF($br.X, $br.Y, $br.Width, $br.Height)) 4
            if ($n -eq 'close') {
                $g.FillPath((New-Object Drawing.SolidBrush(New-Rgb 232 17 35)), $hotPath)
                $fg = [Drawing.Color]::White
            } else {
                $g.FillPath((New-Object Drawing.SolidBrush(New-Rgb 229 231 235)), $hotPath)
                $fg = New-Rgb 60 64 70
            }
        }
        $bp = New-Object Drawing.Pen($fg, 1.4)
        $cx = $br.X + $br.Width / 2; $cy = $br.Y + $br.Height / 2
        switch ($n) {
            'close' {
                $g.DrawLine($bp, ($cx - 4), ($cy - 4), ($cx + 4), ($cy + 4))
                $g.DrawLine($bp, ($cx + 4), ($cy - 4), ($cx - 4), ($cy + 4))
            }
            'min' {
                $g.DrawLine($bp, ($cx - 4), $cy, ($cx + 4), $cy)
            }
            'gear' {
                # gear: 8 teeth around a ring with a hole
                $tp = New-Object Drawing.Pen($fg, 2)
                for ($k = 0; $k -lt 8; $k++) {
                    $a = $k * [Math]::PI / 4
                    $g.DrawLine($tp, [float]($cx + 3.5 * [Math]::Cos($a)), [float]($cy + 3.5 * [Math]::Sin($a)),
                                     [float]($cx + 6 * [Math]::Cos($a)),   [float]($cy + 6 * [Math]::Sin($a)))
                }
                $g.DrawEllipse($bp, [float]($cx - 4), [float]($cy - 4), [float]8, [float]8)
                $g.DrawEllipse($bp, [float]($cx - 1.5), [float]($cy - 1.5), [float]3, [float]3)
            }
        }
    }

    # drop zone
    $zone = New-Object Drawing.RectangleF(10, 30, ($canvas.Width - 21), ($canvas.Height - 41))
    Draw-Zone $g $zone $st
    $zr = [Drawing.Rectangle]::Round($zone)
    [Windows.Forms.TextRenderer]::DrawText($g, $script:text, $fontBody, $zr, $st.text, [Windows.Forms.TextFormatFlags]'HorizontalCenter, VerticalCenter, WordBreak')
})

$tip = New-Object Windows.Forms.ToolTip
$tip.ShowAlways = $true   # the window is usually inactive
$script:restoredAt = [DateTime]::MinValue

# Switch between the normal window and the mini window. The corner nearest
# to the screen edge stays put, and the result is kept on the same screen.
function Set-Mini([bool]$on) {
    $old = $form.Bounds
    $script:mini = $on
    $script:hot = $null
    if ($on) { $size = $miniSize } else { $size = $fullSize }
    $area = [Windows.Forms.Screen]::FromRectangle($old).WorkingArea
    if (($old.X + $old.Width / 2) -gt ($area.X + $area.Width / 2)) { $x = $old.Right - $size } else { $x = $old.X }
    if (($old.Y + $old.Height / 2) -gt ($area.Y + $area.Height / 2)) { $y = $old.Bottom - $size } else { $y = $old.Y }
    $x = [Math]::Max($area.Left, [Math]::Min($x, $area.Right - $size))
    $y = [Math]::Max($area.Top, [Math]::Min($y, $area.Bottom - $size))
    $form.Bounds = New-Object Drawing.Rectangle($x, $y, $size, $size)
    if (-not $on) { $script:restoredAt = [DateTime]::Now }
    if ($on) { $tip.SetToolTip($canvas, $cfg.miniTip) } else { $tip.SetToolTip($canvas, '') }
    $canvas.Invalidate()
    & $savePos
    Write-Log ('mini=' + $on + ' size=' + $form.Size.Width + 'x' + $form.Size.Height)
}

# Applied at once while the slider moves; saved when the settings window closes.
function Set-Transparency([int]$v) {
    $script:transparency = $v
    $form.Opacity = (100 - $v) / 100.0
}

function Get-TransparencyText([int]$v) {
    $t = '{0}: {1} %' -f $cfg.transparencyLabel, $v
    if ($v -eq 0) { $t += $cfg.transparencyOpaque }
    return $t
}

# Settings window: non-modal (a modal one would disable the main window
# and block drops), owned by the main window, at most one at a time.
$script:settings = $null
function Show-Settings {
    if ($null -ne $script:settings -and -not $script:settings.IsDisposed) {
        $script:settings.Activate()
        return
    }
    $dlg = New-Object Windows.Forms.Form
    $dlg.Text = $cfg.settingsTitle
    $dlg.FormBorderStyle = 'FixedToolWindow'
    $dlg.ShowInTaskbar = $false
    $dlg.TopMost = $true
    $dlg.StartPosition = 'Manual'
    $dlg.Font = $fontTitle
    $dlg.BackColor = [Drawing.Color]::White

    $lbl = New-Object Windows.Forms.Label
    $lbl.Text = Get-TransparencyText $script:transparency
    $lbl.AutoSize = $true
    $lbl.Location = New-Object Drawing.Point(14, 12)
    $dlg.Controls.Add($lbl)
    $script:transparencyText = $lbl

    # Slider 0-50 % (1 % steps; PageUp/PageDown or a click on the track: 10 %).
    $bar = New-Object Windows.Forms.TrackBar
    $bar.Minimum = 0
    $bar.Maximum = 50
    $bar.SmallChange = 1
    $bar.LargeChange = 10
    $bar.TickFrequency = 10
    $bar.Value = $script:transparency
    $bar.Location = New-Object Drawing.Point(10, 36)
    $bar.Width = 200
    $bar.Add_ValueChanged({
        param($s, $e)
        Set-Transparency $s.Value
        $script:transparencyText.Text = Get-TransparencyText $s.Value
    })
    $dlg.Controls.Add($bar)

    $btn = New-Object Windows.Forms.Button
    $btn.Text = $cfg.closeLabel
    $btn.Size = New-Object Drawing.Size(80, 26)
    $btn.Location = New-Object Drawing.Point(130, 88)
    $btn.Add_Click({ $script:settings.Close() })
    $dlg.Controls.Add($btn)
    $dlg.CancelButton = $btn
    $dlg.ActiveControl = $bar   # arrow keys move the slider right away
    $dlg.ClientSize = New-Object Drawing.Size(220, 124)
    $dlg.Add_FormClosed({ & $savePos })

    # Place it beside the main window, inside the same screen.
    $area = [Windows.Forms.Screen]::FromRectangle($form.Bounds).WorkingArea
    $x = $form.Left - $dlg.Width - 8
    if ($x -lt $area.Left) { $x = $form.Right + 8 }
    $x = [Math]::Max($area.Left, [Math]::Min($x, $area.Right - $dlg.Width))
    $y = [Math]::Max($area.Top, [Math]::Min($form.Top, $area.Bottom - $dlg.Height))
    $dlg.Location = New-Object Drawing.Point($x, $y)

    $script:settings = $dlg
    $dlg.Show($form)
}

# No title bar: drag anywhere to move; header buttons on the right.
# In mini mode a click (press and release without moving) restores the window.
$canvas.Add_MouseMove({
    param($s, $e)
    $hot = Get-HotButton $e.Location
    if ($hot -ne $script:hot) { $script:hot = $hot; $canvas.Invalidate() }
})
$canvas.Add_MouseLeave({ if ($null -ne $script:hot) { $script:hot = $null; $canvas.Invalidate() } })
$canvas.Add_MouseDown({
    param($s, $e)
    if ($e.Button -ne 'Left') { return }
    # Right after leaving mini mode, a double click's second press could land
    # on the close button of the restored window, so ignore it.
    if (([DateTime]::Now - $script:restoredAt).TotalMilliseconds -lt [Windows.Forms.SystemInformation]::DoubleClickTime) { return }
    switch (Get-HotButton $e.Location) {
        'close' { $form.Close(); return }
        'min'   { Set-Mini $true; return }
        'gear'  { Show-Settings; return }
    }
    $before = $form.Location
    [void][ClaudeDrop.Win32]::ReleaseCapture()
    # WM_NCLBUTTONDOWN + HTCAPTION: let Windows move the window; returns when done
    [void][ClaudeDrop.Win32]::SendMessage($form.Handle, 0xA1, [IntPtr]2, [IntPtr]::Zero)
    if ($script:mini -and $before.Equals($form.Location)) { Set-Mini $false; return }
    & $savePos
})

# Windows 11: rounded corners (attr 33 = 2 round) and a thin border (attr 34).
$form.Add_HandleCreated({
    try {
        $round = 2
        [void][ClaudeDrop.Win32]::DwmSetWindowAttribute($form.Handle, 33, [ref]$round, 4)
        $border = 0x00D8D4D0   # COLORREF 0x00BBGGRR
        [void][ClaudeDrop.Win32]::DwmSetWindowAttribute($form.Handle, 34, [ref]$border, 4)
    } catch { }
})

$script:pending = $null   # 'outlook' or array of file paths

$resetTimer = New-Object Windows.Forms.Timer
$resetTimer.Interval = 2500
$resetTimer.Add_Tick({ $resetTimer.Stop(); Set-State 'idle' $cfg.label })

# Work runs on a timer after the drop returns, so Outlook is not blocked
# in its drag loop while we call it over COM.
$workTimer = New-Object Windows.Forms.Timer
$workTimer.Interval = 150
$workTimer.Add_Tick({
    $workTimer.Stop()
    $job = $script:pending
    $script:pending = $null
    try {
        if ($job -eq 'outlook') { $paths = Save-SelectedMails } else { $paths = $job }
        if (-not $paths -or $paths.Count -eq 0) { throw 'nothing to open' }
        Open-ClaudeSession $paths
        Write-Log ('opened: ' + ($paths -join ' | '))
        Set-State 'done' $cfg.labelDone
    } catch {
        Write-Log ('error: ' + $_.Exception.Message)
        Set-State 'error' $cfg.labelError
    }
    $resetTimer.Start()
})

$onDragEnter = {
    param($s, $e)
    $d = $e.Data
    if ($d.GetDataPresent('FileGroupDescriptorW') -or $d.GetDataPresent('RenPrivateMessages') -or $d.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [Windows.Forms.DragDropEffects]::Copy
        if ($script:state -eq 'idle') { Set-State 'hover' $cfg.label }
    } else {
        $e.Effect = [Windows.Forms.DragDropEffects]::None
    }
}

$onDragDrop = {
    param($s, $e)
    $d = $e.Data
    if ($d.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        # Files and folders from Explorer (any type) are passed as they are.
        $files = @($d.GetData([Windows.Forms.DataFormats]::FileDrop))
        if ($files.Count -eq 0) { Set-State 'idle' $cfg.label; return }
        $script:pending = [string[]]$files
    } else {
        $script:pending = 'outlook'
    }
    Set-State 'busy' $cfg.labelBusy
    $workTimer.Start()
}

$onDragLeave = { if ($script:state -eq 'hover') { Set-State 'idle' $cfg.label } }

# Other topmost windows can push this one back, so re-assert topmost
# every 2 s without taking focus (SetWindowPos HWND_TOPMOST, SWP_NOACTIVATE).
$topTimer = New-Object Windows.Forms.Timer
$topTimer.Interval = 2000
$topTimer.Add_Tick({
    # 0x0013 = SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE
    [void][ClaudeDrop.Win32]::SetWindowPos($form.Handle, [IntPtr](-1), 0, 0, 0, 0, 0x0013)
})
$form.Add_Shown({
    if ($script:mini) { $tip.SetToolTip($canvas, $cfg.miniTip) }
    $topTimer.Start()
})

foreach ($c in @($form, $canvas)) {
    $c.Add_DragEnter($onDragEnter)
    $c.Add_DragDrop($onDragDrop)
    $c.Add_DragLeave($onDragLeave)
}

Write-Log 'started'
[Windows.Forms.Application]::Run($form)
