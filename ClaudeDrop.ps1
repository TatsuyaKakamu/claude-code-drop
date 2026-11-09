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
        $item.SaveAs($p, 3)   # 3 = olMSG
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
$form = New-Object Windows.Forms.Form
$form.Text = $cfg.title
$form.FormBorderStyle = 'None'
$form.TopMost = $true
$form.ShowInTaskbar = $false
$form.BackColor = [Drawing.Color]::FromArgb(251, 251, 251)
$form.ClientSize = New-Object Drawing.Size(184, 96)
$form.StartPosition = 'Manual'
$wa = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$form.Location = New-Object Drawing.Point(($wa.Right - 200), ($wa.Bottom - 120))
$form.AllowDrop = $true

# Remember the window position in window.json (not in git).
$posPath = Join-Path $here 'window.json'
try {
    if (Test-Path $posPath) {
        $pos = [IO.File]::ReadAllText($posPath) | ConvertFrom-Json
        $pt = New-Object Drawing.Point([int]$pos.x, [int]$pos.y)
        $rect = New-Object Drawing.Rectangle($pt, $form.Size)
        $onScreen = $false
        foreach ($scr in [Windows.Forms.Screen]::AllScreens) {
            if ($scr.WorkingArea.IntersectsWith($rect)) { $onScreen = $true }
        }
        if ($onScreen) { $form.Location = $pt }
    }
} catch { }

$savePos = {
    try {
        $json = '{"x":' + $form.Location.X + ',"y":' + $form.Location.Y + '}'
        [IO.File]::WriteAllText($posPath, $json)
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
# a header with the title and a close button, and a dashed drop zone.
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
$script:closeHot = $false

function Set-State([string]$st, [string]$txt) {
    $script:state = $st
    $script:text = $txt
    $canvas.Invalidate()
}

function Get-CloseRect { New-Object Drawing.Rectangle(($canvas.Width - 30), 4, 24, 22) }

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

$canvas = New-Object Windows.Forms.PictureBox
$canvas.Dock = 'Fill'
$canvas.AllowDrop = $true
$form.Controls.Add($canvas)

$canvas.Add_Paint({
    param($s, $e)
    $g = $e.Graphics
    $g.SmoothingMode = 'AntiAlias'
    $st = $styles[$script:state]

    # header: title + close button
    $titleRect = New-Object Drawing.Rectangle(12, 4, ($canvas.Width - 44), 22)
    [Windows.Forms.TextRenderer]::DrawText($g, $cfg.title, $fontTitle, $titleRect, (New-Rgb 95 99 104), [Windows.Forms.TextFormatFlags]'Left, VerticalCenter, EndEllipsis')
    $cr = Get-CloseRect
    $xColor = New-Rgb 120 124 130
    if ($script:closeHot) {
        $hotPath = New-RoundPath (New-Object Drawing.RectangleF($cr.X, $cr.Y, $cr.Width, $cr.Height)) 4
        $g.FillPath((New-Object Drawing.SolidBrush(New-Rgb 232 17 35)), $hotPath)
        $xColor = [Drawing.Color]::White
    }
    $xp = New-Object Drawing.Pen($xColor, 1.4)
    $cx = $cr.X + $cr.Width / 2; $cy = $cr.Y + $cr.Height / 2
    $g.DrawLine($xp, ($cx - 4), ($cy - 4), ($cx + 4), ($cy + 4))
    $g.DrawLine($xp, ($cx + 4), ($cy - 4), ($cx - 4), ($cy + 4))

    # drop zone
    $zone = New-Object Drawing.RectangleF(10, 30, ($canvas.Width - 21), ($canvas.Height - 41))
    $path = New-RoundPath $zone 8
    $g.FillPath((New-Object Drawing.SolidBrush($st.fill)), $path)
    $pen = New-Object Drawing.Pen($st.line, 1.5)
    if ($script:state -ne 'done' -and $script:state -ne 'error') { $pen.DashPattern = [single[]](4, 3) }
    $g.DrawPath($pen, $path)
    $zr = [Drawing.Rectangle]::Round($zone)
    [Windows.Forms.TextRenderer]::DrawText($g, $script:text, $fontBody, $zr, $st.text, [Windows.Forms.TextFormatFlags]'HorizontalCenter, VerticalCenter, WordBreak')
})

# No title bar: drag anywhere to move, click the X to close.
$canvas.Add_MouseMove({
    param($s, $e)
    $hot = (Get-CloseRect).Contains($e.Location)
    if ($hot -ne $script:closeHot) { $script:closeHot = $hot; $canvas.Invalidate() }
})
$canvas.Add_MouseLeave({ if ($script:closeHot) { $script:closeHot = $false; $canvas.Invalidate() } })
$canvas.Add_MouseDown({
    param($s, $e)
    if ($e.Button -ne 'Left') { return }
    if ((Get-CloseRect).Contains($e.Location)) { $form.Close(); return }
    [void][ClaudeDrop.Win32]::ReleaseCapture()
    # WM_NCLBUTTONDOWN + HTCAPTION: let Windows move the window; returns when done
    [void][ClaudeDrop.Win32]::SendMessage($form.Handle, 0xA1, [IntPtr]2, [IntPtr]::Zero)
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
$form.Add_Shown({ $topTimer.Start() })

foreach ($c in @($form, $canvas)) {
    $c.Add_DragEnter($onDragEnter)
    $c.Add_DragDrop($onDragDrop)
    $c.Add_DragLeave($onDragLeave)
}

Write-Log 'started'
[Windows.Forms.Application]::Run($form)
