Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================
#   LOGO - Descarga desde GitHub
# ============================================================
$logoUrl  = "https://raw.githubusercontent.com/syscodi7/Tools/main/sis.png"
$logoPath = "$env:TEMP\syscodi_logo.png"
try { Invoke-WebRequest -Uri $logoUrl -OutFile $logoPath -ErrorAction Stop }
catch { $logoPath = "" }

# ============================================================
#   COLORES CORPORATIVOS
# ============================================================
$cBg      = [Drawing.Color]::FromArgb(15, 25, 50)
$cPanel   = [Drawing.Color]::FromArgb(22, 38, 75)
$cCard    = [Drawing.Color]::FromArgb(30, 50, 100)
$cAccent  = [Drawing.Color]::FromArgb(0, 120, 215)
$cAccent2 = [Drawing.Color]::FromArgb(0, 180, 255)
$cText    = [Drawing.Color]::White
$cSubText = [Drawing.Color]::FromArgb(160, 200, 255)
$cBtn     = [Drawing.Color]::FromArgb(0, 100, 180)
$cBtnHov  = [Drawing.Color]::FromArgb(0, 140, 220)
$cOutput  = [Drawing.Color]::FromArgb(10, 18, 40)
$cBorder  = [Drawing.Color]::FromArgb(0, 120, 215)
$cGreen   = [Drawing.Color]::FromArgb(0, 180, 80)
$cRed     = [Drawing.Color]::FromArgb(220, 60, 60)
$cYellow  = [Drawing.Color]::FromArgb(255, 200, 0)
$cOrange  = [Drawing.Color]::FromArgb(255, 140, 0)

# ============================================================
#   FUNCIONES HELPER  (todas aquí, antes del formulario)
# ============================================================
function New-Tab($titulo) {
    $t = New-Object Windows.Forms.TabPage
    $t.Text = "  $titulo  "
    $t.BackColor = $cBg
    $t.ForeColor = $cText
    $tabs.TabPages.Add($t)
    return $t
}

function New-CorporateButton($texto, $x, $y, $w = 200, $h = 36) {
    $b = New-Object Windows.Forms.Button
    $b.Text = $texto
    $b.Location = New-Object Drawing.Point($x, $y)
    $b.Size = New-Object Drawing.Size($w, $h)
    $b.BackColor = $cBtn
    $b.ForeColor = $cText
    $b.FlatStyle = "Flat"
    $b.FlatAppearance.BorderColor = $cAccent
    $b.FlatAppearance.BorderSize = 1
    $b.Font = New-Object Drawing.Font("Segoe UI", 9)
    $b.Cursor = "Hand"
    return $b
}

function New-SectionLabel($texto, $x, $y, $parent) {
    $lbl = New-Object Windows.Forms.Label
    $lbl.Text = $texto
    $lbl.Location = New-Object Drawing.Point($x, $y)
    $lbl.Size = New-Object Drawing.Size(860, 22)
    $lbl.ForeColor = $cAccent2
    $lbl.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
    $parent.Controls.Add($lbl)
}

function New-UtilPanel($titulo, $subtitulo, $parent, $y, $h = 120) {
    $pnl = New-Object Windows.Forms.Panel
    $pnl.Location = New-Object Drawing.Point(8, $y)
    $pnl.Size = New-Object Drawing.Size(690, $h)
    $pnl.BackColor = [Drawing.Color]::FromArgb(22, 38, 75)
    $parent.Controls.Add($pnl)
    $lbl = New-Object Windows.Forms.Label
    $lbl.Text = $titulo
    $lbl.Location = New-Object Drawing.Point(10, 8)
    $lbl.Size = New-Object Drawing.Size(670, 22)
    $lbl.ForeColor = [Drawing.Color]::FromArgb(0, 180, 255)
    $lbl.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
    $pnl.Controls.Add($lbl)
    $lblSub = New-Object Windows.Forms.Label
    $lblSub.Text = $subtitulo
    $lblSub.Location = New-Object Drawing.Point(10, 32)
    $lblSub.Size = New-Object Drawing.Size(670, 18)
    $lblSub.ForeColor = [Drawing.Color]::FromArgb(160, 200, 255)
    $lblSub.Font = New-Object Drawing.Font("Segoe UI", 8)
    $pnl.Controls.Add($lblSub)
    return $pnl
}

function Install-MsOffCrypto {
    $check = python -c "import msoffcrypto" 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Out "Instalando msoffcrypto-tool..." $cSubText
        python -m pip install msoffcrypto-tool | Out-Null
    }
}

function Install-Pikepdf {
    $check = python -c "import pikepdf" 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Out "Instalando pikepdf..." $cSubText
        python -m pip install pikepdf | Out-Null
    }
}

function Write-Out($msg, $color = $null) {
    $ts = Get-Date -Format "HH:mm:ss"
    $outputBox.SelectionStart = $outputBox.TextLength
    $outputBox.SelectionColor = [Drawing.Color]::FromArgb(80, 120, 180)
    $outputBox.AppendText("`r`n [$ts] ")
    if ($color) { $outputBox.SelectionColor = $color }
    else { $outputBox.SelectionColor = $cAccent2 }
    $outputBox.AppendText($msg)
    $outputBox.ScrollToCaret()
}

function Run-Cmd($cmd) {
    Write-Out "Ejecutando: $cmd" $cSubText
    try {
        $res = Invoke-Expression $cmd 2>&1
        Write-Out ($res -join "`r`n") $cText
    } catch {
        Write-Out "Error: $_" $cRed
    }
}

function Save-Log {
    $dlg = New-Object Windows.Forms.SaveFileDialog
    $dlg.Filter = "Log files (*.log)|*.log|Text files (*.txt)|*.txt"
    $dlg.FileName = "SysCodi_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    if ($dlg.ShowDialog() -eq "OK") {
        $outputBox.Text | Set-Content $dlg.FileName -Encoding UTF8
        Write-Out "Log guardado en: $($dlg.FileName)" $cGreen
    }
}

# ============================================================
#   FORMULARIO PRINCIPAL
# ============================================================
$form = New-Object Windows.Forms.Form
$form.Text = "SysCodi WinTool Pro v2"
$form.Size = New-Object Drawing.Size(1200, 700)
$form.StartPosition = "CenterScreen"
$form.BackColor = $cBg
$form.ForeColor = $cText
$form.Font = New-Object Drawing.Font("Segoe UI", 9)
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false

# ============================================================
#   HEADER
# ============================================================
$header = New-Object Windows.Forms.Panel
$header.Size = New-Object Drawing.Size(1200, 60)
$header.Location = New-Object Drawing.Point(0, 0)
$header.BackColor = $cPanel
$form.Controls.Add($header)
$header.BringToFront()

if (Test-Path $logoPath) {
    $logoPic = New-Object Windows.Forms.PictureBox
    $logoPic.Location = New-Object Drawing.Point(10, 5)
    $logoPic.Size = New-Object Drawing.Size(50, 50)
    $logoPic.SizeMode = "Zoom"
    $logoPic.BackColor = $cPanel
    $logoPic.Image = [Drawing.Image]::FromFile($logoPath)
    $header.Controls.Add($logoPic)
    try {
        $bmp = [Drawing.Bitmap][Drawing.Image]::FromFile($logoPath)
        $icon = [Drawing.Icon]::FromHandle($bmp.GetHicon())
        $form.Icon = $icon
    } catch {}
    $titleX = 70
} else { $titleX = 15 }

$lblTitle = New-Object Windows.Forms.Label
$lblTitle.Text = "SysCodi WinTool Pro v2"
$lblTitle.Font = New-Object Drawing.Font("Segoe UI", 14, [Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = $cAccent2
$lblTitle.Location = New-Object Drawing.Point($titleX, 10)
$lblTitle.Size = New-Object Drawing.Size(500, 30)
$header.Controls.Add($lblTitle)

$lblSub = New-Object Windows.Forms.Label
$lblSub.Text = "Utilidad de sistema avanzada para Windows — v2.0"
$lblSub.Font = New-Object Drawing.Font("Segoe UI", 8)
$lblSub.ForeColor = $cSubText
$lblSub.Location = New-Object Drawing.Point($titleX, 38)
$lblSub.Size = New-Object Drawing.Size(500, 16)
$header.Controls.Add($lblSub)

# Indicador de admin
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$lblAdmin = New-Object Windows.Forms.Label
$lblAdmin.Text = if ($isAdmin) { "  Admin" } else { "  Sin Admin" }
$lblAdmin.ForeColor = if ($isAdmin) { $cGreen } else { $cRed }
$lblAdmin.Font = New-Object Drawing.Font("Segoe UI", 8, [Drawing.FontStyle]::Bold)
$lblAdmin.Location = New-Object Drawing.Point(1080, 20)
$lblAdmin.Size = New-Object Drawing.Size(110, 20)
$header.Controls.Add($lblAdmin)

# ============================================================
#   TAB CONTROL (7 PESTAÑAS)
# ============================================================
$tabs = New-Object Windows.Forms.TabControl
$tabs.Location = New-Object Drawing.Point(5, 65)
$tabs.Size = New-Object Drawing.Size(720, 570)
$tabs.BackColor = $cBg
$tabs.Appearance = "FlatButtons"
$tabs.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
$form.Controls.Add($tabs)

$tabRepair   = New-Tab " Reparación"
$tabApps     = New-Tab " Aplicaciones"
$tabTweaks   = New-Tab " Tweaks"
$tabUtils    = New-Tab " Utilidades"
$tabSecurity = New-Tab " Seguridad"
$tabBackup   = New-Tab " Backup"
$tabInfo     = New-Tab " Sistema"

# ============================================================
#   PANEL DERECHO — CONSOLA
# ============================================================
$rightPanel = New-Object Windows.Forms.Panel
$rightPanel.Location = New-Object Drawing.Point(728, 65)
$rightPanel.Size = New-Object Drawing.Size(455, 570)
$rightPanel.BackColor = [Drawing.Color]::FromArgb(10, 18, 40)
$form.Controls.Add($rightPanel)

$lblConsole = New-Object Windows.Forms.Label
$lblConsole.Text = "  Consola de salida"
$lblConsole.Location = New-Object Drawing.Point(0, 0)
$lblConsole.Size = New-Object Drawing.Size(310, 28)
$lblConsole.ForeColor = $cAccent2
$lblConsole.BackColor = [Drawing.Color]::FromArgb(22, 38, 75)
$lblConsole.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
$lblConsole.TextAlign = "MiddleLeft"
$rightPanel.Controls.Add($lblConsole)

$btnSaveLog = New-Object Windows.Forms.Button
$btnSaveLog.Text = "Guardar"
$btnSaveLog.Location = New-Object Drawing.Point(310, 3)
$btnSaveLog.Size = New-Object Drawing.Size(70, 22)
$btnSaveLog.BackColor = [Drawing.Color]::FromArgb(0, 80, 150)
$btnSaveLog.ForeColor = [Drawing.Color]::White
$btnSaveLog.FlatStyle = "Flat"
$btnSaveLog.Font = New-Object Drawing.Font("Segoe UI", 7)
$btnSaveLog.Add_Click({ Save-Log })
$rightPanel.Controls.Add($btnSaveLog)

$btnClearOutput = New-Object Windows.Forms.Button
$btnClearOutput.Text = "Limpiar"
$btnClearOutput.Location = New-Object Drawing.Point(383, 3)
$btnClearOutput.Size = New-Object Drawing.Size(68, 22)
$btnClearOutput.BackColor = [Drawing.Color]::FromArgb(0, 100, 180)
$btnClearOutput.ForeColor = [Drawing.Color]::White
$btnClearOutput.FlatStyle = "Flat"
$btnClearOutput.Font = New-Object Drawing.Font("Segoe UI", 7)
$btnClearOutput.Add_Click({ $outputBox.Clear(); Write-Out "Consola limpiada." $cSubText })
$rightPanel.Controls.Add($btnClearOutput)

# Buscador en consola
$txtSearch = New-Object Windows.Forms.TextBox
$txtSearch.Location = New-Object Drawing.Point(0, 30)
$txtSearch.Size = New-Object Drawing.Size(340, 22)
$txtSearch.BackColor = [Drawing.Color]::FromArgb(15, 28, 55)
$txtSearch.ForeColor = $cSubText
$txtSearch.BorderStyle = "FixedSingle"
$txtSearch.Font = New-Object Drawing.Font("Consolas", 8)
$txtSearch.Text = "Buscar en consola..."
$txtSearch.Add_Enter({ if ($txtSearch.Text -eq "Buscar en consola...") { $txtSearch.Text = ""; $txtSearch.ForeColor = $cText } })
$txtSearch.Add_Leave({ if ($txtSearch.Text -eq "") { $txtSearch.Text = "Buscar en consola..."; $txtSearch.ForeColor = $cSubText } })
$rightPanel.Controls.Add($txtSearch)

$btnSearch = New-Object Windows.Forms.Button
$btnSearch.Text = "Ir"
$btnSearch.Location = New-Object Drawing.Point(342, 30)
$btnSearch.Size = New-Object Drawing.Size(35, 22)
$btnSearch.BackColor = [Drawing.Color]::FromArgb(0, 100, 180)
$btnSearch.ForeColor = [Drawing.Color]::White
$btnSearch.FlatStyle = "Flat"
$btnSearch.Font = New-Object Drawing.Font("Segoe UI", 7)
$btnSearch.Add_Click({
    $q = $txtSearch.Text.Trim()
    if ($q -and $q -ne "Buscar en consola...") {
        $idx = $outputBox.Text.IndexOf($q, [System.StringComparison]::OrdinalIgnoreCase)
        if ($idx -ge 0) {
            $outputBox.Select($idx, $q.Length)
            $outputBox.ScrollToCaret()
        } else { Write-Out "No encontrado: $q" $cYellow }
    }
})
$rightPanel.Controls.Add($btnSearch)

$outputBox = New-Object Windows.Forms.RichTextBox
$outputBox.Location = New-Object Drawing.Point(0, 54)
$outputBox.Size = New-Object Drawing.Size(455, 516)
$outputBox.BackColor = $cOutput
$outputBox.ForeColor = $cAccent2
$outputBox.Font = New-Object Drawing.Font("Consolas", 9)
$outputBox.ReadOnly = $true
$outputBox.BorderStyle = "None"
$outputBox.Text = "  Listo. Selecciona una opción y ejecuta."
$rightPanel.Controls.Add($outputBox)

# ============================================================
#   TAB 1: REPARACIÓN
# ============================================================
New-SectionLabel " Limpieza" 10 10 $tabRepair

$btnLimpiar = New-CorporateButton "  Limpiar Temporales" 10 35
$btnLimpiar.Add_Click({
    Remove-Item "$env:TEMP\*" -Recurse -Force -EA SilentlyContinue
    Remove-Item "C:\Windows\Temp\*" -Recurse -Force -EA SilentlyContinue
    Write-Out "Temporales eliminados correctamente." $cGreen
})
$tabRepair.Controls.Add($btnLimpiar)

$btnPrefetch = New-CorporateButton "  Limpiar Prefetch" 220 35
$btnPrefetch.Add_Click({
    Remove-Item "C:\Windows\Prefetch\*" -Recurse -Force -EA SilentlyContinue
    Write-Out "Prefetch limpiado." $cGreen
})
$tabRepair.Controls.Add($btnPrefetch)

$btnWUpdate = New-CorporateButton "  Limpiar Caché Windows Update" 430 35
$btnWUpdate.Add_Click({
    Stop-Service wuauserv -Force -EA SilentlyContinue
    Remove-Item "C:\Windows\SoftwareDistribution\Download\*" -Recurse -Force -EA SilentlyContinue
    Start-Service wuauserv -EA SilentlyContinue
    Write-Out "Caché de Windows Update limpiada." $cGreen
})
$tabRepair.Controls.Add($btnWUpdate)

New-SectionLabel " Reparación de Windows" 10 85 $tabRepair

$btnSFC = New-CorporateButton "  SFC /scannow" 10 110
$btnSFC.Add_Click({ Run-Cmd "sfc /scannow" })
$tabRepair.Controls.Add($btnSFC)

$btnDISM = New-CorporateButton "  DISM RestoreHealth" 220 110
$btnDISM.Add_Click({ Run-Cmd "DISM /Online /Cleanup-Image /RestoreHealth" })
$tabRepair.Controls.Add($btnDISM)

$btnChkDsk = New-CorporateButton "  CheckDisk (C:)" 430 110
$btnChkDsk.Add_Click({ Run-Cmd "chkdsk C: /f /r /x" })
$tabRepair.Controls.Add($btnChkDsk)

$btnStoreReset = New-CorporateButton "  Reparar Microsoft Store" 10 158
$btnStoreReset.Add_Click({
    Write-Out "Reparando Microsoft Store..." $cSubText
    Start-Process wsreset.exe
    Write-Out "Store reiniciada. Espera unos segundos." $cGreen
})
$tabRepair.Controls.Add($btnStoreReset)

$btnRestorePoint = New-CorporateButton "  Crear Punto de Restauración" 220 158
$btnRestorePoint.Add_Click({
    Write-Out "Creando punto de restauración..." $cSubText
    try {
        Checkpoint-Computer -Description "SysCodi Backup $(Get-Date -Format 'dd/MM/yyyy')" -RestorePointType MODIFY_SETTINGS
        Write-Out "Punto de restauración creado." $cGreen
    } catch { Write-Out "Error: $_" $cRed }
})
$tabRepair.Controls.Add($btnRestorePoint)

$btnRestoreSys = New-CorporateButton "  Abrir Restaurar Sistema" 430 158
$btnRestoreSys.Add_Click({ Start-Process rstrui.exe })
$tabRepair.Controls.Add($btnRestoreSys)

New-SectionLabel " Red" 10 205 $tabRepair

$btnDNS = New-CorporateButton "  DNS Flush" 10 230
$btnDNS.Add_Click({ Run-Cmd "ipconfig /flushdns" })
$tabRepair.Controls.Add($btnDNS)

$btnNetReset = New-CorporateButton "  Reset Red (netsh)" 220 230
$btnNetReset.Add_Click({
    Run-Cmd "netsh int ip reset"
    Run-Cmd "netsh winsock reset"
    Write-Out "Reinicia el PC para aplicar cambios de red." $cYellow
})
$tabRepair.Controls.Add($btnNetReset)

$btnPuertos = New-CorporateButton "  Ver Puertos Abiertos" 430 230
$btnPuertos.Add_Click({ Run-Cmd "netstat -ano" })
$tabRepair.Controls.Add($btnPuertos)

$btnPing = New-CorporateButton "  Diagnóstico de Red" 10 278
$btnPing.Add_Click({
    Write-Out "--- Diagnóstico de red ---" $cAccent2
    Run-Cmd "ping 8.8.8.8 -n 3"
    Run-Cmd "tracert -d -h 5 8.8.8.8"
    Run-Cmd "Test-NetConnection google.com -Port 443"
})
$tabRepair.Controls.Add($btnPing)

$btnKill80 = New-CorporateButton "  Matar Puerto 80" 220 278
$btnKill80.Add_Click({
    $pids = (netstat -ano | Select-String ":80\s") -replace '.*\s(\d+)$','$1' | Sort-Object -Unique
    foreach ($p in $pids) {
        if ($p -match '^\d+$') {
            Stop-Process -Id $p -Force -EA SilentlyContinue
            Write-Out "PID $p en puerto 80 terminado." $cGreen
        }
    }
})
$tabRepair.Controls.Add($btnKill80)

New-SectionLabel " Servicios" 10 325 $tabRepair

$btnSlowServices = New-CorporateButton "  Ver Servicios Lentos al Inicio" 10 350
$btnSlowServices.Add_Click({
    Write-Out "--- Servicios automáticos activos ---" $cAccent2
    Get-Service | Where-Object { $_.StartType -eq 'Automatic' -and $_.Status -eq 'Running' } |
        Select-Object Name, DisplayName |
        ForEach-Object { Write-Out "$($_.Name) — $($_.DisplayName)" $cText }
})
$tabRepair.Controls.Add($btnSlowServices)

$btnEventLog = New-CorporateButton "  Ver Errores del Sistema" 220 350
$btnEventLog.Add_Click({
    Write-Out "--- Últimos errores del sistema ---" $cAccent2
    try {
        Get-EventLog -LogName System -EntryType Error -Newest 10 |
            ForEach-Object { Write-Out "$($_.TimeGenerated) — $($_.Source): $($_.Message.Substring(0,[Math]::Min(80,$_.Message.Length)))" $cRed }
    } catch { Write-Out "Error al leer log: $_" $cRed }
})
$tabRepair.Controls.Add($btnEventLog)

# ============================================================
#   TAB 2: APLICACIONES
# ============================================================
$txtAppFilter = New-Object Windows.Forms.TextBox
$txtAppFilter.Location = New-Object Drawing.Point(5, 5)
$txtAppFilter.Size = New-Object Drawing.Size(350, 26)
$txtAppFilter.BackColor = [Drawing.Color]::FromArgb(22, 38, 75)
$txtAppFilter.ForeColor = $cSubText
$txtAppFilter.BorderStyle = "FixedSingle"
$txtAppFilter.Font = New-Object Drawing.Font("Segoe UI", 9)
$txtAppFilter.Text = "Buscar aplicación..."
$tabApps.Controls.Add($txtAppFilter)

$scroll = New-Object Windows.Forms.Panel
$scroll.Location = New-Object Drawing.Point(0, 35)
$scroll.Size = New-Object Drawing.Size(875, 430)
$scroll.AutoScroll = $true
$scroll.BackColor = $cBg
$tabApps.Controls.Add($scroll)

$appList = @(
    @{cat="Navegadores";      name="Google Chrome";    cmd="winget install -e --id Google.Chrome"},
    @{cat="Navegadores";      name="Mozilla Firefox";  cmd="winget install -e --id Mozilla.Firefox"},
    @{cat="Navegadores";      name="Brave Browser";    cmd="winget install -e --id Brave.Brave"; foss=$true},
    @{cat="Navegadores";      name="LibreWolf";        cmd="winget install -e --id LibreWolf.LibreWolf"; foss=$true},
    @{cat="Comunicación";     name="Discord";          cmd="winget install -e --id Discord.Discord"},
    @{cat="Comunicación";     name="Telegram";         cmd="winget install -e --id Telegram.TelegramDesktop"; foss=$true},
    @{cat="Comunicación";     name="Slack";            cmd="winget install -e --id SlackTechnologies.Slack"},
    @{cat="Comunicación";     name="Signal";           cmd="winget install -e --id OpenWhisperSystems.Signal"; foss=$true},
    @{cat="Desarrollo";       name="VS Code";          cmd="winget install -e --id Microsoft.VisualStudioCode"},
    @{cat="Desarrollo";       name="Git";              cmd="winget install -e --id Git.Git"; foss=$true},
    @{cat="Desarrollo";       name="Python 3";         cmd="winget install -e --id Python.Python.3"; foss=$true},
    @{cat="Desarrollo";       name="NodeJS LTS";       cmd="winget install -e --id OpenJS.NodeJS.LTS"; foss=$true},
    @{cat="Herramientas";     name="7-Zip";            cmd="winget install -e --id 7zip.7zip"; foss=$true},
    @{cat="Herramientas";     name="VLC";              cmd="winget install -e --id VideoLAN.VLC"; foss=$true},
    @{cat="Herramientas";     name="WinRAR";           cmd="winget install -e --id RARLab.WinRAR"},
    @{cat="Herramientas";     name="Notepad++";        cmd="winget install -e --id Notepad++.Notepad++"; foss=$true},
    @{cat="Herramientas";     name="Everything";       cmd="winget install -e --id voidtools.Everything"; foss=$true},
    @{cat="Herramientas";     name="ShareX";           cmd="winget install -e --id ShareX.ShareX"; foss=$true},
    @{cat="Herramientas";     name="Rufus";            cmd="winget install -e --id Rufus.Rufus"; foss=$true},
    @{cat="Multimedia";       name="OBS Studio";       cmd="winget install -e --id OBSProject.OBSStudio"; foss=$true},
    @{cat="Multimedia";       name="VLC";              cmd="winget install -e --id VideoLAN.VLC"; foss=$true},
    @{cat="Hardware";         name="CrystalDiskInfo";  cmd="winget install -e --id CrystalDewWorld.CrystalDiskInfo"; foss=$true},
    @{cat="Hardware";         name="HWiNFO";           cmd="winget install -e --id REALiX.HWiNFO"},
    @{cat="Hardware";         name="GPU-Z";            cmd="winget install -e --id TechPowerUp.GPU-Z"},
    @{cat="Seguridad";        name="Bitwarden";        cmd="winget install -e --id Bitwarden.Bitwarden"; foss=$true},
    @{cat="Microsoft Office"; name="Office 2019";      cmd="winget install -e --id Microsoft.Office2019.HomeAndBusiness"},
    @{cat="Microsoft Office"; name="Office 2021";      cmd="winget install -e --id Microsoft.Office2021.HomeAndBusiness"},
    @{cat="Microsoft Office"; name="Office 2024";      cmd="winget install -e --id Microsoft.Office2024.HomeAndBusiness"},
    @{cat="Microsoft Office"; name="Microsoft 365";    cmd="winget install -e --id Microsoft.Microsoft365"},
    @{cat="Microsoft Office"; name="OneDrive";         cmd="winget install -e --id Microsoft.OneDrive"},
    @{cat="Microsoft Office"; name="Teams";            cmd="winget install -e --id Microsoft.Teams"}
)

$checkboxes = @()
$yPos = 5; $lastCat = ""; $col = 0

function Rebuild-AppList($filter = "") {
    $scroll.Controls.Clear()
    $script:checkboxes = @()
    $yy = 5; $lcat = ""; $cc = 0
    foreach ($app in $appList) {
        if ($filter -and $app.name -notlike "*$filter*" -and $app.cat -notlike "*$filter*") { continue }
        if ($app.cat -ne $lcat) {
            $cc = 0
            if ($lcat -ne "") { $yy += 10 }
            $lbl = New-Object Windows.Forms.Label
            $lbl.Text = " $($app.cat) "
            $lbl.Location = New-Object Drawing.Point(5, $yy)
            $lbl.Size = New-Object Drawing.Size(860, 20)
            $lbl.ForeColor = $cAccent2
            $lbl.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
            $scroll.Controls.Add($lbl)
            $yy += 22; $lcat = $app.cat
        }
        $cb = New-Object Windows.Forms.CheckBox
        $cb.Text = $app.name
        $cb.Location = New-Object Drawing.Point((5 + $cc * 210), $yy)
        $cb.Size = New-Object Drawing.Size(200, 22)
        $cb.ForeColor = if ($app.foss) { $cAccent2 } else { $cText }
        $cb.BackColor = $cBg
        $cb.Tag = $app.cmd
        $scroll.Controls.Add($cb)
        $script:checkboxes += $cb
        $cc++
        if ($cc -ge 4) { $cc = 0; $yy += 25 }
    }
}

Rebuild-AppList

$txtAppFilter.Add_TextChanged({
    $q = $txtAppFilter.Text.Trim()
    if ($q -eq "Buscar aplicación...") { Rebuild-AppList } else { Rebuild-AppList $q }
})
$txtAppFilter.Add_Enter({ if ($txtAppFilter.Text -eq "Buscar aplicación...") { $txtAppFilter.Text = ""; $txtAppFilter.ForeColor = $cText } })
$txtAppFilter.Add_Leave({ if ($txtAppFilter.Text -eq "") { $txtAppFilter.Text = "Buscar aplicación..."; $txtAppFilter.ForeColor = $cSubText } })

$pnlAppBtns = New-Object Windows.Forms.Panel
$pnlAppBtns.Location = New-Object Drawing.Point(0, 470)
$pnlAppBtns.Size = New-Object Drawing.Size(875, 45)
$pnlAppBtns.BackColor = $cPanel
$tabApps.Controls.Add($pnlAppBtns)

$lblFoss = New-Object Windows.Forms.Label
$lblFoss.Text = " Azul claro = FOSS (Software Libre)"
$lblFoss.ForeColor = $cAccent2
$lblFoss.Location = New-Object Drawing.Point(10, 12)
$lblFoss.Size = New-Object Drawing.Size(220, 20)
$pnlAppBtns.Controls.Add($lblFoss)

$btnListInstalled = New-CorporateButton " Ver Instaladas" 230 5 160 34
$btnListInstalled.Add_Click({
    Write-Out "--- Apps instaladas (winget list) ---" $cAccent2
    $res = winget list 2>&1
    foreach ($line in $res) { Write-Out $line $cText }
})
$pnlAppBtns.Controls.Add($btnListInstalled)

$btnUpgradeAll = New-CorporateButton " Actualizar Todo" 400 5 160 34
$btnUpgradeAll.BackColor = [Drawing.Color]::FromArgb(0, 100, 60)
$btnUpgradeAll.Add_Click({
    Write-Out "Actualizando todas las apps con winget..." $cSubText
    Start-Process powershell -ArgumentList "-NoProfile -Command `"winget upgrade --all --silent`"" -Verb RunAs
    Write-Out "Actualización iniciada en ventana separada." $cGreen
})
$pnlAppBtns.Controls.Add($btnUpgradeAll)

$btnInstallApps = New-CorporateButton " Instalar Seleccionadas" 570 5 200 34
$btnInstallApps.Add_Click({
    $sel = $checkboxes | Where-Object { $_.Checked }
    if ($sel.Count -eq 0) { Write-Out "No seleccionaste ninguna aplicación." $cYellow; return }
    foreach ($cb in $sel) {
        Write-Out "Instalando: $($cb.Text)..." $cSubText
        Start-Process powershell -ArgumentList "-NoProfile -Command `"$($cb.Tag)`"" -Wait
        Write-Out "$($cb.Text) instalado." $cGreen
    }
})
$pnlAppBtns.Controls.Add($btnInstallApps)

# ============================================================
#   TAB 3: TWEAKS
# ============================================================
New-SectionLabel " Rendimiento, privacidad y experiencia" 10 10 $tabTweaks

$tweaks = @(
    @{name=" Plan de energía: alto rendimiento";    cmd='powercfg /s 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'; undo='powercfg /s 381b4222-f694-41f0-9685-ff5bb260df2e'},
    @{name=" Deshabilitar efectos visuales";         cmd='SystemPropertiesPerformance.exe'; undo=''},
    @{name=" Deshabilitar notificaciones";           cmd='reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\PushNotifications" /v ToastEnabled /t REG_DWORD /d 0 /f'; undo='reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\PushNotifications" /v ToastEnabled /t REG_DWORD /d 1 /f'},
    @{name="  Deshabilitar telemetría";              cmd='reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection" /v AllowTelemetry /t REG_DWORD /d 0 /f'; undo='reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection" /v AllowTelemetry /f'},
    @{name=" Deshabilitar Cortana";                  cmd='reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /v AllowCortana /t REG_DWORD /d 0 /f'; undo='reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /v AllowCortana /f'},
    @{name="  Modo juego activado";                  cmd='reg add "HKCU\Software\Microsoft\GameBar" /v AllowAutoGameMode /t REG_DWORD /d 1 /f'; undo='reg add "HKCU\Software\Microsoft\GameBar" /v AllowAutoGameMode /t REG_DWORD /d 0 /f'},
    @{name=" Mostrar extensiones de archivo";        cmd='reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v HideFileExt /t REG_DWORD /d 0 /f'; undo='reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v HideFileExt /t REG_DWORD /d 1 /f'},
    @{name="  Mostrar archivos ocultos";             cmd='reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v Hidden /t REG_DWORD /d 1 /f'; undo='reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v Hidden /t REG_DWORD /d 2 /f'},
    @{name=" Deshabilitar OneDrive al inicio";        cmd='reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v OneDrive /t REG_SZ /d "" /f'; undo=''},
    @{name="  Deshabilitar Xbox Game Bar";           cmd='reg add "HKCU\System\GameConfigStore" /v GameDVR_Enabled /t REG_DWORD /d 0 /f'; undo='reg add "HKCU\System\GameConfigStore" /v GameDVR_Enabled /t REG_DWORD /d 1 /f'},
    @{name=" Activar God Mode en Escritorio";        cmd='$gm="$env:USERPROFILE\Desktop\GodMode.{ED7BA470-8E54-465E-825C-99712043E01C}"; New-Item -ItemType Directory -Path $gm -EA SilentlyContinue'; undo='Remove-Item "$env:USERPROFILE\Desktop\GodMode.{ED7BA470-8E54-465E-825C-99712043E01C}" -EA SilentlyContinue'},
    @{name="  Deshabilitar actualizaciones auto";    cmd='reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v NoAutoUpdate /t REG_DWORD /d 1 /f'; undo='reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v NoAutoUpdate /f'}
)

$yT = 35; $colT = 0; $tweakChecks = @()
foreach ($tw in $tweaks) {
    $cb = New-Object Windows.Forms.CheckBox
    $cb.Text = $tw.name
    $cb.Location = New-Object Drawing.Point((10 + $colT * 350), $yT)
    $cb.Size = New-Object Drawing.Size(340, 24)
    $cb.ForeColor = $cText
    $cb.BackColor = $cBg
    $cb.Tag = $tw.cmd
    $cb.AccessibleDescription = $tw.undo
    $tabTweaks.Controls.Add($cb)
    $tweakChecks += $cb
    $colT++
    if ($colT -ge 2) { $colT = 0; $yT += 28 }
}

$btnApplyTweaks = New-CorporateButton "  Aplicar Tweaks Seleccionados" 10 400 260 38
$btnApplyTweaks.Add_Click({
    $sel = $tweakChecks | Where-Object { $_.Checked }
    if ($sel.Count -eq 0) { Write-Out "No seleccionaste ningún tweak." $cYellow; return }
    foreach ($cb in $sel) {
        Write-Out "Aplicando: $($cb.Text)..." $cSubText
        Invoke-Expression $cb.Tag 2>&1 | Out-Null
        Write-Out "Listo." $cGreen
    }
    Write-Out "Todos los tweaks aplicados. Puede requerir reinicio." $cGreen
})
$tabTweaks.Controls.Add($btnApplyTweaks)

$btnRevertTweaks = New-CorporateButton "  Revertir Tweaks Seleccionados" 280 400 260 38
$btnRevertTweaks.BackColor = [Drawing.Color]::FromArgb(120, 60, 0)
$btnRevertTweaks.Add_Click({
    $sel = $tweakChecks | Where-Object { $_.Checked }
    foreach ($cb in $sel) {
        if ($cb.AccessibleDescription) {
            Write-Out "Revirtiendo: $($cb.Text)..." $cSubText
            Invoke-Expression $cb.AccessibleDescription 2>&1 | Out-Null
            Write-Out "Revertido." $cYellow
        }
    }
})
$tabTweaks.Controls.Add($btnRevertTweaks)

# ============================================================
#   TAB 4: UTILIDADES
# ============================================================
$utilScroll = New-Object Windows.Forms.Panel
$utilScroll.Location = New-Object Drawing.Point(0, 0)
$utilScroll.Size = New-Object Drawing.Size(720, 565)
$utilScroll.AutoScroll = $true
$utilScroll.BackColor = $cBg
$tabUtils.Controls.Add($utilScroll)

# Excel
$pnlExcel = New-UtilPanel "  Quitar contraseña — Excel (.xlsx / .xls)" "Crea una copia sin contraseña en la misma carpeta." $utilScroll 10
$lblExcelPath = New-Object Windows.Forms.Label; $lblExcelPath.Text = "Ningún archivo seleccionado"
$lblExcelPath.Location = New-Object Drawing.Point(10, 55); $lblExcelPath.Size = New-Object Drawing.Size(670, 16)
$lblExcelPath.ForeColor = $cText; $lblExcelPath.Font = New-Object Drawing.Font("Consolas", 7); $pnlExcel.Controls.Add($lblExcelPath)
$btnBrowseExcel = New-CorporateButton "Buscar Excel" 10 75 150 32
$btnBrowseExcel.Add_Click({ $dlg = New-Object Windows.Forms.OpenFileDialog; $dlg.Filter = "Excel (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm"; if ($dlg.ShowDialog() -eq "OK") { $lblExcelPath.Text = $dlg.FileName } })
$pnlExcel.Controls.Add($btnBrowseExcel)
$btnRemoveExcel = New-CorporateButton "Quitar Contraseña" 170 75 180 32
$btnRemoveExcel.Add_Click({
    $path = $lblExcelPath.Text
    if (-not (Test-Path $path)) { Write-Out "Selecciona un archivo Excel primero." $cYellow; return }
    Install-MsOffCrypto
    $out = $path -replace '(\.[^.]+)$','_sin_pass$1'
    $py = "import msoffcrypto`nwith open(r'$path','rb') as f:`n    o=msoffcrypto.OfficeFile(f)`n    o.load_key(password='')`n    with open(r'$out','wb') as fw: o.decrypt(fw)`nprint('OK')"
    $tmp = "$env:TEMP\unlock_excel.py"; $py | Set-Content $tmp -Encoding UTF8
    $res = python $tmp 2>&1
    if ($res -like "*OK*") { Write-Out "Excel desbloqueado: $out" $cGreen } else { Write-Out "Error: $res" $cRed }
})
$pnlExcel.Controls.Add($btnRemoveExcel)

# Word
$pnlWord = New-UtilPanel "  Quitar contraseña — Word (.docx / .doc)" "Crea una copia sin contraseña en la misma carpeta." $utilScroll 140
$lblWordPath = New-Object Windows.Forms.Label; $lblWordPath.Text = "Ningún archivo seleccionado"
$lblWordPath.Location = New-Object Drawing.Point(10, 55); $lblWordPath.Size = New-Object Drawing.Size(670, 16)
$lblWordPath.ForeColor = $cText; $lblWordPath.Font = New-Object Drawing.Font("Consolas", 7); $pnlWord.Controls.Add($lblWordPath)
$btnBrowseWord = New-CorporateButton "Buscar Word" 10 75 150 32
$btnBrowseWord.Add_Click({ $dlg = New-Object Windows.Forms.OpenFileDialog; $dlg.Filter = "Word (*.docx;*.doc;*.docm)|*.docx;*.doc;*.docm"; if ($dlg.ShowDialog() -eq "OK") { $lblWordPath.Text = $dlg.FileName } })
$pnlWord.Controls.Add($btnBrowseWord)
$btnRemoveWord = New-CorporateButton "Quitar Contraseña" 170 75 180 32
$btnRemoveWord.Add_Click({
    $path = $lblWordPath.Text
    if (-not (Test-Path $path)) { Write-Out "Selecciona un archivo Word primero." $cYellow; return }
    Install-MsOffCrypto
    $out = $path -replace '(\.[^.]+)$','_sin_pass$1'
    $py = "import msoffcrypto`nwith open(r'$path','rb') as f:`n    o=msoffcrypto.OfficeFile(f)`n    o.load_key(password='')`n    with open(r'$out','wb') as fw: o.decrypt(fw)`nprint('OK')"
    $tmp = "$env:TEMP\unlock_word.py"; $py | Set-Content $tmp -Encoding UTF8
    $res = python $tmp 2>&1
    if ($res -like "*OK*") { Write-Out "Word desbloqueado: $out" $cGreen } else { Write-Out "Error: $res" $cRed }
})
$pnlWord.Controls.Add($btnRemoveWord)

# PDF
$pnlPdf = New-UtilPanel "  Quitar contraseña — PDF" "Requiere Python + pikepdf. Se instala automáticamente." $utilScroll 270
$lblPdfPath = New-Object Windows.Forms.Label; $lblPdfPath.Text = "Ningún archivo seleccionado"
$lblPdfPath.Location = New-Object Drawing.Point(10, 55); $lblPdfPath.Size = New-Object Drawing.Size(400, 16)
$lblPdfPath.ForeColor = $cText; $lblPdfPath.Font = New-Object Drawing.Font("Consolas", 7); $pnlPdf.Controls.Add($lblPdfPath)
$lblPdfPass = New-Object Windows.Forms.Label; $lblPdfPass.Text = "Contraseña:"; $lblPdfPass.Location = New-Object Drawing.Point(10, 76); $lblPdfPass.Size = New-Object Drawing.Size(80, 20); $lblPdfPass.ForeColor = $cText; $lblPdfPass.Font = New-Object Drawing.Font("Segoe UI", 8); $pnlPdf.Controls.Add($lblPdfPass)
$txtPdfPass = New-Object Windows.Forms.TextBox; $txtPdfPass.Location = New-Object Drawing.Point(93, 74); $txtPdfPass.Size = New-Object Drawing.Size(170, 22); $txtPdfPass.UseSystemPasswordChar = $true; $txtPdfPass.BackColor = [Drawing.Color]::FromArgb(10,18,40); $txtPdfPass.ForeColor = $cText; $pnlPdf.Controls.Add($txtPdfPass)
$btnBrowsePdf = New-CorporateButton "Buscar PDF" 10 105 130 28
$btnBrowsePdf.Add_Click({ $dlg = New-Object Windows.Forms.OpenFileDialog; $dlg.Filter = "PDF (*.pdf)|*.pdf"; if ($dlg.ShowDialog() -eq "OK") { $lblPdfPath.Text = $dlg.FileName } })
$pnlPdf.Controls.Add($btnBrowsePdf)
$btnRemovePdf = New-CorporateButton "Quitar Contraseña PDF" 150 105 200 28
$btnRemovePdf.Add_Click({
    $path = $lblPdfPath.Text; $pass = $txtPdfPass.Text.Trim()
    if (-not (Test-Path $path)) { Write-Out "Selecciona un archivo PDF primero." $cYellow; return }
    Install-Pikepdf
    $out = $path -replace '\.pdf$','_sin_pass.pdf'
    $py = "import pikepdf`ntry:`n    pdf=pikepdf.open(r'$path',password='$pass')`n    pdf.save(r'$out')`n    print('OK')`nexcept Exception as e:`n    print('ERROR:'+str(e))"
    $tmp = "$env:TEMP\unlock_pdf.py"; $py | Set-Content $tmp -Encoding UTF8
    $res = python $tmp 2>&1
    if ($res -like "*OK*") { Write-Out "PDF desbloqueado: $out" $cGreen } else { Write-Out "Error: $res" $cRed }
})
$pnlPdf.Controls.Add($btnRemovePdf)

# Hash
$pnlHash = New-UtilPanel "  Verificador de hashes" "Calcula MD5, SHA1 y SHA256 de cualquier archivo." $utilScroll 420 110
$lblHashPath = New-Object Windows.Forms.Label; $lblHashPath.Text = "Ningún archivo seleccionado"; $lblHashPath.Location = New-Object Drawing.Point(10, 55); $lblHashPath.Size = New-Object Drawing.Size(670, 16); $lblHashPath.ForeColor = $cText; $lblHashPath.Font = New-Object Drawing.Font("Consolas", 7); $pnlHash.Controls.Add($lblHashPath)
$btnBrowseHash = New-CorporateButton "Seleccionar archivo" 10 75 180 28
$btnBrowseHash.Add_Click({ $dlg = New-Object Windows.Forms.OpenFileDialog; if ($dlg.ShowDialog() -eq "OK") { $lblHashPath.Text = $dlg.FileName } })
$pnlHash.Controls.Add($btnBrowseHash)
$btnCalcHash = New-CorporateButton "Calcular Hashes" 200 75 160 28
$btnCalcHash.Add_Click({
    $f = $lblHashPath.Text
    if (-not (Test-Path $f)) { Write-Out "Selecciona un archivo primero." $cYellow; return }
    Write-Out "--- Hashes de: $(Split-Path $f -Leaf) ---" $cAccent2
    Write-Out "MD5   : $((Get-FileHash $f -Algorithm MD5).Hash)" $cText
    Write-Out "SHA1  : $((Get-FileHash $f -Algorithm SHA1).Hash)" $cText
    Write-Out "SHA256: $((Get-FileHash $f -Algorithm SHA256).Hash)" $cText
})
$pnlHash.Controls.Add($btnCalcHash)

# Activación
$pnlActivar = New-UtilPanel "  Activación Windows / Office (MAS)" "irm https://get.activated.win | iex — Proyecto open source MAS" $utilScroll 550 260
$lblSecWin = New-Object Windows.Forms.Label; $lblSecWin.Text = "  Windows"; $lblSecWin.Location = New-Object Drawing.Point(10, 78); $lblSecWin.Size = New-Object Drawing.Size(200, 18); $lblSecWin.ForeColor = $cAccent2; $lblSecWin.Font = New-Object Drawing.Font("Segoe UI", 8, [Drawing.FontStyle]::Bold); $pnlActivar.Controls.Add($lblSecWin)
$winVersions = @("Windows 7","Windows 8.1","Windows 10","Windows 11","Windows Server 2019","Windows Server 2022")
$cbWin = @(); $xW = 10; $yW = 98
foreach ($ver in $winVersions) { $cb = New-Object Windows.Forms.CheckBox; $cb.Text = $ver; $cb.Location = New-Object Drawing.Point($xW, $yW); $cb.Size = New-Object Drawing.Size(160, 20); $cb.ForeColor = $cText; $cb.BackColor = [Drawing.Color]::FromArgb(22,38,75); $cb.Font = New-Object Drawing.Font("Segoe UI", 8); $pnlActivar.Controls.Add($cb); $cbWin += $cb; $xW += 162; if ($xW -gt 490) { $xW = 10; $yW += 22 } }
$lblSecOff = New-Object Windows.Forms.Label; $lblSecOff.Text = "  Office"; $lblSecOff.Location = New-Object Drawing.Point(10, 148); $lblSecOff.Size = New-Object Drawing.Size(200, 18); $lblSecOff.ForeColor = $cGreen; $lblSecOff.Font = New-Object Drawing.Font("Segoe UI", 8, [Drawing.FontStyle]::Bold); $pnlActivar.Controls.Add($lblSecOff)
$offVersions = @("Office 2013","Office 2016","Office 2019","Office 2021","Office 2024","Microsoft 365")
$cbOff = @(); $xO = 10; $yO = 168
foreach ($ver in $offVersions) { $cb = New-Object Windows.Forms.CheckBox; $cb.Text = $ver; $cb.Location = New-Object Drawing.Point($xO, $yO); $cb.Size = New-Object Drawing.Size(160, 20); $cb.ForeColor = $cGreen; $cb.BackColor = [Drawing.Color]::FromArgb(22,38,75); $cb.Font = New-Object Drawing.Font("Segoe UI", 8); $pnlActivar.Controls.Add($cb); $cbOff += $cb; $xO += 162; if ($xO -gt 490) { $xO = 10; $yO += 22 } }
$btnActivar = New-CorporateButton "  Activar Seleccionados" 10 215 200 32
$btnActivar.BackColor = [Drawing.Color]::FromArgb(0, 130, 60)
$btnActivar.FlatAppearance.BorderColor = $cGreen
$btnActivar.Add_Click({
    $selTodos = @($cbWin | Where-Object { $_.Checked } | ForEach-Object { $_.Text }) + @($cbOff | Where-Object { $_.Checked } | ForEach-Object { $_.Text })
    if ($selTodos.Count -eq 0) { Write-Out "Selecciona al menos un producto." $cYellow; return }
    $confirm = [Windows.Forms.MessageBox]::Show("Se activarán: $($selTodos -join ', ')`n`nComando: irm https://get.activated.win | iex`n`n¿Continuar?","Activación",[Windows.Forms.MessageBoxButtons]::YesNo,[Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirm -eq [Windows.Forms.DialogResult]::Yes) {
        Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"irm https://get.activated.win | iex`"" -Verb RunAs
        Write-Out "Script MAS lanzado. Sigue las instrucciones." $cGreen
    }
})
$pnlActivar.Controls.Add($btnActivar)

# ============================================================
#   TAB 5: SEGURIDAD
# ============================================================
New-SectionLabel " Estado del sistema de seguridad" 10 10 $tabSecurity

$btnDefender = New-CorporateButton "  Estado de Defender" 10 35 220 36
$btnDefender.Add_Click({
    Write-Out "--- Windows Defender ---" $cAccent2
    try {
        $status = Get-MpComputerStatus
        Write-Out "Antivirus activo    : $($status.AntivirusEnabled)" $(if($status.AntivirusEnabled){$cGreen}else{$cRed})
        Write-Out "Tiempo real activo  : $($status.RealTimeProtectionEnabled)" $(if($status.RealTimeProtectionEnabled){$cGreen}else{$cRed})
        Write-Out "Última actualización : $($status.AntivirusSignatureLastUpdated)" $cText
        Write-Out "Definiciones        : $($status.AntivirusSignatureVersion)" $cText
    } catch { Write-Out "Error al leer Defender: $_" $cRed }
})
$tabSecurity.Controls.Add($btnDefender)

$btnQuickScan = New-CorporateButton "  Quick Scan Defender" 240 35 200 36
$btnQuickScan.Add_Click({
    Write-Out "Iniciando Quick Scan..." $cSubText
    Start-Process powershell -ArgumentList "-NoProfile -Command `"Start-MpScan -ScanType QuickScan`"" -Verb RunAs
    Write-Out "Scan iniciado en segundo plano." $cGreen
})
$tabSecurity.Controls.Add($btnQuickScan)

$btnFirewall = New-CorporateButton "  Estado del Firewall" 450 35 200 36
$btnFirewall.Add_Click({
    Write-Out "--- Estado del Firewall ---" $cAccent2
    try {
        $profiles = Get-NetFirewallProfile
        foreach ($p in $profiles) {
            $color = if ($p.Enabled) { $cGreen } else { $cRed }
            Write-Out "$($p.Name): $( if($p.Enabled){'ACTIVO'}else{'INACTIVO'} )" $color
        }
    } catch { Run-Cmd "netsh advfirewall show allprofiles state" }
})
$tabSecurity.Controls.Add($btnFirewall)

$btnFirewallOn = New-CorporateButton "  Activar Firewall" 10 85 200 36
$btnFirewallOn.BackColor = [Drawing.Color]::FromArgb(0, 100, 40)
$btnFirewallOn.Add_Click({
    Run-Cmd "netsh advfirewall set allprofiles state on"
    Write-Out "Firewall activado en todos los perfiles." $cGreen
})
$tabSecurity.Controls.Add($btnFirewallOn)

$btnFirewallOff = New-CorporateButton "  Desactivar Firewall" 220 85 200 36
$btnFirewallOff.BackColor = [Drawing.Color]::FromArgb(120, 30, 30)
$btnFirewallOff.Add_Click({
    $c = [Windows.Forms.MessageBox]::Show("¿Seguro que deseas desactivar el Firewall?","Advertencia",[Windows.Forms.MessageBoxButtons]::YesNo,[Windows.Forms.MessageBoxIcon]::Warning)
    if ($c -eq [Windows.Forms.DialogResult]::Yes) {
        Run-Cmd "netsh advfirewall set allprofiles state off"
        Write-Out "Firewall desactivado." $cYellow
    }
})
$tabSecurity.Controls.Add($btnFirewallOff)

New-SectionLabel " Usuarios y cuentas" 10 140 $tabSecurity

$btnListUsers = New-CorporateButton "  Listar Usuarios" 10 165 200 36
$btnListUsers.Add_Click({
    Write-Out "--- Usuarios locales ---" $cAccent2
    Get-LocalUser | ForEach-Object {
        $color = if ($_.Enabled) { $cGreen } else { $cSubText }
        Write-Out "$($_.Name) — $( if($_.Enabled){'Activo'}else{'Desactivado'} ) — Último acceso: $($_.LastLogon)" $color
    }
})
$tabSecurity.Controls.Add($btnListUsers)

$btnBadDevices = New-CorporateButton "  Dispositivos con Error" 220 165 220 36
$btnBadDevices.Add_Click({
    Write-Out "--- Dispositivos con problema ---" $cAccent2
    $devs = Get-PnpDevice -Status Error,Unknown -EA SilentlyContinue
    if ($devs) { $devs | ForEach-Object { Write-Out "$($_.Class): $($_.FriendlyName) — $($_.Status)" $cRed } }
    else { Write-Out "No se encontraron dispositivos con error." $cGreen }
})
$tabSecurity.Controls.Add($btnBadDevices)

$btnDevMgr = New-CorporateButton "  Administrador de Dispositivos" 450 165 240 36
$btnDevMgr.Add_Click({ Start-Process devmgmt.msc })
$tabSecurity.Controls.Add($btnDevMgr)

New-SectionLabel " Certificados y contraseñas" 10 220 $tabSecurity

$btnCerts = New-CorporateButton "  Certificados Caducados" 10 245 220 36
$btnCerts.Add_Click({
    Write-Out "--- Certificados próximos a vencer o caducados ---" $cAccent2
    $hoy = Get-Date
    Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.NotAfter -lt $hoy.AddDays(30) } | ForEach-Object {
        $color = if ($_.NotAfter -lt $hoy) { $cRed } else { $cYellow }
        Write-Out "$($_.Subject) — Vence: $($_.NotAfter.ToString('dd/MM/yyyy'))" $color
    }
    Write-Out "Revisión completada." $cGreen
})
$tabSecurity.Controls.Add($btnCerts)

$btnPolicies = New-CorporateButton "  Ver Políticas de Seguridad" 240 245 220 36
$btnPolicies.Add_Click({ Start-Process secpol.msc })
$tabSecurity.Controls.Add($btnPolicies)

$btnUAC = New-CorporateButton "  Configurar UAC" 470 245 180 36
$btnUAC.Add_Click({ Start-Process UserAccountControlSettings.exe })
$tabSecurity.Controls.Add($btnUAC)

# ============================================================
#   TAB 6: BACKUP
# ============================================================
New-SectionLabel " Backup de archivos personales" 10 10 $tabBackup

$lblBackupDest = New-Object Windows.Forms.Label
$lblBackupDest.Text = "Destino: $env:USERPROFILE\Desktop"
$lblBackupDest.Location = New-Object Drawing.Point(10, 35)
$lblBackupDest.Size = New-Object Drawing.Size(500, 20)
$lblBackupDest.ForeColor = $cSubText
$lblBackupDest.Font = New-Object Drawing.Font("Consolas", 8)
$tabBackup.Controls.Add($lblBackupDest)

$btnBackupDest = New-CorporateButton "  Cambiar Destino" 520 30 180 28
$btnBackupDest.Add_Click({
    $dlg = New-Object Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Selecciona la carpeta de destino para el backup"
    if ($dlg.ShowDialog() -eq "OK") { $lblBackupDest.Text = "Destino: $($dlg.SelectedPath)" }
})
$tabBackup.Controls.Add($btnBackupDest)

$backupFolders = @(
    @{name=" Documentos";  path="$env:USERPROFILE\Documents";  checked=$true},
    @{name=" Escritorio";  path="$env:USERPROFILE\Desktop";    checked=$true},
    @{name=" Descargas";   path="$env:USERPROFILE\Downloads";  checked=$false},
    @{name=" Imágenes";    path="$env:USERPROFILE\Pictures";   checked=$false},
    @{name=" Videos";      path="$env:USERPROFILE\Videos";     checked=$false},
    @{name=" Música";      path="$env:USERPROFILE\Music";      checked=$false}
)

$cbFolders = @(); $xF = 10; $yF = 65
foreach ($bf in $backupFolders) {
    $cb = New-Object Windows.Forms.CheckBox; $cb.Text = $bf.name; $cb.Checked = $bf.checked
    $cb.Location = New-Object Drawing.Point($xF, $yF); $cb.Size = New-Object Drawing.Size(160, 22)
    $cb.ForeColor = $cText; $cb.BackColor = $cBg; $cb.Tag = $bf.path; $tabBackup.Controls.Add($cb)
    $cbFolders += $cb; $xF += 170; if ($xF -gt 680) { $xF = 10; $yF += 26 }
}

$btnBackup = New-CorporateButton "  Crear Backup ZIP" 10 120 200 36
$btnBackup.BackColor = [Drawing.Color]::FromArgb(0, 100, 60)
$btnBackup.Add_Click({
    $dest = ($lblBackupDest.Text -replace "^Destino: ","").Trim()
    $sel = $cbFolders | Where-Object { $_.Checked }
    if ($sel.Count -eq 0) { Write-Out "Selecciona al menos una carpeta." $cYellow; return }
    $zipName = "Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').zip"
    $zipPath = Join-Path $dest $zipName
    Write-Out "Creando backup: $zipPath" $cSubText
    Add-Type -Assembly System.IO.Compression.FileSystem
    $tmp = "$env:TEMP\syscodi_backup_tmp"
    Remove-Item $tmp -Recurse -Force -EA SilentlyContinue
    New-Item $tmp -ItemType Directory | Out-Null
    foreach ($cb in $sel) {
        if (Test-Path $cb.Tag) {
            $folderName = Split-Path $cb.Tag -Leaf
            Copy-Item $cb.Tag "$tmp\$folderName" -Recurse -Force -EA SilentlyContinue
            Write-Out "Copiado: $($cb.Tag)" $cSubText
        }
    }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tmp, $zipPath)
    Remove-Item $tmp -Recurse -Force -EA SilentlyContinue
    Write-Out "Backup creado: $zipPath" $cGreen
})
$tabBackup.Controls.Add($btnBackup)

New-SectionLabel " Backup de drivers y registro" 10 175 $tabBackup

$btnExportDrivers = New-CorporateButton "  Exportar Drivers Instalados" 10 200 230 36
$btnExportDrivers.Add_Click({
    $dlg = New-Object Windows.Forms.FolderBrowserDialog; $dlg.Description = "Carpeta destino para drivers"
    if ($dlg.ShowDialog() -eq "OK") {
        Write-Out "Exportando drivers a: $($dlg.SelectedPath)" $cSubText
        Start-Process powershell -ArgumentList "-NoProfile -Command `"pnputil /export-driver * '$($dlg.SelectedPath)'`"" -Verb RunAs -Wait
        Write-Out "Drivers exportados." $cGreen
    }
})
$tabBackup.Controls.Add($btnExportDrivers)

$btnExportReg = New-CorporateButton "  Exportar Registro (HKCU)" 250 200 230 36
$btnExportReg.Add_Click({
    $dlg = New-Object Windows.Forms.SaveFileDialog; $dlg.Filter = "Registry (*.reg)|*.reg"; $dlg.FileName = "HKCU_Backup_$(Get-Date -Format 'yyyyMMdd').reg"
    if ($dlg.ShowDialog() -eq "OK") {
        Run-Cmd "reg export HKCU `"$($dlg.FileName)`" /y"
        Write-Out "Registro exportado: $($dlg.FileName)" $cGreen
    }
})
$tabBackup.Controls.Add($btnExportReg)

$btnWinBackup = New-CorporateButton "  Abrir Copia de Seguridad Windows" 490 200 230 36
$btnWinBackup.Add_Click({ Start-Process "control /name Microsoft.BackupAndRestore" })
$tabBackup.Controls.Add($btnWinBackup)

# ============================================================
#   TAB 7: SISTEMA
# ============================================================
$infoBox = New-Object Windows.Forms.RichTextBox
$infoBox.Location = New-Object Drawing.Point(5, 5)
$infoBox.Size = New-Object Drawing.Size(700, 320)
$infoBox.BackColor = $cOutput
$infoBox.ForeColor = $cAccent2
$infoBox.Font = New-Object Drawing.Font("Consolas", 9)
$infoBox.ReadOnly = $true
$infoBox.BorderStyle = "None"
$tabInfo.Controls.Add($infoBox)

# Monitor en tiempo real
$lblMonitor = New-Object Windows.Forms.Label
$lblMonitor.Location = New-Object Drawing.Point(5, 330)
$lblMonitor.Size = New-Object Drawing.Size(700, 22)
$lblMonitor.ForeColor = $cAccent2
$lblMonitor.Font = New-Object Drawing.Font("Consolas", 9)
$lblMonitor.Text = "  CPU: --%   RAM: -- GB libres   Disco C: -- GB libres"
$tabInfo.Controls.Add($lblMonitor)

$timerMonitor = New-Object Windows.Forms.Timer
$timerMonitor.Interval = 2000
$timerMonitor.Add_Tick({
    try {
        $os   = Get-CimInstance Win32_OperatingSystem -EA SilentlyContinue
        $cpu  = (Get-CimInstance Win32_Processor -EA SilentlyContinue).LoadPercentage
        $ramFree = [math]::Round($os.FreePhysicalMemory / 1MB, 1)
        $disk = Get-PSDrive C -EA SilentlyContinue
        $diskFree = [math]::Round($disk.Free / 1GB, 1)
        $cpuColor = if ($cpu -gt 80) { "color:red" } else { "color:lime" }
        $lblMonitor.Text = "  CPU: $cpu%   RAM libre: $ramFree GB   Disco C libre: $diskFree GB"
        $lblMonitor.ForeColor = if ($cpu -gt 80) { $cRed } elseif ($cpu -gt 50) { $cYellow } else { $cGreen }
    } catch {}
})

$btnInfo = New-CorporateButton "  Cargar Info del Sistema" 5 360 220 36
$btnInfo.Add_Click({
    $infoBox.Clear()
    $os  = Get-CimInstance Win32_OperatingSystem
    $cpu = Get-CimInstance Win32_Processor
    $mem = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
    $free= [math]::Round($os.FreePhysicalMemory / 1MB, 2)
    $disk= Get-PSDrive C
    $infoBox.AppendText("Sistema Operativo  : $($os.Caption)`r`n")
    $infoBox.AppendText("Versión            : $($os.Version)`r`n")
    $infoBox.AppendText("Arquitectura       : $($os.OSArchitecture)`r`n")
    $infoBox.AppendText("Procesador         : $($cpu.Name)`r`n")
    $infoBox.AppendText("Núcleos            : $($cpu.NumberOfCores) núcleos / $($cpu.NumberOfLogicalProcessors) lógicos`r`n")
    $infoBox.AppendText("RAM Total          : $mem GB`r`n")
    $infoBox.AppendText("RAM Libre          : $free GB`r`n")
    $infoBox.AppendText("Disco C: Libre     : $([math]::Round($disk.Free/1GB,2)) GB de $([math]::Round(($disk.Used+$disk.Free)/1GB,2)) GB`r`n")
    $infoBox.AppendText("Nombre del equipo  : $env:COMPUTERNAME`r`n")
    $infoBox.AppendText("Usuario actual     : $env:USERNAME`r`n")
    $infoBox.AppendText("Ejecutando como Admin: $isAdmin`r`n")
    $timerMonitor.Start()
    Write-Out "Información del sistema cargada. Monitor activo." $cGreen
})
$tabInfo.Controls.Add($btnInfo)

$btnUptime = New-CorporateButton "  Ver Uptime" 235 360 160 36
$btnUptime.Add_Click({
    $boot = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
    $up   = (Get-Date) - $boot
    Write-Out "Uptime: $($up.Days)d $($up.Hours)h $($up.Minutes)m desde $($boot.ToString('dd/MM/yyyy HH:mm'))" $cText
})
$tabInfo.Controls.Add($btnUptime)

$btnUpdates = New-CorporateButton "  Buscar Actualizaciones" 405 360 200 36
$btnUpdates.Add_Click({ Start-Process ms-settings:windowsupdate })
$tabInfo.Controls.Add($btnUpdates)

$btnExportReport = New-CorporateButton "  Exportar Reporte" 615 360 100 36
$btnExportReport.Add_Click({
    $dlg = New-Object Windows.Forms.SaveFileDialog
    $dlg.Filter = "Text (*.txt)|*.txt"
    $dlg.FileName = "Reporte_Sistema_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    if ($dlg.ShowDialog() -eq "OK") {
        $infoBox.Text | Set-Content $dlg.FileName -Encoding UTF8
        Write-Out "Reporte guardado: $($dlg.FileName)" $cGreen
    }
})
$tabInfo.Controls.Add($btnExportReport)

# ============================================================
#   FOOTER
# ============================================================
$footer = New-Object Windows.Forms.Label
$footer.Text = "SysCodi WinTool Pro v2  |  Usa WinGet como gestor de paquetes  |  Ejecutar siempre como Administrador"
$footer.Location = New-Object Drawing.Point(0, 645)
$footer.Size = New-Object Drawing.Size(1200, 20)
$footer.TextAlign = "MiddleCenter"
$footer.ForeColor = $cSubText
$footer.Font = New-Object Drawing.Font("Segoe UI", 7)
$form.Controls.Add($footer)

# Limpiar timer al cerrar
$form.Add_FormClosing({ $timerMonitor.Stop() })

# ============================================================
$form.ShowDialog()
