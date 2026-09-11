<#
NEPLAN-Projektdatei exportieren.

Ein Fenster: Projektdatei(en) auswaehlen, Exportieren. Je Projektdatei entsteht
eine Datenbank <Ort>.mdb mit allen Tabellen und Spalten des NEPLAN-Datenbank-
exports (dazu A_AENDERUNGEN), in C:\Users\<name>\NEPLAN-Export. Eine vorhandene
Datenbank desselben Orts wird ersetzt. NEPLAN wird nicht gebraucht.

Unten rechts stehen Version und Update-Knopf. Beim Start wird im Hintergrund
auf GitHub nach einer neueren Version gesehen (aktualisieren.ps1).

Mit -Still laeuft alles ohne Fenster (fuer Tests und spaeter den SM).
#>
param(
    [string[]]$Dateien,     # vorbelegte Auswahl
    [string]$Ziel,          # ohne Angabe: C:\Users\<name>\NEPLAN-Export
    [switch]$Still,         # ohne Fenster, Ergebnis als Text
    [switch]$Probelauf      # Fenster exportiert die vorbelegte Auswahl und schliesst sich, fuer Tests
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$Titel   = 'NEPLAN-Projektdatei exportieren'
$Basis   = Split-Path -Parent $PSScriptRoot
$Vorlage = Join-Path $Basis 'Vorlage\Neplan-DB_Leer.mdb'
$Starter = Join-Path $Basis 'NEPLAN-Projektdatei exportieren.cmd'
$Updater = Join-Path $PSScriptRoot 'aktualisieren.ps1'
$Merker  = Join-Path $env:APPDATA 'NEPLAN-Export\letzter_ordner.txt'
if (-not $Ziel) { $Ziel = Join-Path $env:USERPROFILE 'NEPLAN-Export' }


function Fehlertext($fehler) {
    $x = $fehler.Exception
    while ($x.InnerException) { $x = $x.InnerException }
    $x.Message
}

function Ist-Gesperrt([string]$pfad) {
    if (-not (Test-Path -LiteralPath $pfad)) { return $false }
    try { $s = [IO.File]::Open($pfad, 'Open', 'ReadWrite', 'None'); $s.Close(); $false } catch { $true }
}

# Liest jede Projektdatei und schreibt sie in <Ziel>\<Ort>.mdb
function Exportiere([string[]]$dateien) {
    if ([Environment]::Is64BitProcess) { throw 'Bitte über "NEPLAN-Projektdatei exportieren.cmd" starten, nicht direkt.' }
    if (-not ('NeplanLeser.Vollexport' -as [type])) {
        $quellen = 'NepprjLeser.cs', 'MdbSchreiber.cs', 'Regeln.cs', 'Vollexport.cs' | ForEach-Object { Join-Path $PSScriptRoot $_ }
        Add-Type -Path $quellen -ReferencedAssemblies System.Data, System.Xml
    }
    if (-not (Test-Path -LiteralPath $Ziel)) { New-Item -ItemType Directory -Path $Ziel -Force | Out-Null }
    $zeilen = @()
    $fertig = @()
    foreach ($prj in $dateien) {
        $ort = [IO.Path]::GetFileNameWithoutExtension($prj)
        $mdb = Join-Path $Ziel "$ort.mdb"
        try {
            if (Ist-Gesperrt $mdb) { throw "$ort.mdb ist noch geöffnet, im SM oder in Access? Bitte schließen und noch einmal exportieren." }
            $v = [NeplanLeser.Vollexport]::Lies($prj)
            Copy-Item -LiteralPath $Vorlage -Destination $mdb -Force
            (Get-Item -LiteralPath $mdb).Attributes = 'Normal'
            Unblock-File -LiteralPath $mdb    # kein Herkunftsvermerk "aus dem Internet", sonst warnt Office
            $anz = $v.Schreibe($mdb)
            $n = @{}
            foreach ($t in 'BUSBAR', 'LINE', 'DISC_SWITCH') { $n[$t] = if ($anz.ContainsKey($t)) { $anz[$t] } else { 0 } }
            $zeile = '{0}: {1} Stationen, {2} Leitungen, {3} Trennschalter' -f $ort, $n.BUSBAR, $n.LINE, $n.DISC_SWITCH
            if ($v.Planstand) { $zeile += ", Planstand $($v.Planstand.ToString('dd.MM.yyyy'))" }
            $zeilen += $zeile
            foreach ($h in $v.Hinweise) { $zeilen += "   Hinweis: $h" }
            $fertig += $mdb
        } catch {
            $zeilen += "${ort}: nicht exportiert. $(Fehlertext $_)"
        }
    }
    # Datenbanken sofort freigeben, damit man sie oeffnen kann, waehrend das Fenster noch offen ist
    [System.Data.OleDb.OleDbConnection]::ReleaseObjectPool()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $kopf = if ($fertig.Count -eq 1) { "Fertig, 1 Datenbank in $Ziel" } else { "Fertig, $($fertig.Count) Datenbanken in $Ziel" }
    [pscustomobject]@{ Text = (@($kopf, '') + $zeilen) -join "`r`n"; Fertig = $fertig }
}


if ($Still) {
    (Exportiere $Dateien).Text
    return
}


# --------------------------------------------------------------------------- #
# Fenster
# --------------------------------------------------------------------------- #
[Windows.Forms.Application]::EnableVisualStyles()
try {
    . $Updater
    $script:auswahl = @()
    $script:ersteDatei = $null
    $script:exportLaeuft = $false

    $form = New-Object Windows.Forms.Form
    $form.Text = $Titel
    $form.Size = New-Object Drawing.Size(560, 440)
    $form.MinimumSize = New-Object Drawing.Size(420, 340)
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object Drawing.Font('Segoe UI', 9)
    $form.AllowDrop = $true

    $tab = New-Object Windows.Forms.TableLayoutPanel
    $tab.Dock = 'Fill'
    $tab.Padding = New-Object Windows.Forms.Padding(10)
    $tab.ColumnCount = 1
    $tab.RowCount = 5
    foreach ($s in 'AutoSize', 'Percent', 'AutoSize', 'Percent', 'AutoSize') {
        $rs = New-Object Windows.Forms.RowStyle($s)
        if ($s -eq 'Percent') { $rs.Height = 50 }
        [void]$tab.RowStyles.Add($rs)
    }
    $form.Controls.Add($tab)

    $btnWahl = New-Object Windows.Forms.Button
    $btnWahl.Text = 'Projektdateien auswählen …'
    $btnWahl.Dock = 'Fill'; $btnWahl.Height = 34
    $tab.Controls.Add($btnWahl, 0, 0)

    $lst = New-Object Windows.Forms.ListBox
    $lst.Dock = 'Fill'; $lst.IntegralHeight = $false; $lst.HorizontalScrollbar = $true
    $lst.AllowDrop = $true
    $tab.Controls.Add($lst, 0, 1)

    $btnExport = New-Object Windows.Forms.Button
    $btnExport.Text = 'Exportieren'
    $btnExport.Dock = 'Fill'; $btnExport.Height = 42
    $btnExport.Font = New-Object Drawing.Font('Segoe UI', 10, [Drawing.FontStyle]::Bold)
    $btnExport.Enabled = $false
    $btnExport.Margin = New-Object Windows.Forms.Padding(0, 8, 0, 8)
    $tab.Controls.Add($btnExport, 0, 2)

    $aus = New-Object Windows.Forms.TextBox
    $aus.Multiline = $true; $aus.ReadOnly = $true; $aus.ScrollBars = 'Vertical'; $aus.Dock = 'Fill'
    $aus.Text = "Eine oder mehrere .nepprj auswählen (oder hier hineinziehen) und Exportieren klicken.`r`nDie Datenbanken landen in $Ziel."
    $tab.Controls.Add($aus, 0, 3)

    # Fusszeile: links "Ordner oeffnen", rechts Version und Update-Knopf
    $fuss = New-Object Windows.Forms.TableLayoutPanel
    $fuss.Dock = 'Fill'; $fuss.AutoSize = $true
    $fuss.ColumnCount = 3; $fuss.RowCount = 1
    $fuss.Margin = New-Object Windows.Forms.Padding(0, 4, 0, 0)
    [void]$fuss.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle -ArgumentList ([Windows.Forms.SizeType]::Percent), 100))
    [void]$fuss.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle -ArgumentList ([Windows.Forms.SizeType]::AutoSize)))
    [void]$fuss.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle -ArgumentList ([Windows.Forms.SizeType]::AutoSize)))
    $tab.Controls.Add($fuss, 0, 4)

    $lnk = New-Object Windows.Forms.LinkLabel
    $lnk.Text = 'Ordner öffnen'
    $lnk.AutoSize = $true; $lnk.Anchor = 'Left'
    $fuss.Controls.Add($lnk, 0, 0)

    $lblVer = New-Object Windows.Forms.Label
    $lblVer.AutoSize = $true; $lblVer.Anchor = 'Right'
    $lblVer.ForeColor = [Drawing.Color]::Gray
    $lblVer.Margin = New-Object Windows.Forms.Padding(0, 0, 6, 0)
    $fuss.Controls.Add($lblVer, 1, 0)

    $btnUpd = New-Object Windows.Forms.Button
    $btnUpd.Text = 'Nach Updates suchen'
    $btnUpd.AutoSize = $true; $btnUpd.Anchor = 'Right'
    $btnUpd.Margin = New-Object Windows.Forms.Padding(0)
    $fuss.Controls.Add($btnUpd, 2, 0)
    $tipp = New-Object Windows.Forms.ToolTip

    function Setze-Auswahl([string[]]$neu) {
        $script:auswahl = @($neu | Where-Object { [IO.Path]::GetExtension($_) -eq '.nepprj' })
        $lst.Items.Clear()
        foreach ($f in $script:auswahl) { [void]$lst.Items.Add("$([IO.Path]::GetFileName($f))    ($(Split-Path -Parent $f))") }
        $btnExport.Enabled = $script:auswahl.Count -gt 0
        $btnExport.Text = if ($script:auswahl.Count -gt 1) { "$($script:auswahl.Count) Projektdateien exportieren" } else { 'Exportieren' }
        if ($script:auswahl.Count -eq 0 -and $neu) { $aus.Text = 'Das waren keine .nepprj-Dateien.' }
    }

    $btnWahl.add_Click({
        $d = New-Object Windows.Forms.OpenFileDialog
        $d.Title = 'NEPLAN-Projektdatei(en) auswählen, mit Strg oder Umschalt auch mehrere'
        $d.Filter = 'NEPLAN-Projekt (*.nepprj)|*.nepprj'
        $d.Multiselect = $true
        if (Test-Path -LiteralPath $Merker) {
            $letzter = [IO.File]::ReadAllText($Merker, [Text.Encoding]::UTF8).Trim()
            if ($letzter -and (Test-Path -LiteralPath $letzter)) { $d.InitialDirectory = $letzter }
        }
        if ($d.ShowDialog($form) -ne 'OK') { return }
        Setze-Auswahl $d.FileNames
        try {
            $ordner = Split-Path -Parent $Merker
            if (-not (Test-Path -LiteralPath $ordner)) { New-Item -ItemType Directory -Path $ordner -Force | Out-Null }
            [IO.File]::WriteAllText($Merker, (Split-Path -Parent $d.FileNames[0]), (New-Object Text.UTF8Encoding $true))
        } catch { }
    })

    $ziehen = {
        param($s, $a)
        if ($a.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) { $a.Effect = 'Copy' }
    }
    $fallen = {
        param($s, $a)
        Setze-Auswahl ([string[]]$a.Data.GetData([Windows.Forms.DataFormats]::FileDrop))
    }
    $form.add_DragEnter($ziehen); $lst.add_DragEnter($ziehen)
    $form.add_DragDrop($fallen); $lst.add_DragDrop($fallen)

    $btnExport.add_Click({
        $form.Cursor = 'WaitCursor'
        $btnExport.Enabled = $false
        $script:exportLaeuft = $true
        $btnUpd.Enabled = $false          # waehrend des Exports keine Dateien tauschen
        $aus.Text = 'Wird exportiert …'
        $form.Refresh()
        try {
            $r = Exportiere $script:auswahl
            $aus.Text = $r.Text
            if ($r.Fertig.Count -gt 0) { $script:ersteDatei = $r.Fertig[0] }
        } catch {
            $aus.Text = "Das hat nicht geklappt:`r`n$(Fehlertext $_)"
        } finally {
            $script:exportLaeuft = $false
            $btnExport.Enabled = $script:auswahl.Count -gt 0
            $btnUpd.Enabled = -not $script:suche
            $form.Cursor = 'Default'
        }
    })

    # Immer sichtbar: nach einem Export mit der neuen Datenbank markiert, sonst einfach der Ordner
    $lnk.add_LinkClicked({
        if ($script:ersteDatei -and (Test-Path -LiteralPath $script:ersteDatei)) {
            Start-Process explorer.exe "/select,`"$($script:ersteDatei)`""
        } else {
            if (-not (Test-Path -LiteralPath $Ziel)) { New-Item -ItemType Directory -Path $Ziel -Force | Out-Null }
            Start-Process explorer.exe "`"$Ziel`""
        }
    })

    # ---------------------------------------------------------------- Update
    $script:lokal = Lokale-Version $Basis
    $script:neu = $null
    $script:suche = $null
    function Versionstext { if ($script:lokal) { "Version $($script:lokal)" } else { 'Version unbekannt' } }
    $lblVer.Text = Versionstext

    # Die Abfrage laeuft in einem eigenen Runspace, das Fenster bleibt dabei bedienbar
    function Starte-Suche([bool]$vonHand) {
        if ($script:suche) { return }
        $btnUpd.Enabled = $false
        if ($vonHand) { $lblVer.ForeColor = [Drawing.Color]::Gray; $lblVer.Text = 'Suche nach Updates …' }
        $ps = [powershell]::Create()
        [void]$ps.AddScript('param($skript, $basis) $ErrorActionPreference = "Stop"; . $skript; Suche-Update $basis')
        [void]$ps.AddArgument($Updater)
        [void]$ps.AddArgument($Basis)
        $script:suche = [pscustomobject]@{ PS = $ps; Lauf = $ps.BeginInvoke(); VonHand = $vonHand }
        $uhr.Start()
    }

    $uhr = New-Object Windows.Forms.Timer
    $uhr.Interval = 250
    $uhr.add_Tick({
        $s = $script:suche
        if (-not $s -or -not $s.Lauf.IsCompleted) { return }
        $uhr.Stop()
        $script:suche = $null
        try {
            $r = @($s.PS.EndInvoke($s.Lauf)) | Select-Object -First 1
            if (-not $r) { throw 'Keine Antwort von GitHub.' }
            $lblVer.ForeColor = [Drawing.Color]::Gray
            $tipp.SetToolTip($lblVer, '')
            if ($r.Neuer) {
                $script:neu = $r
                $btnUpd.Text = "Update auf $($r.Neu)"
                $btnUpd.FlatStyle = 'Flat'
                $btnUpd.BackColor = [Drawing.Color]::FromArgb(0, 103, 184)
                $btnUpd.ForeColor = [Drawing.Color]::White
                $lblVer.Text = Versionstext
                $tipp.SetToolTip($btnUpd, $(if ($r.Hinweis) { "Neu in $($r.Neu): $($r.Hinweis)" } else { "Version $($r.Neu) ist da." }))
            } else {
                $lblVer.Text = "$(Versionstext), aktuell"
            }
        } catch {
            $grund = Fehlertext $_
            if ($s.VonHand) { $lblVer.ForeColor = [Drawing.Color]::Firebrick; $lblVer.Text = 'GitHub nicht erreichbar' }
            else { $lblVer.Text = Versionstext }
            $tipp.SetToolTip($lblVer, "Updateprüfung nicht möglich: $grund")
        } finally {
            $s.PS.Dispose()
            $btnUpd.Enabled = -not $script:exportLaeuft
        }
    })

    $btnUpd.add_Click({
        if (-not $script:neu) { Starte-Suche $true; return }
        $r = $script:neu
        $frage = "Version $($r.Neu) installieren?"
        if ($r.Hinweis) { $frage += "`r`n`r`nNeu: $($r.Hinweis)" }
        $frage += "`r`n`r`nDas Fenster schließt sich dabei und startet neu."
        if ([Windows.Forms.MessageBox]::Show($form, $frage, $Titel, 'OKCancel', 'Question') -ne 'OK') { return }
        $form.Cursor = 'WaitCursor'
        $btnUpd.Enabled = $false
        $btnExport.Enabled = $false
        $lblVer.Text = 'Wird aktualisiert …'
        $form.Refresh()
        try {
            [void](Installiere-Update $Basis $r.Neu)
            Start-Process -FilePath $Starter
            $form.Close()
        } catch {
            $lblVer.Text = Versionstext
            [Windows.Forms.MessageBox]::Show($form, "Das Update hat nicht geklappt, es bleibt bei $(Versionstext).`r`n`r`n$(Fehlertext $_)", $Titel, 'OK', 'Warning') | Out-Null
            $btnUpd.Enabled = $true
            $btnExport.Enabled = $script:auswahl.Count -gt 0
        } finally {
            $form.Cursor = 'Default'
        }
    })

    $form.add_Shown({ Starte-Suche $false })

    if ($Dateien) { Setze-Auswahl $Dateien }
    if ($Probelauf) {
        # Fuer Tests: erst die Updatepruefung abwarten, dann exportieren und schliessen
        $timer = New-Object Windows.Forms.Timer
        $timer.Interval = 1000
        $timer.add_Tick({ if ($script:suche) { return }; $timer.Stop(); if ($btnExport.Enabled) { $btnExport.PerformClick() }; $form.Close() })
        $timer.Start()
    }
    [void]$form.ShowDialog()
    if ($Probelauf) {
        "Knopf: '$($btnExport.Text)', Link auf: $($script:ersteDatei)"
        "Fußzeile: '$($lblVer.Text)' | Update-Knopf: '$($btnUpd.Text)' | Hinweis: '$($tipp.GetToolTip($lblVer))$($tipp.GetToolTip($btnUpd))'"
        $aus.Text
    }
} catch {
    [Windows.Forms.MessageBox]::Show("Das Werkzeug ist auf einen Fehler gestoßen:`r`n`r`n$(Fehlertext $_)", $Titel, 'OK', 'Error') | Out-Null
    if ($Probelauf) { throw }
}
