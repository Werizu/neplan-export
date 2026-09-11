<#
Export ohne NEPLAN gegen den NEPLAN-Export pruefen.

Man zieht eine Datenbank, die NEPLAN exportiert hat, auf "Pruefen\Leser gegen
NEPLAN pruefen". Fuer jedes Netz darin sucht das Werkzeug die Projektdatei,
exportiert sie selbst ohne NEPLAN und vergleicht beide Datenbanken Zelle fuer
Zelle ueber alle Tabellen. So laesst sich an jedem Ort pruefen, ob der Export
ohne NEPLAN dasselbe liefert.

Die Projektdatei wird gesucht unter dem Pfad aus der INFOTABLE, sonst mit
gleichem Namen in dem Ordner, aus dem zuletzt exportiert wurde. Die selbst
erzeugte Datenbank bleibt zum Nachsehen im Berichtsordner liegen
(<Netz>_ohne_NEPLAN.mdb).
#>
param(
    [string]$Datenbank,
    [string]$Projektordner,
    [string]$BerichtOrdner
)

$ErrorActionPreference = 'Stop'
$Basis = Split-Path -Parent $PSScriptRoot
$Vorlage = Join-Path $Basis 'Vorlage\Neplan-DB_Leer.mdb'
if (-not $BerichtOrdner) { $BerichtOrdner = Join-Path $env:USERPROFILE 'NEPLAN-Export\Berichte' }

function Fehlertext($fehler) {
    $x = $fehler.Exception
    while ($x.InnerException) { $x = $x.InnerException }
    $x.Message
}

Write-Host ''
Write-Host 'Export ohne NEPLAN gegen den NEPLAN-Export prüfen' -ForegroundColor Cyan

if ([Environment]::Is64BitProcess) { Write-Host 'Bitte über "Leser gegen NEPLAN prüfen.cmd" starten, nicht direkt.' -ForegroundColor Red; return }
if (-not $Datenbank) {
    Add-Type -AssemblyName System.Windows.Forms
    $d = New-Object Windows.Forms.OpenFileDialog
    $d.Title = 'Datenbank wählen, die NEPLAN exportiert hat'
    $d.Filter = 'Access-Datenbank (*.mdb)|*.mdb'
    if ($d.ShowDialog() -ne 'OK') { Write-Host 'Abgebrochen.'; return }
    $Datenbank = $d.FileName
}
if (-not (Test-Path -LiteralPath $Datenbank)) { Write-Host "Die Datenbank gibt es nicht: $Datenbank" -ForegroundColor Red; return }
if (-not $Projektordner) {
    # Der Ordner, aus dem zuletzt Projektdateien exportiert wurden
    $merker = Join-Path $env:APPDATA 'NEPLAN-Export\letzter_ordner.txt'
    if (Test-Path -LiteralPath $merker) { $Projektordner = [IO.File]::ReadAllText($merker, [Text.Encoding]::UTF8).Trim() }
}

try {
    $quellen = 'NepprjLeser.cs', 'MdbSchreiber.cs', 'Regeln.cs', 'Vollexport.cs' | ForEach-Object { Join-Path $PSScriptRoot $_ }
    Add-Type -Path $quellen -ReferencedAssemblies System.Data, System.Xml
    $c = [NeplanLeser.Mdb]::Oeffne($Datenbank, $false)
} catch {
    Write-Host (Fehlertext $_) -ForegroundColor Red
    return
}
try {
    function Tabelle([string]$sql) {
        $a = New-Object Data.OleDb.OleDbDataAdapter $sql, $c
        $t = New-Object Data.DataTable
        [void]$a.Fill($t)
        , $t
    }
    $netze = @((Tabelle 'SELECT DISTINCT NETNAME FROM TOPOLOGY').Rows | ForEach-Object { "$($_.NETNAME)".Trim() } | Where-Object { $_ })
    $info = @()
    if ([NeplanLeser.Mdb]::HatTabelle($c, 'INFOTABLE')) { $info = @((Tabelle 'SELECT * FROM INFOTABLE').Rows) }
} finally {
    $c.Close()
    [NeplanLeser.Mdb]::Freigeben()
}

Write-Host "Datenbank: $Datenbank"
Write-Host "Projektordner: $(if ($Projektordner) { $Projektordner } else { '(keiner eingestellt)' })"
if ($netze.Count -eq 0) { Write-Host 'In dieser Datenbank steht kein Netz.' -ForegroundColor Yellow; return }
$vomLeser = @($info | Where-Object { "$($_.VERSION)" -like 'nepprj-Leser*' })
if ($info.Count -gt 0 -and $vomLeser.Count -eq $info.Count) {
    Write-Host 'Diese Datenbank hat das Werkzeug selbst geschrieben, nicht NEPLAN. Zum Prüfen bitte einen NEPLAN-Export nehmen.' -ForegroundColor Yellow
    return
}
if (-not (Test-Path -LiteralPath $BerichtOrdner)) { New-Item -ItemType Directory -Path $BerichtOrdner -Force | Out-Null }

$bericht = New-Object Collections.Generic.List[string]
$bericht.Add("Export ohne NEPLAN gegen NEPLAN-Export, $(Get-Date -Format 'dd.MM.yyyy HH:mm')")
$bericht.Add("NEPLAN-Datenbank: $Datenbank")
$ergebnis = @()

foreach ($netz in $netze) {
    $kandidaten = @()
    foreach ($r in $info) {
        if ("$($r.NETNAME)".Trim() -ieq $netz -and "$($r.PROJECT)".Trim()) { $kandidaten += "$($r.PROJECT)".Trim() }
    }
    if ($Projektordner) {
        foreach ($k in @($kandidaten)) { $kandidaten += Join-Path $Projektordner ([IO.Path]::GetFileName($k)) }
        $kandidaten += Join-Path $Projektordner "$netz.nepprj"
    }
    $prj = $kandidaten | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1

    Write-Host ''
    Write-Host "== $netz" -ForegroundColor Cyan
    $bericht.Add('')
    $bericht.Add("== $netz")
    if (-not $prj) {
        $t = 'Projektdatei nicht gefunden. Gesucht: ' + (($kandidaten | Select-Object -Unique) -join ', ')
        Write-Host $t -ForegroundColor Yellow
        $bericht.Add($t)
        $ergebnis += [pscustomobject]@{ Netz = $netz; Ergebnis = 'Projektdatei fehlt' }
        continue
    }
    $bericht.Add("Projektdatei: $prj")
    Write-Host "Projektdatei: $prj"
    try {
        $v = [NeplanLeser.Vollexport]::Lies($prj)
        $probe = Join-Path $BerichtOrdner "$($netz)_ohne_NEPLAN.mdb"
        Copy-Item -LiteralPath $Vorlage -Destination $probe -Force
        (Get-Item -LiteralPath $probe).Attributes = 'Normal'
        Unblock-File -LiteralPath $probe
        [void]$v.Schreibe($probe)
        $bericht.Add("Export ohne NEPLAN: $probe")
        $z = [NeplanLeser.Zellvergleich]::Vergleiche($Datenbank, $probe, $netz)
        foreach ($zeile in $z.Zeilen) { Write-Host $zeile; $bericht.Add($zeile) }
        foreach ($h in $v.Hinweise) { Write-Host "Hinweis: $h" -ForegroundColor Yellow; $bericht.Add("Hinweis: $h") }
        $exp = @($info | Where-Object { "$($_.NETNAME)".Trim() -ieq $netz }) | Select-Object -Last 1
        if ($exp -and $exp.EXPORT_TIME -is [datetime] -and (Get-Item -LiteralPath $prj).LastWriteTime -gt $exp.EXPORT_TIME.AddMinutes(1)) {
            $t = 'Achtung: Die Projektdatei wurde nach dem NEPLAN-Export geändert. Abweichungen können daher kommen.'
            Write-Host $t -ForegroundColor Yellow
            $bericht.Add($t)
        }
        if ($z.Fehlt + $z.Falsch + $z.ZeilenZuviel -eq 0) { $urteil = 'stimmt vollständig' }
        else {
            $urteil = '{0:0.0} % gleich, {1} fehlen, {2} falsch' -f (100.0 * $z.Gleich / [math]::Max(1, $z.Zellen)), $z.Fehlt, $z.Falsch
            if ($z.ZeilenZuviel) { $urteil += ", $($z.ZeilenZuviel) Zeilen zu viel" }
        }
        $ergebnis += [pscustomobject]@{ Netz = $netz; Ergebnis = $urteil }
    } catch {
        $t = "Nicht geprüft: $(Fehlertext $_)"
        Write-Host $t -ForegroundColor Red
        $bericht.Add($t)
        $ergebnis += [pscustomobject]@{ Netz = $netz; Ergebnis = 'Fehler beim Lesen' }
    }
}

Write-Host ''
Write-Host 'Zusammenfassung (alle Tabellen, Wert für Wert):' -ForegroundColor Cyan
$bericht.Add('')
$bericht.Add('Zusammenfassung (alle Tabellen, Wert für Wert):')
foreach ($r in $ergebnis) {
    $t = '{0,-16} {1}' -f $r.Netz, $r.Ergebnis
    Write-Host $t -ForegroundColor $(if ($r.Ergebnis -eq 'stimmt vollständig') { 'Green' } else { 'Yellow' })
    $bericht.Add($t)
}
$datei = Join-Path $BerichtOrdner ('Leserpruefung_{0:yyyy-MM-dd_HHmmss}.txt' -f (Get-Date))
[IO.File]::WriteAllLines($datei, [string[]]$bericht, (New-Object Text.UTF8Encoding $true))
Write-Host "Bericht gespeichert: $datei"
