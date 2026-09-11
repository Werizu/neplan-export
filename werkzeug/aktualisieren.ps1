<#
Automatische Aktualisierung ueber GitHub, fuer das Fenster in exportieren.ps1.

Die installierte Version steht in manifest.json im Programmordner. Nachgesehen
wird in der manifest.json im Zweig main des oeffentlichen Repos, ueber
raw.githubusercontent.com und bewusst NICHT ueber die GitHub-API: die zaehlt
ohne Anmeldung 60 Anfragen je Stunde und Adresse, und der Firmenproxy geht ueber
wechselnde Ausgangsadressen, dann bliebe das Update still aus.
Geholt wird von der Versionsmarke v<Version>, nicht von main, denn main kann
bei raw bis zu fuenf Minuten alt sein. Erst werden alle geaenderten Dateien
geladen und per SHA-256 geprueft, dann erst getauscht. Scheitert das Tauschen,
wird zurueckgerollt und es bleibt bei der alten Version.
#>

$UpdateRepo = 'Werizu/neplan-export'
$UpdateRoh = "https://raw.githubusercontent.com/$UpdateRepo"
if ($env:NEPLAN_EXPORT_UPDATEQUELLE) { $UpdateRoh = $env:NEPLAN_EXPORT_UPDATEQUELLE.TrimEnd('/') }   # nur fuer Tests
$UpdateDateiarten = '.ps1', '.cs', '.cmd', '.txt', '.mdb'


function Hole-Bytes([string]$url, [int]$zeit = 15000) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $rq = [Net.HttpWebRequest]::Create($url)
    $rq.Timeout = $zeit
    $rq.ReadWriteTimeout = $zeit
    $rq.UserAgent = 'NEPLAN-Export-Update'
    $rq.Proxy = [Net.WebRequest]::GetSystemWebProxy()
    $rq.Proxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
    $rq.Headers['Cache-Control'] = 'no-cache'
    $rs = $rq.GetResponse()
    try {
        $ms = New-Object IO.MemoryStream
        $rs.GetResponseStream().CopyTo($ms)
        , $ms.ToArray()
    } finally {
        $rs.Close()
    }
}

function Text-Aus([byte[]]$b) { [Text.Encoding]::UTF8.GetString($b).TrimStart([char]0xFEFF) }

function Pruefsumme([byte[]]$b) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { -join ($sha.ComputeHash($b) | ForEach-Object { $_.ToString('x2') }) } finally { $sha.Dispose() }
}

function Lokale-Version([string]$basis) {
    $m = Join-Path $basis 'manifest.json'
    if (-not (Test-Path -LiteralPath $m)) { return $null }
    try { [version]((Text-Aus ([IO.File]::ReadAllBytes($m)) | ConvertFrom-Json).version) } catch { $null }
}

# Fragt die Version auf GitHub ab. Neuer ist wahr, wenn sie hoeher ist als die installierte.
function Suche-Update([string]$basis) {
    $fern = Text-Aus (Hole-Bytes "$UpdateRoh/main/manifest.json" 8000) | ConvertFrom-Json
    $neu = [version]$fern.version
    $alt = Lokale-Version $basis
    [pscustomobject]@{ Neu = $neu; Alt = $alt; Hinweis = "$($fern.hinweis)"; Neuer = (($null -eq $alt) -or ($neu -gt $alt)) }
}

# Das Manifest kommt aus dem Netz, und was ankommt, wird beim naechsten Start ausgefuehrt:
# nur relative Pfade innerhalb des Programmordners und nur die vorgesehenen Dateiarten.
function Pruefe-Pfad([string]$basis, [string]$pfad) {
    if (-not $pfad -or $pfad.Contains('..') -or $pfad.Contains(':') -or $pfad.StartsWith('/') -or $pfad.StartsWith('\')) {
        throw "Unzulässiger Pfad im Manifest: $pfad"
    }
    if ($UpdateDateiarten -notcontains [IO.Path]::GetExtension($pfad).ToLowerInvariant()) { throw "Unzulässige Dateiart im Manifest: $pfad" }
    $wurzel = [IO.Path]::GetFullPath($basis).TrimEnd('\') + '\'
    $voll = [IO.Path]::GetFullPath((Join-Path $wurzel ($pfad -replace '/', '\')))
    if (-not $voll.StartsWith($wurzel, [StringComparison]::OrdinalIgnoreCase)) { throw "Pfad außerhalb des Programmordners: $pfad" }
    $voll
}

# Installiert die Version von ihrer Versionsmarke. Liefert die Zahl der getauschten Dateien.
function Installiere-Update([string]$basis, [version]$version) {
    $marke = "$UpdateRoh/v$version"
    $manBytes = Hole-Bytes "$marke/manifest.json"
    $man = Text-Aus $manBytes | ConvertFrom-Json
    if ([version]$man.version -ne $version) { throw "Das Manifest der Marke v$version nennt Version $($man.version)." }

    # 1. Alles laden und pruefen, nichts veraendern
    $plan = @()
    foreach ($d in @($man.dateien)) {
        $ziel = Pruefe-Pfad $basis $d.pfad
        if ((Test-Path -LiteralPath $ziel) -and (Pruefsumme ([IO.File]::ReadAllBytes($ziel))) -eq "$($d.sha256)".ToLowerInvariant()) { continue }
        $url = $marke + '/' + ((($d.pfad -split '/') | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/')
        $b = Hole-Bytes $url 60000
        if ((Pruefsumme $b) -ne "$($d.sha256)".ToLowerInvariant()) { throw "Prüfsumme stimmt nicht: $($d.pfad). Es wurde nichts geändert." }
        $plan += [pscustomobject]@{ Ziel = $ziel; Daten = $b }
    }

    # 2. Neben das Ziel legen (*.neu), dann tauschen; bei Fehler zurueckrollen
    $getauscht = New-Object Collections.ArrayList
    try {
        foreach ($p in $plan) {
            $ordner = Split-Path -Parent $p.Ziel
            if (-not (Test-Path -LiteralPath $ordner)) { New-Item -ItemType Directory -Path $ordner -Force | Out-Null }
            [IO.File]::WriteAllBytes($p.Ziel + '.neu', $p.Daten)
        }
        foreach ($p in $plan) {
            $alt = $p.Ziel + '.alt'
            if (Test-Path -LiteralPath $alt) { Remove-Item -LiteralPath $alt -Force }
            $hatteAlt = Test-Path -LiteralPath $p.Ziel
            if ($hatteAlt) { [IO.File]::Move($p.Ziel, $alt) }
            [void]$getauscht.Add([pscustomobject]@{ Ziel = $p.Ziel; Alt = $(if ($hatteAlt) { $alt } else { $null }) })
            [IO.File]::Move($p.Ziel + '.neu', $p.Ziel)
        }
    } catch {
        foreach ($g in $getauscht) {
            try {
                if (Test-Path -LiteralPath $g.Ziel) { Remove-Item -LiteralPath $g.Ziel -Force }
                if ($g.Alt) { [IO.File]::Move($g.Alt, $g.Ziel) }
            } catch { }
        }
        foreach ($p in $plan) { if (Test-Path -LiteralPath ($p.Ziel + '.neu')) { Remove-Item -LiteralPath ($p.Ziel + '.neu') -Force } }
        throw
    }
    foreach ($g in $getauscht) { if ($g.Alt) { try { Remove-Item -LiteralPath $g.Alt -Force } catch { } } }

    # 3. Zuletzt die Versionsangabe, erst jetzt gilt die neue Version als installiert
    [IO.File]::WriteAllBytes((Join-Path $basis 'manifest.json'), $manBytes)
    $plan.Count
}
