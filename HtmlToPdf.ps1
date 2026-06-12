#Requires -Version 5.1
<#
.SYNOPSIS
    HtmlToPdf.ps1 – erzeugt aus HTML-Dateien PDF-Dokumente.

.DESCRIPTION
    Liest eine oder mehrere HTML-Dateien ein, baut die HTML-Struktur mit der
    Bibliothek PsXml.ps1 vollständig in einen DOM ein (Voll-DOM-Modus),
    fügt dort eine Seiten-Einrichtung (@page-Regel für Format + Ausrichtung,
    optional <meta charset>) ein, serialisiert das Dokument zurück und druckt
    es anschließend mit einer headless Browser-Engine nach PDF.

    Format ("Ax") und Ausrichtung werden VORHER per Parameter festgelegt und
    als explizite Millimeter-Maße in eine @page-Regel geschrieben. Dadurch
    funktionieren alle A-Formate (A0–A8) zuverlässig – Chromium kennt als
    @page-Größennamen sonst nur A3/A4/A5.

    Render-Engine
        Auto-Erkennung in der Reihenfolge Edge -> Chrome -> wkhtmltopdf,
        überschreibbar mit -Engine / -EnginePath.
          * Chromium (Edge/Chrome): Format/Ausrichtung über die eingebettete
            @page-CSS-Regel; gedruckt via --headless --print-to-pdf.
          * wkhtmltopdf: Format/Ausrichtung zusätzlich über die CLI
            (--page-width/--page-height), da @page hier ignoriert wird.

    Voll-DOM-Modus – Voraussetzung
        PsXml ist ein XML-Parser. Die Eingabe muss daher wohlgeformtes
        (X)HTML sein:
          * Leere Elemente selbstschließend: <br/>, <img .../>, <meta .../>.
          * In <script>/<style> bricht ein rohes '<' das Parsen; dort als
            CDATA kapseln (XHTML-Muster: //<![CDATA[ ... //]]>). Zeichen wie
            >, ", ' und & in CSS/JS bleiben dagegen erhalten (sie werden nach
            dem Serialisieren wieder entschärft).
        Bei Parse-Fehlern nennt das Skript Zeile/Spalte aus PsXml.

.PARAMETER Path
    Eine oder mehrere HTML-Dateien, Verzeichnisse (es werden *.html/*.htm
    verarbeitet) oder Platzhalter-Pfade. Pipeline-fähig.

.PARAMETER OutputPath
    Zielpfad. Als Verzeichnis gilt er nur, wenn er bereits als Ordner
    existiert ODER mit '\' bzw. '/' endet (z. B. '.\pdf\'); dann landet je
    Eingabe eine gleichnamige .pdf darin. Andernfalls ist es ein Dateiname –
    fehlt die .pdf-Endung, wird sie ergänzt. Ohne Angabe entsteht die PDF
    neben der Eingabedatei (gleicher Name, .pdf).

.PARAMETER InputEncoding
    Zeichenkodierung der HTML-Eingabe. 'Auto' (Standard) erkennt die
    Kodierung per BOM und nimmt sonst UTF-8 an. Für ältere Dateien ohne BOM
    (z. B. ANSI/Windows-1252 mit Umlauten) 'Latin1' angeben, sonst entstehen
    falsche Zeichen.

.PARAMETER Format
    Papierformat A0–A8 (Standard: A4).

.PARAMETER Orientation
    Ausrichtung Portrait (Hochformat) oder Landscape (Querformat),
    Standard: Portrait.

.PARAMETER Margin
    Seitenrand als CSS-Längenangabe für die @page-Regel (Standard: 10mm).
    Bei wkhtmltopdf wird der Wert auf alle vier Seiten angewandt.

.PARAMETER Engine
    Auto (Standard), Edge, Chrome oder wkhtmltopdf.

.PARAMETER EnginePath
    Vollständiger Pfad zur Engine-EXE; übersteuert die Auto-Erkennung. Die
    Art (Chromium/wkhtmltopdf) wird aus dem Dateinamen abgeleitet.

.PARAMETER TimeoutSeconds
    Höchstlaufzeit eines Druckvorgangs in Sekunden (Standard: 120).

.PARAMETER KeepIntermediateHtml
    Behält die aufbereitete Zwischen-HTML-Datei und gibt ihren Pfad mit aus
    (zur Fehlersuche).

.EXAMPLE
    . .\HtmlToPdf.ps1
    Convert-HtmlToPdf -Path .\bericht.html -Format A4 -Orientation Portrait

.EXAMPLE
    # Querformat A3, ganzes Verzeichnis
    Convert-HtmlToPdf -Path .\html\ -Format A3 -Orientation Landscape -OutputPath .\pdf\

.EXAMPLE
    # Direkter Skriptaufruf (ohne Dot-Sourcing)
    .\HtmlToPdf.ps1 -Path .\rechnung.html -Format A5 -Margin 15mm

.EXAMPLE
    # Pipeline (dot-sourcen und die Funktion verwenden)
    . .\HtmlToPdf.ps1
    Get-ChildItem .\html\*.html | Convert-HtmlToPdf -Format A4 -OutputPath .\pdf\

.NOTES
    Datei kann per Dot-Sourcing (". .\HtmlToPdf.ps1") geladen werden,
    um die Funktion Convert-HtmlToPdf bereitzustellen, oder direkt als Skript
    ausgeführt werden. PsXml.ps1 muss im selben Verzeichnis liegen.
#>
# Hinweis: Path ist auf SKRIPT-Ebene bewusst NICHT verpflichtend, damit das
# Dot-Sourcing (". .\HtmlToPdf.ps1") ohne Argumente funktioniert. Die
# Funktion Convert-HtmlToPdf erzwingt Path weiterhin.
[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('FullName', 'PSPath')]
    [string[]]$Path,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [ValidateSet('A0', 'A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8')]
    [string]$Format = 'A4',

    [Parameter()]
    [ValidateSet('Portrait', 'Landscape')]
    [string]$Orientation = 'Portrait',

    [Parameter()]
    [string]$Margin = '10mm',

    [Parameter()]
    [ValidateSet('Auto', 'UTF8', 'UTF8BOM', 'UTF16LE', 'UTF16BE', 'ASCII', 'Latin1', 'Default')]
    [string]$InputEncoding = 'Auto',

    [Parameter()]
    [ValidateSet('Auto', 'Edge', 'Chrome', 'wkhtmltopdf')]
    [string]$Engine = 'Auto',

    [Parameter()]
    [string]$EnginePath,

    [Parameter()]
    [ValidateRange(1, 3600)]
    [int]$TimeoutSeconds = 120,

    [Parameter()]
    [switch]$KeepIntermediateHtml
)

# PsXml.ps1 muss VOR den eigenen Funktionen geladen werden. Wichtig:
# Im restlichen Skript werden KEINE PsXml-Typliterale (z. B.
# [PsXmlNodeType]) verwendet, da diese sonst schon beim Parsen aufgelöst
# würden – die Klassen stehen aber erst nach diesem Dot-Sourcing bereit.
# Statt dessen werden Strings ('Element', 'Data', ...) per Coercion und die
# mitgelieferten Funktionen (New-PsXmlDocument, ConvertTo-PsXmlString)
# genutzt.
$script:H2P_PsXmlPath = '.\PsXml.ps1'
if (-not (Test-Path -LiteralPath $script:H2P_PsXmlPath)) {
    throw "PsXml.ps1 wurde nicht gefunden (relativ erwartet: $script:H2P_PsXmlPath; bitte aus dem Skript-Ordner starten)."
}
. $script:H2P_PsXmlPath

# ============================================================================
#  Hilfsfunktionen
# ============================================================================

function Get-Html2PdfPageSizeMm {
    <#
    .SYNOPSIS
        Liefert @(BreiteMm, HöheMm) für ein A-Format und eine Ausrichtung.
    #>
    [CmdletBinding()]
    [OutputType([int[]])]
    param(
        [Parameter(Mandatory)][string]$Format,
        [Parameter(Mandatory)][string]$Orientation
    )
    # A-Reihe, Hochformat Breite x Höhe in Millimetern (ISO 216).
    $table = @{
        A0 = @(841, 1189); A1 = @(594, 841); A2 = @(420, 594); A3 = @(297, 420)
        A4 = @(210, 297);  A5 = @(148, 210); A6 = @(105, 148); A7 = @(74, 105)
        A8 = @(52, 74)
    }
    $dim = $table[$Format]
    if ($null -eq $dim) { throw "Unbekanntes Format '$Format'." }
    $w = $dim[0]; $h = $dim[1]
    if ($Orientation -eq 'Landscape') { return @($h, $w) }
    return @($w, $h)
}

function Resolve-Html2PdfEngine {
    <#
    .SYNOPSIS
        Bestimmt die zu nutzende Render-Engine (Pfad + Art).
    .OUTPUTS
        PSCustomObject mit Name, Kind ('Chromium'|'wkhtmltopdf') und Path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Engine,
        [Parameter()][string]$EnginePath
    )

    function New-EngineInfo([string]$name, [string]$kind, [string]$path) {
        [PSCustomObject]@{ Name = $name; Kind = $kind; Path = $path }
    }

    # Kandidatenpfade je Engine.
    $edgePaths = @(
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe')
    )
    $chromePaths = @(
        (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe'),
        (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe')
    )
    $wkPaths = @(
        (Join-Path $env:ProgramFiles 'wkhtmltopdf\bin\wkhtmltopdf.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'wkhtmltopdf\bin\wkhtmltopdf.exe')
    )

    function Find-First([string[]]$paths) {
        foreach ($p in $paths) {
            if (-not [string]::IsNullOrEmpty($p) -and (Test-Path -LiteralPath $p)) { return $p }
        }
        return $null
    }

    # Expliziter Pfad hat Vorrang.
    if ($PSBoundParameters.ContainsKey('EnginePath') -and -not [string]::IsNullOrEmpty($EnginePath)) {
        if (-not (Test-Path -LiteralPath $EnginePath)) {
            throw "EnginePath nicht gefunden: $EnginePath"
        }
        $leaf = [System.IO.Path]::GetFileNameWithoutExtension($EnginePath).ToLowerInvariant()
        if ($leaf -like '*wkhtmltopdf*') { return New-EngineInfo 'wkhtmltopdf' 'wkhtmltopdf' $EnginePath }
        $name = if ($leaf -like '*msedge*') { 'Edge' } elseif ($leaf -like '*chrome*') { 'Chrome' } else { 'Chromium' }
        return New-EngineInfo $name 'Chromium' $EnginePath
    }

    switch ($Engine) {
        'Edge' {
            $p = Find-First $edgePaths
            if ($p) { return New-EngineInfo 'Edge' 'Chromium' $p }
            throw 'Microsoft Edge wurde nicht gefunden.'
        }
        'Chrome' {
            $p = Find-First $chromePaths
            if ($p) { return New-EngineInfo 'Chrome' 'Chromium' $p }
            throw 'Google Chrome wurde nicht gefunden.'
        }
        'wkhtmltopdf' {
            $p = Find-First $wkPaths
            if (-not $p) { $p = (Get-Command 'wkhtmltopdf.exe' -ErrorAction SilentlyContinue).Source }
            if ($p) { return New-EngineInfo 'wkhtmltopdf' 'wkhtmltopdf' $p }
            throw 'wkhtmltopdf wurde nicht gefunden.'
        }
        default {
            # Auto: Edge -> Chrome -> wkhtmltopdf
            $p = Find-First $edgePaths
            if ($p) { return New-EngineInfo 'Edge' 'Chromium' $p }
            $p = Find-First $chromePaths
            if ($p) { return New-EngineInfo 'Chrome' 'Chromium' $p }
            $p = Find-First $wkPaths
            if (-not $p) { $p = (Get-Command 'wkhtmltopdf.exe' -ErrorAction SilentlyContinue).Source }
            if ($p) { return New-EngineInfo 'wkhtmltopdf' 'wkhtmltopdf' $p }
            throw 'Keine Render-Engine gefunden (weder Edge, Chrome noch wkhtmltopdf).'
        }
    }
}

function Set-Html2PdfEmptyElementGuards {
    <#
    .SYNOPSIS
        Hängt an leere Nicht-Void-Elemente ein leeres Data-Kind an, damit der
        XML-Drucker sie als <tag></tag> statt selbstschließend (<tag/>) ausgibt.
    .DESCRIPTION
        PsXmls Drucker schließt jedes kinderlose, wertlose Element selbst.
        Bei HTML ist das nur für Void-Elemente (br, img, meta, ...) korrekt.
        Für alle anderen (z. B. <script src=...></script>, <div></div>,
        <style></style>) würde der HTML-Parser das selbstschließende Tag als
        OFFEN behandeln und den restlichen Seiteninhalt verschlucken. Diese
        rekursive DOM-Korrektur verhindert das vor dem Serialisieren.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Node,
        [Parameter(Mandatory)]$Document,
        [Parameter(Mandatory)][hashtable]$VoidElements
    )
    $child = $Node.FirstNode()
    if ($null -eq $child) {
        if ("$($Node.Type)" -eq 'Element' -and $Node.Value.Length -eq 0 -and
            -not $VoidElements.ContainsKey($Node.Name.ToLowerInvariant())) {
            $Node.AppendNode($Document.AllocateNode('Data', '', ''))
        }
        return
    }
    while ($null -ne $child) {
        if ("$($child.Type)" -eq 'Element') {
            Set-Html2PdfEmptyElementGuards -Node $child -Document $Document -VoidElements $VoidElements
        }
        $child = $child.NextSibling()
    }
}

function Restore-Html2PdfRawTextContent {
    <#
    .SYNOPSIS
        Hebt die XML-Entity-Maskierung im Inhalt von <style>/<script> wieder
        auf. Diese sind in HTML „raw text"-Elemente: ihr Inhalt wird vom
        Browser NICHT entity-dekodiert, darf also nicht maskiert sein.
    .DESCRIPTION
        PsXmls Drucker maskiert in jedem Textknoten < > " ' &. In CSS/JS
        sind das aber normale Zeichen (font-family: "Arial"; a > b; x && y).
        Würde der Browser z. B. font-family: &quot;Arial&quot; lesen, bräche
        die Regel. Da im serialisierten Inhalt kein rohes '<' steht (alles ist
        maskiert), ist das erste </style> bzw. </script> eindeutig das Ende.
        Hinweis: Ein rohes '<' in der QUELLE eines <script>/<style> bricht
        weiterhin bereits das Parsen (XML) – dafür ist CDATA nötig.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Html)

    $unescape = {
        param($m)
        $inner = $m.Groups[2].Value
        # Reihenfolge wichtig: &amp; zuletzt, damit z. B. literales &lt;
        # (gedruckt als &amp;lt;) nicht fälschlich zu < wird.
        $inner = $inner.Replace('&lt;', '<').Replace('&gt;', '>').Replace('&quot;', '"').Replace('&apos;', "'").Replace('&amp;', '&')
        return $m.Groups[1].Value + $inner + $m.Groups[3].Value
    }
    $opts = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor `
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    $Html = [regex]::Replace($Html, '(<style\b[^>]*>)(.*?)(</style>)', $unescape, $opts)
    $Html = [regex]::Replace($Html, '(<script\b[^>]*>)(.*?)(</script>)', $unescape, $opts)
    return $Html
}

function Add-Html2PdfPageSetup {
    <#
    .SYNOPSIS
        Parst HTML mit PsXml, fügt @page-Setup + <meta charset> in den
        <head> ein und liefert das fertige HTML als Text zurück (Voll-DOM).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][string]$Html,
        [Parameter(Mandatory)][int]$WidthMm,
        [Parameter(Mandatory)][int]$HeightMm,
        [Parameter(Mandatory)][string]$Margin
    )

    # Knotentypen + Strukturnamen bewusst als Strings (Coercion), siehe Kopf.
    $flags = 'DeclarationNode, CommentNodes, DoctypeNode, PiNodes'
    try {
        $doc = New-PsXmlDocument -Xml $Html -Flags $flags
    }
    catch {
        $ex = $_.Exception
        $detail = $ex.Message
        if ($null -ne $ex.PSObject.Properties['Line']) {
            $detail = "$($ex.What) (Zeile $($ex.Line), Spalte $($ex.Column), Offset $($ex.Offset))"
        }
        throw ("HTML konnte im Voll-DOM-Modus nicht geparst werden: $detail. " +
            'Die Eingabe muss wohlgeformtes (X)HTML sein (leere Elemente ' +
            'selbstschließend, <, > und & in <script>/<style> als CDATA).')
    }

    # <html> suchen (ASCII-case-insensitiv); ggf. Grundgerüst aufbauen.
    $htmlEl = $doc.FirstNode('html', $false)
    if ($null -ne $htmlEl) {
        $head = $htmlEl.FirstNode('head', $false)
        if ($null -eq $head) {
            $head = $doc.AllocateNode('Element', 'head')
            $htmlEl.PrependNode($head)
        }
    }
    else {
        # Kein <html>: vollständiges Gerüst bauen und Inhalt in <body> hängen.
        $htmlEl = $doc.AllocateNode('Element', 'html')
        $head = $doc.AllocateNode('Element', 'head')
        $body = $doc.AllocateNode('Element', 'body')
        $htmlEl.AppendNode($head)
        $htmlEl.AppendNode($body)

        $existing = @()
        $n = $doc.FirstNode()
        while ($null -ne $n) { $existing += $n; $n = $n.NextSibling() }
        foreach ($m in $existing) {
            $t = "$($m.Type)"
            if ($t -eq 'Doctype' -or $t -eq 'Declaration') { continue }
            $doc.RemoveNode($m)
            $body.AppendNode($m)
        }
        $doc.AppendNode($htmlEl)
    }

    # <meta charset> sicherstellen und auf utf-8 normalisieren: die
    # Zwischendatei wird immer als UTF-8 (mit BOM) geschrieben, daher darf
    # keine abweichende Deklaration (z. B. ISO-8859-1) stehen bleiben.
    $charsetMeta = $null
    $hc = $head.FirstNode('meta', $false)
    while ($null -ne $hc) {
        if (-not [string]::IsNullOrEmpty($hc.GetAttributeValue('charset'))) { $charsetMeta = $hc; break }
        $hc = $hc.NextSibling('meta', $false)
    }
    if ($null -ne $charsetMeta) {
        if ($charsetMeta.GetAttributeValue('charset') -ne 'utf-8') {
            [void]$charsetMeta.SetAttribute('charset', 'utf-8')
        }
    }
    else {
        $meta = $doc.AllocateNode('Element', 'meta')
        $meta.AppendAttribute($doc.AllocateAttribute('charset', 'utf-8'))
        $head.PrependNode($meta)
    }

    # @page-Setup als letztes <style> in den <head> (gewinnt in der Kaskade).
    # CSS bewusst ohne <, >, &, ", ' – wird sonst beim Drucken maskiert.
    $css = "@page { size: ${WidthMm}mm ${HeightMm}mm; margin: $Margin; }"
    $style = $doc.AllocateNode('Element', 'style')
    $style.AppendAttribute($doc.AllocateAttribute('type', 'text/css'))
    $style.AppendNode($doc.AllocateNode('Data', '', $css))
    $head.AppendNode($style)

    # Leere Nicht-Void-Elemente vor dem Serialisieren absichern (s. o.).
    $void = @{}
    foreach ($v in 'area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input',
        'link', 'meta', 'param', 'source', 'track', 'wbr') { $void[$v] = $true }
    Set-Html2PdfEmptyElementGuards -Node $doc -Document $doc -VoidElements $void

    # Serialisieren und die Entity-Maskierung in <style>/<script> aufheben.
    $out = ConvertTo-PsXmlString -Node $doc -NoIndenting
    return (Restore-Html2PdfRawTextContent -Html $out)
}

function ConvertTo-Html2PdfArgLine {
    <#
    .SYNOPSIS
        Setzt ein Argument-Array zu einer Windows-Kommandozeile zusammen und
        quotiert nach den Regeln von CommandLineToArgvW (PS 5.1/.NET Framework
        kennt ProcessStartInfo.ArgumentList nicht).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowEmptyCollection()][string[]]$Arguments)
    $q = [char]34
    $bs = [char]92
    $sb = [System.Text.StringBuilder]::new()
    foreach ($arg in $Arguments) {
        if ($sb.Length -gt 0) { [void]$sb.Append(' ') }
        if ($arg.Length -gt 0 -and $arg.IndexOfAny(@(' ', "`t", $q)) -lt 0) {
            [void]$sb.Append($arg)
            continue
        }
        [void]$sb.Append($q)
        $i = 0
        while ($i -lt $arg.Length) {
            $slashes = 0
            while ($i -lt $arg.Length -and $arg[$i] -eq $bs) { $slashes++; $i++ }
            if ($i -eq $arg.Length) {
                [void]$sb.Append($bs, $slashes * 2)   # vor schließendem Quote verdoppeln
            }
            elseif ($arg[$i] -eq $q) {
                [void]$sb.Append($bs, $slashes * 2 + 1); [void]$sb.Append($q); $i++
            }
            else {
                [void]$sb.Append($bs, $slashes); [void]$sb.Append($arg[$i]); $i++
            }
        }
        [void]$sb.Append($q)
    }
    return $sb.ToString()
}

function Stop-Html2PdfProcessTree {
    <#
    .SYNOPSIS
        Beendet einen Prozess samt aller Kindprozesse (taskkill /T). Nötig,
        weil .NET-Framework-Process.Kill() nur den direkten Prozess beendet –
        Chromium ist mehrprozessig und Kinder hielten sonst die Pipe-Handles.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][int]$ProcessId)
    try {
        $tk = Join-Path $env:WINDIR 'System32\taskkill.exe'
        if (-not (Test-Path -LiteralPath $tk)) { $tk = 'taskkill.exe' }
        & $tk '/PID' $ProcessId '/T' '/F' 2>&1 | Out-Null
    }
    catch { }
}

function Invoke-Html2PdfProcess {
    <#
    .SYNOPSIS
        Startet eine EXE mit Argumenten, wartet (mit Timeout) und liefert
        Exitcode + StdOut/StdErr zurück. Sowohl der Prozess-Exit als auch das
        Auslesen der Streams sind zeitlich begrenzt – ein hängendes
        Chromium-Kind kann so nicht zum Dauer-Deadlock führen.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Exe,
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$Arguments,
        [Parameter(Mandatory)][int]$TimeoutSeconds
    )
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Exe
    $psi.Arguments = ConvertTo-Html2PdfArgLine -Arguments $Arguments
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    $pid2 = $proc.Id
    # Streams asynchron lesen, bevor gewartet wird – verhindert Deadlocks,
    # wenn die Engine ihre Puffer füllt.
    $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
    $stderrTask = $proc.StandardError.ReadToEndAsync()

    if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
        Stop-Html2PdfProcessTree -ProcessId $pid2   # ganzen Baum beenden
        throw "Zeitüberschreitung nach $TimeoutSeconds s (Engine reagiert nicht)."
    }

    # Der Hauptprozess ist beendet. ReadToEndAsync erreicht EOF aber erst,
    # wenn AUCH alle Kindprozesse (GPU/Crashpad) die geerbten Schreib-Handles
    # schließen. Daher die Stream-Reads NICHT unbegrenzt blockieren lassen.
    $tasks = [System.Threading.Tasks.Task[]]@($stdoutTask, $stderrTask)
    try { [void][System.Threading.Tasks.Task]::WaitAll($tasks, 5000) } catch { }
    if (-not ($stdoutTask.IsCompleted -and $stderrTask.IsCompleted)) {
        # Ein Kind hält die Pipe noch: Baum beenden, kurz nachfassen.
        Stop-Html2PdfProcessTree -ProcessId $pid2
        try { [void][System.Threading.Tasks.Task]::WaitAll($tasks, 2000) } catch { }
    }
    $done = [System.Threading.Tasks.TaskStatus]::RanToCompletion
    [PSCustomObject]@{
        ExitCode = $proc.ExitCode
        StdOut   = if ($stdoutTask.Status -eq $done) { $stdoutTask.Result } else { '' }
        StdErr   = if ($stderrTask.Status -eq $done) { $stderrTask.Result } else { '' }
    }
}

function Invoke-Html2PdfChromium {
    <#
    .SYNOPSIS
        Druckt eine lokale HTML-Datei mit Edge/Chrome headless nach PDF.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Exe,
        [Parameter(Mandatory)][string]$HtmlPath,
        [Parameter(Mandatory)][string]$PdfPath,
        [Parameter(Mandatory)][int]$TimeoutSeconds
    )
    # Eigenes Profilverzeichnis: verhindert das Andocken an eine bereits
    # laufende Edge/Chrome-Instanz (sonst kehrt der Aufruf ohne PDF zurück).
    $profileDir = Join-Path $env:TEMP ('h2p_' + [Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $profileDir -Force
    try {
        $url = ([Uri]$HtmlPath).AbsoluteUri
        $arguments = @(
            '--headless=new',
            '--disable-gpu',
            # Kopf-/Fußzeilen unterdrücken (Schreibweisen mehrerer Versionen,
            # unbekannte Schalter ignoriert Chromium gefahrlos):
            '--no-pdf-header-footer',
            '--print-to-pdf-no-header',
            '--run-all-compositor-stages-before-draw',
            '--no-first-run',
            '--no-default-browser-check',
            '--disable-extensions',
            "--user-data-dir=$profileDir",
            "--print-to-pdf=$PdfPath",
            $url
        )
        $r = Invoke-Html2PdfProcess -Exe $Exe -Arguments $arguments -TimeoutSeconds $TimeoutSeconds
        # Hinweis: Chromium liefert auch bei Erfolg gelegentlich Exitcode != 0
        # bei rein informativen Log-Meldungen. Maßgeblich ist die erzeugte PDF;
        # der Exitcode dient nur der Diagnose im Fehlerfall.
        if (-not (Test-Path -LiteralPath $PdfPath) -or (Get-Item -LiteralPath $PdfPath).Length -eq 0) {
            throw "Es wurde keine (gültige) PDF erzeugt (Exitcode $($r.ExitCode)). $($r.StdErr)"
        }
    }
    finally {
        # Aufräumen mit kurzem Retry: nach einem Timeout-Baum-Kill können
        # Kindprozesse das Profil noch einen Moment sperren.
        for ($try = 0; $try -lt 3 -and (Test-Path -LiteralPath $profileDir); $try++) {
            Remove-Item -LiteralPath $profileDir -Recurse -Force -ErrorAction SilentlyContinue
            if (Test-Path -LiteralPath $profileDir) { Start-Sleep -Milliseconds 150 }
        }
    }
}

function Invoke-Html2PdfWkhtml {
    <#
    .SYNOPSIS
        Druckt eine lokale HTML-Datei mit wkhtmltopdf nach PDF. Format und
        Ausrichtung kommen hier über die CLI (wkhtmltopdf ignoriert @page).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Exe,
        [Parameter(Mandatory)][string]$HtmlPath,
        [Parameter(Mandatory)][string]$PdfPath,
        [Parameter(Mandatory)][int]$WidthMm,
        [Parameter(Mandatory)][int]$HeightMm,
        [Parameter(Mandatory)][string]$Margin,
        [Parameter(Mandatory)][int]$TimeoutSeconds
    )
    $arguments = @(
        '--quiet',
        '--enable-local-file-access',
        '--encoding', 'utf-8',
        '--page-width', "${WidthMm}mm",
        '--page-height', "${HeightMm}mm",
        '--margin-top', $Margin, '--margin-right', $Margin,
        '--margin-bottom', $Margin, '--margin-left', $Margin,
        $HtmlPath, $PdfPath
    )
    $r = Invoke-Html2PdfProcess -Exe $Exe -Arguments $arguments -TimeoutSeconds $TimeoutSeconds
    if ($r.ExitCode -ne 0) {
        throw "wkhtmltopdf endete mit Exitcode $($r.ExitCode). $($r.StdErr)"
    }
    if (-not (Test-Path -LiteralPath $PdfPath) -or (Get-Item -LiteralPath $PdfPath).Length -eq 0) {
        throw 'wkhtmltopdf hat keine (gültige) PDF erzeugt.'
    }
}

# ============================================================================
#  Hauptfunktion
# ============================================================================

function Convert-HtmlToPdf {
    <#
    .SYNOPSIS
        Erzeugt aus HTML-Dateien PDF-Dokumente (siehe Skript-Kopf).
    #>
    [CmdletBinding()]
    [OutputType([System.IO.FileInfo])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName', 'PSPath')]
        [string[]]$Path,

        [Parameter()]
        [string]$OutputPath,

        [Parameter()]
        [ValidateSet('A0', 'A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8')]
        [string]$Format = 'A4',

        [Parameter()]
        [ValidateSet('Portrait', 'Landscape')]
        [string]$Orientation = 'Portrait',

        [Parameter()]
        [string]$Margin = '10mm',

        [Parameter()]
        [ValidateSet('Auto', 'UTF8', 'UTF8BOM', 'UTF16LE', 'UTF16BE', 'ASCII', 'Latin1', 'Default')]
        [string]$InputEncoding = 'Auto',

        [Parameter()]
        [ValidateSet('Auto', 'Edge', 'Chrome', 'wkhtmltopdf')]
        [string]$Engine = 'Auto',

        [Parameter()]
        [string]$EnginePath,

        [Parameter()]
        [ValidateRange(1, 3600)]
        [int]$TimeoutSeconds = 120,

        [Parameter()]
        [switch]$KeepIntermediateHtml
    )

    begin {
        # Relativen EnginePath am PowerShell-Speicherort auflösen, sonst sucht
        # Process.Start im Prozess-Arbeitsverzeichnis (nach 'cd' abweichend).
        if ($PSBoundParameters.ContainsKey('EnginePath') -and -not [string]::IsNullOrEmpty($EnginePath)) {
            $EnginePath = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($EnginePath)
        }
        $engineInfo = Resolve-Html2PdfEngine -Engine $Engine -EnginePath $EnginePath
        $size = Get-Html2PdfPageSizeMm -Format $Format -Orientation $Orientation
        $widthMm = $size[0]; $heightMm = $size[1]
        Write-Verbose ("Engine: {0} ({1}) – {2}" -f $engineInfo.Name, $engineInfo.Kind, $engineInfo.Path)
        Write-Verbose ("Format: {0} {1} = {2}mm x {3}mm, Rand {4}" -f $Format, $Orientation, $widthMm, $heightMm, $Margin)

        # Zielangabe interpretieren: Datei vs. Verzeichnis.
        # Verzeichnis NUR bei vorhandenem Container oder abschließendem
        # Trennzeichen ('\'/'/'). Sonst eine Datei – fehlt die .pdf-Endung,
        # wird sie ergänzt (robust auch bei Punkten im Namen, z. B.
        # 'v1.2-bericht' -> 'v1.2-bericht.pdf').
        $outIsDirectory = $false
        $outFile = $null
        if ($PSBoundParameters.ContainsKey('OutputPath') -and -not [string]::IsNullOrEmpty($OutputPath)) {
            $resolvedOut = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($OutputPath)
            if ((Test-Path -LiteralPath $resolvedOut -PathType Container) -or
                $OutputPath.EndsWith('\') -or $OutputPath.EndsWith('/')) {
                $outIsDirectory = $true
                if (-not (Test-Path -LiteralPath $resolvedOut)) { $null = New-Item -ItemType Directory -Path $resolvedOut -Force }
            }
            else {
                $outFile = $resolvedOut
                if (-not $outFile.ToLowerInvariant().EndsWith('.pdf')) { $outFile += '.pdf' }
            }
        }

        $inputs = [System.Collections.Generic.List[string]]::new()
    }

    process {
        foreach ($p in $Path) {
            if ([string]::IsNullOrWhiteSpace($p)) { continue }
            # Platzhalter und Verzeichnisse auflösen.
            $items = $null
            if (Test-Path -LiteralPath $p) {
                $items = Get-Item -LiteralPath $p
            }
            else {
                $items = Get-Item -Path $p -ErrorAction SilentlyContinue
            }
            if ($null -eq $items) {
                Write-Error "Eingabe nicht gefunden: $p"
                continue
            }
            foreach ($item in $items) {
                if ($item.PSIsContainer) {
                    # Nur *.html/*.htm aus dem Verzeichnis (nicht rekursiv).
                    Get-ChildItem -LiteralPath $item.FullName -File |
                        Where-Object { $_.Extension -eq '.html' -or $_.Extension -eq '.htm' } |
                        ForEach-Object { $inputs.Add($_.FullName) }
                }
                else {
                    $inputs.Add($item.FullName)
                }
            }
        }
    }

    end {
        # Case-insensitiv deduplizieren (Windows-Pfade); Trenner normalisieren.
        $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $inputList = [System.Collections.Generic.List[string]]::new()
        foreach ($f in $inputs) {
            $full = [System.IO.Path]::GetFullPath($f)
            if ($seen.Add($full)) { [void]$inputList.Add($full) }
        }
        $count = $inputList.Count
        if ($count -eq 0) { Write-Warning 'Keine HTML-Eingabedateien gefunden.'; return }
        if ($outFile -and $count -gt 1) {
            throw 'OutputPath benennt eine einzelne Datei, es liegen aber mehrere Eingaben vor. Bitte ein Verzeichnis angeben.'
        }

        foreach ($inFull in $inputList) {
            try {
                # Zielpfad bestimmen.
                if ($outFile) {
                    $pdfPath = $outFile
                }
                elseif ($outIsDirectory) {
                    $base = [System.IO.Path]::GetFileNameWithoutExtension($inFull)
                    $pdfPath = Join-Path $resolvedOut ($base + '.pdf')
                }
                else {
                    $pdfPath = [System.IO.Path]::ChangeExtension($inFull, '.pdf')
                }
                $pdfDir = [System.IO.Path]::GetDirectoryName($pdfPath)
                if (-not (Test-Path -LiteralPath $pdfDir)) { $null = New-Item -ItemType Directory -Path $pdfDir -Force }

                Write-Verbose "Verarbeite: $inFull -> $pdfPath"

                # 1) HTML lesen. 'Auto' = BOM-Erkennung, sonst UTF-8; eine
                #    konkrete Kodierung (z. B. Latin1) übersteuert das.
                if ($InputEncoding -eq 'Auto') {
                    $rawHtml = [System.IO.File]::ReadAllText($inFull)
                }
                else {
                    $rawHtml = [System.IO.File]::ReadAllText($inFull, (Resolve-PsXmlEncoding -Name $InputEncoding))
                }

                # 2) Voll-DOM mit PsXml: @page-Setup einfügen.
                $processed = Add-Html2PdfPageSetup -Html $rawHtml -WidthMm $widthMm -HeightMm $heightMm -Margin $Margin

                # 3) Aufbereitetes HTML in temporäre Datei (UTF-8 mit BOM).
                $tmpHtml = Join-Path $env:TEMP ('h2p_' + [Guid]::NewGuid().ToString('N') + '.html')
                [System.IO.File]::WriteAllText($tmpHtml, $processed, [System.Text.UTF8Encoding]::new($true))

                # In eine frische Temp-PDF rendern und erst bei Erfolg über das
                # Ziel schieben: erkennt fehlgeschlagenes Schreiben zuverlässig
                # (statt eine veraltete PDF als Erfolg zu werten) und meldet ein
                # gesperrtes/geöffnetes Ziel klar.
                $tmpPdf = Join-Path $pdfDir ('.h2p_' + [Guid]::NewGuid().ToString('N') + '.pdf')
                try {
                    # 4) Drucken.
                    if ($engineInfo.Kind -eq 'wkhtmltopdf') {
                        Invoke-Html2PdfWkhtml -Exe $engineInfo.Path -HtmlPath $tmpHtml -PdfPath $tmpPdf `
                            -WidthMm $widthMm -HeightMm $heightMm -Margin $Margin -TimeoutSeconds $TimeoutSeconds
                    }
                    else {
                        Invoke-Html2PdfChromium -Exe $engineInfo.Path -HtmlPath $tmpHtml -PdfPath $tmpPdf `
                            -TimeoutSeconds $TimeoutSeconds
                    }

                    # 5) Atomar über das Ziel schieben.
                    try {
                        Move-Item -LiteralPath $tmpPdf -Destination $pdfPath -Force -ErrorAction Stop
                    }
                    catch {
                        throw "Ziel-PDF konnte nicht geschrieben werden (geöffnet/gesperrt?): $pdfPath. $($_.Exception.Message)"
                    }
                }
                finally {
                    Remove-Item -LiteralPath $tmpPdf -Force -ErrorAction SilentlyContinue
                    if ($KeepIntermediateHtml) {
                        $keepPath = [System.IO.Path]::ChangeExtension($pdfPath, '.processed.html')
                        try {
                            Move-Item -LiteralPath $tmpHtml -Destination $keepPath -Force -ErrorAction Stop
                            Write-Verbose "Zwischen-HTML: $keepPath"
                        }
                        catch {
                            Write-Verbose "Zwischen-HTML (Verschieben fehlgeschlagen, Originalpfad): $tmpHtml"
                        }
                    }
                    else {
                        Remove-Item -LiteralPath $tmpHtml -Force -ErrorAction SilentlyContinue
                    }
                }

                Get-Item -LiteralPath $pdfPath
            }
            catch {
                Write-Error ("Fehler bei '{0}': {1}" -f $inFull, $_.Exception.Message)
            }
        }
    }
}

# ============================================================================
#  Direkter Skriptaufruf (nicht beim Dot-Sourcing)
# ============================================================================
# Nur ausführen, wenn das Skript direkt aufgerufen wurde (InvocationName ist
# beim Dot-Sourcing '.') UND ein Pfad übergeben wurde. Für Pipeline-Eingaben
# das Skript dot-sourcen und Convert-HtmlToPdf direkt verwenden.
if ($MyInvocation.InvocationName -ne '.' -and $PSBoundParameters.ContainsKey('Path')) {
    Convert-HtmlToPdf @PSBoundParameters
}
