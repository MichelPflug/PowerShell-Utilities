# PowerShell-Toolkit

Eine allgemeine Sammlung nützlicher PowerShell-Skripte, Hilfsklassen und kleiner Werkzeuge für unterschiedliche technische Aufgaben. Das Repository startet mit einer XML-Klasse und einem HTML-zu-PDF-Konverter und wird schrittweise um weitere wiederverwendbare Skripte erweitert.

## Inhalt

| Datei | Beschreibung |
| --- | --- |
| `PsXml.ps1` | XML-Bibliothek für PowerShell mit Klassen und Funktionen zum Einlesen, Bearbeiten, Durchsuchen und Schreiben von XML-Dokumenten. Die API orientiert sich am Stil von `rapidxml`. |
| `HtmlToPdf.ps1` | Konverter für wohlgeformtes HTML/XHTML nach PDF. Unterstützt Papierformate, Ausrichtung, Ränder, Encoding-Auswahl und verschiedene Render-Engines. |

## Voraussetzungen

- Windows PowerShell 5.0 oder neuer
- Für `HtmlToPdf.ps1`: `PsXml.ps1` im selben Verzeichnis
- Für die PDF-Erzeugung: Microsoft Edge, Google Chrome oder `wkhtmltopdf`

## Schnellstart

### XML verarbeiten

```powershell
. .\PsXml.ps1

$doc = New-PsXmlDocument -Xml '<root><item id="1">Demo</item></root>'
$item = $doc.FirstNode('root').FirstNode('item')
$item.Value
```

### HTML nach PDF konvertieren

```powershell
.\HtmlToPdf.ps1 -Path .\bericht.html -Format A4 -Orientation Portrait
```

Oder als Funktion nach Dot-Sourcing:

```powershell
. .\HtmlToPdf.ps1
Convert-HtmlToPdf -Path .\bericht.html -OutputPath .\pdf\ -Format A4
```

## Hinweise

- Die Skripte sind als wiederverwendbare Bausteine gedacht und können direkt in eigene Projekte übernommen oder per Dot-Sourcing geladen werden.
- `PsXml.ps1` arbeitet mit wohlgeformtem XML.
- `HtmlToPdf.ps1` erwartet wohlgeformtes HTML/XHTML, da die HTML-Struktur intern über `PsXml.ps1` verarbeitet wird.
- Der PDF-Konverter erkennt verfügbare Engines automatisch in der Reihenfolge Edge, Chrome und `wkhtmltopdf`.

## Geplante Erweiterungen

- weitere PowerShell-Hilfsklassen
- wiederverwendbare Automatisierungsskripte
- Beispiele und kleine Werkzeuge für Entwicklungs- und Analyseaufgaben
- optionale Dokumentation zu einzelnen Skripten

## Lizenz

Das Repository ist für eine möglichst freie Nutzung vorgesehen. Die verbindliche Lizenz ergibt sich aus der `LICENSE`-Datei des Repositories.
