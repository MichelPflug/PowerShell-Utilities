#Requires -Version 5.0
<#
.SYNOPSIS
    PsXml.ps1 – XML-Bibliothek im Stil von rapidxml für PowerShell 5.1+.

.DESCRIPTION
    Stellt Klassen und Funktionen zum Einlesen (Parsen), Verarbeiten und
    Abspeichern von XML-Daten bereit. API und Verhalten sind ein
    originalgetreuer Port der C++-Bibliothek rapidxml (./rapidxml):

      Klassen
        PsXmlDocument   – Dokument (ist selbst ein Knoten), Parse()/Clear()/
                             AllocateNode()/AllocateAttribute()/CloneNode()/Save()
        PsXmlNode       – Knoten mit FirstNode()/LastNode()/NextSibling()/
                             PreviousSibling()/FirstAttribute()/AppendNode()/
                             InsertNode()/RemoveNode()/ToXmlString() u. v. m.
        PsXmlAttribute  – Attribut mit NextAttribute()/PreviousAttribute()
        PsXmlParseError – Ausnahme mit Offset/Zeile/Spalte und der
                             rapidxml-Originalmeldung in .What

      Enums
        PsXmlNodeType   – Document, Element, Data, Cdata, Comment,
                             Declaration, Doctype, Pi   (Werte wie rapidxml)
        PsXmlParseFlags – alle parse_*-Flags von rapidxml (gleiche Werte)
        PsXmlPrintFlags – None, NoIndenting (print_no_indenting)

      Funktionen (Werkzeuge)
        New-PsXmlDocument      – leeres Dokument oder aus XML-Text
        Import-PsXml           – XML-Datei einlesen und parsen
        Export-PsXml           – Dokument/Knoten in Datei schreiben
        ConvertTo-PsXmlString  – Knoten als XML-Text serialisieren
        Select-PsXmlNode       – einfache Pfadsuche (z. B. 'a/b/*/c')

    Abweichungen gegenüber rapidxml (bedingt durch .NET-Strings):
      * Es gibt keinen Memory-Pool; Strings verwaltet die .NET-GC.
        AllocateNode/AllocateAttribute existieren nur zur API-Parität.
      * NoStringTerminators und der Nicht-Destruktiv-Aspekt von
        NonDestructive sind wirkungslos (der Quelltext wird nie verändert).
        NoEntityTranslation, NoUtf8 usw. wirken wie im Original.
      * first_node(name) mit leerem Namen: '' bzw. $null wirkt hier als
        Wildcard (wie der parameterlose Aufruf).
      * Statt assert() werfen Manipulationsfehler ArgumentException/
        InvalidOperationException; Parse-Fehler werfen PsXmlParseError
        mit denselben Meldungstexten wie rapidxml ('expected <', …).

.EXAMPLE
    . .\PsXml.ps1
    $doc   = New-PsXmlDocument -Xml '<buecher><buch id="1"><titel>PowerShell</titel></buch></buecher>'
    $titel = $doc.FirstNode('buecher').FirstNode('buch').FirstNode('titel').Value
    $id    = $doc.FirstNode('buecher').FirstNode('buch').GetAttributeValue('id')

.EXAMPLE
    # Datei einlesen, verarbeiten, zurückschreiben
    . .\PsXml.ps1
    $doc = Import-PsXml -Path .\daten.xml -Flags ([PsXmlParseFlags]::Full)
    foreach ($eintrag in $doc.FirstNode('liste').ChildNodes('eintrag')) {
        $eintrag.SetAttribute('geprueft', 'ja')
    }
    Export-PsXml -Node $doc -Path .\daten_neu.xml

.EXAMPLE
    # Dokument von Grund auf neu erstellen
    . .\PsXml.ps1
    $doc  = New-PsXmlDocument
    $root = $doc.AllocateNode([PsXmlNodeType]::Element, 'wurzel')
    $doc.AppendNode($root)
    $kind = $doc.AllocateNode([PsXmlNodeType]::Element, 'kind', 'Hallo & Tschüss')
    $root.AppendNode($kind)
    $kind.AppendAttribute($doc.AllocateAttribute('nr', '1'))
    $doc.ToXmlString([PsXmlPrintFlags]::NoIndenting)
    # -> <wurzel><kind nr="1">Hallo &amp; Tschüss</kind></wurzel>
    # ToXmlString() ohne Flags liefert dieselbe Struktur mit Tab-Einrückung.

.EXAMPLE
    # Pfadsuche und Pipeline
    . .\PsXml.ps1
    $doc = Import-PsXml .\katalog.xml
    Select-PsXmlNode $doc 'katalog/cd/titel' | ForEach-Object { $_.Value }

.NOTES
    Die Datei muss per Dot-Sourcing geladen werden (". .\PsXml.ps1"),
    damit Klassen UND Funktionen in der Sitzung verfügbar sind.
    Parse-Flags sind kombinierbar, z. B.:
        $flags = [PsXmlParseFlags]'CommentNodes, DeclarationNode'
    Sehr tief verschachteltes XML (> ca. 1000 Ebenen) kann wie beim
    C++-Original die Rekursionsgrenze (Stack) erreichen.
#>

# ============================================================================
#  Enums
# ============================================================================

# Knotentypen, Werte identisch zu rapidxml::node_type.
enum PsXmlNodeType {
    Document    = 0
    Element     = 1
    Data        = 2
    Cdata       = 3
    Comment     = 4
    Declaration = 5
    Doctype     = 6
    Pi          = 7
}

# Parse-Flags, Werte identisch zu den rapidxml-Konstanten parse_*.
[Flags()]
enum PsXmlParseFlags {
    Default             = 0x000
    NoDataNodes         = 0x001
    NoElementValues     = 0x002
    NoStringTerminators = 0x004   # wirkungslos (kein In-Situ-Parsen in .NET)
    NoEntityTranslation = 0x008
    NoUtf8              = 0x010
    DeclarationNode     = 0x020
    CommentNodes        = 0x040
    DoctypeNode         = 0x080
    PiNodes             = 0x100
    ValidateClosingTags = 0x200
    TrimWhitespace      = 0x400
    NormalizeWhitespace = 0x800
    NonDestructive      = 0x00C
    Fastest             = 0x00D
    Full                = 0x3E0
}

# Druck-Flags, Werte identisch zu rapidxml_print.hpp.
[Flags()]
enum PsXmlPrintFlags {
    None        = 0
    NoIndenting = 1
}

# ============================================================================
#  Ausnahme: PsXmlParseError  (Pendant zu rapidxml::parse_error)
# ============================================================================

class PsXmlParseError : System.Exception {
    # Originalmeldung von rapidxml, z. B. 'expected <'
    [string]$What
    # 0-basierter Zeichen-Offset im Quelltext (Pendant zu where())
    [int]$Offset
    # 1-basierte Zeile/Spalte, aus dem Offset berechnet
    [int]$Line
    [int]$Column

    PsXmlParseError([string]$what, [string]$text, [int]$offset) : base(
        [PsXmlParseError]::BuildMessage($what, $text, $offset)) {
        $this.What   = $what
        $this.Offset = $offset
        $lc = [PsXmlParseError]::LineColumnOf($text, $offset)
        $this.Line   = $lc[0]
        $this.Column = $lc[1]
    }

    hidden static [string] BuildMessage([string]$what, [string]$text, [int]$offset) {
        $lc = [PsXmlParseError]::LineColumnOf($text, $offset)
        return ('PsXml-Parse-Fehler: {0} (Zeile {1}, Spalte {2}, Offset {3})' -f $what, $lc[0], $lc[1], $offset)
    }

    hidden static [int[]] LineColumnOf([string]$text, [int]$at) {
        if ($null -eq $text) { return @(1, 1) }
        $pos = $at
        if ($pos -gt $text.Length) { $pos = $text.Length }
        $ln = 1
        $lnStart = 0
        for ($i = 0; $i -lt $pos; $i++) {
            if ($text[$i] -eq [char]10) { $ln++; $lnStart = $i + 1 }
        }
        return @($ln, ($pos - $lnStart + 1))
    }
}

# ============================================================================
#  Basisklasse: PsXmlBase  (Pendant zu rapidxml::xml_base)
# ============================================================================

class PsXmlBase {
    [string]$Name  = ''
    [string]$Value = ''
    hidden [PsXmlNode]$_parent = $null

    # Elternknoten oder $null (Pendant zu parent()).
    [PsXmlNode] Parent() { return $this._parent }

    # Namensvergleich wie rapidxml::internal::compare:
    # case-sensitive = Ordinal; case-insensitive faltet nur ASCII a-z.
    # Vergleich über [int]-Codepoints, da PowerShells -eq/-ge/-le auf [char]
    # case-insensitiv arbeitet und so Nicht-ASCII-Paare fälschlich matchen würde.
    static [bool] NamesEqual([string]$a, [string]$b, [bool]$caseSensitive) {
        if ($null -eq $a) { $a = '' }
        if ($null -eq $b) { $b = '' }
        if ($caseSensitive) { return [string]::Equals($a, $b, [System.StringComparison]::Ordinal) }
        if ($a.Length -ne $b.Length) { return $false }
        for ($i = 0; $i -lt $a.Length; $i++) {
            $ca = [int]$a[$i]
            $cb = [int]$b[$i]
            if ($ca -ge 97 -and $ca -le 122) { $ca -= 32 }
            if ($cb -ge 97 -and $cb -le 122) { $cb -= 32 }
            if ($ca -ne $cb) { return $false }
        }
        return $true
    }
}

# ============================================================================
#  PsXmlAttribute  (Pendant zu rapidxml::xml_attribute)
# ============================================================================

class PsXmlAttribute : PsXmlBase {
    hidden [PsXmlAttribute]$_prevAttr = $null
    hidden [PsXmlAttribute]$_nextAttr = $null

    PsXmlAttribute() { }
    PsXmlAttribute([string]$name) { $this.Name = $name }
    PsXmlAttribute([string]$name, [string]$value) { $this.Name = $name; $this.Value = $value }

    # Dokument, zu dem das Attribut gehört, oder $null (Pendant zu document()).
    [PsXmlDocument] Document() {
        if ($null -eq $this._parent) { return $null }
        return $this._parent.Document()
    }

    [PsXmlAttribute] PreviousAttribute() {
        if ($null -ne $this._parent) { return $this._prevAttr }
        return $null
    }
    [PsXmlAttribute] PreviousAttribute([string]$name) { return $this.PreviousAttribute($name, $true) }
    [PsXmlAttribute] PreviousAttribute([string]$name, [bool]$caseSensitive) {
        if ([string]::IsNullOrEmpty($name)) { return $this.PreviousAttribute() }
        $a = $this._prevAttr
        while ($null -ne $a) {
            if ([PsXmlBase]::NamesEqual($a.Name, $name, $caseSensitive)) { return $a }
            $a = $a._prevAttr
        }
        return $null
    }

    [PsXmlAttribute] NextAttribute() {
        if ($null -ne $this._parent) { return $this._nextAttr }
        return $null
    }
    [PsXmlAttribute] NextAttribute([string]$name) { return $this.NextAttribute($name, $true) }
    [PsXmlAttribute] NextAttribute([string]$name, [bool]$caseSensitive) {
        if ([string]::IsNullOrEmpty($name)) { return $this.NextAttribute() }
        $a = $this._nextAttr
        while ($null -ne $a) {
            if ([PsXmlBase]::NamesEqual($a.Name, $name, $caseSensitive)) { return $a }
            $a = $a._nextAttr
        }
        return $null
    }

    # Liefert das Attribut so, wie es der Drucker serialisieren würde.
    [string] ToString() {
        $sb = [System.Text.StringBuilder]::new()
        $sb.Append($this.Name).Append('=')
        [PsXmlPrinter]::AppendQuotedAttributeValue($sb, $this.Value)
        return $sb.ToString()
    }
}

# ============================================================================
#  PsXmlNode  (Pendant zu rapidxml::xml_node)
# ============================================================================

class PsXmlNode : PsXmlBase {
    [PsXmlNodeType]$Type = [PsXmlNodeType]::Element

    hidden [PsXmlNode]$_firstNode   = $null
    hidden [PsXmlNode]$_lastNode    = $null
    hidden [PsXmlNode]$_prevSibling = $null
    hidden [PsXmlNode]$_nextSibling = $null
    hidden [PsXmlAttribute]$_firstAttribute = $null
    hidden [PsXmlAttribute]$_lastAttribute  = $null

    PsXmlNode() { }
    PsXmlNode([PsXmlNodeType]$type) { $this.Type = $type }
    PsXmlNode([PsXmlNodeType]$type, [string]$name) { $this.Type = $type; $this.Name = $name }
    PsXmlNode([PsXmlNodeType]$type, [string]$name, [string]$value) { $this.Type = $type; $this.Name = $name; $this.Value = $value }

    # ---- Navigation -------------------------------------------------------

    # Dokument, zu dem der Knoten gehört, oder $null (Pendant zu document()).
    [PsXmlDocument] Document() {
        $n = $this
        while ($null -ne $n._parent) { $n = $n._parent }
        if ($n -is [PsXmlDocument]) { return [PsXmlDocument]$n }
        return $null
    }

    [PsXmlNode] FirstNode() { return $this._firstNode }
    [PsXmlNode] FirstNode([string]$name) { return $this.FirstNode($name, $true) }
    [PsXmlNode] FirstNode([string]$name, [bool]$caseSensitive) {
        if ([string]::IsNullOrEmpty($name)) { return $this._firstNode }
        $n = $this._firstNode
        while ($null -ne $n) {
            if ([PsXmlBase]::NamesEqual($n.Name, $name, $caseSensitive)) { return $n }
            $n = $n._nextSibling
        }
        return $null
    }

    [PsXmlNode] LastNode() { return $this._lastNode }
    [PsXmlNode] LastNode([string]$name) { return $this.LastNode($name, $true) }
    [PsXmlNode] LastNode([string]$name, [bool]$caseSensitive) {
        if ([string]::IsNullOrEmpty($name)) { return $this._lastNode }
        $n = $this._lastNode
        while ($null -ne $n) {
            if ([PsXmlBase]::NamesEqual($n.Name, $name, $caseSensitive)) { return $n }
            $n = $n._prevSibling
        }
        return $null
    }

    [PsXmlNode] PreviousSibling() {
        if ($null -ne $this._parent) { return $this._prevSibling }
        return $null
    }
    [PsXmlNode] PreviousSibling([string]$name) { return $this.PreviousSibling($name, $true) }
    [PsXmlNode] PreviousSibling([string]$name, [bool]$caseSensitive) {
        if ([string]::IsNullOrEmpty($name)) { return $this.PreviousSibling() }
        if ($null -eq $this._parent) { return $null }
        $n = $this._prevSibling
        while ($null -ne $n) {
            if ([PsXmlBase]::NamesEqual($n.Name, $name, $caseSensitive)) { return $n }
            $n = $n._prevSibling
        }
        return $null
    }

    [PsXmlNode] NextSibling() {
        if ($null -ne $this._parent) { return $this._nextSibling }
        return $null
    }
    [PsXmlNode] NextSibling([string]$name) { return $this.NextSibling($name, $true) }
    [PsXmlNode] NextSibling([string]$name, [bool]$caseSensitive) {
        if ([string]::IsNullOrEmpty($name)) { return $this.NextSibling() }
        if ($null -eq $this._parent) { return $null }
        $n = $this._nextSibling
        while ($null -ne $n) {
            if ([PsXmlBase]::NamesEqual($n.Name, $name, $caseSensitive)) { return $n }
            $n = $n._nextSibling
        }
        return $null
    }

    [PsXmlAttribute] FirstAttribute() { return $this._firstAttribute }
    [PsXmlAttribute] FirstAttribute([string]$name) { return $this.FirstAttribute($name, $true) }
    [PsXmlAttribute] FirstAttribute([string]$name, [bool]$caseSensitive) {
        if ([string]::IsNullOrEmpty($name)) { return $this._firstAttribute }
        $a = $this._firstAttribute
        while ($null -ne $a) {
            if ([PsXmlBase]::NamesEqual($a.Name, $name, $caseSensitive)) { return $a }
            $a = $a._nextAttr
        }
        return $null
    }

    [PsXmlAttribute] LastAttribute() { return $this._lastAttribute }
    [PsXmlAttribute] LastAttribute([string]$name) { return $this.LastAttribute($name, $true) }
    [PsXmlAttribute] LastAttribute([string]$name, [bool]$caseSensitive) {
        if ([string]::IsNullOrEmpty($name)) { return $this._lastAttribute }
        $a = $this._lastAttribute
        while ($null -ne $a) {
            if ([PsXmlBase]::NamesEqual($a.Name, $name, $caseSensitive)) { return $a }
            $a = $a._prevAttr
        }
        return $null
    }

    # ---- Komfort (PowerShell-Erweiterungen) -------------------------------

    # Alle Kindknoten als Array (für Pipeline/foreach).
    [object[]] ChildNodes() {
        $list = [System.Collections.Generic.List[object]]::new()
        $n = $this._firstNode
        while ($null -ne $n) { $list.Add($n); $n = $n._nextSibling }
        return $list.ToArray()
    }
    # Alle Kindknoten mit passendem Namen (Ordinal, case-sensitive);
    # leerer Name wirkt wie der parameterlose Aufruf (Wildcard).
    [object[]] ChildNodes([string]$name) {
        if ([string]::IsNullOrEmpty($name)) { return $this.ChildNodes() }
        $list = [System.Collections.Generic.List[object]]::new()
        $n = $this._firstNode
        while ($null -ne $n) {
            if ([PsXmlBase]::NamesEqual($n.Name, $name, $true)) { $list.Add($n) }
            $n = $n._nextSibling
        }
        return $list.ToArray()
    }

    # Alle Attribute als Array.
    [object[]] GetAttributes() {
        $list = [System.Collections.Generic.List[object]]::new()
        $a = $this._firstAttribute
        while ($null -ne $a) { $list.Add($a); $a = $a._nextAttr }
        return $list.ToArray()
    }

    # Attributwert lesen; '' wenn nicht vorhanden (bzw. eigener Standardwert).
    [string] GetAttributeValue([string]$name) { return $this.GetAttributeValue($name, '') }
    [string] GetAttributeValue([string]$name, [string]$default) {
        $a = $this.FirstAttribute($name)
        if ($null -ne $a) { return $a.Value }
        return $default
    }

    # Attribut setzen (vorhandenes aktualisieren oder neues anhängen).
    [PsXmlAttribute] SetAttribute([string]$name, [string]$value) {
        $a = $this.FirstAttribute($name)
        if ($null -ne $a) { $a.Value = $value; return $a }
        $a = [PsXmlAttribute]::new($name, $value)
        $this.AppendAttribute($a)
        return $a
    }

    # Pendants zu rapidxml::count_children / count_attributes (O(n)).
    [int] CountChildren() {
        $count = 0
        $n = $this._firstNode
        while ($null -ne $n) { $count++; $n = $n._nextSibling }
        return $count
    }
    [int] CountAttributes() {
        $count = 0
        $a = $this._firstAttribute
        while ($null -ne $a) { $count++; $a = $a._nextAttr }
        return $count
    }

    # ---- Manipulation: Kindknoten -----------------------------------------

    hidden [void] ValidateInsertableNode([PsXmlNode]$child) {
        if ($null -eq $child) {
            throw [System.ArgumentNullException]::new('child', 'Es wurde kein Knoten übergeben.')
        }
        if ($null -ne $child._parent) {
            throw [System.InvalidOperationException]::new('Der Knoten hat bereits einen Elternknoten. Zuerst mit RemoveNode() entfernen.')
        }
        if ($child.Type -eq [PsXmlNodeType]::Document) {
            throw [System.InvalidOperationException]::new('Ein Dokumentknoten kann nicht als Kind eingefügt werden.')
        }
        $a = $this
        while ($null -ne $a) {
            if ([object]::ReferenceEquals($a, $child)) {
                throw [System.InvalidOperationException]::new('Der Knoten kann nicht in seinen eigenen Teilbaum eingefügt werden.')
            }
            $a = $a._parent
        }
    }

    [void] PrependNode([PsXmlNode]$child) {
        $this.ValidateInsertableNode($child)
        if ($null -ne $this._firstNode) {
            $child._nextSibling = $this._firstNode
            $this._firstNode._prevSibling = $child
        }
        else {
            $child._nextSibling = $null
            $this._lastNode = $child
        }
        $this._firstNode = $child
        $child._parent = $this
        $child._prevSibling = $null
    }

    [void] AppendNode([PsXmlNode]$child) {
        $this.ValidateInsertableNode($child)
        if ($null -ne $this._firstNode) {
            $child._prevSibling = $this._lastNode
            $this._lastNode._nextSibling = $child
        }
        else {
            $child._prevSibling = $null
            $this._firstNode = $child
        }
        $this._lastNode = $child
        $child._parent = $this
        $child._nextSibling = $null
    }

    # Fügt child vor 'where' ein; where = $null wirkt wie AppendNode (wie rapidxml).
    [void] InsertNode([PsXmlNode]$where, [PsXmlNode]$child) {
        if ($null -eq $where) { $this.AppendNode($child); return }
        if (-not [object]::ReferenceEquals($where._parent, $this)) {
            throw [System.ArgumentException]::new('Der Knoten "where" ist kein Kind dieses Knotens.', 'where')
        }
        if ([object]::ReferenceEquals($where, $this._firstNode)) { $this.PrependNode($child); return }
        $this.ValidateInsertableNode($child)
        $child._prevSibling = $where._prevSibling
        $child._nextSibling = $where
        $where._prevSibling._nextSibling = $child
        $where._prevSibling = $child
        $child._parent = $this
    }

    [void] RemoveFirstNode() {
        if ($null -eq $this._firstNode) {
            throw [System.InvalidOperationException]::new('Der Knoten hat keine Kindknoten.')
        }
        $child = $this._firstNode
        $this._firstNode = $child._nextSibling
        if ($null -ne $child._nextSibling) { $child._nextSibling._prevSibling = $null }
        else { $this._lastNode = $null }
        $child._parent = $null
        $child._prevSibling = $null
        $child._nextSibling = $null
    }

    [void] RemoveLastNode() {
        if ($null -eq $this._firstNode) {
            throw [System.InvalidOperationException]::new('Der Knoten hat keine Kindknoten.')
        }
        $child = $this._lastNode
        if ($null -ne $child._prevSibling) {
            $this._lastNode = $child._prevSibling
            $child._prevSibling._nextSibling = $null
        }
        else {
            $this._firstNode = $null
            $this._lastNode = $null
        }
        $child._parent = $null
        $child._prevSibling = $null
        $child._nextSibling = $null
    }

    [void] RemoveNode([PsXmlNode]$where) {
        if ($null -eq $where -or -not [object]::ReferenceEquals($where._parent, $this)) {
            throw [System.ArgumentException]::new('Der Knoten "where" ist kein Kind dieses Knotens.', 'where')
        }
        if ([object]::ReferenceEquals($where, $this._firstNode)) { $this.RemoveFirstNode(); return }
        if ([object]::ReferenceEquals($where, $this._lastNode)) { $this.RemoveLastNode(); return }
        $where._prevSibling._nextSibling = $where._nextSibling
        $where._nextSibling._prevSibling = $where._prevSibling
        $where._parent = $null
        $where._prevSibling = $null
        $where._nextSibling = $null
    }

    [void] RemoveAllNodes() {
        $n = $this._firstNode
        while ($null -ne $n) {
            $next = $n._nextSibling
            $n._parent = $null
            $n._prevSibling = $null
            $n._nextSibling = $null
            $n = $next
        }
        $this._firstNode = $null
        $this._lastNode = $null
    }

    # ---- Manipulation: Attribute ------------------------------------------

    hidden [void] ValidateInsertableAttribute([PsXmlAttribute]$attribute) {
        if ($null -eq $attribute) {
            throw [System.ArgumentNullException]::new('attribute', 'Es wurde kein Attribut übergeben.')
        }
        if ($null -ne $attribute._parent) {
            throw [System.InvalidOperationException]::new('Das Attribut gehört bereits zu einem Knoten. Zuerst mit RemoveAttribute() entfernen.')
        }
    }

    [void] PrependAttribute([PsXmlAttribute]$attribute) {
        $this.ValidateInsertableAttribute($attribute)
        if ($null -ne $this._firstAttribute) {
            $attribute._nextAttr = $this._firstAttribute
            $this._firstAttribute._prevAttr = $attribute
        }
        else {
            $attribute._nextAttr = $null
            $this._lastAttribute = $attribute
        }
        $this._firstAttribute = $attribute
        $attribute._parent = $this
        $attribute._prevAttr = $null
    }

    [void] AppendAttribute([PsXmlAttribute]$attribute) {
        $this.ValidateInsertableAttribute($attribute)
        if ($null -ne $this._firstAttribute) {
            $attribute._prevAttr = $this._lastAttribute
            $this._lastAttribute._nextAttr = $attribute
        }
        else {
            $attribute._prevAttr = $null
            $this._firstAttribute = $attribute
        }
        $this._lastAttribute = $attribute
        $attribute._parent = $this
        $attribute._nextAttr = $null
    }

    # Fügt attribute vor 'where' ein; where = $null wirkt wie AppendAttribute.
    [void] InsertAttribute([PsXmlAttribute]$where, [PsXmlAttribute]$attribute) {
        if ($null -eq $where) { $this.AppendAttribute($attribute); return }
        if (-not [object]::ReferenceEquals($where._parent, $this)) {
            throw [System.ArgumentException]::new('Das Attribut "where" gehört nicht zu diesem Knoten.', 'where')
        }
        if ([object]::ReferenceEquals($where, $this._firstAttribute)) { $this.PrependAttribute($attribute); return }
        $this.ValidateInsertableAttribute($attribute)
        $attribute._prevAttr = $where._prevAttr
        $attribute._nextAttr = $where
        $where._prevAttr._nextAttr = $attribute
        $where._prevAttr = $attribute
        $attribute._parent = $this
    }

    [void] RemoveFirstAttribute() {
        if ($null -eq $this._firstAttribute) {
            throw [System.InvalidOperationException]::new('Der Knoten hat keine Attribute.')
        }
        $a = $this._firstAttribute
        $this._firstAttribute = $a._nextAttr
        if ($null -ne $a._nextAttr) { $a._nextAttr._prevAttr = $null }
        else { $this._lastAttribute = $null }
        $a._parent = $null
        $a._prevAttr = $null
        $a._nextAttr = $null
    }

    [void] RemoveLastAttribute() {
        if ($null -eq $this._firstAttribute) {
            throw [System.InvalidOperationException]::new('Der Knoten hat keine Attribute.')
        }
        $a = $this._lastAttribute
        if ($null -ne $a._prevAttr) {
            $this._lastAttribute = $a._prevAttr
            $a._prevAttr._nextAttr = $null
        }
        else {
            $this._firstAttribute = $null
            $this._lastAttribute = $null
        }
        $a._parent = $null
        $a._prevAttr = $null
        $a._nextAttr = $null
    }

    [void] RemoveAttribute([PsXmlAttribute]$where) {
        if ($null -eq $where -or -not [object]::ReferenceEquals($where._parent, $this)) {
            throw [System.ArgumentException]::new('Das Attribut "where" gehört nicht zu diesem Knoten.', 'where')
        }
        if ([object]::ReferenceEquals($where, $this._firstAttribute)) { $this.RemoveFirstAttribute(); return }
        if ([object]::ReferenceEquals($where, $this._lastAttribute)) { $this.RemoveLastAttribute(); return }
        $where._prevAttr._nextAttr = $where._nextAttr
        $where._nextAttr._prevAttr = $where._prevAttr
        $where._parent = $null
        $where._prevAttr = $null
        $where._nextAttr = $null
    }

    [void] RemoveAllAttributes() {
        $a = $this._firstAttribute
        while ($null -ne $a) {
            $next = $a._nextAttr
            $a._parent = $null
            $a._prevAttr = $null
            $a._nextAttr = $null
            $a = $next
        }
        $this._firstAttribute = $null
        $this._lastAttribute = $null
    }

    # ---- Serialisierung ----------------------------------------------------

    [string] ToXmlString() { return [PsXmlPrinter]::Print($this, 0) }
    [string] ToXmlString([PsXmlPrintFlags]$flags) { return [PsXmlPrinter]::Print($this, [int]$flags) }

    [string] ToString() { return $this.ToXmlString() }
}

# ============================================================================
#  PsXmlDocument  (Pendant zu rapidxml::xml_document)
# ============================================================================

class PsXmlDocument : PsXmlNode {

    PsXmlDocument() : base([PsXmlNodeType]::Document) { }
    PsXmlDocument([string]$xml) : base([PsXmlNodeType]::Document) { $this.Parse($xml) }
    PsXmlDocument([string]$xml, [PsXmlParseFlags]$flags) : base([PsXmlNodeType]::Document) { $this.Parse($xml, $flags) }

    # Parst XML-Text. Ein vorhandener Baum wird vorher verworfen.
    # Mehrere Wurzelelemente werden (wie bei rapidxml) akzeptiert.
    [void] Parse([string]$xml) { $this.Parse($xml, [PsXmlParseFlags]::Default) }
    [void] Parse([string]$xml, [PsXmlParseFlags]$flags) {
        $this.RemoveAllNodes()
        $this.RemoveAllAttributes()
        if ($null -eq $xml) { $xml = '' }
        $parser = [PsXmlParser]::new($xml, [int]$flags, $this)
        $parser.ParseDocument()
    }

    # Verwirft den gesamten Baum (Pendant zu clear()).
    [void] Clear() {
        $this.RemoveAllNodes()
        $this.RemoveAllAttributes()
    }

    # Knoten-/Attribut-Fabriken (Pendants zu allocate_node/allocate_attribute;
    # ein Memory-Pool ist in .NET nicht nötig).
    [PsXmlNode] AllocateNode([PsXmlNodeType]$type) { return [PsXmlNode]::new($type) }
    [PsXmlNode] AllocateNode([PsXmlNodeType]$type, [string]$name) { return [PsXmlNode]::new($type, $name) }
    [PsXmlNode] AllocateNode([PsXmlNodeType]$type, [string]$name, [string]$value) { return [PsXmlNode]::new($type, $name, $value) }
    [PsXmlAttribute] AllocateAttribute([string]$name) { return [PsXmlAttribute]::new($name) }
    [PsXmlAttribute] AllocateAttribute([string]$name, [string]$value) { return [PsXmlAttribute]::new($name, $value) }

    # Tiefe Kopie eines Knotens (Pendant zu clone_node). Mit $result kann in
    # einen vorhandenen Knoten geklont werden, z. B. in ein anderes Dokument:
    #   $ziel = [PsXmlDocument]::new(); $quelle.CloneNode($quellDoc, $ziel)
    [PsXmlNode] CloneNode([PsXmlNode]$source) { return [PsXmlDocument]::CloneInto($source, $null) }
    [PsXmlNode] CloneNode([PsXmlNode]$source, [PsXmlNode]$result) { return [PsXmlDocument]::CloneInto($source, $result) }

    hidden static [PsXmlNode] CloneInto([PsXmlNode]$source, [PsXmlNode]$result) {
        if ($null -eq $source) {
            throw [System.ArgumentNullException]::new('source', 'Es wurde kein Quellknoten übergeben.')
        }
        if ($null -ne $result) {
            $result.RemoveAllAttributes()
            $result.RemoveAllNodes()
            $result.Type = $source.Type
        }
        else {
            $result = [PsXmlNode]::new($source.Type)
        }
        $result.Name  = $source.Name
        $result.Value = $source.Value
        $child = $source.FirstNode()
        while ($null -ne $child) {
            $result.AppendNode([PsXmlDocument]::CloneInto($child, $null))
            $child = $child.NextSibling()
        }
        $a = $source.FirstAttribute()
        while ($null -ne $a) {
            $result.AppendAttribute([PsXmlAttribute]::new($a.Name, $a.Value))
            $a = $a.NextAttribute()
        }
        return $result
    }

    # Dokument in eine Datei schreiben (UTF-8 ohne BOM).
    [void] Save([string]$path) { $this.Save($path, [PsXmlPrintFlags]::None) }
    [void] Save([string]$path, [PsXmlPrintFlags]$flags) {
        $resolved = [PsXmlDocument]::ResolvePath($path)
        $xml = [PsXmlPrinter]::Print($this, [int]$flags)
        [System.IO.File]::WriteAllText($resolved, $xml, [System.Text.UTF8Encoding]::new($false))
    }

    # Dokument aus einer Datei laden (Pendant zu rapidxml::file + parse;
    # Zeichenkodierung wird per BOM erkannt, sonst UTF-8).
    static [PsXmlDocument] FromFile([string]$path) { return [PsXmlDocument]::FromFile($path, [PsXmlParseFlags]::Default) }
    static [PsXmlDocument] FromFile([string]$path, [PsXmlParseFlags]$flags) {
        $resolved = [PsXmlDocument]::ResolvePath($path)
        $text = [System.IO.File]::ReadAllText($resolved)
        return [PsXmlDocument]::new($text, $flags)
    }

    # Relative Pfade beziehen sich auf den PowerShell-Speicherort, nicht auf
    # das Prozess-Arbeitsverzeichnis.
    hidden static [string] ResolvePath([string]$path) {
        if ([string]::IsNullOrEmpty($path)) {
            throw [System.ArgumentNullException]::new('path', 'Es wurde kein Dateipfad übergeben.')
        }
        if ([System.IO.Path]::IsPathRooted($path)) { return [System.IO.Path]::GetFullPath($path) }
        $base = (Get-Location).ProviderPath
        return [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($base, $path))
    }
}

# ============================================================================
#  PsXmlParser  (interner Arbeiter; Port des rapidxml-Parsers)
# ============================================================================

class PsXmlParser {
    hidden [string]$Text
    hidden [int]$Pos
    hidden [int]$Len
    hidden [PsXmlDocument]$Doc

    hidden [bool]$FNoDataNodes
    hidden [bool]$FNoElementValues
    hidden [bool]$FNoEntityTranslation
    hidden [bool]$FNoUtf8
    hidden [bool]$FDeclarationNode
    hidden [bool]$FCommentNodes
    hidden [bool]$FDoctypeNode
    hidden [bool]$FPiNodes
    hidden [bool]$FValidateClosingTags
    hidden [bool]$FTrimWhitespace
    hidden [bool]$FNormalizeWhitespace

    # Stoppzeichen-Tabellen (NUL steht für das rapidxml-Stringende).
    hidden static [char[]]$WsChars            = " `t`n`r".ToCharArray()
    hidden static [char[]]$NodeNameStops      = "`0 `t`n`r/>?".ToCharArray()
    hidden static [char[]]$AttNameStops       = "`0 `t`n`r!/<=>?".ToCharArray()
    hidden static [char[]]$TextStops          = "`0<".ToCharArray()
    hidden static [char[]]$TextStopsEnt       = "`0<&".ToCharArray()
    hidden static [char[]]$TextStopsNorm      = "`0<& `t`n`r".ToCharArray()
    hidden static [char[]]$TextStopsNormNoEnt = "`0< `t`n`r".ToCharArray()
    hidden static [char[]]$AttValStopsSq      = "`0'".ToCharArray()
    hidden static [char[]]$AttValStopsSqEnt   = "`0'&".ToCharArray()
    hidden static [char[]]$AttValStopsDq      = "`0`"".ToCharArray()
    hidden static [char[]]$AttValStopsDqEnt   = "`0`"&".ToCharArray()
    hidden static [char[]]$DoctypeStops       = "`0[>".ToCharArray()
    hidden static [char[]]$BracketStops       = "`0[]".ToCharArray()

    PsXmlParser([string]$text, [int]$flags, [PsXmlDocument]$doc) {
        $this.Text = $text
        $this.Len  = $text.Length
        $this.Pos  = 0
        $this.Doc  = $doc
        $this.FNoDataNodes         = ($flags -band 0x001) -ne 0
        $this.FNoElementValues     = ($flags -band 0x002) -ne 0
        $this.FNoEntityTranslation = ($flags -band 0x008) -ne 0
        $this.FNoUtf8              = ($flags -band 0x010) -ne 0
        $this.FDeclarationNode     = ($flags -band 0x020) -ne 0
        $this.FCommentNodes        = ($flags -band 0x040) -ne 0
        $this.FDoctypeNode         = ($flags -band 0x080) -ne 0
        $this.FPiNodes             = ($flags -band 0x100) -ne 0
        $this.FValidateClosingTags = ($flags -band 0x200) -ne 0
        $this.FTrimWhitespace      = ($flags -band 0x400) -ne 0
        $this.FNormalizeWhitespace = ($flags -band 0x800) -ne 0
    }

    # ---- Grundbausteine ----------------------------------------------------

    hidden [char] CharAt([int]$index) {
        if ($index -lt $this.Len) { return $this.Text[$index] }
        return [char]0
    }

    hidden static [bool] IsWs([char]$c) {
        return ($c -eq [char]32 -or $c -eq [char]9 -or $c -eq [char]10 -or $c -eq [char]13)
    }

    hidden [void] SkipWhitespace() {
        $t = $this.Text
        $p = $this.Pos
        $n = $this.Len
        while ($p -lt $n) {
            $c = $t[$p]
            if ($c -eq [char]32 -or $c -eq [char]9 -or $c -eq [char]10 -or $c -eq [char]13) { $p++ }
            else { break }
        }
        $this.Pos = $p
    }

    # IndexOfAny, das das Stringende wie das NUL-Terminator-Zeichen behandelt.
    hidden [int] FindAny([char[]]$stops, [int]$from) {
        $i = $this.Text.IndexOfAny($stops, $from)
        if ($i -lt 0) { return $this.Len }
        return $i
    }

    # Sucht $token ab $from. Ein eingebettetes NUL davor gilt wie bei
    # rapidxml als Datenende ('unexpected end of data'), ebenso ein
    # fehlendes Token.
    hidden [int] FindToken([string]$token, [int]$from) {
        $idx = $this.Text.IndexOf($token, $from, [System.StringComparison]::Ordinal)
        $stop = $this.Len
        if ($idx -ge 0) { $stop = $idx }
        $nul = $this.Text.IndexOf([char]0, $from, $stop - $from)
        if ($nul -ge 0) { $this.Fail('unexpected end of data', $nul) }
        if ($idx -lt 0) { $this.Fail('unexpected end of data', $this.Len) }
        return $idx
    }

    hidden [void] Fail([string]$what) {
        throw [PsXmlParseError]::new($what, $this.Text, $this.Pos)
    }
    hidden [void] Fail([string]$what, [int]$at) {
        throw [PsXmlParseError]::new($what, $this.Text, $at)
    }

    # [int]-Vergleiche, da PowerShell-Operatoren auf [char] case-insensitiv sind.
    hidden static [int] DigitValue([char]$c) {
        $i = [int]$c
        if ($i -ge 48 -and $i -le 57)  { return $i - 48 }   # 0-9
        if ($i -ge 65 -and $i -le 70)  { return $i - 55 }   # A-F
        if ($i -ge 97 -and $i -le 102) { return $i - 87 }   # a-f
        return 255
    }

    hidden [bool] IsAttNameChar([char]$c) {
        if ($c -eq [char]0  -or $c -eq [char]32 -or $c -eq [char]9 -or
            $c -eq [char]10 -or $c -eq [char]13 -or
            $c -eq [char]'!' -or $c -eq [char]'/' -or $c -eq [char]'<' -or
            $c -eq [char]'=' -or $c -eq [char]'>' -or $c -eq [char]'?') { return $false }
        return $true
    }

    # ---- Dokumentebene -----------------------------------------------------

    [void] ParseDocument() {
        # UTF-8-BOM überspringen (als Bytes EF BB BF oder als Zeichen U+FEFF)
        if ($this.Len -ge 3 -and $this.Text[0] -eq [char]0xEF -and $this.Text[1] -eq [char]0xBB -and $this.Text[2] -eq [char]0xBF) {
            $this.Pos = 3
        }
        elseif ($this.Len -ge 1 -and $this.Text[0] -eq [char]0xFEFF) {
            $this.Pos = 1
        }
        while ($true) {
            $this.SkipWhitespace()
            if ($this.CharAt($this.Pos) -eq [char]0) { break }
            if ($this.Text[$this.Pos] -ne [char]'<') { $this.Fail('expected <') }
            $this.Pos++
            $node = $this.ParseNode()
            if ($null -ne $node) { $this.Doc.AppendNode($node) }
        }
    }

    # ---- Knoten-Dispatch (nach konsumiertem '<') ----------------------------

    hidden [PsXmlNode] ParseNode() {
        $c = $this.CharAt($this.Pos)
        if ($c -eq [char]'?') {
            $this.Pos++
            $c0 = $this.CharAt($this.Pos)
            $c1 = $this.CharAt($this.Pos + 1)
            $c2 = $this.CharAt($this.Pos + 2)
            $c3 = $this.CharAt($this.Pos + 3)
            if (($c0 -eq [char]'x' -or $c0 -eq [char]'X') -and
                ($c1 -eq [char]'m' -or $c1 -eq [char]'M') -and
                ($c2 -eq [char]'l' -or $c2 -eq [char]'L') -and
                [PsXmlParser]::IsWs($c3)) {
                $this.Pos += 4
                return $this.ParseXmlDeclaration()
            }
            return $this.ParsePi()
        }
        if ($c -eq [char]'!') {
            $c1 = $this.CharAt($this.Pos + 1)
            if ($c1 -eq [char]'-' -and $this.CharAt($this.Pos + 2) -eq [char]'-') {
                $this.Pos += 3
                return $this.ParseComment()
            }
            if ($c1 -eq [char]'[' -and ($this.Len - $this.Pos) -ge 8 -and
                [string]::CompareOrdinal($this.Text, $this.Pos, '![CDATA[', 0, 8) -eq 0) {
                $this.Pos += 8
                return $this.ParseCdata()
            }
            if ($c1 -eq [char]'D' -and ($this.Len - $this.Pos) -ge 8 -and
                [string]::CompareOrdinal($this.Text, $this.Pos, '!DOCTYPE', 0, 8) -eq 0 -and
                [PsXmlParser]::IsWs($this.CharAt($this.Pos + 8))) {
                $this.Pos += 9
                return $this.ParseDoctype()
            }
            # Unbekanntes '<!...>' wird bis zum ersten '>' übersprungen.
            $idx = $this.FindToken('>', $this.Pos + 1)
            $this.Pos = $idx + 1
            return $null
        }
        return $this.ParseElement()
    }

    # ---- Einzelne Knotentypen ----------------------------------------------

    hidden [PsXmlNode] ParseXmlDeclaration() {
        if (-not $this.FDeclarationNode) {
            $idx = $this.FindToken('?>', $this.Pos)
            $this.Pos = $idx + 2
            return $null
        }
        $decl = [PsXmlNode]::new([PsXmlNodeType]::Declaration)
        $this.SkipWhitespace()
        $this.ParseNodeAttributes($decl)
        if ($this.CharAt($this.Pos) -ne [char]'?' -or $this.CharAt($this.Pos + 1) -ne [char]'>') {
            $this.Fail('expected ?>')
        }
        $this.Pos += 2
        return $decl
    }

    hidden [PsXmlNode] ParseComment() {
        $idx = $this.FindToken('-->', $this.Pos)
        if (-not $this.FCommentNodes) {
            $this.Pos = $idx + 3
            return $null
        }
        $n = [PsXmlNode]::new([PsXmlNodeType]::Comment)
        $n.Value = $this.Text.Substring($this.Pos, $idx - $this.Pos)
        $this.Pos = $idx + 3
        return $n
    }

    hidden [PsXmlNode] ParseCdata() {
        $idx = $this.FindToken(']]>', $this.Pos)
        if ($this.FNoDataNodes) {
            $this.Pos = $idx + 3
            return $null
        }
        $n = [PsXmlNode]::new([PsXmlNodeType]::Cdata)
        $n.Value = $this.Text.Substring($this.Pos, $idx - $this.Pos)
        $this.Pos = $idx + 3
        return $n
    }

    hidden [PsXmlNode] ParseDoctype() {
        $start = $this.Pos
        $i = $this.Pos
        $end = -1
        while ($true) {
            $j = $this.Text.IndexOfAny([PsXmlParser]::DoctypeStops, $i)
            if ($j -lt 0) { $this.Fail('unexpected end of data', $this.Len) }
            $c = $this.Text[$j]
            if ($c -eq [char]0) { $this.Fail('unexpected end of data', $j) }
            if ($c -eq [char]'>') { $end = $j; break }
            # '[' – interne Teilmenge mit Klammertiefe überspringen
            $depth = 1
            $k = $j + 1
            while ($depth -gt 0) {
                $m = $this.Text.IndexOfAny([PsXmlParser]::BracketStops, $k)
                if ($m -lt 0) { $this.Fail('unexpected end of data', $this.Len) }
                $bc = $this.Text[$m]
                if ($bc -eq [char]0) { $this.Fail('unexpected end of data', $m) }
                if ($bc -eq [char]'[') { $depth++ } else { $depth-- }
                $k = $m + 1
            }
            $i = $k
        }
        if ($this.FDoctypeNode) {
            $n = [PsXmlNode]::new([PsXmlNodeType]::Doctype)
            $n.Value = $this.Text.Substring($start, $end - $start)
            $this.Pos = $end + 1
            return $n
        }
        $this.Pos = $end + 1
        return $null
    }

    hidden [PsXmlNode] ParsePi() {
        if ($this.FPiNodes) {
            $nameStart = $this.Pos
            $nameEnd = $this.FindAny([PsXmlParser]::NodeNameStops, $this.Pos)
            if ($nameEnd -eq $nameStart) { $this.Fail('expected PI target') }
            $n = [PsXmlNode]::new([PsXmlNodeType]::Pi)
            $n.Name = $this.Text.Substring($nameStart, $nameEnd - $nameStart)
            $this.Pos = $nameEnd
            $this.SkipWhitespace()
            $idx = $this.FindToken('?>', $this.Pos)
            $n.Value = $this.Text.Substring($this.Pos, $idx - $this.Pos)
            $this.Pos = $idx + 2
            return $n
        }
        $idx = $this.FindToken('?>', $this.Pos)
        $this.Pos = $idx + 2
        return $null
    }

    hidden [PsXmlNode] ParseElement() {
        $n = [PsXmlNode]::new([PsXmlNodeType]::Element)
        $nameStart = $this.Pos
        $nameEnd = $this.FindAny([PsXmlParser]::NodeNameStops, $this.Pos)
        if ($nameEnd -eq $nameStart) { $this.Fail('expected element name') }
        $n.Name = $this.Text.Substring($nameStart, $nameEnd - $nameStart)
        $this.Pos = $nameEnd
        $this.SkipWhitespace()
        $this.ParseNodeAttributes($n)
        $c = $this.CharAt($this.Pos)
        if ($c -eq [char]'>') {
            $this.Pos++
            $this.ParseNodeContents($n)
        }
        elseif ($c -eq [char]'/') {
            $this.Pos++
            if ($this.CharAt($this.Pos) -ne [char]'>') { $this.Fail('expected >') }
            $this.Pos++
        }
        else {
            $this.Fail('expected >')
        }
        return $n
    }

    hidden [void] ParseNodeContents([PsXmlNode]$node) {
        while ($true) {
            $contentsStart = $this.Pos
            $this.SkipWhitespace()
            $c = $this.CharAt($this.Pos)
            if ($c -eq [char]'<') {
                if ($this.CharAt($this.Pos + 1) -eq [char]'/') {
                    # Schließendes Tag
                    $this.Pos += 2
                    if ($this.FValidateClosingTags) {
                        $nameStart = $this.Pos
                        $nameEnd = $this.FindAny([PsXmlParser]::NodeNameStops, $this.Pos)
                        $closing = $this.Text.Substring($nameStart, $nameEnd - $nameStart)
                        if (-not [string]::Equals($closing, $node.Name, [System.StringComparison]::Ordinal)) {
                            # rapidxml meldet die Position NACH dem Namen
                            $this.Fail('invalid closing tag name', $nameEnd)
                        }
                        $this.Pos = $nameEnd
                    }
                    else {
                        $this.Pos = $this.FindAny([PsXmlParser]::NodeNameStops, $this.Pos)
                    }
                    $this.SkipWhitespace()
                    if ($this.CharAt($this.Pos) -ne [char]'>') { $this.Fail('expected >') }
                    $this.Pos++
                    return
                }
                $this.Pos++
                $child = $this.ParseNode()
                if ($null -ne $child) { $node.AppendNode($child) }
            }
            elseif ($c -eq [char]0) {
                $this.Fail('unexpected end of data')
            }
            else {
                $this.ParseAndAppendData($node, $contentsStart)
            }
        }
    }

    hidden [void] ParseNodeAttributes([PsXmlNode]$node) {
        while ($this.IsAttNameChar($this.CharAt($this.Pos))) {
            $nameStart = $this.Pos
            $nameEnd = $this.FindAny([PsXmlParser]::AttNameStops, $this.Pos)
            $attr = [PsXmlAttribute]::new()
            $attr.Name = $this.Text.Substring($nameStart, $nameEnd - $nameStart)
            $node.AppendAttribute($attr)
            $this.Pos = $nameEnd
            $this.SkipWhitespace()
            if ($this.CharAt($this.Pos) -ne [char]'=') { $this.Fail('expected =') }
            $this.Pos++
            $this.SkipWhitespace()
            $quote = $this.CharAt($this.Pos)
            if ($quote -ne [char]39 -and $quote -ne [char]34) { $this.Fail("expected ' or `"") }
            $this.Pos++
            $attr.Value = $this.ScanAttributeValue($quote)
            if ($this.CharAt($this.Pos) -ne $quote) { $this.Fail("expected ' or `"") }
            $this.Pos++
            $this.SkipWhitespace()
        }
    }

    # ---- Text-/Wertverarbeitung --------------------------------------------

    hidden [void] ParseAndAppendData([PsXmlNode]$node, [int]$contentsStart) {
        # Ohne TrimWhitespace gehört der führende Leerraum mit zum Datenwert.
        if (-not $this.FTrimWhitespace) { $this.Pos = $contentsStart }
        $value = $this.ScanData()
        if ($this.FTrimWhitespace) {
            if ($this.FNormalizeWhitespace) {
                # Läufe wurden bereits zu je einem ' ' kondensiert
                if ($value.Length -gt 0 -and $value[$value.Length - 1] -eq [char]32) {
                    $value = $value.Substring(0, $value.Length - 1)
                }
            }
            else {
                $value = $value.TrimEnd([PsXmlParser]::WsChars)
            }
        }
        # Elementwert = erster Datenabschnitt (spätere überschreiben nicht)
        if (-not $this.FNoElementValues -and $node.Value.Length -eq 0) { $node.Value = $value }
        if (-not $this.FNoDataNodes) {
            $d = [PsXmlNode]::new([PsXmlNodeType]::Data)
            $d.Value = $value
            $node.AppendNode($d)
        }
    }

    hidden [string] ScanData() {
        $translate = -not $this.FNoEntityTranslation
        $normalize = $this.FNormalizeWhitespace
        if (-not $translate -and -not $normalize) {
            $idx = $this.FindAny([PsXmlParser]::TextStops, $this.Pos)
            $v = $this.Text.Substring($this.Pos, $idx - $this.Pos)
            $this.Pos = $idx
            return $v
        }
        $stops = [PsXmlParser]::TextStopsEnt
        if ($normalize) {
            if ($translate) { $stops = [PsXmlParser]::TextStopsNorm }
            else            { $stops = [PsXmlParser]::TextStopsNormNoEnt }
        }
        $sb = [System.Text.StringBuilder]::new()
        while ($true) {
            $idx = $this.FindAny($stops, $this.Pos)
            if ($idx -gt $this.Pos) { $sb.Append($this.Text, $this.Pos, $idx - $this.Pos) }
            $this.Pos = $idx
            $c = $this.CharAt($idx)
            if ($c -eq [char]0 -or $c -eq [char]'<') { break }
            if ($normalize -and [PsXmlParser]::IsWs($c)) {
                $sb.Append(' ')
                $this.SkipWhitespace()
                continue
            }
            $this.TranslateEntity($sb)
        }
        return $sb.ToString()
    }

    hidden [string] ScanAttributeValue([char]$quote) {
        $translate = -not $this.FNoEntityTranslation
        $stops = $null
        if ($quote -eq [char]34) {
            if ($translate) { $stops = [PsXmlParser]::AttValStopsDqEnt }
            else            { $stops = [PsXmlParser]::AttValStopsDq }
        }
        else {
            if ($translate) { $stops = [PsXmlParser]::AttValStopsSqEnt }
            else            { $stops = [PsXmlParser]::AttValStopsSq }
        }
        $sb = $null
        while ($true) {
            $idx = $this.FindAny($stops, $this.Pos)
            $c = $this.CharAt($idx)
            if ($c -ne [char]'&') {
                # Schlusszeichen (Quote) oder Datenende – Prüfung macht der Aufrufer
                if ($null -eq $sb) {
                    $v = $this.Text.Substring($this.Pos, $idx - $this.Pos)
                    $this.Pos = $idx
                    return $v
                }
                $sb.Append($this.Text, $this.Pos, $idx - $this.Pos)
                $this.Pos = $idx
                return $sb.ToString()
            }
            if ($null -eq $sb) { $sb = [System.Text.StringBuilder]::new() }
            $sb.Append($this.Text, $this.Pos, $idx - $this.Pos)
            $this.Pos = $idx
            $this.TranslateEntity($sb)
        }
        return ''
    }

    # Prüft case-sensitiv (ordinal), ob $token an Position $index steht.
    hidden [bool] MatchesAt([int]$index, [string]$token) {
        if ($index + $token.Length -gt $this.Len) { return $false }
        return [string]::CompareOrdinal($this.Text, $index, $token, 0, $token.Length) -eq 0
    }

    # Übersetzt eine Entity an der aktuellen Position (Text[Pos] = '&').
    # Unbekannte Sequenzen werden wie bei rapidxml unverändert übernommen.
    # Vergleiche strikt case-sensitiv: '&AMP;' u. ä. bleiben verbatim.
    hidden [void] TranslateEntity([System.Text.StringBuilder]$sb) {
        $p = $this.Pos
        $c1 = [int]$this.CharAt($p + 1)
        if ($c1 -eq 97) {        # 'a'
            if ($this.MatchesAt($p, '&amp;'))  { $sb.Append('&'); $this.Pos = $p + 5; return }
            if ($this.MatchesAt($p, '&apos;')) { $sb.Append("'"); $this.Pos = $p + 6; return }
        }
        elseif ($c1 -eq 113) {   # 'q'
            if ($this.MatchesAt($p, '&quot;')) { $sb.Append('"'); $this.Pos = $p + 6; return }
        }
        elseif ($c1 -eq 103) {   # 'g'
            if ($this.MatchesAt($p, '&gt;'))   { $sb.Append('>'); $this.Pos = $p + 4; return }
        }
        elseif ($c1 -eq 108) {   # 'l'
            if ($this.MatchesAt($p, '&lt;'))   { $sb.Append('<'); $this.Pos = $p + 4; return }
        }
        elseif ($c1 -eq 35) {    # '#' – numerische Entity
            # Ziffern wie rapidxml über die Digit-Tabelle (auch im Dezimal-
            # modus gelten A-F/a-f als Ziffern 10-15); Überlauf wickelt wie
            # ein 32-Bit-unsigned-long (MSVC) modulo 2^32.
            $code = [long]0
            $q = 0
            # Hinweis: Modulo statt -band, da PS 5.1 das Literal 0xFFFFFFFF
            # als [int]-1 parst und -band damit wirkungslos wäre.
            if ([int]$this.CharAt($p + 2) -eq 120) {   # nur kleines 'x' wie rapidxml
                $q = $p + 3
                while ($true) {
                    $d = [PsXmlParser]::DigitValue($this.CharAt($q))
                    if ($d -eq 255) { break }
                    $code = ($code * 16 + $d) % 4294967296
                    $q++
                }
            }
            else {
                $q = $p + 2
                while ($true) {
                    $d = [PsXmlParser]::DigitValue($this.CharAt($q))
                    if ($d -eq 255) { break }
                    $code = ($code * 10 + $d) % 4294967296
                    $q++
                }
            }
            # Wie rapidxml: erst Zeichen einfügen (wirft ggf. 'invalid numeric
            # character entity'), danach das Semikolon verlangen.
            $this.Pos = $q
            $this.AppendCodedChar($sb, $code)
            if ([int]$this.CharAt($q) -ne 59) { $this.Fail('expected ;') }
            $this.Pos = $q + 1
            return
        }
        # Keine bekannte Entity: '&' unverändert übernehmen
        $sb.Append('&')
        $this.Pos = $p + 1
    }

    hidden [void] AppendCodedChar([System.Text.StringBuilder]$sb, [long]$code) {
        if ($this.FNoUtf8) {
            # rapidxml-Verhalten bei parse_no_utf8: nur das niederwertige Byte
            $sb.Append([char]($code -band 0xFF))
            return
        }
        if ($code -lt 0x10000) {
            $sb.Append([char]$code)
            return
        }
        if ($code -lt 0x110000) {
            $sb.Append([char]::ConvertFromUtf32([int]$code))
            return
        }
        $this.Fail('invalid numeric character entity')
    }
}

# ============================================================================
#  PsXmlPrinter  (interner Arbeiter; Port von rapidxml_print.hpp)
# ============================================================================

class PsXmlPrinter {
    hidden static [char[]]$EscapeChars = "<>'`"&".ToCharArray()

    # Einrückung wie rapidxml: ein Tab pro Ebene, entfällt bei NoIndenting.
    hidden static [void] Indent([System.Text.StringBuilder]$sb, [int]$flags, [int]$indentLevel) {
        if (($flags -band 1) -eq 0 -and $indentLevel -gt 0) {
            $sb.Append([char]9, $indentLevel)
        }
    }

    static [string] Print([PsXmlNode]$node, [int]$flags) {
        if ($null -eq $node) {
            throw [System.ArgumentNullException]::new('node', 'Es wurde kein Knoten übergeben.')
        }
        $sb = [System.Text.StringBuilder]::new()
        [PsXmlPrinter]::PrintNode($sb, $node, $flags, 0)
        return $sb.ToString()
    }

    hidden static [void] PrintNode([System.Text.StringBuilder]$sb, [PsXmlNode]$node, [int]$flags, [int]$indent) {
        $type = $node.Type
        if ($type -eq [PsXmlNodeType]::Document) {
            $child = $node.FirstNode()
            while ($null -ne $child) {
                [PsXmlPrinter]::PrintNode($sb, $child, $flags, $indent)
                $child = $child.NextSibling()
            }
        }
        elseif ($type -eq [PsXmlNodeType]::Element) {
            [PsXmlPrinter]::PrintElement($sb, $node, $flags, $indent)
        }
        elseif ($type -eq [PsXmlNodeType]::Data) {
            [PsXmlPrinter]::Indent($sb, $flags, $indent)
            [PsXmlPrinter]::AppendEscaped($sb, $node.Value, [char]0)
        }
        elseif ($type -eq [PsXmlNodeType]::Cdata) {
            [PsXmlPrinter]::Indent($sb, $flags, $indent)
            $sb.Append('<![CDATA[').Append($node.Value).Append(']]>')
        }
        elseif ($type -eq [PsXmlNodeType]::Comment) {
            [PsXmlPrinter]::Indent($sb, $flags, $indent)
            $sb.Append('<!--').Append($node.Value).Append('-->')
        }
        elseif ($type -eq [PsXmlNodeType]::Declaration) {
            [PsXmlPrinter]::Indent($sb, $flags, $indent)
            $sb.Append('<?xml')
            [PsXmlPrinter]::PrintAttributes($sb, $node)
            $sb.Append('?>')
        }
        elseif ($type -eq [PsXmlNodeType]::Doctype) {
            [PsXmlPrinter]::Indent($sb, $flags, $indent)
            $sb.Append('<!DOCTYPE ').Append($node.Value).Append('>')
        }
        elseif ($type -eq [PsXmlNodeType]::Pi) {
            [PsXmlPrinter]::Indent($sb, $flags, $indent)
            $sb.Append('<?').Append($node.Name).Append(' ').Append($node.Value).Append('?>')
        }
        # Wie rapidxml: nach jedem Knoten ein Zeilenumbruch (sofern eingerückt wird)
        if (($flags -band 1) -eq 0) { $sb.Append("`n") }
    }

    hidden static [void] PrintElement([System.Text.StringBuilder]$sb, [PsXmlNode]$node, [int]$flags, [int]$indent) {
        [PsXmlPrinter]::Indent($sb, $flags, $indent)
        $sb.Append('<').Append($node.Name)
        [PsXmlPrinter]::PrintAttributes($sb, $node)
        $first = $node.FirstNode()
        if ($node.Value.Length -eq 0 -and $null -eq $first) {
            # Kinderlos und ohne Wert: selbstschließend
            $sb.Append('/>')
            return
        }
        $sb.Append('>')
        if ($null -eq $first) {
            # Nur eigener Wert, inline
            [PsXmlPrinter]::AppendEscaped($sb, $node.Value, [char]0)
        }
        elseif ($null -eq $first.NextSibling() -and $first.Type -eq [PsXmlNodeType]::Data) {
            # Genau ein Daten-Kind: dessen Wert inline
            [PsXmlPrinter]::AppendEscaped($sb, $first.Value, [char]0)
        }
        else {
            # Kinder mit voller Einrückung
            if (($flags -band 1) -eq 0) { $sb.Append("`n") }
            $child = $first
            while ($null -ne $child) {
                [PsXmlPrinter]::PrintNode($sb, $child, $flags, $indent + 1)
                $child = $child.NextSibling()
            }
            [PsXmlPrinter]::Indent($sb, $flags, $indent)
        }
        $sb.Append('</').Append($node.Name).Append('>')
    }

    hidden static [void] PrintAttributes([System.Text.StringBuilder]$sb, [PsXmlNode]$node) {
        $a = $node.FirstAttribute()
        while ($null -ne $a) {
            $sb.Append(' ').Append($a.Name).Append('=')
            [PsXmlPrinter]::AppendQuotedAttributeValue($sb, $a.Value)
            $a = $a.NextAttribute()
        }
    }

    # Quote-Wahl wie rapidxml: enthält der Wert ein '"', wird mit '...'
    # umschlossen (und '"' bleibt roh), sonst mit "..." (und ' bleibt roh).
    static [void] AppendQuotedAttributeValue([System.Text.StringBuilder]$sb, [string]$value) {
        if ($null -eq $value) { $value = '' }
        if ($value.IndexOf([char]34) -ge 0) {
            $sb.Append([char]39)
            [PsXmlPrinter]::AppendEscaped($sb, $value, [char]34)
            $sb.Append([char]39)
        }
        else {
            $sb.Append([char]34)
            [PsXmlPrinter]::AppendEscaped($sb, $value, [char]39)
            $sb.Append([char]34)
        }
    }

    # Escaping wie copy_and_expand_chars: < > ' " & werden ersetzt,
    # das Zeichen $noExpand bleibt unverändert.
    hidden static [void] AppendEscaped([System.Text.StringBuilder]$sb, [string]$s, [char]$noExpand) {
        if ([string]::IsNullOrEmpty($s)) { return }
        $pos = 0
        while ($true) {
            $idx = $s.IndexOfAny([PsXmlPrinter]::EscapeChars, $pos)
            if ($idx -lt 0) {
                $sb.Append($s, $pos, $s.Length - $pos)
                break
            }
            if ($idx -gt $pos) { $sb.Append($s, $pos, $idx - $pos) }
            $c = $s[$idx]
            if ($c -eq $noExpand) { $sb.Append($c) }
            elseif ($c -eq [char]'<') { $sb.Append('&lt;') }
            elseif ($c -eq [char]'>') { $sb.Append('&gt;') }
            elseif ($c -eq [char]39)  { $sb.Append('&apos;') }
            elseif ($c -eq [char]34)  { $sb.Append('&quot;') }
            else { $sb.Append('&amp;') }
            $pos = $idx + 1
        }
    }
}

# ============================================================================
#  Funktionen (Werkzeuge)
# ============================================================================

function New-PsXmlDocument {
    <#
    .SYNOPSIS
        Erstellt ein neues (optional sofort geparstes) PsXml-Dokument.
    .PARAMETER Xml
        XML-Text, der sofort geparst wird. Ohne Angabe entsteht ein leeres Dokument.
    .PARAMETER Flags
        Parse-Flags, z. B. [PsXmlParseFlags]::Full oder
        [PsXmlParseFlags]'CommentNodes, TrimWhitespace'.
    .EXAMPLE
        $doc = New-PsXmlDocument -Xml '<a><b>1</b></a>'
    #>
    [CmdletBinding()]
    [OutputType([PsXmlDocument])]
    param(
        [Parameter(Position = 0, ValueFromPipeline)]
        [AllowEmptyString()]
        [string]$Xml,

        [Parameter()]
        [PsXmlParseFlags]$Flags = [PsXmlParseFlags]::Default
    )
    process {
        $doc = [PsXmlDocument]::new()
        if ($PSBoundParameters.ContainsKey('Xml') -and -not [string]::IsNullOrEmpty($Xml)) {
            $doc.Parse($Xml, $Flags)
        }
        return $doc
    }
}

function Import-PsXml {
    <#
    .SYNOPSIS
        Liest eine XML-Datei ein und liefert ein PsXml-Dokument
        (Pendant zu rapidxml::file + parse).
    .PARAMETER Path
        Pfad zur XML-Datei (relativ zum aktuellen PowerShell-Speicherort).
    .PARAMETER Flags
        Parse-Flags (Standard: [PsXmlParseFlags]::Default).
    .PARAMETER Encoding
        Zeichenkodierung der Datei. 'Auto' (Standard) erkennt die Kodierung
        per BOM und nimmt sonst UTF-8 an.
    .EXAMPLE
        $doc = Import-PsXml -Path .\daten.xml -Flags ([PsXmlParseFlags]::Full)
    #>
    [CmdletBinding()]
    [OutputType([PsXmlDocument])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName', 'PSPath')]
        [string]$Path,

        [Parameter()]
        [PsXmlParseFlags]$Flags = [PsXmlParseFlags]::Default,

        [Parameter()]
        [ValidateSet('Auto', 'UTF8', 'UTF8BOM', 'UTF16LE', 'UTF16BE', 'ASCII', 'Latin1', 'Default')]
        [string]$Encoding = 'Auto'
    )
    process {
        $resolved = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)
        if ($Encoding -eq 'Auto') {
            $text = [System.IO.File]::ReadAllText($resolved)
        }
        else {
            $text = [System.IO.File]::ReadAllText($resolved, (Resolve-PsXmlEncoding -Name $Encoding))
        }
        $doc = [PsXmlDocument]::new()
        $doc.Parse($text, $Flags)
        return $doc
    }
}

function Export-PsXml {
    <#
    .SYNOPSIS
        Serialisiert ein PsXml-Dokument (oder einen Teilbaum) in eine Datei.
    .PARAMETER Node
        Das Dokument oder der Knoten, der geschrieben werden soll.
    .PARAMETER Path
        Zielpfad (relativ zum aktuellen PowerShell-Speicherort).
    .PARAMETER NoIndenting
        Ausgabe ohne Einrückung und Zeilenumbrüche (print_no_indenting).
    .PARAMETER Encoding
        Zeichenkodierung der Ausgabedatei (Standard: UTF8 ohne BOM).
        Hinweis: Eine encoding-Angabe in einer <?xml ...?>-Deklaration wird
        nicht automatisch angepasst.
    .EXAMPLE
        Export-PsXml -Node $doc -Path .\ausgabe.xml
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [Alias('Document')]
        [PsXmlNode]$Node,

        [Parameter(Mandatory, Position = 1)]
        [string]$Path,

        [Parameter()]
        [switch]$NoIndenting,

        [Parameter()]
        [ValidateSet('UTF8', 'UTF8BOM', 'UTF16LE', 'UTF16BE', 'ASCII', 'Latin1', 'Default')]
        [string]$Encoding = 'UTF8'
    )
    process {
        $flags = 0
        if ($NoIndenting) { $flags = 1 }
        $xml = [PsXmlPrinter]::Print($Node, $flags)
        $resolved = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)
        [System.IO.File]::WriteAllText($resolved, $xml, (Resolve-PsXmlEncoding -Name $Encoding))
    }
}

function ConvertTo-PsXmlString {
    <#
    .SYNOPSIS
        Serialisiert einen PsXml-Knoten als XML-Text
        (Pendant zu rapidxml::print).
    .PARAMETER Node
        Der zu serialisierende Knoten (Dokument, Element, ...).
    .PARAMETER NoIndenting
        Ausgabe ohne Einrückung und Zeilenumbrüche (print_no_indenting).
    .EXAMPLE
        $doc | ConvertTo-PsXmlString -NoIndenting
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [PsXmlNode]$Node,

        [Parameter()]
        [switch]$NoIndenting
    )
    process {
        $flags = 0
        if ($NoIndenting) { $flags = 1 }
        return [PsXmlPrinter]::Print($Node, $flags)
    }
}

function Select-PsXmlNode {
    <#
    .SYNOPSIS
        Einfache Pfadsuche über Elementnamen, z. B. 'katalog/cd/titel'.
    .DESCRIPTION
        Jedes Pfadsegment steigt eine Ebene in den Kindknoten hinab.
        '*' steht für ein beliebiges Element. Zurückgegeben werden alle
        Knoten, die dem vollständigen Pfad entsprechen (kein XPath).
    .PARAMETER Node
        Startknoten (z. B. das Dokument).
    .PARAMETER Path
        Pfad aus Elementnamen, getrennt durch '/'.
    .PARAMETER CaseInsensitive
        Namen ohne Beachtung der Groß-/Kleinschreibung vergleichen
        (wie rapidxml nur für ASCII).
    .EXAMPLE
        Select-PsXmlNode $doc 'katalog/cd/titel' | ForEach-Object Value
    #>
    [CmdletBinding()]
    [OutputType([PsXmlNode])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [PsXmlNode]$Node,

        [Parameter(Mandatory, Position = 1)]
        [string]$Path,

        [Parameter()]
        [switch]$CaseInsensitive
    )
    process {
        $caseSensitive = -not $CaseInsensitive.IsPresent
        $current = [System.Collections.Generic.List[object]]::new()
        $current.Add($Node)
        foreach ($segment in $Path.Trim('/').Split('/')) {
            if ($segment -eq '') { continue }
            $next = [System.Collections.Generic.List[object]]::new()
            foreach ($n in $current) {
                $child = $n.FirstNode()
                while ($null -ne $child) {
                    if ($child.Type -eq [PsXmlNodeType]::Element) {
                        if ($segment -eq '*' -or [PsXmlBase]::NamesEqual($child.Name, $segment, $caseSensitive)) {
                            $next.Add($child)
                        }
                    }
                    $child = $child.NextSibling()
                }
            }
            $current = $next
            if ($current.Count -eq 0) { break }
        }
        return $current
    }
}

function Resolve-PsXmlEncoding {
    <#
    .SYNOPSIS
        Interner Helfer: bildet einen Kodierungsnamen auf System.Text.Encoding ab.
    #>
    [CmdletBinding()]
    [OutputType([System.Text.Encoding])]
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )
    switch ($Name) {
        'UTF8'    { return [System.Text.UTF8Encoding]::new($false) }
        'UTF8BOM' { return [System.Text.UTF8Encoding]::new($true) }
        'UTF16LE' { return [System.Text.Encoding]::Unicode }
        'UTF16BE' { return [System.Text.Encoding]::BigEndianUnicode }
        'ASCII'   { return [System.Text.Encoding]::ASCII }
        'Latin1'  { return [System.Text.Encoding]::GetEncoding(28591) }
        'Default' { return [System.Text.Encoding]::Default }
    }
    return [System.Text.UTF8Encoding]::new($false)
}
