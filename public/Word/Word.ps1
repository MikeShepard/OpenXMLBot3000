using namespace DocumentFormat.OpenXml.Wordprocessing
using namespace DocumentFormat.OpenXml.Packaging
using namespace DocumentFormat.OpenXml
function Get-DocumentParagraph {
    param($Filename)
        
    $newdoc = [WordprocessingDocument]::Open($filename, $false)
    $body = $newdoc.MainDocumentPart.Document.body
    , $body.Clone()
}


function WordDocument {
    Param($filename, [Scriptblock]$contents = {})
    $doc = [WordprocessingDocument]::Create($filename, [WordprocessingDocumentType]::Document)
    $mainPart = $doc.AddMainDocumentPart()
    $mainPart.Document = new-object Document 
    $temp = OpenXMLItem -contents $contents  -object $doc.MainDocumentPart.Document
    [pscustomobject]@{type = 'Document'; O = $doc}
}

function Body {
    Param([Scriptblock]$tablecontents = {})
    $t = OpenXMLItem 'Body' $tablecontents
    [pscustomobject]@{type = 'Body'; O = $t}
}
function Table {
    Param([Scriptblock]$tablecontents = {})
    $t = OpenXMLItem 'Table' $tablecontents
    [pscustomobject]@{type = 'Table'; O = $t}
}
 
function TableRow {
    Param([Scriptblock]$rowcontents = {})
    $r = OpenXMLItem 'TableRow' $rowcontents
    [pscustomobject]@{type = 'Row'; O = $r}
}
 
function TableCell {
    Param($name, [Scriptblock]$cellcontents = {})
    $c = OpenXMLItem 'TableCell' $cellcontents
    [pscustomobject]@{type = 'Cell'; O = $c}
}
 
function TableBorders {
    Param([Scriptblock]$cellcontents = {})
    $c = OpenXMLItem 'TableBorders' $cellcontents
    [pscustomobject]@{type = 'TableBorders'; O = $c}
}
function TableProperties {
    Param([Scriptblock]$cellcontents = {})
    $c = OpenXMLItem 'TableProperties' $cellcontents
    [pscustomobject]@{type = 'TableProperties'; O = $c}
}
 
function TableStyle {
    Param($type)
    $c = OpenXMLItem TableStyle
    if ($type) {
        $c.Val = $type
    }
    [pscustomobject]@{type = 'TableStyle'; O = $c}
}
function TableLook {
    Param($val, $firstRow, $lastRow, $firstColumn, $lastColumn, $NohBand, $noVBand)
    $c = OpenXMLItem TableLook
    if ($val) {
        $c.Val = $val
    }
    if ($firstRow) {
        $c.FirstRow = $firstRow
    }
    if ($lastRow) {
        $c.lastRow = $lastRow
    }
    if ($firstColumn) {
        $c.firstRow = $firstColumn
    }
    if ($nohband) {
        $c.NoHorizontalBand = $nohband
    }
    if ($novband) {
        $c.NoVerticalBand = $novband
    }
  
    [pscustomobject]@{type = 'TableLook'; O = $c}
}

function TableLayout {
    Param([TableLayoutValues]$type)
    $c = OpenXMLItem TableLayout
    if ($type) {
        $c.type = [TableLayoutValues]$type
    }
    [pscustomobject]@{type = 'TableLayout'; O = $c}
}

function TableGrid {
    Param([Scriptblock]$contents = {})
    $c = OpenXMLItem TableGrid $contents
    
    [pscustomobject]@{type = 'TableGrid'; O = $c}
}

function GridColumn {
    Param([Scriptblock]$contents = {}, $width)
    $c = OpenXMLItem GridColumn $contents
    if ($width) {
        $c.width = $width
    }
    [pscustomobject]@{type = 'GridColumn'; O = $c}
}

function ParagraphStyleID {
    [CmdletBinding()]
    Param($style)
    $c = OpenXMLItem ParagraphStyleID 
    if ($style) {
        $c.val = $style
    }
    [pscustomobject]@{type = 'ParagraphStyleID'; O = $c}
}
function RunStyle {
    [CmdletBinding()]
    Param($style)
    $c = OpenXMLItem RunStyle 
    if ($val) {
        $c.style = $style
    }
    [pscustomobject]@{type = 'RunStyle'; O = $c}
}
                     
function FontSize {
    [CmdletBinding()]
    Param($val)
    $c = OpenXMLItem FontSize 
    if ($val) {
        $c.val = $val
    }
    [pscustomobject]@{type = 'FontSize'; O = $c}
}

function FontSizeComplexScript {
    [CmdletBinding()]
    Param($val)
    $c = OpenXMLItem FontSizeComplexScript 
    if ($val) {
        $c.val = $val
    }
    [pscustomobject]@{type = 'FontSizeComplexScript'; O = $c}
}
function GenericBorder {
    Param($bordertype, $val, $size)
    $c = OpenXMLItem $bordertype 
    if ($val) {
        $c.Val = [BorderValues]::$Val
    }
    if ($size) {
        $c.Size = [UInt32Value]::FromUInt32($size)
    }
 
    [pscustomobject]@{type = $borderType; O = $c}
}
 
function TopBorder {
    [CmdletBinding()]
    Param($val, $size)
    GenericBorder TopBorder @PSBoundParameters
}
 
function BottomBorder {
    [CmdletBinding()]
    Param($val, $size)
    GenericBorder BottomBorder @PSBoundParameters
}
 
function LeftBorder {
    [CmdletBinding()]
    Param($val, $size)
    GenericBorder LeftBorder @PSBoundParameters
}
 
function RightBorder {
    [CmdletBinding()]
    Param($val, $size)
    GenericBorder RightBorder @PSBoundParameters
}

function InsideHorizontalBorder {
    [CmdletBinding()]
    Param($val, $size)
    GenericBorder InsideHorizontalBorder @PSBoundParameters
}
 
function InsideVerticalBorder {
    [CmdletBinding()]
    Param($val, $size)
    GenericBorder InsideVerticalBorder @PSBoundParameters
}
function Paragraph {
    Param([Scriptblock]$contents = {})
    $c = OpenXMLItem Paragraph $contents
    
    [pscustomobject]@{type = 'Paragraph'; O = $c}
}
function Run {
    Param([Scriptblock]$contents = {})
    $c = OpenXMLItem Run $contents
    
    [pscustomobject]@{type = 'Run'; O = $c}
}

function Text {
    Param($text)
    $c = OpenXMLItem Text 
    $c.Text = $text
    [pscustomobject]@{type = 'Text'; O = $c}
}
function Bold {
    Param()
    $c = OpenXMLItem Bold
    [pscustomobject]@{type = 'Bold'; O = $c}
}
  
function PageBreakBefore {
    Param()
    $c = OpenXMLItem PageBreakBefore
    [pscustomobject]@{type = 'PageBreakBefore'; O = $c}
}

function RunProperties {
    Param([Scriptblock]$contents = {})
    $c = OpenXMLItem RunProperties $contents
    [pscustomobject]@{type = 'RunProperties'; O = $c}
}
function ParagraphProperties {
    Param([Scriptblock]$contents = {})
    $c = OpenXMLItem ParagraphProperties $contents
    [pscustomobject]@{type = 'ParagraphProperties'; O = $c}
}
