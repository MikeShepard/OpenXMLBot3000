using namespace DocumentFormat.OpenXml.Packaging
using namespace DocumentFormat.OpenXml
using namespace DocumentFormat.OpenXml.Spreadsheet


function ExcelDocument {
    Param($filename, [Scriptblock]$contents = {})
    $doc = [SpreadSheetDocument]::Create($filename, [SpreadSheetDocumentType]::Workbook)
    $workbookPart = $doc.AddWorkbookPart()
    $workbookPart.Workbook = new-object DocumentFormat.OpenXml.Spreadsheet.Workbook
    $worksheetPart = $workbookPart | Invoke-GenericMethod -MethodName AddNewPart -GenericType DocumentFormat.OpenXml.Packaging.WorksheetPart #p.AddNewPart<WorksheetPart>();
    $temp = OpenXMLItem -contents $contents  -object $doc.WorkbookPart.WorkBook
    [pscustomobject]@{type = 'WorkBook'; O = $doc}
}
  

function Sheets {
    Param([Scriptblock]$tablecontents = {})
    $t = OpenXMLItem 'DocumentFormat.OpenXml.Spreadsheet.Sheets' $tablecontents
    [pscustomobject]@{type = 'Sheets'; O = $t}
}

function Sheet {
    Param([Scriptblock]$tablecontents = {})
    $t = OpenXMLItem 'DocumentFormat.OpenXml.Spreadsheet.Sheet' $tablecontents
    [pscustomobject]@{type = 'Sheet'; O = $t}
 }

function Row {
    Param($name, [Scriptblock]$contents = {})
    $t = OpenXMLItem 'DocumentFormat.OpenXml.Spreadsheet.Row' $contents
    $t.RowIndex=$name
    [pscustomobject]@{type = 'Row'; O = $t}

}
function Cell {
    Param($name, [Scriptblock]$contents = {})
    $t = OpenXMLItem 'DocumentFormat.OpenXml.Spreadsheet.Cell' $contents
    $t.CellReference=$name
    [pscustomobject]@{type = 'Cell'; O = $t}
}

function SheetData {
    Param()
    $t = OpenXMLItem 'DocumentFormat.OpenXml.Spreadsheet.SheetData' 
    [pscustomobject]@{type = 'SheetData'; O = $t}
}

function CellValue {
    Param($Value)
    $t = OpenXMLItem 'DocumentFormat.OpenXml.Spreadsheet.CellValue' 
    $t.Text=$Value
    [pscustomobject]@{type = 'CellValue'; O = $t}
}