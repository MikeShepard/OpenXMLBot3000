import-module OpenXMLBot3000 -force


$doc = ExcelDocument 'c:\temp\testspreadsheet.xlsx' {
    Sheets {
        Sheet {
                     SheetData {
            #             Row 1 {
            #                 Cell A1 {
            #                     CellValue 
            #                 }             
            #             }
                     }
        }
    }
} 

$doc.o.Close()