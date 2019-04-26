import-module OpenXMLBot3000 -force


$doc=WordDocument 'c:\temp\testdoc.docx' {
    Body {
        Paragraph {
            Run {
                Text 'Hello World'
            }
        }
    }
} 

$doc.o.Close()