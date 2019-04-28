import-module OpenXMLBot3000 -force


$doc=WordDocument 'c:\temp\testdoc.docx' {
    Body {
        Paragraph {
             Run {
                RunProperties {
                    Underline DashDotHeavy
                }
                   Text 'Hello World'
            }
        }
    }
} 

$doc.o.Close() 