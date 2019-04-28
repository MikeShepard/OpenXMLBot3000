add-type -path $PSScriptRoot\DocumentFormat.OpenXml.dll

$code=@'
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;

    namespace OpenXMLBot3000 {
    public class OpenXML 
    {
        public static void ReplaceTokens(string filename,Dictionary<string, string> tokens)
        {
          using (WordprocessingDocument w = WordprocessingDocument.Open(filename, true)){

            var texts= w.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>(); 
            foreach(var text in texts){
              foreach (var d in tokens) {
                if(text.Text.Contains(d.Key)){
                  //System.Console.WriteLine("Replacing {0} with {1} in {2}",d.Key,d.Value,text.Text);
                  text.Text=text.Text.Replace(d.Key,d.Value);
                }
              }
            }
            w.MainDocumentPart.PutXDocument();
          }
        }
        public static System.Collections.Generic.IEnumerable<DocumentFormat.OpenXml.Wordprocessing.Text> GetTexts(WordprocessingDocument w){
          return w.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>();
        }
        public static System.Collections.Generic.IEnumerable<Paragraph> GetParagraphs(WordprocessingDocument w){
          return w.MainDocumentPart.Document.Body.Descendants<Paragraph>();
        }
        public static System.Collections.Generic.IEnumerable<DocumentFormat.OpenXml.Wordprocessing.Run> GetRuns(WordprocessingDocument w){
          return w.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>();
        }
    }
    public static class OpenXMLBotExtensions{
      public static WorksheetPart AddWorksheetPart(this DocumentFormat.OpenXml.Packaging.WorkbookPart p){
        return p.AddNewPart<WorksheetPart>();
      }
      public static void PutXDocument(this OpenXmlPart part)
      {
          if (part == null) throw new ArgumentNullException("part");

          XDocument partXDocument = part.GetXDocument();
          if (partXDocument != null)
          {
              using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
              using (System.Xml.XmlWriter partXmlWriter = System.Xml.XmlWriter.Create(partStream))
                  partXDocument.Save(partXmlWriter);
          }
      }
      public static XDocument GetXDocument(this OpenXmlPart part)
        {
            if (part == null) throw new ArgumentNullException("part");

            XDocument partXDocument = part.Annotation<XDocument>();
            if (partXDocument != null) return partXDocument;

            using (Stream partStream = part.GetStream())
            {
                if (partStream.Length == 0)
                {
                    partXDocument = new XDocument();
                    partXDocument.Declaration = new XDeclaration("1.0", "UTF-8", "yes");
                }
                else
                {
                    using (XmlReader partXmlReader = XmlReader.Create(partStream))
                        partXDocument = XDocument.Load(partXmlReader);
                }
            }

            part.AddAnnotation(partXDocument);
            return partXDocument;
        }

    }
}
'@

add-type -TypeDefinition $code -ReferencedAssemblies WindowsBase, 'System.Xml.Linq','System.Xml',$PSScriptRoot\DocumentFormat.OpenXml.dll
