using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DRW = DocumentFormat.OpenXml.Drawing;
using System.IO;

namespace CB_Document
{
    public enum CBDocumentBasicParts : short { wdMainPart = 1, wdHeaderPart = 2, wdFooterPart = 3, wdFootnotesPart = 4, wdStylesPart = 5 };
    class CBDocument
    {
        private CBDocumentBasicParts wdparts;
        private WordprocessingDocument newdoc;
        private WordprocessingDocument templatedoc;
        private Drawing signature_image;
        private string pathtemplatedoc;
        private string pathnewdoc;

        public CBDocument()
        {
            pathnewdoc = @"C:\Users\ajones\Documents\Visual Studio 2015\Operation Kyuzo\Prototypes and Study\Documents Generated\Simple Doc.docx";
        }

        public CBDocument(string Path_Template)
        {
            pathtemplatedoc = Path_Template;
        }

        /* Won't compile because the signature of course is the same (string x) 
         * x = Path_Template
         * x = Path_NewDoc
         * Different variables but both are string variables passed in so 
         * same signature.  Chaining contructors?
         */
        /*
        public CBDocument(string Path_NewDoc)
        {
            pathnewdoc = Path_NewDoc;
        }
        */

        public CBDocument(string Path_Template, string Path_NewDoc)
        {
            pathtemplatedoc = Path_Template;
            pathnewdoc = Path_NewDoc;
        }

        public void CreateNewSimpleDoc()
        {
            using (WordprocessingDocument wrdNewDoc = WordprocessingDocument.Create(pathnewdoc, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mdpNewDoc = wrdNewDoc.AddMainDocumentPart();
                mdpNewDoc.Document = new Document(new Body(new Paragraph(new Run(new Text("This is a Simple New Document")))));
            }
        }
    }
}
