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
    public enum relshpids : short { wdrelMainDocPart = 1, wdrelHeaderPart = 2, wdrelFooterPart = 3, wdrelFootnotePart = 4, wdrelStylesPart = 5, wdrelGlossaryDocPart = 6 };
    class CBDocument
    {
        private string pathtemplatedoc;
        private string pathnewdoc;
        private const string TEMPLATE_NAME = "SoleSourceLetter v54.dotx";
        private const string DOCNAME_GENERATED = "CB Generated Document.docx";

        public CBDocument()
        {
            pathtemplatedoc = @"C:\Users\ajones\Documents\Visual Studio 2015\Operation Kyuzo\Prototypes and Study\Templates\";
            pathnewdoc = @"C:\Users\ajones\Documents\Visual Studio 2015\Operation Kyuzo\Prototypes and Study\Documents Generated\";
        }

        public CBDocument(string Path_Template)
        {
            pathtemplatedoc = Path_Template;
            pathnewdoc = @"C:\Users\ajones\Documents\Visual Studio 2015\Operation Kyuzo\Prototypes and Study\Documents Generated\";
        }

        public CBDocument(string Path_Template, string Path_NewDoc)
        {
            pathtemplatedoc = Path_Template;
            pathnewdoc = Path_NewDoc;
        }

        public void CreateNewSimpleDoc()
        {
            const string DOCNAME_SIMPLE = "Simple Doc.docx";
            using (WordprocessingDocument wrdNewDoc = WordprocessingDocument.Create(pathnewdoc + DOCNAME_SIMPLE, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mdpNewDoc = wrdNewDoc.AddMainDocumentPart();
                mdpNewDoc.Document = new Document(new Body(new Paragraph(new Run(new Text("This is a New Simple Document")))));
            }
        }

        public void CreateCBDocumentFromCBTemplate()
        {
            // Get the template
            using (WordprocessingDocument SourceTemplate = WordprocessingDocument.Open(pathtemplatedoc + TEMPLATE_NAME, false))
            // Create a new document
            using (WordprocessingDocument NewDoc = WordprocessingDocument.Create(pathnewdoc + DOCNAME_GENERATED, WordprocessingDocumentType.Document))
            {
                /* Get the main boilerplate text from the template for the new document */
                MainDocumentPart mdpNewDoc = NewDoc.AddMainDocumentPart();
                mdpNewDoc.FeedData(SourceTemplate.MainDocumentPart.GetStream());
                NewDoc.ChangeIdOfPart(mdpNewDoc, "rId" + (short)relshpids.wdrelMainDocPart);

                foreach (IdPartPair sourceTempMainDocPart in SourceTemplate.MainDocumentPart.Parts)
                {
                    // Transfer a select set of parts to the new document from the template
                    string PartName = "partname";
                    PartName = sourceTempMainDocPart.OpenXmlPart.GetType().Name;
                    if (IsCBDocumentPart(PartName))
                    {
                        OpenXmlPart NewPart = mdpNewDoc.AddPart(sourceTempMainDocPart.OpenXmlPart, AssignRelID(PartName));
                        NewPart.FeedData(sourceTempMainDocPart.OpenXmlPart.GetStream());
                    }
                }

                // Update the relationship ids
                // Get section and update Header and Footer relationship IDs
                Document doc = mdpNewDoc.Document;
                SectionProperties secprps = doc.Body.Elements<SectionProperties>().First<SectionProperties>();
                foreach (HeaderFooterReferenceType hfref in secprps.Elements<HeaderFooterReferenceType>())
                {
                    if (hfref.LocalName.Contains("header"))
                    {
                        hfref.Id = "rId" + (short)relshpids.wdrelHeaderPart;
                    }

                    if (hfref.LocalName.Contains("footer"))
                    {
                        hfref.Id = "rId" + (short)relshpids.wdrelFooterPart;
                    }
                }
            }
        }

        public void GetPicFromGlossary()
        {
            using (WordprocessingDocument wrdTemplate = WordprocessingDocument.Open(pathtemplatedoc + TEMPLATE_NAME, false))
            {
                GlossaryDocumentPart glsDocPart = wrdTemplate.MainDocumentPart.GlossaryDocumentPart;
                if (glsDocPart !=null)
                {
                    foreach (ImagePart img in glsDocPart.ImageParts)
                    {
                        Console.WriteLine("Relationship Id ==> {0}\tUri ==> {1}", glsDocPart.GetIdOfPart(img), img.Uri);
                    }
                }
            }
        }

        private bool IsCBDocumentPart(string CBDocumentPart)
        {
            switch (CBDocumentPart)
            {
                case "HeaderPart":
                case "FooterPart":
                case "FootnotesPart":
                case "StyleDefinitionsPart":
                case "ImagePart":
                case "GlossaryDocumentPart":
                    return true;
                default:
                    return false;
            }
        }

        private string AssignRelID(string CBDocumentPart)
        {
            const string RELID_PREFIX = "rId";
            switch (CBDocumentPart)
            {
                case "HeaderPart":
                    return RELID_PREFIX + (short)relshpids.wdrelHeaderPart;
                case "FooterPart":
                    return RELID_PREFIX + (short)relshpids.wdrelFooterPart;
                case "FootnotesPart":
                    return RELID_PREFIX + (short)relshpids.wdrelFootnotePart;
                case "StyleDefinitionsPart":
                    return RELID_PREFIX + (short)relshpids.wdrelStylesPart;
                case "GlossaryDocumentPart":
                    return RELID_PREFIX + (short)relshpids.wdrelGlossaryDocPart;
                default:
                    return RELID_PREFIX + 0;
            }
        }
    }
}
