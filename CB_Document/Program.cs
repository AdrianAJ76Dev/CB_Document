using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CB_Document
{
    class Program
    {
        static void Main(string[] args)
        {
            CBDocument newdoc = new CBDocument();
            /*
            Console.WriteLine("Creating Simple New Document!");
            newdoc.CreateNewSimpleDoc();
            Console.WriteLine("Creating New Document from Template!");
            newdoc.CreateCBDocumentFromCBTemplate();
            Console.WriteLine("Done!");
            newdoc.GetPicFromGlossary();
            newdoc.InvestigateGlossaryDocumentPart();
            newdoc.InvestigateTemplate();
            newdoc.CreateDocUseOuterXml();
            */
            newdoc.InsertAutoText("CS_Signature");
            Console.ReadLine();
        }
    }
}
