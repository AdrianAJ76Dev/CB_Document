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
            Console.WriteLine("Creating Simple New Document!");
            CBDocument newdoc = new CBDocument();
            newdoc.CreateNewSimpleDoc();
            Console.WriteLine("Done!");
            Console.ReadLine();
        }
    }
}
