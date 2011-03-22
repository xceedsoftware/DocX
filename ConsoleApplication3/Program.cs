using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Xml.Linq;
using Novacode;

namespace ConsoleApplication3
{
    class Program
    {
        static void Main(string[] args)
        {
            String filename = @"C:\Users\cathal\Downloads\ExternalAssessmentT.docx";

            using (DocX doc = DocX.Load(filename))
            {
                doc.ReplaceText("{Company Name}", "Penn Inc.", false);
                doc.ReplaceText("{Primary Contact}", "John Smith", false);

                doc.AddCoreProperty("dc:subject", "CLE-OP55555");
       
                doc.SaveAs(@"C:\Users\cathal\Downloads\ExternalAssessmentT2.docx");
            }
        }
    }
}
