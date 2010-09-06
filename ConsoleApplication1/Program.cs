using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Novacode;
using System.IO.Packaging;
using System.Diagnostics;
using System.Xml.Linq;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            using (DocX document = DocX.Load(@"D:\Source Control\DocX\UnitTests\documents\test.docx"))
            {
                Paragraph p = document.Paragraphs[0];
                p.ReplaceText("foo", "bar", false);

                document.SaveAs("output2.docx");
            }
        }
    }
}
