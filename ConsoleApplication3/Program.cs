using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Novacode;
using System.Drawing;

namespace ConsoleApplication3
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a document.
            using (DocX document = DocX.Load(@"Test.docx"))
            {
                document.ReplaceText("Hio", "World");
                document.SaveAs("Test2.docx");
            }// Release this document from memory.
        }
    }
}
