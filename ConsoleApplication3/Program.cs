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
            // Create a new document.
            using (DocX document = DocX.Create(@"Test.docx"))
            {
                document.InsertParagraph("I cant believe this took so long.");

                // Save the document.
                document.Save();
            }
        }
    }
}
