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
                // Add Headers to the document.
                document.AddHeaders();

                // Get the default Header.
                Header header = document.Headers.odd;

                // Insert a Paragraph into the Header.
                Paragraph p0 = header.InsertParagraph();

                // Append place holders for PageNumber and PageCount into the Header.
                // Word will replace these with the correct value foreach Page.
                p0.Append("Page (");
                p0.AppendPageNumber(PageNumberFormat.normal);
                p0.Append(" of ");
                p0.AppendPageCount(PageNumberFormat.normal);
                p0.Append(")");

                p0.ReplaceText("Page (", "Monster <");
                p0.ReplaceText(")", ">");

                // Save the document.
                document.Save();
            }
        }
    }
}
