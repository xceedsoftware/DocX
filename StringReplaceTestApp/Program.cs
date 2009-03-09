using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Novacode;
using System.Text.RegularExpressions;
using System.IO;

namespace StringReplaceTestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            File.Copy(@"Test.docx", "Manipulated.docx", true);

            // Load the document that you want to manipulate
            DocX document = DocX.Load(@"Manipulated.docx");
            
            // Loop through the paragraphs in the document
            foreach (Paragraph p in document.Paragraphs)
            {
                /* 
                 * Replace each instance of the string pear with the string banana.
                 * Specifying true as the third argument informs DocX to track the
                 * changes made by this replace. The fourth argument tells DocX to
                 * ignore case when matching the string pear.
                 */

                p.Replace("pear", "banana", true, RegexOptions.IgnoreCase);
            }

            // File will be saved to \StringReplaceTestApp\bin\Debug
            document.Save();
        }
    }
}
