using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Xceed.Document.NET.Src;
using Xceed.Words.NET.Examples;

namespace Xceed.Words.NET.Examples
{
    public class EndnoteSample
    {
        #region Private Members

        private const string FootnoteSampleOutputDirectory = Program.SampleDirectory + @"FootnotesEndnotes\Output\";

        #endregion

        #region Constructors

        static EndnoteSample()
        {
            if (!Directory.Exists(FootnoteSampleOutputDirectory))
            {
                Directory.CreateDirectory(FootnoteSampleOutputDirectory);
            }
        }

        #endregion

        #region Public Methods

        public static void SimpleEndnote()
        {
            Console.WriteLine("\tSimpleEndnote()");
            string[] noteBrackets = new[] { "[", "]" };
            using (var document = DocX.Create(FootnoteSampleOutputDirectory + @"SimpleEndnote.docx"))
            {
                // Insert a Paragraph into this document.
                var p = document.InsertParagraph();

                // Append some text and add formatting.
                p.Append("This is a simple paragraph with an endnote.");
                // Append a footnote
                Endnote fn = new Endnote(document, "Make note of this source information.");
                fn.Apply(p);

                // new page, new para, append an endnote in the middle of text
                // with the optional []s around the number
                document.InsertSectionPageBreak();
                p = document.InsertParagraph();
                p.Append("This is another example, with brackets to set off the note number,");
                fn = new Endnote(document, "This source information is also noteworthy, and the note is made extra long in order to illustrate the default style of hanging indent; a human can easily edit the style in the output document (that's WHY we use styles!).", noteBrackets);
                fn.Apply(p);
                p.Append(" and so on to the end of the sentence.");

                // Save this document to disk.
                document.Save();
                Console.WriteLine("\tCreated: SimpleEndnote.docx\n");
            }

        }
        #endregion
    }
}
