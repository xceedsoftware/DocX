using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Xceed.Document.NET;
using Xceed.Document.NET.Src; 

namespace Xceed.Words.NET.Examples
{
    public class FootnoteSample
    {
        #region Private Members

        private const string FootnoteSampleOutputDirectory = Program.SampleDirectory + @"FootnotesEndnotes\Output\";

        #endregion

        #region Constructors

        static FootnoteSample()
        {
            if (!Directory.Exists(FootnoteSample.FootnoteSampleOutputDirectory))
            {
                Directory.CreateDirectory(FootnoteSample.FootnoteSampleOutputDirectory);
            }
        }

        #endregion

        #region Public Methods

        public static void SimpleFootnote()
        {
            Console.WriteLine("\tSimpleFootnote()");
            string[] noteBrackets = new[] {"[", "]"};
            using (var document = DocX.Create(FootnoteSampleOutputDirectory + @"SimpleFootnote.docx"))
            {
                // Insert a Paragraph into this document.
                var p = document.InsertParagraph();

                // Append some text and add formatting.
                p.Append("This is a simple paragraph with a footnote.");
                // Append a footnote
                Footnote fn = new Footnote(document, "Make note of this source information.");
                fn.Apply(p);

                // new para, append a footnote in the middle of text
                // with the optional []s around the number
                p = document.InsertParagraph();
                p.Append("This is another example, with brackets to set off the note number,");
                fn = new Footnote(document, "This source information is also noteworthy, and the note is made extra long in order to illustrate the default style of hanging indent; a human can easily edit the style in the output document (that's WHY we use styles!).", noteBrackets);
                fn.Apply(p);
                p.Append(" and so on to the end of the sentence.");

                // and another on the next page
                document.InsertSectionPageBreak();
                p = document.InsertParagraph();
                p.Append("This is another example, to show that footnotes appear,");
                fn = new Footnote(document, "This source is the best authority.");
                fn.Apply(p);
                p.Append(" on the same page not at the end.");

                // Save this document to disk.
                document.Save();
                Console.WriteLine("\tCreated: SimpleFootnote.docx\n");
            }

        }
        public static void BookmarkedFootnote()
        {
            Console.WriteLine("\tBookmarkedFootnote()");
            //Footnote.BookmarkReferencePattern = "See note {0}.";
            string[] noteBrackets = new[] { "[", "]" };
            using (var document = DocX.Create(FootnoteSampleOutputDirectory + @"BookmarkedFootnote.docx"))
            {
                // Insert a Paragraph into this document.
                var p = document.InsertParagraph();

                // Append some text and add formatting.
                p.Append("This is a simple paragraph with a footnote.");
                // Append a footnote
                Footnote fnRef = new Footnote(document, "Make note of this source information.");
                fnRef.Apply(p, true);

                // new para, append a footnote in the middle of text
                // with the optional []s around the number
                p = document.InsertParagraph();
                p.Append("This is another example, with brackets to set off the note number,");
                Footnote fn = new Footnote(document, "This source information is also noteworthy, and the note is made extra long in order to illustrate the default style of hanging indent; a human can easily edit the style in the output document (that's WHY we use styles!).", noteBrackets);
                fn.Apply(p);
                p.Append(" and so on to the end of the sentence.");

                // and another on the next page
                document.InsertSectionPageBreak();
                p = document.InsertParagraph();
                p.Append("This is another example, to show that footnotes appear,");
                fn = new Footnote(document, "This source is the best authority.");
                fn.Apply(p);
                p.Append(" on the same page not at the end.");

                // this shows how to include note reference and hyperlink in a footnote
                fn = new Footnote(document)
                    .AppendText("See note ")
                    .AppendNoteRef(fnRef)
                    .AppendText(". See also, ")
                    .AppendHyperlink("http://www.google.com")
                    .AppendText(".");
                fn.Apply(p);

                // Save this document to disk.
                document.Save();
                Console.WriteLine("\tCreated: BookmarkedFootnote.docx\n");
            }

        }
        #endregion
    }
}
