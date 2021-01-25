using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Xceed.Document.NET;
using Xceed.Document.NET.Src;
using Xceed.Words.NET.Examples;

namespace Xceed.Words.NET.Examples
{
    public class IndexSample
    {
        #region Private Members

        private const string IndexSampleOutputDirectory = Program.SampleDirectory + @"Index\Output\";

        #endregion

        #region Constructors

        static IndexSample()
        {
            if (!Directory.Exists(IndexSampleOutputDirectory))
            {
                Directory.CreateDirectory(IndexSampleOutputDirectory);
            }
        }

        #endregion

        #region Public Methods
        public static void SimpleIndex()
        {
            Console.WriteLine("\tSimpleIndex()");

            using (var document = DocX.Create(IndexSampleOutputDirectory + @"SimpleIndex.docx"))
            {
                // Insert a Paragraph into this document.
                var p = document.InsertParagraph();

                // Append some text and index entries.
                p.Append("This is a simple paragraph about John Smith");
                p.AppendField(new IndexEntry(document) {IndexValue = "Smith:John"}.Build());
                p.Append(" and his buddy Abe Jones");
                p.AppendField(new IndexEntry(document) { IndexValue = "Jones:Abraham" }.Build());

                document.InsertSectionPageBreak();
                p = document.InsertParagraph("We have a lot more to say about that Jones!");
                p.AppendField(new IndexEntry(document) { IndexValue = "Jones:Abraham" }.Build());
                p.Append(" He was quite a character.");

                document.InsertSectionPageBreak();
                p = document.InsertParagraph();
                p = document.InsertParagraph("Index of Names", false, new Formatting(){Bold = true});
                p = document.InsertParagraph();
                p.AppendField(new IndexField(document) {Columns = 2}.Build());

                // Save this document to disk.
                document.Save();
                Console.WriteLine("\tCreated: SimpleIndex.docx");
                Console.WriteLine("\t\tNB to show index, open doc and hit ctrl-a then F9\n");
            }

        }
        public static void MultiIndex()
        {
            Console.WriteLine("\tMultiIndex()");

            string nameType = "names";
            string placeType = "places";

            using (var document = DocX.Create(IndexSampleOutputDirectory + @"MultiIndex.docx"))
            {
                // Insert a Paragraph into this document.
                var p = document.InsertParagraph();

                // Append some text and index entries.
                p.Append("This is a simple paragraph about John Smith");
                p.AppendField(new IndexEntry(document) { IndexValue = "Smith:John", IndexName = nameType}.Build());
                p.Append(" of Jackson Hole, Wyoming");
                p.AppendField(new IndexEntry(document) { IndexValue = "Wyoming:Teton County:Jackson Hole", IndexName = placeType }.Build());
                p.Append(" and his buddy Abe Jones");
                p.AppendField(new IndexEntry(document) { IndexValue = "Jones:Abraham", IndexName = nameType }.Build());
                p.Append(".");

                document.InsertSectionPageBreak();
                p = document.InsertParagraph("We have a lot more to say about that Jones!");
                p.AppendField(new IndexEntry(document) { IndexValue = "Jones:Abraham", IndexName = nameType }.Build());
                p.Append(" He was quite a character. He came to Wyoming");
                p.AppendField(new IndexEntry(document) { IndexValue = "Wyoming", IndexName = placeType }.Build());
                p.Append(" from New London.");
                p.AppendField(new IndexEntry(document) { IndexValue = "Connecticut:New London County:New London", IndexName = placeType }.Build());


                document.InsertSectionPageBreak();
                p = document.InsertParagraph();
                p = document.InsertParagraph("Index of Names", false, new Formatting() { Bold = true });
                p = document.InsertParagraph();
                p.AppendField(new IndexField(document) { Columns = 2, IndexName=nameType }.Build());

                document.InsertSectionPageBreak();
                p = document.InsertParagraph();
                p = document.InsertParagraph("Index of Places", false, new Formatting() { Bold = true });
                p = document.InsertParagraph();
                p.AppendField(new IndexField(document) { Columns = 1, IndexName = placeType}.Build());

                // Save this document to disk.
                document.Save();
                Console.WriteLine("\tCreated: MultiIndex.docx");
                Console.WriteLine("\t\tNB to show index, open doc and hit ctrl-a then F9\n");
            }

        }
        #endregion

    }
}
