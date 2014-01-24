using System;
using Novacode;
using System.Linq;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests
{
    [TestClass]
    public class InsertAtBookmark
    {
        [TestMethod]
        public void Inserting_at_bookmark_should_add_text_in_paragraph()
        {
            using (var document = DocX.Create(""))
            {
                document.InsertParagraph("Hello ");
                document.InsertBookmark("bookmark1");
                document.InsertParagraph("!");

                document.InsertAtBookmark("world", "bookmark1");

                Assert.AreEqual("Hello world!", document.Text);

            }
        }

        [TestMethod]
        public void Inserting_at_bookmark_should_add_text_in_header()
        {
            using (var document = DocX.Create(""))
            {
                document.AddHeaders();
                var header = document.Headers.even;
                header.InsertParagraph("Hello ");
                header.InsertBookmark("bookmark1");
                header.InsertParagraph("!");

                document.InsertAtBookmark("world", "bookmark1");

                Assert.AreEqual("Hello world!", String.Join("", header.Paragraphs.Select(x=>x.Text)));

            }
        }

        [TestMethod]
        public void Inserting_at_bookmark_should_add_text_in_footer()
        {
            using (var document = DocX.Create(""))
            {
                document.AddHeaders();
                var footer = document.Headers.even;
                footer.InsertParagraph("Hello ");
                footer.InsertBookmark("bookmark1");
                footer.InsertParagraph("!");

                document.InsertAtBookmark("world", "bookmark1");

                Assert.AreEqual("Hello world!", String.Join("", footer.Paragraphs.Select(x => x.Text)));

            }
        }
    }
}
