﻿using System.Linq;
using Novacode;
using NUnit.Framework;

namespace UnitTests
{
    [TestFixture]
    public class InsertAtBookmark
    {
        [Test]
        public void Inserting_at_bookmark_should_add_text_in_footer()
        {
            using (DocX document = DocX.Create(""))
            {
                document.AddHeaders();
                Header footer = document.Headers.even;
                footer.InsertParagraph("Hello ");
                footer.InsertBookmark("bookmark1");
                footer.InsertParagraph("!");

                document.InsertAtBookmark("world", "bookmark1");

                Assert.AreEqual("Hello world!", string.Join("", footer.Paragraphs.Select(x => x.Text)));
            }
        }

        [Test]
        public void Inserting_at_bookmark_should_add_text_in_header()
        {
            using (DocX document = DocX.Create(""))
            {
                document.AddHeaders();
                Header header = document.Headers.even;
                header.InsertParagraph("Hello ");
                header.InsertBookmark("bookmark1");
                header.InsertParagraph("!");

                document.InsertAtBookmark("world", "bookmark1");

                Assert.AreEqual("Hello world!", string.Join("", header.Paragraphs.Select(x => x.Text)));
            }
        }

        [Test]
        public void Inserting_at_bookmark_should_add_text_in_paragraph()
        {
            using (DocX document = DocX.Create(""))
            {
                document.InsertParagraph("Hello ");
                document.InsertBookmark("bookmark1");
                document.InsertParagraph("!");

                document.InsertAtBookmark("world", "bookmark1");

                Assert.AreEqual("Hello world!", document.Text);
            }
        }
    }
}