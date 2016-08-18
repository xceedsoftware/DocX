using NUnit.Framework;
using Novacode;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UnitTests
{
    [TestFixture]
    public class AppendBookmark
    {
        [Test]
        public void Bookmark_should_be_appended()
        {
            using (var doc = DocX.Create(""))
            {
                var paragraph = doc.InsertParagraph("A paragraph");
                paragraph.AppendBookmark("bookmark");
                var bookmarks = paragraph.GetBookmarks();
                Assert.AreEqual(1, bookmarks.Count());
            }
        }

        [Test]
        public void Bookmark_should_be_named_correctly()
        {
            using (var doc = DocX.Create(""))
            {
                var paragraph = doc.InsertParagraph("A paragraph");
                paragraph.AppendBookmark("bookmark");
                var bookmarks = paragraph.GetBookmarks();
                Assert.AreEqual("bookmark", bookmarks.First().Name);
            }
        }

        [Test]
        public void Bookmark_should_reference_paragraph()
        {
            using (var doc = DocX.Create(""))
            {
                var paragraph = doc.InsertParagraph("A paragraph");
                paragraph.AppendBookmark("bookmark");
                var bookmarks = paragraph.GetBookmarks();
                Assert.AreEqual(paragraph, bookmarks.First().Paragraph);
            }
        }

    }
}
