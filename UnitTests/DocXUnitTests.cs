using System;
using Novacode;
using System.IO;
using System.Linq;
using System.Drawing;
using System.Xml.Linq;
using System.Reflection;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.ObjectModel;

namespace UnitTests
{
    /// <summary>
    /// Summary description for DocXUnitTest
    /// </summary>
    [TestClass]
    public class DocXUnitTests
    {
        
        private readonly string _directoryDocuments;
        private const string FileTemp = "temp.docx";
        private readonly string _directoryWithFiles;

        const string package_part_document = "/word/document.xml";

        public DocXUnitTests()
        {
            string directoryExecutingAssembly = Assembly.GetExecutingAssembly().Location; // Get the fullpath to the executing assembly.
            List<string> steps = directoryExecutingAssembly.Split('\\').ToList(); // The documents directory
            steps.RemoveRange(steps.Count() - 2, 2);
            _directoryDocuments = String.Join("\\", steps) + "\\documents\\";
            Setup(_directoryDocuments); // prepare temp documents directory 
            RelativeDirectory rd = new RelativeDirectory(); // prepares the files for testing
            rd.Up(3);
            _directoryWithFiles = rd.Path + @"\UnitTests\documents\";
        }
        private static void Setup(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
        [TestMethod]
        public void Test_Pattern_Replacement()
        {
            Dictionary<string, string> testPatterns = new Dictionary<string, string>()
            {
                {"COURT NAME","Fred Frump"}, 
                {"Case Number","cr-md-2011-1234567"}
            };

            using (DocX replaceDoc = DocX.Load(_directoryWithFiles + "ReplaceTests.docx"))
            {
                foreach (var t in replaceDoc.Tables)
                {   // each table has 1 row and 3 columns
                    Assert.IsTrue(t.Rows[0].Cells.Count == 3);
                    Assert.IsTrue(t.ColumnCount == 3);
                    Assert.IsTrue(t.Rows.Count == 1);
                    Assert.IsTrue(t.RowCount == 1);
                }

                // Make sure the origional strings are in the document.
                Assert.IsTrue(replaceDoc.FindAll("<COURT NAME>").Count == 2);
                Assert.IsTrue(replaceDoc.FindAll("<Case Number>").Count == 2);

                // There are only two patterns, even though each pattern is used more than once
                Assert.IsTrue(replaceDoc.FindUniqueByPattern(@"<[\w \=]{4,}>", RegexOptions.IgnoreCase).Count == 2);

                // Make sure the new strings are not in the document.
                Assert.IsTrue(replaceDoc.FindAll("Fred Frump").Count == 0);
                Assert.IsTrue(replaceDoc.FindAll("cr-md-2011-1234567").Count == 0);

                // Do the replacing
                foreach (var p in testPatterns)
                    replaceDoc.ReplaceText("<" + p.Key + ">", p.Value, false, RegexOptions.IgnoreCase);

                // Make sure the origional string are no longer in the document.
                Assert.IsTrue(replaceDoc.FindAll("<COURT NAME>").Count == 0);
                Assert.IsTrue(replaceDoc.FindAll("<Case Number>").Count == 0);

                // Make sure the new strings are now in the document.
                Assert.IsTrue(replaceDoc.FindAll("FRED FRUMP").Count == 2);
                Assert.IsTrue(replaceDoc.FindAll("cr-md-2011-1234567").Count == 2);

                // Make sure the replacement worked.
                Assert.IsTrue(replaceDoc.Text == "\t\t\t\t\t\t\t\t\t\t\t\t\t\tThese two tables should look identical:\t\t\t\t\t\t\t\t\t\t\t\t\t\t\n\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tSTATE OF IOWA,\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tPlaintiff,\t\t\t\t\t\t\t\t\t\t\t\t\t\tvs.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tFRED FRUMP,\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tDefendant.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tCase No.: cr-md-2011-1234567\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tORDER SETTING ASIDE DEFAULT JUDGMENT\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\n\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tSTATE OF IOWA,\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tPlaintiff,\t\t\t\t\t\t\t\t\t\t\t\t\t\tvs.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tFRED FRUMP,\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tDefendant.\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tCase No.: cr-md-2011-1234567\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\tORDER SETTING ASIDE DEFAULT JUDGMENT\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t");
            }

        }

        [TestMethod]
        public void Test_CustomProperty_Add()
        {
            // Load a document.
            using (DocX document = DocX.Create("CustomProperty_Add.docx"))
            {
                Assert.IsTrue(document.CustomProperties.Count == 0);

                document.AddCustomProperty(new CustomProperty("fname", "cathal"));

                Assert.IsTrue(document.CustomProperties.Count == 1);
                Assert.IsTrue(document.CustomProperties.ContainsKey("fname"));
                Assert.IsTrue((String)document.CustomProperties["fname"].Value == "cathal");

                document.AddCustomProperty(new CustomProperty("age", 24));

                Assert.IsTrue(document.CustomProperties.Count == 2);
                Assert.IsTrue(document.CustomProperties.ContainsKey("age"));
                Assert.IsTrue((int)document.CustomProperties["age"].Value == 24);

                document.AddCustomProperty(new CustomProperty("male", true));

                Assert.IsTrue(document.CustomProperties.Count == 3);
                Assert.IsTrue(document.CustomProperties.ContainsKey("male"));
                Assert.IsTrue((bool)document.CustomProperties["male"].Value == true);

                document.AddCustomProperty(new CustomProperty("newyear2012", new DateTime(2012, 1, 1)));

                Assert.IsTrue(document.CustomProperties.Count == 4);
                Assert.IsTrue(document.CustomProperties.ContainsKey("newyear2012"));
                Assert.IsTrue((DateTime)document.CustomProperties["newyear2012"].Value == new DateTime(2012, 1, 1));

                document.AddCustomProperty(new CustomProperty("fav_num", 3.141592));

                Assert.IsTrue(document.CustomProperties.Count == 5);
                Assert.IsTrue(document.CustomProperties.ContainsKey("fav_num"));
                Assert.IsTrue((double)document.CustomProperties["fav_num"].Value == 3.141592);
            }
        }

        [TestMethod]
        public void Test_EverybodyHasAHome_Loaded()
        {
            // Load a document.
            using (DocX document = DocX.Load(_directoryWithFiles + "EverybodyHasAHome.docx"))
            {
                // Main document tests.
                string document_xml_file = document.mainPart.Uri.OriginalString;
                Assert.IsTrue(document.Paragraphs[0].mainPart.Uri.OriginalString.Equals(document_xml_file));
                Assert.IsTrue(document.Tables[0].mainPart.Uri.OriginalString.Equals(document_xml_file));
                Assert.IsTrue(document.Tables[0].Rows[0].mainPart.Uri.OriginalString.Equals(document_xml_file));
                Assert.IsTrue(document.Tables[0].Rows[0].Cells[0].mainPart.Uri.OriginalString.Equals(document_xml_file));
                Assert.IsTrue(document.Tables[0].Rows[0].Cells[0].Paragraphs[0].mainPart.Uri.OriginalString.Equals(document_xml_file));

                // header first
                Header header_first = document.Headers.first;
                string header_first_xml_file = header_first.mainPart.Uri.OriginalString;

                Assert.IsTrue(header_first.Paragraphs[0].mainPart.Uri.OriginalString.Equals(header_first_xml_file));
                Assert.IsTrue(header_first.Tables[0].mainPart.Uri.OriginalString.Equals(header_first_xml_file));
                Assert.IsTrue(header_first.Tables[0].Rows[0].mainPart.Uri.OriginalString.Equals(header_first_xml_file));
                Assert.IsTrue(header_first.Tables[0].Rows[0].Cells[0].mainPart.Uri.OriginalString.Equals(header_first_xml_file));
                Assert.IsTrue(header_first.Tables[0].Rows[0].Cells[0].Paragraphs[0].mainPart.Uri.OriginalString.Equals(header_first_xml_file));

                // header odd
                Header header_odd = document.Headers.odd;
                string header_odd_xml_file = header_odd.mainPart.Uri.OriginalString;

                Assert.IsTrue(header_odd.mainPart.Uri.OriginalString.Equals(header_odd_xml_file));
                Assert.IsTrue(header_odd.Paragraphs[0].mainPart.Uri.OriginalString.Equals(header_odd_xml_file));
                Assert.IsTrue(header_odd.Tables[0].mainPart.Uri.OriginalString.Equals(header_odd_xml_file));
                Assert.IsTrue(header_odd.Tables[0].Rows[0].mainPart.Uri.OriginalString.Equals(header_odd_xml_file));
                Assert.IsTrue(header_odd.Tables[0].Rows[0].Cells[0].mainPart.Uri.OriginalString.Equals(header_odd_xml_file));
                Assert.IsTrue(header_odd.Tables[0].Rows[0].Cells[0].Paragraphs[0].mainPart.Uri.OriginalString.Equals(header_odd_xml_file));

                // header even
                Header header_even = document.Headers.even;
                string header_even_xml_file = header_even.mainPart.Uri.OriginalString;

                Assert.IsTrue(header_even.mainPart.Uri.OriginalString.Equals(header_even_xml_file));
                Assert.IsTrue(header_even.Paragraphs[0].mainPart.Uri.OriginalString.Equals(header_even_xml_file));
                Assert.IsTrue(header_even.Tables[0].mainPart.Uri.OriginalString.Equals(header_even_xml_file));
                Assert.IsTrue(header_even.Tables[0].Rows[0].mainPart.Uri.OriginalString.Equals(header_even_xml_file));
                Assert.IsTrue(header_even.Tables[0].Rows[0].Cells[0].mainPart.Uri.OriginalString.Equals(header_even_xml_file));
                Assert.IsTrue(header_even.Tables[0].Rows[0].Cells[0].Paragraphs[0].mainPart.Uri.OriginalString.Equals(header_even_xml_file));

                // footer first
                Footer footer_first = document.Footers.first;
                string footer_first_xml_file = footer_first.mainPart.Uri.OriginalString;

                Assert.IsTrue(footer_first.mainPart.Uri.OriginalString.Equals(footer_first_xml_file));
                Assert.IsTrue(footer_first.Paragraphs[0].mainPart.Uri.OriginalString.Equals(footer_first_xml_file));
                Assert.IsTrue(footer_first.Tables[0].mainPart.Uri.OriginalString.Equals(footer_first_xml_file));
                Assert.IsTrue(footer_first.Tables[0].Rows[0].mainPart.Uri.OriginalString.Equals(footer_first_xml_file));
                Assert.IsTrue(footer_first.Tables[0].Rows[0].Cells[0].mainPart.Uri.OriginalString.Equals(footer_first_xml_file));
                Assert.IsTrue(footer_first.Tables[0].Rows[0].Cells[0].Paragraphs[0].mainPart.Uri.OriginalString.Equals(footer_first_xml_file));

                // footer odd
                Footer footer_odd = document.Footers.odd;
                string footer_odd_xml_file = footer_odd.mainPart.Uri.OriginalString;

                Assert.IsTrue(footer_odd.mainPart.Uri.OriginalString.Equals(footer_odd_xml_file));
                Assert.IsTrue(footer_odd.Paragraphs[0].mainPart.Uri.OriginalString.Equals(footer_odd_xml_file));
                Assert.IsTrue(footer_odd.Tables[0].mainPart.Uri.OriginalString.Equals(footer_odd_xml_file));
                Assert.IsTrue(footer_odd.Tables[0].Rows[0].mainPart.Uri.OriginalString.Equals(footer_odd_xml_file));
                Assert.IsTrue(footer_odd.Tables[0].Rows[0].Cells[0].mainPart.Uri.OriginalString.Equals(footer_odd_xml_file));
                Assert.IsTrue(footer_odd.Tables[0].Rows[0].Cells[0].Paragraphs[0].mainPart.Uri.OriginalString.Equals(footer_odd_xml_file));

                // footer even
                Footer footer_even = document.Footers.even;
                string footer_even_xml_file = footer_even.mainPart.Uri.OriginalString;

                Assert.IsTrue(footer_even.mainPart.Uri.OriginalString.Equals(footer_even_xml_file));
                Assert.IsTrue(footer_even.Paragraphs[0].mainPart.Uri.OriginalString.Equals(footer_even_xml_file));
                Assert.IsTrue(footer_even.Tables[0].mainPart.Uri.OriginalString.Equals(footer_even_xml_file));
                Assert.IsTrue(footer_even.Tables[0].Rows[0].mainPart.Uri.OriginalString.Equals(footer_even_xml_file));
                Assert.IsTrue(footer_even.Tables[0].Rows[0].Cells[0].mainPart.Uri.OriginalString.Equals(footer_even_xml_file));
                Assert.IsTrue(footer_even.Tables[0].Rows[0].Cells[0].Paragraphs[0].mainPart.Uri.OriginalString.Equals(footer_even_xml_file));
            }
        }

        [TestMethod]
        public void Test_Insert_Picture_ParagraphBeforeSelf()
        {
            // Create test document.
            using (DocX document = DocX.Create(_directoryDocuments + "Test.docx"))
            {
                // Add an Image to this document.
                Novacode.Image img = document.AddImage(_directoryWithFiles + "purple.png");

                // Create a Picture from this Image.
                Picture pic = img.CreatePicture();
                Assert.IsNotNull(pic);

                // Main document.
                Paragraph p0 = document.InsertParagraph("Hello");
                Paragraph p1 = p0.InsertParagraphBeforeSelf("again");
                p1.InsertPicture(pic, 3);

                // Save this document.
                document.Save();
            }
        }

        [TestMethod]
        public void Test_Insert_Picture_ParagraphAfterSelf()
        {
            // Create test document.
            using (DocX document = DocX.Create(_directoryDocuments + "Test.docx"))
            {
                // Add an Image to this document.
                Novacode.Image img = document.AddImage(_directoryWithFiles + "purple.png");

                // Create a Picture from this Image.
                Picture pic = img.CreatePicture();
                Assert.IsNotNull(pic);

                // Main document.
                Paragraph p0 = document.InsertParagraph("Hello");
                Paragraph p1 = p0.InsertParagraphAfterSelf("again");
                p1.InsertPicture(pic, 3);

                // Save this document.
                document.Save();
            }
        }

        [TestMethod]
        public void Test_EverybodyHasAHome_Created()
        {
            // Create a new document.
            using (DocX document = DocX.Create("Test.docx"))
            {
                // Create a Table.
                Table t = document.AddTable(3, 3);
                t.Design = TableDesign.TableGrid;

                // Insert a Paragraph and a Table into the main document.
                document.InsertParagraph();
                document.InsertTable(t);

                // Insert a Paragraph and a Table into every Header.
                document.AddHeaders();
                document.Headers.odd.InsertParagraph();
                document.Headers.odd.InsertTable(t);
                document.Headers.even.InsertParagraph();
                document.Headers.even.InsertTable(t);
                document.Headers.first.InsertParagraph();
                document.Headers.first.InsertTable(t);

                // Insert a Paragraph and a Table into every Footer.
                document.AddFooters();
                document.Footers.odd.InsertParagraph();
                document.Footers.odd.InsertTable(t);
                document.Footers.even.InsertParagraph();
                document.Footers.even.InsertTable(t);
                document.Footers.first.InsertParagraph();
                document.Footers.first.InsertTable(t);

                // Main document tests.
                string document_xml_file = document.mainPart.Uri.OriginalString;
                Assert.IsTrue(document.Paragraphs[0].mainPart.Uri.OriginalString.Equals(document_xml_file));
                Assert.IsTrue(document.Tables[0].mainPart.Uri.OriginalString.Equals(document_xml_file));
                Assert.IsTrue(document.Tables[0].Rows[0].mainPart.Uri.OriginalString.Equals(document_xml_file));
                Assert.IsTrue(document.Tables[0].Rows[0].Cells[0].mainPart.Uri.OriginalString.Equals(document_xml_file));
                Assert.IsTrue(document.Tables[0].Rows[0].Cells[0].Paragraphs[0].mainPart.Uri.OriginalString.Equals(document_xml_file));

                // header first
                Header header_first = document.Headers.first;
                string header_first_xml_file = header_first.mainPart.Uri.OriginalString;

                Assert.IsTrue(header_first.Paragraphs[0].mainPart.Uri.OriginalString.Equals(header_first_xml_file));
                Assert.IsTrue(header_first.Tables[0].mainPart.Uri.OriginalString.Equals(header_first_xml_file));
                Assert.IsTrue(header_first.Tables[0].Rows[0].mainPart.Uri.OriginalString.Equals(header_first_xml_file));
                Assert.IsTrue(header_first.Tables[0].Rows[0].Cells[0].mainPart.Uri.OriginalString.Equals(header_first_xml_file));
                Assert.IsTrue(header_first.Tables[0].Rows[0].Cells[0].Paragraphs[0].mainPart.Uri.OriginalString.Equals(header_first_xml_file));

                // header odd
                Header header_odd = document.Headers.odd;
                string header_odd_xml_file = header_odd.mainPart.Uri.OriginalString;

                Assert.IsTrue(header_odd.mainPart.Uri.OriginalString.Equals(header_odd_xml_file));
                Assert.IsTrue(header_odd.Paragraphs[0].mainPart.Uri.OriginalString.Equals(header_odd_xml_file));
                Assert.IsTrue(header_odd.Tables[0].mainPart.Uri.OriginalString.Equals(header_odd_xml_file));
                Assert.IsTrue(header_odd.Tables[0].Rows[0].mainPart.Uri.OriginalString.Equals(header_odd_xml_file));
                Assert.IsTrue(header_odd.Tables[0].Rows[0].Cells[0].mainPart.Uri.OriginalString.Equals(header_odd_xml_file));
                Assert.IsTrue(header_odd.Tables[0].Rows[0].Cells[0].Paragraphs[0].mainPart.Uri.OriginalString.Equals(header_odd_xml_file));

                // header even
                Header header_even = document.Headers.even;
                string header_even_xml_file = header_even.mainPart.Uri.OriginalString;

                Assert.IsTrue(header_even.mainPart.Uri.OriginalString.Equals(header_even_xml_file));
                Assert.IsTrue(header_even.Paragraphs[0].mainPart.Uri.OriginalString.Equals(header_even_xml_file));
                Assert.IsTrue(header_even.Tables[0].mainPart.Uri.OriginalString.Equals(header_even_xml_file));
                Assert.IsTrue(header_even.Tables[0].Rows[0].mainPart.Uri.OriginalString.Equals(header_even_xml_file));
                Assert.IsTrue(header_even.Tables[0].Rows[0].Cells[0].mainPart.Uri.OriginalString.Equals(header_even_xml_file));
                Assert.IsTrue(header_even.Tables[0].Rows[0].Cells[0].Paragraphs[0].mainPart.Uri.OriginalString.Equals(header_even_xml_file));

                // footer first
                Footer footer_first = document.Footers.first;
                string footer_first_xml_file = footer_first.mainPart.Uri.OriginalString;

                Assert.IsTrue(footer_first.mainPart.Uri.OriginalString.Equals(footer_first_xml_file));
                Assert.IsTrue(footer_first.Paragraphs[0].mainPart.Uri.OriginalString.Equals(footer_first_xml_file));
                Assert.IsTrue(footer_first.Tables[0].mainPart.Uri.OriginalString.Equals(footer_first_xml_file));
                Assert.IsTrue(footer_first.Tables[0].Rows[0].mainPart.Uri.OriginalString.Equals(footer_first_xml_file));
                Assert.IsTrue(footer_first.Tables[0].Rows[0].Cells[0].mainPart.Uri.OriginalString.Equals(footer_first_xml_file));
                Assert.IsTrue(footer_first.Tables[0].Rows[0].Cells[0].Paragraphs[0].mainPart.Uri.OriginalString.Equals(footer_first_xml_file));

                // footer odd
                Footer footer_odd = document.Footers.odd;
                string footer_odd_xml_file = footer_odd.mainPart.Uri.OriginalString;

                Assert.IsTrue(footer_odd.mainPart.Uri.OriginalString.Equals(footer_odd_xml_file));
                Assert.IsTrue(footer_odd.Paragraphs[0].mainPart.Uri.OriginalString.Equals(footer_odd_xml_file));
                Assert.IsTrue(footer_odd.Tables[0].mainPart.Uri.OriginalString.Equals(footer_odd_xml_file));
                Assert.IsTrue(footer_odd.Tables[0].Rows[0].mainPart.Uri.OriginalString.Equals(footer_odd_xml_file));
                Assert.IsTrue(footer_odd.Tables[0].Rows[0].Cells[0].mainPart.Uri.OriginalString.Equals(footer_odd_xml_file));
                Assert.IsTrue(footer_odd.Tables[0].Rows[0].Cells[0].Paragraphs[0].mainPart.Uri.OriginalString.Equals(footer_odd_xml_file));

                // footer even
                Footer footer_even = document.Footers.even;
                string footer_even_xml_file = footer_even.mainPart.Uri.OriginalString;

                Assert.IsTrue(footer_even.mainPart.Uri.OriginalString.Equals(footer_even_xml_file));
                Assert.IsTrue(footer_even.Paragraphs[0].mainPart.Uri.OriginalString.Equals(footer_even_xml_file));
                Assert.IsTrue(footer_even.Tables[0].mainPart.Uri.OriginalString.Equals(footer_even_xml_file));
                Assert.IsTrue(footer_even.Tables[0].Rows[0].mainPart.Uri.OriginalString.Equals(footer_even_xml_file));
                Assert.IsTrue(footer_even.Tables[0].Rows[0].Cells[0].mainPart.Uri.OriginalString.Equals(footer_even_xml_file));
                Assert.IsTrue(footer_even.Tables[0].Rows[0].Cells[0].Paragraphs[0].mainPart.Uri.OriginalString.Equals(footer_even_xml_file));
            }
        }

        [TestMethod]
        public void Test_Document_AddImage_FromDisk()
        {
            using (DocX document = DocX.Create(_directoryDocuments + "test_add_images.docx"))
            {
                // Add a png to into this document
                Novacode.Image png = document.AddImage(_directoryWithFiles + "purple.png");
                Assert.IsTrue(document.Images.Count == 1);
                Assert.IsTrue(Path.GetExtension(png.pr.TargetUri.OriginalString) == ".png");

                // Add a tiff into to this document
                Novacode.Image tif = document.AddImage(_directoryWithFiles + "yellow.tif");
                Assert.IsTrue(document.Images.Count == 2);
                Assert.IsTrue(Path.GetExtension(tif.pr.TargetUri.OriginalString) == ".tif");

                // Add a gif into to this document
                Novacode.Image gif = document.AddImage(_directoryWithFiles + "orange.gif");
                Assert.IsTrue(document.Images.Count == 3);
                Assert.IsTrue(Path.GetExtension(gif.pr.TargetUri.OriginalString) == ".gif");

                // Add a jpg into to this document
                Novacode.Image jpg = document.AddImage(_directoryWithFiles + "green.jpg");
                Assert.IsTrue(document.Images.Count == 4);
                Assert.IsTrue(Path.GetExtension(jpg.pr.TargetUri.OriginalString) == ".jpg");

                // Add a bitmap to this document
                Novacode.Image bitmap = document.AddImage(_directoryWithFiles + "red.bmp");
                Assert.IsTrue(document.Images.Count == 5);
                // Word does not allow bmp make sure it was inserted as a png.
                Assert.IsTrue(Path.GetExtension(bitmap.pr.TargetUri.OriginalString) == ".png");
            }
        }

        [TestMethod]
        public void Test_Document_AddImage_FromStream()
        {
            using (DocX document = DocX.Create(_directoryDocuments + "test_add_images.docx"))
            {
                // DocX will always insert Images that come from Streams as jpeg.

                // Add a png to into this document
                Novacode.Image png = document.AddImage(new FileStream(_directoryWithFiles + "purple.png", FileMode.Open));
                Assert.IsTrue(document.Images.Count == 1);
                Assert.IsTrue(Path.GetExtension(png.pr.TargetUri.OriginalString) == ".jpeg");

                // Add a tiff into to this document
                Novacode.Image tif = document.AddImage(new FileStream(_directoryWithFiles + "yellow.tif", FileMode.Open));
                Assert.IsTrue(document.Images.Count == 2);
                Assert.IsTrue(Path.GetExtension(tif.pr.TargetUri.OriginalString) == ".jpeg");

                // Add a gif into to this document
                Novacode.Image gif = document.AddImage(new FileStream(_directoryWithFiles + "orange.gif", FileMode.Open));
                Assert.IsTrue(document.Images.Count == 3);
                Assert.IsTrue(Path.GetExtension(gif.pr.TargetUri.OriginalString) == ".jpeg");

                // Add a jpg into to this document
                Novacode.Image jpg = document.AddImage(new FileStream(_directoryWithFiles + "green.jpg", FileMode.Open));
                Assert.IsTrue(document.Images.Count == 4);
                Assert.IsTrue(Path.GetExtension(jpg.pr.TargetUri.OriginalString) == ".jpeg");

                // Add a bitmap to this document
                Novacode.Image bitmap = document.AddImage(new FileStream(_directoryWithFiles + "red.bmp", FileMode.Open));
                Assert.IsTrue(document.Images.Count == 5);
                // Word does not allow bmp make sure it was inserted as a png.
                Assert.IsTrue(Path.GetExtension(bitmap.pr.TargetUri.OriginalString) == ".jpeg");
            }
        }

        [TestMethod]
        public void Test_Tables()
        {
            using (DocX document = DocX.Load(_directoryWithFiles + "Tables.docx"))
            {
                // There is only one Paragraph at the document level.
                Assert.IsTrue(document.Paragraphs.Count() == 13);

                // There is only one Table in this document.
                Assert.IsTrue(document.Tables.Count() == 1);

                // Extract the only Table.
                Table t0 = document.Tables[0];

                // This table has 12 Paragraphs.
                Assert.IsTrue(t0.Paragraphs.Count() == 12);
            }
        }

        [TestMethod]
        public void Test_Images()
        {
            using (DocX document = DocX.Load(_directoryWithFiles + "Images.docx"))
            {
                // Extract Images from Document.
                List<Novacode.Image> document_images = document.Images;

                // Make sure there are 3 Images in this document.
                Assert.IsTrue(document_images.Count() == 3);

                // Extract the headers from this document.
                Headers headers = document.Headers;
                Header header_first = headers.first;
                Header header_odd = headers.odd;
                Header header_even = headers.even;

                #region Header_First
                // Extract Images from the first Header.
                List<Novacode.Image> header_first_images = header_first.Images;

                // Make sure there is 1 Image in the first header.
                Assert.IsTrue(header_first_images.Count() == 1);
                #endregion

                #region Header_Odd
                // Extract Images from the odd Header.
                List<Novacode.Image> header_odd_images = header_odd.Images;

                // Make sure there is 1 Image in the first header.
                Assert.IsTrue(header_odd_images.Count() == 1);
                #endregion

                #region Header_Even
                // Extract Images from the odd Header.
                List<Novacode.Image> header_even_images = header_even.Images;

                // Make sure there is 1 Image in the first header.
                Assert.IsTrue(header_even_images.Count() == 1);
                #endregion
            }
        }

        [TestMethod]
        public void Test_Insert_Picture()
        {
            // Load test document.
            using (DocX document = DocX.Create(_directoryDocuments + "Test.docx"))
            {
                // Add Headers and Footers into this document.
                document.AddHeaders();
                document.AddFooters();
                document.DifferentFirstPage = true;
                document.DifferentOddAndEvenPages = true;

                // Add an Image to this document.
                Novacode.Image img = document.AddImage(_directoryWithFiles + "purple.png");

                // Create a Picture from this Image.
                Picture pic = img.CreatePicture();

                // Main document.
                Paragraph p0 = document.InsertParagraph("Hello");
                p0.InsertPicture(pic, 3);

                // Header first.
                Paragraph p1 = document.Headers.first.InsertParagraph("----");
                p1.InsertPicture(pic, 2);

                // Header odd.
                Paragraph p2 = document.Headers.odd.InsertParagraph("----");
                p2.InsertPicture(pic, 2);

                // Header even.
                Paragraph p3 = document.Headers.even.InsertParagraph("----");
                p3.InsertPicture(pic, 2);

                // Footer first.
                Paragraph p4 = document.Footers.first.InsertParagraph("----");
                p4.InsertPicture(pic, 2);

                // Footer odd.
                Paragraph p5 = document.Footers.odd.InsertParagraph("----");
                p5.InsertPicture(pic, 2);

                // Footer even.
                Paragraph p6 = document.Footers.even.InsertParagraph("----");
                p6.InsertPicture(pic, 2);

                // Save this document.
                document.Save();
            }
        }

        [TestMethod]
        public void Test_Insert_Hyperlink()
        {
            // Load test document.
            using (DocX document = DocX.Create(_directoryDocuments + "Test.docx"))
            {
                // Add Headers and Footers into this document.
                document.AddHeaders();
                document.AddFooters();
                document.DifferentFirstPage = true;
                document.DifferentOddAndEvenPages = true;

                // Add a Hyperlink into this document.
                Hyperlink h = document.AddHyperlink("google", new Uri("http://www.google.com"));

                // Main document.
                Paragraph p0 = document.InsertParagraph("Hello");
                p0.InsertHyperlink(h, 3);

                // Header first.
                Paragraph p1 = document.Headers.first.InsertParagraph("----");
                p1.InsertHyperlink(h, 3);

                // Header odd.
                Paragraph p2 = document.Headers.odd.InsertParagraph("----");
                p2.InsertHyperlink(h, 3);

                // Header even.
                Paragraph p3 = document.Headers.even.InsertParagraph("----");
                p3.InsertHyperlink(h, 3);

                // Footer first.
                Paragraph p4 = document.Footers.first.InsertParagraph("----");
                p4.InsertHyperlink(h, 3);

                // Footer odd.
                Paragraph p5 = document.Footers.odd.InsertParagraph("----");
                p5.InsertHyperlink(h, 3);

                // Footer even.
                Paragraph p6 = document.Footers.even.InsertParagraph("----");
                p6.InsertHyperlink(h, 3);

                // Save this document.
                document.Save();
            }
        }

        [TestMethod]
        public void Test_Get_Set_Hyperlink()
        {
            // Load test document.
            using (DocX document = DocX.Load(_directoryWithFiles + "Hyperlinks.docx"))
            {
                // Hyperlinks in the document.
                Assert.IsTrue(document.Hyperlinks.Count == 3);
                Assert.IsTrue(document.Hyperlinks[0].Text == "page1");
                Assert.IsTrue(document.Hyperlinks[0].Uri.AbsoluteUri == "http://www.page1.com/");
                Assert.IsTrue(document.Hyperlinks[1].Text == "page2");
                Assert.IsTrue(document.Hyperlinks[1].Uri.AbsoluteUri == "http://www.page2.com/");
                Assert.IsTrue(document.Hyperlinks[2].Text == "page3");
                Assert.IsTrue(document.Hyperlinks[2].Uri.AbsoluteUri == "http://www.page3.com/");

                // Change the Hyperlinks and check that it has in fact changed.
                document.Hyperlinks[0].Text = "somethingnew";
                document.Hyperlinks[0].Uri = new Uri("http://www.google.com/");
                Assert.IsTrue(document.Hyperlinks[0].Text == "somethingnew");
                Assert.IsTrue(document.Hyperlinks[0].Uri.AbsoluteUri == "http://www.google.com/");
                document.Hyperlinks[1].Text = "somethingnew";
                document.Hyperlinks[1].Uri = new Uri("http://www.google.com/");
                Assert.IsTrue(document.Hyperlinks[1].Text == "somethingnew");
                Assert.IsTrue(document.Hyperlinks[1].Uri.AbsoluteUri == "http://www.google.com/");
                document.Hyperlinks[2].Text = "somethingnew";
                document.Hyperlinks[2].Uri = new Uri("http://www.google.com/");
                Assert.IsTrue(document.Hyperlinks[2].Text == "somethingnew");
                Assert.IsTrue(document.Hyperlinks[2].Uri.AbsoluteUri == "http://www.google.com/");

                Assert.IsTrue(document.Headers.first.Hyperlinks.Count == 1);
                Assert.IsTrue(document.Headers.first.Hyperlinks[0].Text == "header-first");
                Assert.IsTrue(document.Headers.first.Hyperlinks[0].Uri.AbsoluteUri == "http://www.header-first.com/");

                // Change the Hyperlinks and check that it has in fact changed.
                document.Headers.first.Hyperlinks[0].Text = "somethingnew";
                document.Headers.first.Hyperlinks[0].Uri = new Uri("http://www.google.com/");
                Assert.IsTrue(document.Headers.first.Hyperlinks[0].Text == "somethingnew");
                Assert.IsTrue(document.Headers.first.Hyperlinks[0].Uri.AbsoluteUri == "http://www.google.com/");

                Assert.IsTrue(document.Headers.odd.Hyperlinks.Count == 1);
                Assert.IsTrue(document.Headers.odd.Hyperlinks[0].Text == "header-odd");
                Assert.IsTrue(document.Headers.odd.Hyperlinks[0].Uri.AbsoluteUri == "http://www.header-odd.com/");

                // Change the Hyperlinks and check that it has in fact changed.
                document.Headers.odd.Hyperlinks[0].Text = "somethingnew";
                document.Headers.odd.Hyperlinks[0].Uri = new Uri("http://www.google.com/");
                Assert.IsTrue(document.Headers.odd.Hyperlinks[0].Text == "somethingnew");
                Assert.IsTrue(document.Headers.odd.Hyperlinks[0].Uri.AbsoluteUri == "http://www.google.com/");

                Assert.IsTrue(document.Headers.even.Hyperlinks.Count == 1);
                Assert.IsTrue(document.Headers.even.Hyperlinks[0].Text == "header-even");
                Assert.IsTrue(document.Headers.even.Hyperlinks[0].Uri.AbsoluteUri == "http://www.header-even.com/");

                // Change the Hyperlinks and check that it has in fact changed.
                document.Headers.even.Hyperlinks[0].Text = "somethingnew";
                document.Headers.even.Hyperlinks[0].Uri = new Uri("http://www.google.com/");
                Assert.IsTrue(document.Headers.even.Hyperlinks[0].Text == "somethingnew");
                Assert.IsTrue(document.Headers.even.Hyperlinks[0].Uri.AbsoluteUri == "http://www.google.com/");

                Assert.IsTrue(document.Footers.first.Hyperlinks.Count == 1);
                Assert.IsTrue(document.Footers.first.Hyperlinks[0].Text == "footer-first");
                Assert.IsTrue(document.Footers.first.Hyperlinks[0].Uri.AbsoluteUri == "http://www.footer-first.com/");

                // Change the Hyperlinks and check that it has in fact changed.
                document.Footers.first.Hyperlinks[0].Text = "somethingnew";
                document.Footers.first.Hyperlinks[0].Uri = new Uri("http://www.google.com/");
                Assert.IsTrue(document.Footers.first.Hyperlinks[0].Text == "somethingnew");
                Assert.IsTrue(document.Footers.first.Hyperlinks[0].Uri.AbsoluteUri == "http://www.google.com/");

                Assert.IsTrue(document.Footers.odd.Hyperlinks.Count == 1);
                Assert.IsTrue(document.Footers.odd.Hyperlinks[0].Text == "footer-odd");
                Assert.IsTrue(document.Footers.odd.Hyperlinks[0].Uri.AbsoluteUri == "http://www.footer-odd.com/");

                // Change the Hyperlinks and check that it has in fact changed.
                document.Footers.odd.Hyperlinks[0].Text = "somethingnew";
                document.Footers.odd.Hyperlinks[0].Uri = new Uri("http://www.google.com/");
                Assert.IsTrue(document.Footers.odd.Hyperlinks[0].Text == "somethingnew");
                Assert.IsTrue(document.Footers.odd.Hyperlinks[0].Uri.AbsoluteUri == "http://www.google.com/");

                Assert.IsTrue(document.Footers.even.Hyperlinks.Count == 1);
                Assert.IsTrue(document.Footers.even.Hyperlinks[0].Text == "footer-even");
                Assert.IsTrue(document.Footers.even.Hyperlinks[0].Uri.AbsoluteUri == "http://www.footer-even.com/");

                // Change the Hyperlinks and check that it has in fact changed.
                document.Footers.even.Hyperlinks[0].Text = "somethingnew";
                document.Footers.even.Hyperlinks[0].Uri = new Uri("http://www.google.com/");
                Assert.IsTrue(document.Footers.even.Hyperlinks[0].Text == "somethingnew");
                Assert.IsTrue(document.Footers.even.Hyperlinks[0].Uri.AbsoluteUri == "http://www.google.com/");
            }
        }

        [TestMethod]
        public void Test_Append_Hyperlink()
        {
            // Load test document.
            using (DocX document = DocX.Create(_directoryDocuments + "Test.docx"))
            {
                // Add Headers and Footers into this document.
                document.AddHeaders();
                document.AddFooters();
                document.DifferentFirstPage = true;
                document.DifferentOddAndEvenPages = true;

                // Add a Hyperlink to this document.
                Hyperlink h = document.AddHyperlink("google", new Uri("http://www.google.com"));

                // Main document.
                Paragraph p0 = document.InsertParagraph("----");
                p0.AppendHyperlink(h);
                Assert.IsTrue(p0.Text == "----google");

                // Header first.
                Paragraph p1 = document.Headers.first.InsertParagraph("----");
                p1.AppendHyperlink(h);
                Assert.IsTrue(p1.Text == "----google");

                // Header odd.
                Paragraph p2 = document.Headers.odd.InsertParagraph("----");
                p2.AppendHyperlink(h);
                Assert.IsTrue(p2.Text == "----google");

                // Header even.
                Paragraph p3 = document.Headers.even.InsertParagraph("----");
                p3.AppendHyperlink(h);
                Assert.IsTrue(p3.Text == "----google");

                // Footer first.
                Paragraph p4 = document.Footers.first.InsertParagraph("----");
                p4.AppendHyperlink(h);
                Assert.IsTrue(p4.Text == "----google");

                // Footer odd.
                Paragraph p5 = document.Footers.odd.InsertParagraph("----");
                p5.AppendHyperlink(h);
                Assert.IsTrue(p5.Text == "----google");

                // Footer even.
                Paragraph p6 = document.Footers.even.InsertParagraph("----");
                p6.AppendHyperlink(h);
                Assert.IsTrue(p6.Text == "----google");

                // Save the document.
                document.Save();
            }
        }

        [TestMethod]
        public void Test_Append_Picture()
        {
            // Create test document.
            using (DocX document = DocX.Create(_directoryDocuments + "Test.docx"))
            {
                // Add Headers and Footers into this document.
                document.AddHeaders();
                document.AddFooters();
                document.DifferentFirstPage = true;
                document.DifferentOddAndEvenPages = true;

                // Add an Image to this document.
                Novacode.Image img = document.AddImage(_directoryWithFiles + "purple.png");

                // Create a Picture from this Image.
                Picture pic = img.CreatePicture();

                // Main document.
                Paragraph p0 = document.InsertParagraph();
                p0.AppendPicture(pic);

                // Header first.
                Paragraph p1 = document.Headers.first.InsertParagraph();
                p1.AppendPicture(pic);

                // Header odd.
                Paragraph p2 = document.Headers.odd.InsertParagraph();
                p2.AppendPicture(pic);

                // Header even.
                Paragraph p3 = document.Headers.even.InsertParagraph();
                p3.AppendPicture(pic);

                // Footer first.
                Paragraph p4 = document.Footers.first.InsertParagraph();
                p4.AppendPicture(pic);

                // Footer odd.
                Paragraph p5 = document.Footers.odd.InsertParagraph();
                p5.AppendPicture(pic);

                // Footer even.
                Paragraph p6 = document.Footers.even.InsertParagraph();
                p6.AppendPicture(pic);

                // Save the document.
                document.Save();
            }
        }

        [TestMethod]
        public void Test_Move_Picture_Load()
        {
            // Load test document.
            using (DocX document = DocX.Load(_directoryWithFiles + "MovePicture.docx"))
            {
                // Extract the first Picture from the first Paragraph.
                Picture picture = document.Paragraphs.First().Pictures.First();

                // Move it into the first Header.
                Header header_first = document.Headers.first;
                header_first.Paragraphs.First().AppendPicture(picture);

                // Move it into the even Header.
                Header header_even = document.Headers.even;
                header_even.Paragraphs.First().AppendPicture(picture);

                // Move it into the odd Header.
                Header header_odd = document.Headers.odd;
                header_odd.Paragraphs.First().AppendPicture(picture);

                // Move it into the first Footer.
                Footer footer_first = document.Footers.first;
                footer_first.Paragraphs.First().AppendPicture(picture);

                // Move it into the even Footer.
                Footer footer_even = document.Footers.even;
                footer_even.Paragraphs.First().AppendPicture(picture);

                // Move it into the odd Footer.
                Footer footer_odd = document.Footers.odd;
                footer_odd.Paragraphs.First().AppendPicture(picture);

                // Save this as MovedPicture.docx
                document.SaveAs(_directoryDocuments + "MovedPicture.docx");
            }
        }

        [TestMethod]
        public void Test_Paragraph_InsertHyperlink()
        {
            // Create a new document
            using (DocX document = DocX.Create("Test.docx"))
            {
                // Add a Hyperlink to this document.
                Hyperlink h = document.AddHyperlink("link", new Uri("http://www.google.com"));

                // Simple
                Paragraph p1 = document.InsertParagraph("AC");
                p1.InsertHyperlink(h); Assert.IsTrue(p1.Text == "linkAC");
                p1.InsertHyperlink(h, p1.Text.Length); Assert.IsTrue(p1.Text == "linkAClink");
                p1.InsertHyperlink(h, p1.Text.IndexOf("C")); Assert.IsTrue(p1.Text == "linkAlinkClink");

                // Difficult
                Paragraph p2 = document.InsertParagraph("\tA\tC\t");
                p2.InsertHyperlink(h); Assert.IsTrue(p2.Text == "link\tA\tC\t");
                p2.InsertHyperlink(h, p2.Text.Length); Assert.IsTrue(p2.Text == "link\tA\tC\tlink");
                p2.InsertHyperlink(h, p2.Text.IndexOf("C")); Assert.IsTrue(p2.Text == "link\tA\tlinkC\tlink");

                // Contrived
                // Add a contrived Hyperlink to this document.
                Hyperlink h2 = document.AddHyperlink("\tlink\t", new Uri("http://www.google.com"));
                Paragraph p3 = document.InsertParagraph("\tA\tC\t");
                p3.InsertHyperlink(h2); Assert.IsTrue(p3.Text == "\tlink\t\tA\tC\t");
                p3.InsertHyperlink(h2, p3.Text.Length); Assert.IsTrue(p3.Text == "\tlink\t\tA\tC\t\tlink\t");
                p3.InsertHyperlink(h2, p3.Text.IndexOf("C")); Assert.IsTrue(p3.Text == "\tlink\t\tA\t\tlink\tC\t\tlink\t");
            }
        }

        [TestMethod]
        public void Test_Paragraph_RemoveHyperlink()
        {
            // Create a new document
            using (DocX document = DocX.Create("Test.docx"))
            {
                // Add a Hyperlink to this document.
                Hyperlink h = document.AddHyperlink("link", new Uri("http://www.google.com"));

                // Simple
                Paragraph p1 = document.InsertParagraph("AC");
                p1.InsertHyperlink(h); Assert.IsTrue(p1.Text == "linkAC");
                p1.InsertHyperlink(h, p1.Text.Length); Assert.IsTrue(p1.Text == "linkAClink");
                p1.InsertHyperlink(h, p1.Text.IndexOf("C")); Assert.IsTrue(p1.Text == "linkAlinkClink");

                // Try and remove a Hyperlink using a negative index.
                // This should throw an exception.
                try
                {
                    p1.RemoveHyperlink(-1);
                    Assert.Fail();
                }
                catch (ArgumentException) { }
                catch (Exception) { Assert.Fail(); }

                // Try and remove a Hyperlink at an index greater than the last.
                // This should throw an exception.
                try
                {
                    p1.RemoveHyperlink(3);
                    Assert.Fail();
                }
                catch (ArgumentException) { }
                catch (Exception) { Assert.Fail(); }

                p1.RemoveHyperlink(0); Assert.IsTrue(p1.Text == "AlinkClink");
                p1.RemoveHyperlink(1); Assert.IsTrue(p1.Text == "AlinkC");
                p1.RemoveHyperlink(0); Assert.IsTrue(p1.Text == "AC");
            }
        }

        [TestMethod]
        public void Test_Paragraph_ReplaceText()
        {
            // Create a new document
            using (DocX document = DocX.Create("Test.docx"))
            {
                // Simple
                Paragraph p1 = document.InsertParagraph("Apple Pear Apple Apple Pear Apple");
                p1.ReplaceText("Apple", "Orange"); Assert.IsTrue(p1.Text == "Orange Pear Orange Orange Pear Orange");
                p1.ReplaceText("Pear", "Apple"); Assert.IsTrue(p1.Text == "Orange Apple Orange Orange Apple Orange");
                p1.ReplaceText("Orange", "Pear"); Assert.IsTrue(p1.Text == "Pear Apple Pear Pear Apple Pear");

                // Try and replace text that dosen't exist in the Paragraph.
                string old = p1.Text;
                p1.ReplaceText("foo", "bar"); Assert.IsTrue(p1.Text.Equals(old));

                // Difficult
                Paragraph p2 = document.InsertParagraph("Apple Pear Apple Apple Pear Apple");
                p2.ReplaceText(" ", "\t"); Assert.IsTrue(p2.Text == "Apple\tPear\tApple\tApple\tPear\tApple");
                p2.ReplaceText("\tApple\tApple", ""); Assert.IsTrue(p2.Text == "Apple\tPear\tPear\tApple");
                p2.ReplaceText("Apple\tPear\t", ""); Assert.IsTrue(p2.Text == "Pear\tApple");
                p2.ReplaceText("Pear\tApple", ""); Assert.IsTrue(p2.Text == "");
            }
        }

         [TestMethod]
        public void Test_Paragraph_ReplaceTextInGivenFormat()
        {
            // Load a document.
            using (DocX document = DocX.Load(_directoryWithFiles + "VariousTextFormatting.docx"))
            {
                // Removing red text highlighted with yellow
                var formatting = new Formatting();
                formatting.FontColor = Color.Blue;
                // IMPORTANT: default constructor of Formatting sets up language property - set it to NULL to be language independent
                var desiredFormat = new Formatting() { Language = null, FontColor = Color.FromArgb(255, 0, 0), Highlight = Highlight.yellow };
                var replaced = string.Empty;
                foreach (var p in document.Paragraphs)
                {
                    if (p.Text == "Text highlighted with yellow")
                    {
                        p.ReplaceText("Text highlighted with yellow", "New text highlighted with yellow", false, RegexOptions.None, null, desiredFormat, MatchFormattingOptions.ExactMatch);
                        replaced += p.Text;
                    }
                }

                Assert.AreEqual("New text highlighted with yellow", replaced);

                // Removing red text with no other formatting (ExactMatch)
                desiredFormat = new Formatting() { Language = null, FontColor = Color.FromArgb(255, 0, 0) };
                var count = 0;
                foreach (var p in document.Paragraphs)
                {
                    p.ReplaceText("Text", "Replaced text", false, RegexOptions.None, null, desiredFormat, MatchFormattingOptions.ExactMatch);
                    if (p.Text.StartsWith("Replaced text"))
                    {
                        ++count;
                    }
                }

                Assert.AreEqual(1, count);

                // Removing just red text with any other formatting (SubsetMatch)
                desiredFormat = new Formatting() { Language = null, FontColor = Color.FromArgb(255, 0, 0) };
                count = 0;
                foreach (var p in document.Paragraphs)
                {
                    p.ReplaceText("Text", "Replaced text", false, RegexOptions.None, null, desiredFormat, MatchFormattingOptions.SubsetMatch);
                    if (p.Text.StartsWith("Replaced text"))
                    {
                        ++count;
                    }
                }

                Assert.AreEqual(2, count);
            }
        }

        [TestMethod]
        public void Test_Paragraph_RemoveText()
        {
            // Create a new document
            using (DocX document = DocX.Create("Test.docx"))
            {
                // Simple
                //<p>
                //    <r><t>HellWorld</t></r>
                //</p>
                Paragraph p1 = document.InsertParagraph("HelloWorld");
                p1.RemoveText(0, 1); Assert.IsTrue(p1.Text == "elloWorld");
                p1.RemoveText(p1.Text.Length - 1, 1); Assert.IsTrue(p1.Text == "elloWorl");
                p1.RemoveText(p1.Text.IndexOf("o"), 1); Assert.IsTrue(p1.Text == "ellWorl");

                // Try and remove text at an index greater than the last.
                // This should throw an exception.
                try
                {
                    p1.RemoveText(p1.Text.Length, 1);
                    Assert.Fail();
                }
                catch (ArgumentOutOfRangeException) { }
                catch (Exception) { Assert.Fail(); }

                // Try and remove text at a negative index.
                // This should throw an exception.
                try
                {
                    p1.RemoveText(-1, 1);
                    Assert.Fail();
                }
                catch (ArgumentOutOfRangeException) { }
                catch (Exception) { Assert.Fail(); }

                // Difficult
                //<p>
                //    <r><t>A</t></r>
                //    <r><t>B</t></r>
                //    <r><t>C</t></r>
                //</p>
                Paragraph p2 = document.InsertParagraph("A\tB\tC");
                p2.RemoveText(0, 1); Assert.IsTrue(p2.Text == "\tB\tC");
                p2.RemoveText(p2.Text.Length - 1, 1); Assert.IsTrue(p2.Text == "\tB\t");
                p2.RemoveText(p2.Text.IndexOf("B"), 1); Assert.IsTrue(p2.Text == "\t\t");
                p2.RemoveText(0, 1); Assert.IsTrue(p2.Text == "\t");
                p2.RemoveText(0, 1); Assert.IsTrue(p2.Text == "");

                // Contrived 1
                //<p>
                //    <r>
                //        <t>A</t>
                //        <t>B</t>
                //        <t>C</t>
                //    </r>
                //</p>
                Paragraph p3 = document.InsertParagraph("");
                p3.Xml = XElement.Parse
                (
                    @"<w:p xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                        <w:pPr />
                        <w:r>
                            <w:rPr />
                            <w:t>A</w:t>
                            <w:t>B</w:t>
                            <w:t>C</w:t>
                        </w:r>
                    </w:p>"
                );

                p3.RemoveText(0, 1); Assert.IsTrue(p3.Text == "BC");
                p3.RemoveText(p3.Text.Length - 1, 1); Assert.IsTrue(p3.Text == "B");
                p3.RemoveText(0, 1); Assert.IsTrue(p3.Text == "");

                // Contrived 2
                //<p>
                //    <r>
                //        <t>A</t>
                //        <t>B</t>
                //        <t>C</t>
                //    </r>
                //</p>
                Paragraph p4 = document.InsertParagraph("");
                p4.Xml = XElement.Parse
                (
                    @"<w:p xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                        <w:pPr />
                        <w:r>
                            <w:rPr />
                            <tab />
                            <w:t>A</w:t>
                            <tab />
                        </w:r>
                        <w:r>
                            <w:rPr />
                            <tab />
                            <w:t>B</w:t>
                            <tab />
                        </w:r>
                    </w:p>"
                );

                p4.RemoveText(0, 1); Assert.IsTrue(p4.Text == "A\t\tB\t");
                p4.RemoveText(1, 1); Assert.IsTrue(p4.Text == "A\tB\t");
                p4.RemoveText(p4.Text.Length - 1, 1); Assert.IsTrue(p4.Text == "A\tB");
                p4.RemoveText(1, 1); Assert.IsTrue(p4.Text == "AB");
                p4.RemoveText(p4.Text.Length - 1, 1); Assert.IsTrue(p4.Text == "A");
                p4.RemoveText(p4.Text.Length - 1, 1); Assert.IsTrue(p4.Text == "");
            }
        }

        [TestMethod]
        public void Test_Document_RemoveTextInGivenFormat()
        {
            // Load a document.
            using (DocX document = DocX.Load(_directoryWithFiles + "VariousTextFormatting.docx"))
            {
                var formatting = new Formatting();
                formatting.FontColor = Color.Blue;
                // IMPORTANT: default constructor of Formatting sets up language property - set it to NULL to be language independent
                formatting.Language = null;
                var deletedCount = document.RemoveTextInGivenFormat(formatting);
                Assert.AreEqual(2, deletedCount);

                deletedCount = document.RemoveTextInGivenFormat(new Formatting() { Highlight = Highlight.yellow, Language = null });
                Assert.AreEqual(2, deletedCount);

                deletedCount = document.RemoveTextInGivenFormat(new Formatting() { Highlight = Highlight.blue, Language = null, FontColor = Color.FromArgb(0, 255, 0) });
                Assert.AreEqual(1, deletedCount);

                deletedCount = document.RemoveTextInGivenFormat(new Formatting() { Language = null, FontColor = Color.FromArgb(123, 123, 123) }, MatchFormattingOptions.ExactMatch);
                Assert.AreEqual(2, deletedCount);
            }
        }

        [TestMethod]
        public void Test_Paragraph_InsertText()
        {
            // Create a new document
            using (DocX document = DocX.Create("Test.docx"))
            {
                // Simple
                //<p>
                //    <r><t>HelloWorld</t></r>
                //</p>
                Paragraph p1 = document.InsertParagraph("HelloWorld");
                p1.InsertText(0, "-"); Assert.IsTrue(p1.Text == "-HelloWorld");
                p1.InsertText(p1.Text.Length, "-"); Assert.IsTrue(p1.Text == "-HelloWorld-");
                p1.InsertText(p1.Text.IndexOf("W"), "-"); Assert.IsTrue(p1.Text == "-Hello-World-");

                // Try and insert text at an index greater than the last + 1.
                // This should throw an exception.
                try
                {
                    p1.InsertText(p1.Text.Length + 1, "-");
                    Assert.Fail();
                }
                catch (ArgumentOutOfRangeException) { }
                catch (Exception) { Assert.Fail(); }

                // Try and insert text at a negative index.
                // This should throw an exception.
                try
                {
                    p1.InsertText(-1, "-");
                    Assert.Fail();
                }
                catch (ArgumentOutOfRangeException) { }
                catch (Exception) { Assert.Fail(); }

                // Difficult
                //<p>
                //    <r><t>A</t></r>
                //    <r><t>B</t></r>
                //    <r><t>C</t></r>
                //</p>
                Paragraph p2 = document.InsertParagraph("A\tB\tC");
                p2.InsertText(0, "-"); Assert.IsTrue(p2.Text == "-A\tB\tC");
                p2.InsertText(p2.Text.Length, "-"); Assert.IsTrue(p2.Text == "-A\tB\tC-");
                p2.InsertText(p2.Text.IndexOf("B"), "-"); Assert.IsTrue(p2.Text == "-A\t-B\tC-");
                p2.InsertText(p2.Text.IndexOf("C"), "-"); Assert.IsTrue(p2.Text == "-A\t-B\t-C-");

                // Contrived 1
                //<p>
                //    <r>
                //        <t>A</t>
                //        <t>B</t>
                //        <t>C</t>
                //    </r>
                //</p>
                Paragraph p3 = document.InsertParagraph("");
                p3.Xml = XElement.Parse
                (
                    @"<w:p xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                        <w:pPr />
                        <w:r>
                            <w:rPr />
                            <w:t>A</w:t>
                            <w:t>B</w:t>
                            <w:t>C</w:t>
                        </w:r>
                    </w:p>"
                );

                p3.InsertText(0, "-"); Assert.IsTrue(p3.Text == "-ABC");
                p3.InsertText(p3.Text.Length, "-"); Assert.IsTrue(p3.Text == "-ABC-");
                p3.InsertText(p3.Text.IndexOf("B"), "-"); Assert.IsTrue(p3.Text == "-A-BC-");
                p3.InsertText(p3.Text.IndexOf("C"), "-"); Assert.IsTrue(p3.Text == "-A-B-C-");

                // Contrived 2
                //<p>
                //    <r>
                //        <t>A</t>
                //        <t>B</t>
                //        <t>C</t>
                //    </r>
                //</p>
                Paragraph p4 = document.InsertParagraph("");
                p4.Xml = XElement.Parse
                (
                    @"<w:p xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                        <w:pPr />
                        <w:r>
                            <w:rPr />
                            <w:t>A</w:t>
                            <w:t>B</w:t>
                            <w:t>C</w:t>
                        </w:r>
                    </w:p>"
                );

                p4.InsertText(0, "\t"); Assert.IsTrue(p4.Text == "\tABC");
                p4.InsertText(p4.Text.Length, "\t"); Assert.IsTrue(p4.Text == "\tABC\t");
                p4.InsertText(p4.Text.IndexOf("B"), "\t"); Assert.IsTrue(p4.Text == "\tA\tBC\t");
                p4.InsertText(p4.Text.IndexOf("C"), "\t"); Assert.IsTrue(p4.Text == "\tA\tB\tC\t");
            }
        }

        [TestMethod]
        public void Test_Document_Paragraphs()
        {
            // Load the document 'Paragraphs.docx'
            using (DocX document = DocX.Load(_directoryWithFiles + "Paragraphs.docx"))
            {
                // Extract the Paragraphs from this document.
                ReadOnlyCollection<Paragraph> paragraphs = document.Paragraphs;

                // There should be 3 Paragraphs in this document.
                Assert.IsTrue(paragraphs.Count() == 3);

                // Extract the 3 Paragraphs.
                Paragraph p1 = paragraphs[0];
                Paragraph p2 = paragraphs[1];
                Paragraph p3 = paragraphs[2];

                // Extract their Text properties.
                string p1_text = p1.Text;
                string p2_text = p2.Text;
                string p3_text = p3.Text;

                // Test their Text properties against absolutes.
                Assert.IsTrue(p1_text == "Paragraph 1");
                Assert.IsTrue(p2_text == "Paragraph 2");
                Assert.IsTrue(p3_text == "Paragraph 3");

                // Its important that each Paragraph knows the PackagePart it belongs to.
                 foreach (var paragraph in document.Paragraphs)
                {
                    Assert.IsTrue(paragraph.PackagePart.Uri.ToString() == package_part_document);
                }
 

                // Test the saving of the document.
                document.SaveAs(FileTemp);
            }

            // Delete the tempory file.
            File.Delete(FileTemp);
        }

        [TestMethod]
        public void Test_Table_mainPart_bug9526()
        {
            using (DocX document = DocX.Create("test.docx"))
            {
                Hyperlink h = document.AddHyperlink("follow me", new Uri("http://www.google.com"));
                Table t = document.AddTable(2, 3);
                int cc = t.ColumnCount;

                Paragraph p = t.Rows[0].Cells[0].Paragraphs[0];
                p.AppendHyperlink(h);
            }
        }

        [TestMethod]
        public void Test_Table_InsertRowAndColumn()
        {
            // Create a table
            using (DocX document = DocX.Create(_directoryDocuments + "Tables2.docx"))
            {
                // Add a Table to a document.
                Table t = document.AddTable(2, 2);
                t.Design = TableDesign.TableGrid;

                t.Rows[0].Cells[0].Paragraphs[0].InsertText("X");
                t.Rows[0].Cells[1].Paragraphs[0].InsertText("X");
                t.Rows[1].Cells[0].Paragraphs[0].InsertText("X");
                t.Rows[1].Cells[1].Paragraphs[0].InsertText("X");

                // Insert the Table into the main section of the document.
                Table t1 = document.InsertTable(t);
                // ... and add a column and a row
                t1.InsertRow(1);
                t1.InsertColumn(1);

                // Save the document.
                document.Save();
            }

            // Check table
            using (DocX document = DocX.Load(_directoryDocuments + "Tables2.docx"))
            {
                // Get a table from a document
                Table t = document.Tables[0];

                // Check that the table is equal this: 
                // X - X
                // - - -
                // X - X
                Assert.AreEqual("X", t.Rows[0].Cells[0].Paragraphs[0].Text);
                Assert.AreEqual("X", t.Rows[2].Cells[0].Paragraphs[0].Text);
                Assert.AreEqual("X", t.Rows[0].Cells[2].Paragraphs[0].Text);
                Assert.AreEqual("X", t.Rows[2].Cells[2].Paragraphs[0].Text);
                Assert.IsTrue(String.IsNullOrEmpty(t.Rows[1].Cells[0].Paragraphs[0].Text));
                Assert.IsTrue(String.IsNullOrEmpty(t.Rows[1].Cells[1].Paragraphs[0].Text));
                Assert.IsTrue(String.IsNullOrEmpty(t.Rows[1].Cells[2].Paragraphs[0].Text));
                Assert.IsTrue(String.IsNullOrEmpty(t.Rows[0].Cells[1].Paragraphs[0].Text));
                Assert.IsTrue(String.IsNullOrEmpty(t.Rows[2].Cells[1].Paragraphs[0].Text));
            }
        }

        [TestMethod]
        public void Test_Document_ApplyTemplate()
        {
            using (MemoryStream documentStream = new MemoryStream())
            {
                using (DocX document = DocX.Create(documentStream))
                {
                    document.ApplyTemplate(_directoryWithFiles + "Template.dotx");
                    document.Save();

                    Header firstHeader = document.Headers.first;
                    Header oddHeader = document.Headers.odd;
                    Header evenHeader = document.Headers.even;

                    Footer firstFooter = document.Footers.first;
                    Footer oddFooter = document.Footers.odd;
                    Footer evenFooter = document.Footers.even;

                    Assert.IsTrue(firstHeader.Paragraphs.Count == 1, "More than one paragraph in header.");
                    Assert.IsTrue(firstHeader.Paragraphs[0].Text.Equals("First page header"), "Header isn't retrieved from template.");

                    Assert.IsTrue(oddHeader.Paragraphs.Count == 1, "More than one paragraph in header.");
                    Assert.IsTrue(oddHeader.Paragraphs[0].Text.Equals("Odd page header"), "Header isn't retrieved from template.");

                    Assert.IsTrue(evenHeader.Paragraphs.Count == 1, "More than one paragraph in header.");
                    Assert.IsTrue(evenHeader.Paragraphs[0].Text.Equals("Even page header"), "Header isn't retrieved from template.");

                    Assert.IsTrue(firstFooter.Paragraphs.Count == 1, "More than one paragraph in footer.");
                    Assert.IsTrue(firstFooter.Paragraphs[0].Text.Equals("First page footer"), "Footer isn't retrieved from template.");

                    Assert.IsTrue(oddFooter.Paragraphs.Count == 1, "More than one paragraph in footer.");
                    Assert.IsTrue(oddFooter.Paragraphs[0].Text.Equals("Odd page footer"), "Footer isn't retrieved from template.");

                    Assert.IsTrue(evenFooter.Paragraphs.Count == 1, "More than one paragraph in footer.");
                    Assert.IsTrue(evenFooter.Paragraphs[0].Text.Equals("Even page footer"), "Footer isn't retrieved from template.");

                    Paragraph firstParagraph = document.Paragraphs[0];
                    Assert.IsTrue(firstParagraph.StyleName.Equals("DocXSample"), "First paragraph isn't of style from template.");
                }
            }
        }

        [TestMethod]
        public void Test_ParentContainer_When_Creating_Doc()
        {
            using (DocX document = DocX.Create("Test.docx"))
            {
                document.AddHeaders();
                Paragraph p1 = document.Headers.first.InsertParagraph("Test");

                Assert.IsTrue(p1.ParentContainer == ContainerType.Header);
            }
        }

        [TestMethod]
        public void Test_Section_Count_When_Creating_Doc()
        {
            //This adds a section break - so insert paragraphs, and follow it up by a section break/paragraph
            using (DocX document = DocX.Create("TestSectionCount.docx"))
            {
                document.InsertSection();

                var sections = document.GetSections();

                Assert.AreEqual(sections.Count(), 2);
            }

        }

        [TestMethod]
        public void Test_Sections_And_Paragraphs_When_Creating_Doc()
        {
            //This adds a section break - so insert paragraphs, and follow it up by a section break/paragraph
            using (DocX document = DocX.Create("TestSectionAndParagraph.docx"))
            {
                //Add 2 paras and a break
                document.InsertParagraph("First Para");
                document.InsertParagraph("Second Para");
                document.InsertSection();
                document.InsertParagraph("This is default para");

                var sections = document.GetSections();

                Assert.AreEqual(sections.Count(), 2);
            }


        }

        [TestMethod]
        public void Test_ParentContainer_When_Reading_Doc()
        {
            using (DocX document = DocX.Load(_directoryWithFiles + "Tables.docx"))
            {
                ReadOnlyCollection<Paragraph> paragraphs = document.Paragraphs;

                Paragraph p1 = paragraphs[0];

                Assert.IsTrue(p1.ParentContainer == ContainerType.Cell);
            }

        }

        [TestMethod]
        public void Test_Section_Count_When_Reading_Doc()
        {
            using (DocX document = DocX.Load(_directoryWithFiles + "testdoc_SectionsWithSectionBreaks.docx"))
            {
                var sections = document.GetSections();

                Assert.AreEqual(sections.Count(), 4);
            }

        }

        [TestMethod]
        public void Test_Section_Paragraph_Count_Match_When_Reading_Doc()
        {
            using (DocX document = DocX.Load(_directoryWithFiles + "testdoc_SectionsWithSectionBreaksMultiParagraph.docx"))
            {

                var sections = document.GetSections();

                Assert.AreEqual(sections[0].SectionParagraphs.Count, 2);
                Assert.AreEqual(sections[1].SectionParagraphs.Count, 1);
                Assert.AreEqual(sections[2].SectionParagraphs.Count, 2);
                Assert.AreEqual(sections[3].SectionParagraphs.Count, 1);
            }
        }

        [TestMethod]
        public void Test_Section_Paragraph_Content_Match_When_Reading_Doc()
        {
            using (DocX document = DocX.Load(_directoryWithFiles + "testdoc_SectionsWithSectionBreaks.docx"))
            {

                var sections = document.GetSections();

                Assert.IsTrue(sections[0].SectionParagraphs[0].Text.Contains("Section 1"));
                Assert.IsTrue(sections[1].SectionParagraphs[0].Text.Contains("Section 2"));
                Assert.IsTrue(sections[2].SectionParagraphs[0].Text.Contains("Section 3"));
                Assert.IsTrue(sections[3].SectionParagraphs[0].Text.Contains("Section 4"));

            }
        }

        [TestMethod]
        public void Test_Ordered_List_When_Reading_Doc()
        {
            using (DocX document = DocX.Load(_directoryWithFiles + "testdoc_OrderedList.docx"))
            {

                var sections = document.GetSections();

                Assert.IsTrue(sections[0].SectionParagraphs[0].IsListItem);
                Assert.IsTrue(sections[0].SectionParagraphs[1].IsListItem);
                Assert.IsTrue(sections[0].SectionParagraphs[2].IsListItem);

                Assert.AreEqual(sections[0].SectionParagraphs[0].ListItemType, ListItemType.Numbered);
                Assert.AreEqual(sections[0].SectionParagraphs[1].ListItemType, ListItemType.Numbered);
                Assert.AreEqual(sections[0].SectionParagraphs[2].ListItemType, ListItemType.Numbered);
            }
        }

        [TestMethod]
        public void Test_Unordered_List_When_Reading_Doc()
        {
            using (DocX document = DocX.Load(_directoryWithFiles + "testdoc_UnorderedList.docx"))
            {

                var sections = document.GetSections();

                Assert.IsTrue(sections[0].SectionParagraphs[0].IsListItem);
                Assert.IsTrue(sections[0].SectionParagraphs[1].IsListItem);
                Assert.IsTrue(sections[0].SectionParagraphs[2].IsListItem);

                Assert.AreEqual(sections[0].SectionParagraphs[0].ListItemType, ListItemType.Bulleted);
                Assert.AreEqual(sections[0].SectionParagraphs[1].ListItemType, ListItemType.Bulleted);
                Assert.AreEqual(sections[0].SectionParagraphs[2].ListItemType, ListItemType.Bulleted);
            }
        }

        [TestMethod]
        public void Test_Ordered_Unordered_Lists_When_Reading_Doc()
        {
            using (DocX document = DocX.Load(_directoryWithFiles + "testdoc_OrderedUnorderedLists.docx"))
            {

                var sections = document.GetSections();

                Assert.IsTrue(sections[0].SectionParagraphs[0].IsListItem);
                Assert.IsTrue(sections[0].SectionParagraphs[1].IsListItem);
                Assert.IsTrue(sections[0].SectionParagraphs[2].IsListItem);

                Assert.AreEqual(sections[0].SectionParagraphs[0].ListItemType, ListItemType.Numbered);
                Assert.AreEqual(sections[0].SectionParagraphs[1].ListItemType, ListItemType.Numbered);
                Assert.AreEqual(sections[0].SectionParagraphs[2].ListItemType, ListItemType.Numbered);

                Assert.IsTrue(sections[0].SectionParagraphs[3].IsListItem);
                Assert.IsTrue(sections[0].SectionParagraphs[4].IsListItem);
                Assert.IsTrue(sections[0].SectionParagraphs[5].IsListItem);

                Assert.AreEqual(sections[0].SectionParagraphs[3].ListItemType, ListItemType.Bulleted);
                Assert.AreEqual(sections[0].SectionParagraphs[4].ListItemType, ListItemType.Bulleted);
                Assert.AreEqual(sections[0].SectionParagraphs[5].ListItemType, ListItemType.Bulleted);

            }
        }

        [TestMethod]
        public void WhenCreatingAnOrderedListTheListXmlShouldHaveNumberedListItemType()
        {

            using (DocX document = DocX.Create("TestListXmlNumbered.docx"))
            {
                const int level = 0;
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var list = document.AddList("First Item", level, ListItemType.Numbered);
                document.InsertList(list);

                var listNumPropNode = document.mainDoc.Descendants().First(s => s.Name.LocalName == "numPr");

                var numId = listNumPropNode.Descendants().First(s => s.Name.LocalName == "numId");
                var abstractNum = list.GetAbstractNum(int.Parse(numId.GetAttribute(w + "val")));
                var lvl = abstractNum.Descendants().First(d => d.Name.LocalName == "lvl" && d.GetAttribute(w + "ilvl").Equals(level.ToString()));
                var numFormat = lvl.Descendants().First(d => d.Name.LocalName == "numFmt");

                Assert.AreEqual(numFormat.GetAttribute(w + "val").ToLower(), "decimal");
            }

        }

        [TestMethod]
        public void WhenCreatingAnUnOrderedListTheListXmlShouldHaveBulletListItemType()
        {

            using (DocX document = DocX.Create("TestListXmlBullet.docx"))
            {
                var list = document.AddList("First Item", 0, ListItemType.Bulleted);
                document.InsertList(list);

                var listNumPropNode = document.mainDoc.Descendants().First(s => s.Name.LocalName == "numPr");

                var numId = listNumPropNode.Descendants().First(s => s.Name.LocalName == "numId");

                Assert.AreEqual(numId.Attribute(DocX.w + "val").Value, "1");
            }
        }

        [TestMethod]
        public void WhenCreatingAListWithTextTheListXmlShouldHaveTheCorrectRunItemText()
        {
            using (DocX document = DocX.Create("TestListCreate.docx"))
            {
                const string listText = "RunText";
                var list = document.AddList(listText, 0, ListItemType.Bulleted);
                document.InsertList(list);

                var listNumPropNode = document.mainDoc.Descendants().First(s => s.Name.LocalName == "numPr");

                var runTextNode = document.mainDoc.Descendants().First(s => s.Name.LocalName == "t");

                Assert.IsNotNull(listNumPropNode);
                Assert.AreEqual(list.Items.First().runs.First().Value, runTextNode.Value);
                Assert.AreEqual(listText, runTextNode.Value);
            }
        }

        [TestMethod]
        public void WhenCreatingAnOrderedListTheListShouldHaveNumberedListItemType()
        {

            using (DocX document = DocX.Create("TestListCreateOrderedList.docx"))
            {
                var list = document.AddList("First Item");

                Assert.AreEqual(list.ListType, ListItemType.Numbered);
            }

        }

        [TestMethod]
        public void WhenCreatingAnUnOrderedListTheListShouldHaveBulletListItemType()
        {

            using (DocX document = DocX.Create("TestListCreateUnorderedList.docx"))
            {
                var list = document.AddList("First Item", 0, ListItemType.Bulleted);

                Assert.AreEqual(list.ListType, ListItemType.Bulleted);
            }

        }

        [TestMethod]
        public void WhenCreatingAListWithTextTheListShouldHaveTheCorrectRunItemText()
        {

            using (DocX document = DocX.Create("TestListCreateRunText.docx"))
            {
                var list = document.AddList("RunText", 0, ListItemType.Bulleted);
                document.InsertList(list);

                Assert.AreEqual(list.Items.First().runs.First().Value, "RunText");
            }
        }

        [TestMethod]
        public void WhenCreatingAListTheNumberingShouldGetSaved()
        {

        }

        [TestMethod]
        public void WhenCreatingAListTheListStyleShouldExistOrBeCreated()
        {

            using (DocX document = DocX.Create("TestListStyle.docx"))
            {
                var style = document.AddStylesForList();

                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

                bool listStyleExists =
                    (
                      from s in style.Element(w + "styles").Elements()
                      let styleId = s.Attribute(XName.Get("styleId", w.NamespaceName))
                      where (styleId != null && styleId.Value == "ListParagraph")
                      select s
                    ).Any();

                Assert.IsTrue(listStyleExists);

            }
        }

        [TestMethod]
        public void ANewListItemShouldCreateAnAbstractNumberingEntry()
        {
            using (DocX document = DocX.Create("TestNumbering.docx"))
            {
                var numbering = document.numbering.Descendants().Where(d => d.Name.LocalName == "abstractNum");
                Assert.IsFalse(numbering.Any());

                document.AddList("List Text");

                numbering = document.numbering.Descendants().Where(d => d.Name.LocalName == "abstractNum");
                Assert.IsTrue(numbering.Any());
            }
        }

        [TestMethod]
        public void ANewListItemShouldCreateANewNumEntry()
        {
            using (DocX document = DocX.Create("TestNumEntry.docx"))
            {
                var numbering = document.numbering.Descendants().Where(d => d.Name.LocalName == "num");
                Assert.IsFalse(numbering.Any());

                document.AddList("List Text");

                numbering = document.numbering.Descendants().Where(d => d.Name.LocalName == "num");
                Assert.IsTrue(numbering.Any());
            }
        }

        [TestMethod]
        public void CreateNewNumberingNumIdShouldAddNumberingDataToTheDocument()
        {
            using (DocX document = DocX.Create("TestCreateNumbering.docx"))
            {
                var numbering = document.numbering.Descendants().Where(d => d.Name.LocalName == "num");
                Assert.IsFalse(numbering.Any());
                var list = document.AddList("", 0, ListItemType.Bulleted);
                document.InsertList(list);

                numbering = document.numbering.Descendants().Where(d => d.Name.LocalName == "num");
                Assert.IsTrue(numbering.Any());
            }
        }

        [TestMethod]
        public void CreateNewNumberingNumIdShouldAddNumberingAbstractDataToTheDocument()
        {
            using (DocX document = DocX.Create("TestCreateNumberingAbstract.docx"))
            {
                var numbering = document.numbering.Descendants().Where(d => d.Name.LocalName == "abstractNum");
                Assert.IsFalse(numbering.Any());
                var list = document.AddList("", 0, ListItemType.Bulleted);
                document.InsertList(list);

                numbering = document.numbering.Descendants().Where(d => d.Name.LocalName == "abstractNum");
                Assert.IsTrue(numbering.Any());
            }
        }

        [TestMethod]
        public void IfPreviousElementIsAListThenAddingANewListContinuesThePreviousList()
        {
            using (DocX document = DocX.Create("TestAddListToPreviousList.docx"))
            {
                var list = document.AddList("List Text");
                document.AddListItem(list, "List Text2");
                document.InsertList(list);

                var lvlNodes = document.mainDoc.Descendants().Where(s => s.Name.LocalName == "ilvl").ToList();
                var numIdNodes = document.mainDoc.Descendants().Where(s => s.Name.LocalName == "numId").ToList();

                Assert.AreEqual(lvlNodes.Count(), 2);
                Assert.AreEqual(numIdNodes.Count(), 2);

                var prevLvlNode = lvlNodes[0];
                var newLvlNode = lvlNodes[1];

                Assert.AreEqual(prevLvlNode.Attribute(DocX.w + "val").Value, newLvlNode.Attribute(DocX.w + "val").Value);

                var prevNumIdNode = numIdNodes[0];
                var newNumIdNode = numIdNodes[1];

                Assert.AreEqual(prevNumIdNode.Attribute(DocX.w + "val").Value, newNumIdNode.Attribute(DocX.w + "val").Value);
                document.Save();
            }

        }

        [TestMethod]
        public void WhenADocumentHasListsTheListPropertyReturnsTheCorrectNumberOfLists()
        {

            using (DocX document = DocX.Load(_directoryWithFiles + "testdoc_OrderedUnorderedLists.docx"))
            {
                var lists = document.Lists;

                Assert.AreEqual(lists.Count, 2);

            }
        }

        [TestMethod]
        public void WhenADocumentIsCreatedWithAListItemThatHasASpecifiedStartNumber()
        {
            using (DocX document = DocX.Create("CreateListItemFromDifferentStartValue.docx"))
            {
                var list = document.AddList("Test", 0, ListItemType.Numbered, 5);
                document.AddListItem(list, "NewElement");

                var numbering = document.numbering.Descendants().Where(d => d.Name.LocalName == "abstractNum");
                var level = numbering.Descendants().First(el => el.Name.LocalName == "lvl");
                var start = level.Descendants().First(el => el.Name.LocalName == "start");
                Assert.AreEqual(start.GetAttribute(DocX.w + "val"), 5.ToString());
            }
        }

        [TestMethod]
        public void WhenANumberedAndBulletedListIsCreatedThereShouldBeTwoNumberingXmls()
        {
            using (DocX document = DocX.Create("NumberAndBulletListInOne.docx"))
            {
                var numberList = document.AddList("Test");
                document.AddListItem(numberList, "Second Numbered Item");

                var bulletedList = document.AddList("Bullet", 0, ListItemType.Bulleted);
                document.AddListItem(bulletedList, "Second bullet item");

                document.InsertList(numberList);
                document.InsertList(bulletedList);

                var abstractNums = document.numbering.Descendants().Where(d => d.Name.LocalName == "abstractNum");
                Assert.AreEqual(abstractNums.Count(), 2);

            }
        }

        [TestMethod]
        public void WhenICreateAnEmptyListAndAddEntriesToIt()
        {
            using (DocX document = DocX.Create("CreateEmptyListAndAddItemsToIt.docx"))
            {
                var list = document.AddList();
                Assert.AreEqual(list.Items.Count, 0);

                document.AddListItem(list, "Test item 1.");
                document.AddListItem(list, "Test item 2.");
                Assert.AreEqual(list.Items.Count, 2);
            }
        }

        [TestMethod]
        public void WhenICreateAHeaderItShouldHaveAStyle()
        {
            using (var document = DocX.Create("CreateHeaderElement.docx"))
            {
                document.InsertParagraph("Header Text 1").StyleName = "Header1";
                Assert.IsNotNull(document.styles.Root.Descendants().FirstOrDefault(d => d.GetAttribute(DocX.w + "styleId").ToLowerInvariant() == "heading1"));
                document.Save();
            }
        }


        [TestMethod]
        public void WhileReadingWhenTextIsBoldItalicUnderlineItShouldReadTheProperFormatting()
        {
            using (DocX document = DocX.Load(_directoryWithFiles + "FontFormat.docx"))
            {
                var underlinedTextFormatting = document.Paragraphs[0].MagicText[0].formatting;
                var boldTextFormatting = document.Paragraphs[0].MagicText[2].formatting;
                var italicTextFormatting = document.Paragraphs[0].MagicText[4].formatting;
                var boldItalicUnderlineTextFormatting = document.Paragraphs[0].MagicText[6].formatting;

                Assert.IsTrue(boldTextFormatting.Bold);
                Assert.IsTrue(italicTextFormatting.Italic);
                Assert.AreEqual(underlinedTextFormatting.UnderlineStyle, UnderlineStyle.singleLine);
                Assert.IsTrue(boldItalicUnderlineTextFormatting.Bold);
                Assert.IsTrue(boldItalicUnderlineTextFormatting.Italic);
                Assert.AreEqual(boldItalicUnderlineTextFormatting.UnderlineStyle, UnderlineStyle.singleLine);
            }
        }


        [TestMethod]
        public void WhileWritingWhenTextIsBoldItalicUnderlineItShouldReadTheProperFormatting()
        {
            using (DocX document = DocX.Create("FontFormatWrite.docx"))
            {

                Paragraph p = document.InsertParagraph();
                p.Append("This is bold.").Bold().Append("This is underlined.").UnderlineStyle(UnderlineStyle.singleLine).
                Append("This is italic.").Italic().Append("This is boldItalicUnderlined").Italic().Bold().UnderlineStyle(UnderlineStyle.singleLine);

                var boldTextFormatting = document.Paragraphs[0].MagicText[0].formatting;
                var underlinedTextFormatting = document.Paragraphs[0].MagicText[1].formatting;
                var italicTextFormatting = document.Paragraphs[0].MagicText[2].formatting;
                var boldItalicUnderlineTextFormatting = document.Paragraphs[0].MagicText[3].formatting;

                Assert.IsTrue(boldTextFormatting.Bold);
                Assert.IsTrue(italicTextFormatting.Italic);
                Assert.AreEqual(underlinedTextFormatting.UnderlineStyle, UnderlineStyle.singleLine);
                Assert.IsTrue(boldItalicUnderlineTextFormatting.Bold);
                Assert.IsTrue(boldItalicUnderlineTextFormatting.Italic);
                Assert.AreEqual(boldItalicUnderlineTextFormatting.UnderlineStyle, UnderlineStyle.singleLine);
            }
        }

        [TestMethod]
        public void InsertingANextPageBreakShouldAddADocumentSection()
        {
            using (DocX document = DocX.Create("SectionPageBreak.docx"))
            {
                document.InsertSectionPageBreak();

                var sections = document.GetSections();
                Assert.AreEqual(sections.Count, 2);
                document.Save();
            }
        }

      [TestMethod]
      public void InsertANextPageBreakWithParagraphTextsShouldAddProperParagraphsToProperSections()
      {
        using (DocX document = DocX.Create("SectionPageBreakWithParagraphs.docx"))
        {
          document.InsertParagraph("First paragraph.");
          document.InsertParagraph("Second paragraph.");
          document.InsertSectionPageBreak();
          document.InsertParagraph("Third paragraph.");
          document.InsertParagraph("Fourth paragraph.");

          var sections = document.GetSections();
          Assert.AreEqual(sections.Count, 2);

          Assert.AreEqual(sections[0].SectionParagraphs.Count(p => !string.IsNullOrWhiteSpace(p.Text)), 2);
          Assert.AreEqual(sections[1].SectionParagraphs.Count(p => !string.IsNullOrWhiteSpace(p.Text)), 2);
          document.Save();
        }
      }

      [TestMethod]
        public void WhenAFontFamilyIsSpecifiedForAParagraphItShouldSetTheFontOfTheParagraphTextToTheFontFamily()
        {
          using (DocX document = DocX.Create("FontTest.docx"))
          {
            Paragraph p = document.InsertParagraph();

            p.Append("Hello World").Font(new FontFamily("Symbol"));

            Assert.AreEqual(p.MagicText[0].formatting.FontFamily.Name, "Symbol");

            document.Save();
          }
        }

      [TestMethod]
      public void Test_Paragraph_RemoveTextManyLetters()
      {
          using (DocX document = DocX.Create(@"docs\HelloWorldRemovingManyLetters.docx"))
          {

              Paragraph p3 = document.InsertParagraph("");
              p3.Xml = XElement.Parse(
                  @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                    <w:pPr>
                    <w:ind />
                    </w:pPr>
                    <w:r>
                    <w:t>Based on the previous screening criteria, you qualify to participate in this particular survey. At the completion of the survey, you will be notified that your responses have been received and honoraria information will be captured for future payment. Thank you in advance for taking the time to participate with us. ^f('xMinutes').get() == 'xx' ? """" : ""</w:t>
                    </w:r>
                    <w:r>
                    <w:rPr>
                        <w:lang w:val=""pl-PL"" />
                    </w:rPr>
                    <w:t xml:space=""preserve"">This survey should take </w:t>
                    </w:r>
                    <w:r>
                    <w:t>"" + f('xMinutes').get() + ""</w:t>
                    </w:r>
                    <w:r>
                    <w:rPr>
                        <w:lang w:val=""pl-PL"" />
                    </w:rPr>
                    <w:t xml:space=""preserve""> minutes.  </w:t>
                    </w:r>
                    <w:r>
                    <w:t>""^Participants completing this survey will receive the honorarium designated in the invitation you have received. &lt;BR&gt;&lt;BR&gt;If you leave the survey prior to finishing it, you may return to your last question by visiting the same link provided in your email invitation (please be certain to use the same email that you used a moment ago to register for this study). If you have any questions or concerns about this study, please contact us at &lt;a href=""mailto:blabla@blabla.com?Subject=^f('sName')^ PD:^f('pdID')^""&gt;blabla@blabla.com&lt;/a&gt;. Thank you.</w:t>
                    </w:r>
                    </w:p>");

              int l1 = p3.Text.Length; //960
              p3.RemoveText(318, 99);
              int l2 = p3.Text.Length; //should be 861
              Assert.AreEqual(l1 - 99, l2);
          }
      }

      [TestMethod]
      public void Test_Table_RemoveParagraphs()
      {
          MemoryStream memoryStream;
          DocX document;
 
          memoryStream = new MemoryStream();
          document = DocX.Create(memoryStream);
          // Add a Table into the document.
          Table table = document.AddTable(1, 4); // 1 row, 4 cells
          table.Design = TableDesign.TableGrid;
          table.Alignment = Alignment.center;
          // Edit row
          var row = table.Rows[0];

          // Fill 1st paragraph
          row.Cells[0].Paragraphs.ElementAt(0).Append("Paragraph 1");
          // Fill 2nd paragraph
          var secondParagraph = row.Cells[0].InsertParagraph("Paragraph 2");

          // Check number of paragraphs
          Assert.AreEqual(2, row.Cells[0].Paragraphs.Count());

          // Remove 1st paragraph
          var deleted = row.Cells[0].RemoveParagraphAt(0);
          Assert.IsTrue(deleted);
          // Check number of paragraphs
          Assert.AreEqual(1, row.Cells[0].Paragraphs.Count());

          // Remove 3rd (nonexisting) paragraph
          deleted = row.Cells[0].RemoveParagraphAt(3);
          Assert.IsFalse(deleted);
          //check number of paragraphs
          Assert.AreEqual(1, row.Cells[0].Paragraphs.Count());

          // Remove secondParagraph (this time the only one) paragraph
          deleted = row.Cells[0].RemoveParagraph(secondParagraph);
          Assert.IsTrue(deleted);
          Assert.AreEqual(0, row.Cells[0].Paragraphs.Count());

          // Remove last paragraph once again - this time this paragraph does not exists
          deleted = row.Cells[0].RemoveParagraph(secondParagraph);
          Assert.IsFalse(deleted);
          Assert.AreEqual(0, row.Cells[0].Paragraphs.Count());
      }


    }
}
       