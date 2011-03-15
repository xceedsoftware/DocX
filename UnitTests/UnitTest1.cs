using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Novacode;
using System.Reflection;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Xml.Linq;
using System.IO.Packaging;

namespace UnitTests
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class UnitTest1
    {
        // Get the fullpath to the executing assembly.
        string directory_executing_assembly;
        string directory_documents;
        string file_temp = "temp.docx";

        const string package_part_document = "/word/document.xml";

        public UnitTest1()
        {
            directory_executing_assembly = Assembly.GetExecutingAssembly().Location;

            // The documents directory
            List<string> steps = directory_executing_assembly.Split('\\').ToList();
            steps.RemoveRange(steps.Count() - 3, 3);
            directory_documents = String.Join("\\", steps) + "\\documents\\";
        }

        [TestMethod]
        public void Test_EverybodyHasAHome_Loaded()
        {
            // Load a document.
            using (DocX document = DocX.Load(directory_documents + "EverybodyHasAHome.docx"))
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
            using (DocX document = DocX.Create(directory_documents + "Test.docx"))
            {
                // Add an Image to this document.
                Novacode.Image img = document.AddImage(directory_documents + "purple.png");

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
            using (DocX document = DocX.Create(directory_documents + "Test.docx"))
            {
                // Add an Image to this document.
                Novacode.Image img = document.AddImage(directory_documents + "purple.png");

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
            using (DocX document = DocX.Create(directory_documents + "test_add_images.docx"))
            {
                // Add a png to into this document
                Novacode.Image png = document.AddImage(directory_documents + "purple.png");
                Assert.IsTrue(document.Images.Count == 1);
                Assert.IsTrue(Path.GetExtension(png.pr.TargetUri.OriginalString) == ".png");

                // Add a tiff into to this document
                Novacode.Image tif = document.AddImage(directory_documents + "yellow.tif");
                Assert.IsTrue(document.Images.Count == 2);
                Assert.IsTrue(Path.GetExtension(tif.pr.TargetUri.OriginalString) == ".tif");

                // Add a gif into to this document
                Novacode.Image gif = document.AddImage(directory_documents + "orange.gif");
                Assert.IsTrue(document.Images.Count == 3);
                Assert.IsTrue(Path.GetExtension(gif.pr.TargetUri.OriginalString) == ".gif");

                // Add a jpg into to this document
                Novacode.Image jpg = document.AddImage(directory_documents + "green.jpg");
                Assert.IsTrue(document.Images.Count == 4);
                Assert.IsTrue(Path.GetExtension(jpg.pr.TargetUri.OriginalString) == ".jpg");

                // Add a bitmap to this document
                Novacode.Image bitmap = document.AddImage(directory_documents + "red.bmp");
                Assert.IsTrue(document.Images.Count == 5);
                // Word does not allow bmp make sure it was inserted as a png.
                Assert.IsTrue(Path.GetExtension(bitmap.pr.TargetUri.OriginalString) == ".png");
            }
        }

        [TestMethod]
        public void Test_Document_AddImage_FromStream()
        {
            using (DocX document = DocX.Create(directory_documents + "test_add_images.docx"))
            {
                // DocX will always insert Images that come from Streams as jpeg.

                // Add a png to into this document
                Novacode.Image png = document.AddImage(new FileStream(directory_documents + "purple.png", FileMode.Open));
                Assert.IsTrue(document.Images.Count == 1);
                Assert.IsTrue(Path.GetExtension(png.pr.TargetUri.OriginalString) == ".jpeg");

                // Add a tiff into to this document
                Novacode.Image tif = document.AddImage(new FileStream(directory_documents + "yellow.tif", FileMode.Open));
                Assert.IsTrue(document.Images.Count == 2);
                Assert.IsTrue(Path.GetExtension(tif.pr.TargetUri.OriginalString) == ".jpeg");

                // Add a gif into to this document
                Novacode.Image gif = document.AddImage(new FileStream(directory_documents + "orange.gif", FileMode.Open));
                Assert.IsTrue(document.Images.Count == 3);
                Assert.IsTrue(Path.GetExtension(gif.pr.TargetUri.OriginalString) == ".jpeg");

                // Add a jpg into to this document
                Novacode.Image jpg = document.AddImage(new FileStream(directory_documents + "green.jpg", FileMode.Open));
                Assert.IsTrue(document.Images.Count == 4);
                Assert.IsTrue(Path.GetExtension(jpg.pr.TargetUri.OriginalString) == ".jpeg");

                // Add a bitmap to this document
                Novacode.Image bitmap = document.AddImage(new FileStream(directory_documents + "red.bmp", FileMode.Open));
                Assert.IsTrue(document.Images.Count == 5);
                // Word does not allow bmp make sure it was inserted as a png.
                Assert.IsTrue(Path.GetExtension(bitmap.pr.TargetUri.OriginalString) == ".jpeg");
            }
        }

        [TestMethod]
        public void Test_Tables()
        {
            using (DocX document = DocX.Load(directory_documents + "Tables.docx"))
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
            using (DocX document = DocX.Load(directory_documents + "Images.docx"))
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
            using (DocX document = DocX.Create(directory_documents + "Test.docx"))
            {
                // Add Headers and Footers into this document.
                document.AddHeaders();
                document.AddFooters();
                document.DifferentFirstPage = true;
                document.DifferentOddAndEvenPages = true;

                // Add an Image to this document.
                Novacode.Image img = document.AddImage(directory_documents + "purple.png");

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
            using (DocX document = DocX.Create(directory_documents + "Test.docx"))
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
        public void Test_Append_Hyperlink()
        {
            // Load test document.
            using (DocX document = DocX.Create(directory_documents + "Test.docx"))
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
            using (DocX document = DocX.Create(directory_documents + "Test.docx"))
            {
                // Add Headers and Footers into this document.
                document.AddHeaders();
                document.AddFooters();
                document.DifferentFirstPage = true;
                document.DifferentOddAndEvenPages = true;

                // Add an Image to this document.
                Novacode.Image img = document.AddImage(directory_documents + "purple.png");

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
            using (DocX document = DocX.Load(directory_documents + "MovePicture.docx"))
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
                document.SaveAs(directory_documents + "MovedPicture.docx");
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
                p1.InsertHyperlink(h);                       Assert.IsTrue(p1.Text == "linkAC");
                p1.InsertHyperlink(h, p1.Text.Length);       Assert.IsTrue(p1.Text == "linkAClink");
                p1.InsertHyperlink(h, p1.Text.IndexOf("C")); Assert.IsTrue(p1.Text == "linkAlinkClink");

                // Difficult
                Paragraph p2 = document.InsertParagraph("\tA\tC\t");
                p2.InsertHyperlink(h);                       Assert.IsTrue(p2.Text == "link\tA\tC\t");
                p2.InsertHyperlink(h, p2.Text.Length);       Assert.IsTrue(p2.Text == "link\tA\tC\tlink");
                p2.InsertHyperlink(h, p2.Text.IndexOf("C")); Assert.IsTrue(p2.Text == "link\tA\tlinkC\tlink");

                // Contrived
                // Add a contrived Hyperlink to this document.
                Hyperlink h2 = document.AddHyperlink("\tlink\t", new Uri("http://www.google.com"));
                Paragraph p3 = document.InsertParagraph("\tA\tC\t");
                p3.InsertHyperlink(h2);                    Assert.IsTrue(p3.Text == "\tlink\t\tA\tC\t");
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
                catch (ArgumentException) {}
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
            using (DocX document = DocX.Load(directory_documents + "Paragraphs.docx"))
            {
                // Extract the Paragraphs from this document.
                List<Paragraph> paragraphs = document.Paragraphs;

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
                document.Paragraphs.ForEach(p => Assert.IsTrue(p.PackagePart.Uri.ToString() == package_part_document));

                // Test the saving of the document.
                document.SaveAs(file_temp);
            }

            // Delete the tempory file.
            File.Delete(file_temp);
        }

        [TestMethod]
        public void Test_Table_InsertRow()
        {
            using (DocX document = DocX.Create(directory_documents + "Tables2.docx"))
            {
                // Add a Table to a document.
                Table t = document.AddTable(2, 2);
                t.Design = TableDesign.TableGrid;

                // Insert the Table into the main section of the document.
                Table t1 = document.InsertTable(t);
                t1.InsertRow();
                t1.InsertColumn();

                // Save the document.
                document.Save();
            }
        }

      [TestMethod]
      public void Test_Document_ApplyTemplate()
      {
        using (MemoryStream documentStream = new MemoryStream())
        {
          using (DocX document = DocX.Create(documentStream))
          {
            document.ApplyTemplate(directory_documents + "Template.dotx");
            document.Save();
            Header firstHeader = document.Headers.first;
            Header oddHeader = document.Headers.odd;
            Header evenHeader = document.Headers.even;

            Footer firstFooter = document.Footers.first;
            Footer oddFooter = document.Footers.odd;
            Footer evenFooter = document.Footers.even;

            Assert.IsTrue(firstHeader.Paragraphs.Count==1, "More than one paragraph in header.");
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
    }
}
