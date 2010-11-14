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
        public void Test_Tables()
        {
            using (DocX document = DocX.Load(directory_documents + "Tables.docx"))
            {
                // There is only one Paragraph at the document level.
                Assert.IsTrue(document.Paragraphs.Count() == 1);

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

        // Write the string "Hello World" into this Image.
        private static void CoolExample(Novacode.Image i, Stream s, string str)
        {
            // Write "Hello World" into this Image.
            Bitmap b = new Bitmap(s);

            /* 
             * Get the Graphics object for this Bitmap.
             * The Graphics object provides functions for drawing.
             */
            Graphics g = Graphics.FromImage(b);

            // Draw the string "Hello World".
            g.DrawString
            (
                str,
                new Font("Tahoma", 20),
                Brushes.Blue,
                new PointF(0, 0)
            );

            // Save this Bitmap back into the document using a Create\Write stream.
            b.Save(i.GetStream(FileMode.Create, FileAccess.Write), ImageFormat.Png);
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
                p1.InsertHyperlink(0, h);                    Assert.IsTrue(p1.Text == "linkAC");
                p1.InsertHyperlink(p1.Text.Length, h);       Assert.IsTrue(p1.Text == "linkAClink");
                p1.InsertHyperlink(p1.Text.IndexOf("C"), h); Assert.IsTrue(p1.Text == "linkAlinkClink");

                // Difficult
                Paragraph p2 = document.InsertParagraph("\tA\tC\t");
                p2.InsertHyperlink(0, h);                    Assert.IsTrue(p2.Text == "link\tA\tC\t");
                p2.InsertHyperlink(p2.Text.Length, h);       Assert.IsTrue(p2.Text == "link\tA\tC\tlink");
                p2.InsertHyperlink(p2.Text.IndexOf("C"), h); Assert.IsTrue(p2.Text == "link\tA\tlinkC\tlink");

                // Contrived
                // Add a contrived Hyperlink to this document.
                Hyperlink h2 = document.AddHyperlink("\tlink\t", new Uri("http://www.google.com"));
                Paragraph p3 = document.InsertParagraph("\tA\tC\t");
                p3.InsertHyperlink(0, h2);                    Assert.IsTrue(p3.Text == "\tlink\t\tA\tC\t");
                p3.InsertHyperlink(p3.Text.Length, h2);       Assert.IsTrue(p3.Text == "\tlink\t\tA\tC\t\tlink\t");
                p3.InsertHyperlink(p3.Text.IndexOf("C"), h2); Assert.IsTrue(p3.Text == "\tlink\t\tA\t\tlink\tC\t\tlink\t");
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
            // This document contains a run with two text next to each other.
            // <run>
            //   <text>Hello World</text>
            //   <text>foo</text>
            // </run>
            using (DocX document = DocX.Load(@"C:\Users\Cathal\Desktop\Bug.docx"))
            {
                Paragraph p = document.Paragraphs[0];
                Assert.IsTrue(p.Text == "Hello worldfoo");
                p.RemoveText("Hello world".Length, 3, false);
                Assert.IsTrue(p.Text == "Hello world");
            }

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

                // Create a string to append to each Paragraph.
                string appended_text = "foo bar foo";

                // Test the appending of text to each Paragraph.
                Assert.IsTrue(p1.Append(appended_text).Text == p1_text + appended_text);
                Assert.IsTrue(p2.Append(appended_text).Text == p2_text + appended_text);
                Assert.IsTrue(p3.Append(appended_text).Text == p3_text + appended_text);

                // Test FindAll
                List<int> p1_foos = p1.FindAll("foo");
                Assert.IsTrue(p1_foos.Count() == 2 && p1_foos[0] == 11 && p1_foos[1] == 19);

                // Test ReplaceText
                p2.ReplaceText("foo", "bar", false);

                Assert.IsTrue(p2.Text == "Paragraph 2bar bar bar");

                // Test RemoveText
                p3.RemoveText(1, 3, false);
                Assert.IsTrue(p3.Text == "Pgraph 3foo bar foo");

                // Its important that each Paragraph knows the PackagePart it belongs to.
                document.Paragraphs.ForEach(p => Assert.IsTrue(p.PackagePart.Uri.ToString() == package_part_document));

                // Test the saving of the document.
                document.SaveAs(file_temp);
            }

            // Delete the tempory file.
            File.Delete(file_temp);
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
