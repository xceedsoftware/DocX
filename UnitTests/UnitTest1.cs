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
    }
}
