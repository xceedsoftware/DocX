using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Novacode;
using WindowsBitmap = System.Drawing.Bitmap;
using WindowsBrushes = System.Drawing.Brushes;
using WindowsColor = System.Drawing.Color;
using WindowsFont = System.Drawing.Font;
using WindowsFontFamily = System.Drawing.FontFamily;
using WindowsGraphics = System.Drawing.Graphics;
using WindowsImageFormat = System.Drawing.Imaging.ImageFormat;

namespace Examples
{
    class Program
    {
        private static Border BlankBorder = new Border(BorderStyle.Tcbs_none, 0, 0, WindowsColor.White);

        static void Main(string[] args)
        {
            Setup();
            Examples();
        }

        static void Examples()
        {
            // Easy
            Console.WriteLine("\nRunning Easy Examples");
            HelloWorld();
            HelloWorldKeepLinesTogether();
            HelloWorldKeepWithNext();
            HelloWorldAdvancedFormatting();
            HelloWorldProtectedDocument();
            HelloWorldAddPictureToWord();
            HelloWorldInsertHorizontalLine();
            RightToLeft();
            Indentation();
            HeadersAndFooters();
            HyperlinksImagesTables();
            AddList();
            Equations();
            Bookmarks();
            BookmarksReplaceTextOfBookmarkKeepingFormat();
            BarChart();
            PieChart();
            LineChart();
            Chart3D();
            DocumentMargins();
            CreateTableWithTextDirection();
            CreateTableRowsFromTemplate();
            AddToc();
            AddTocByReference();

            // Intermediate
            Console.WriteLine("\nRunning Intermediate Examples");
            CreateInvoice();
            HyperlinksImagesTablesWithLists();
            HeadersAndFootersWithImagesAndTables();
            HeadersAndFootersWithImagesAndTablesUsingInsertPicture();
            DocumentsWithListsFontChange();
            DocumentHeading();
            LargeTable();
            TableWithSpecifiedWidths();
            //Contents();

            // Advanced
            Console.WriteLine("\nRunning Advanced Examples");
            ProgrammaticallyManipulateImbeddedImage();
            ReplaceTextParallel();

            Console.WriteLine("\nPress any key to exit.");
            Console.ReadKey();
        }

        private static void Setup()
        {
            if (!Directory.Exists("docs"))
            {
                Directory.CreateDirectory("docs");
            }
        }

        #region Charts

        private class ChartData
        {
            public String Mounth { get; set; }
            public Double Money { get; set; }

            public static List<ChartData> CreateCompanyList1()
            {
                List<ChartData> company1 = new List<ChartData>();
                company1.Add(new ChartData() { Mounth = "January", Money = 100 });
                company1.Add(new ChartData() { Mounth = "February", Money = 120 });
                company1.Add(new ChartData() { Mounth = "March", Money = 140 });
                return company1;
            }

            public static List<ChartData> CreateCompanyList2()
            {
                List<ChartData> company2 = new List<ChartData>();
                company2.Add(new ChartData() { Mounth = "January", Money = 80 });
                company2.Add(new ChartData() { Mounth = "February", Money = 160 });
                company2.Add(new ChartData() { Mounth = "March", Money = 130 });
                return company2;
            }
        }

        private static void BarChart()
        {
            Console.WriteLine("\tBarChart()");
            // Create new document. 
            using (DocX document = DocX.Create(@"docs\BarChart.docx"))
            {
                // Create chart.
                BarChart c = new BarChart();
                c.BarDirection = BarDirection.Column;
                c.BarGrouping = BarGrouping.Standard;
                c.showVal = true;
                c.GapWidth = 400;
                c.AddLegend(ChartLegendPosition.Bottom, false);

                // Create data.
                List<ChartData> company1 = ChartData.CreateCompanyList1();
                List<ChartData> company2 = ChartData.CreateCompanyList2();

                // Create and add series
                Series s1 = new Series("Microsoft");
                s1.Color = WindowsColor.GreenYellow;
                s1.Bind(company1, "Mounth", "Money");
                c.AddSeries(s1);
                Series s2 = new Series("Apple");
                s2.Bind(company2, "Mounth", "Money");
                c.AddSeries(s2);

                // Insert chart into document
                document.InsertParagraph("Diagram").FontSize(20);
                document.InsertChart(c);
                document.Save();
            }
            Console.WriteLine("\tCreated: docs\\BarChart.docx\n");
        }

        private static void PieChart()
        {
            Console.WriteLine("\tPieChart()");
            // Create new document. 
            using (DocX document = DocX.Create(@"docs\PieChart.docx"))
            {
                // Create chart.
                PieChart c = new PieChart();
                c.AddLegend(ChartLegendPosition.Bottom, false);

                // Create data.
                List<ChartData> company2 = ChartData.CreateCompanyList2();

                // Create and add series
                Series s = new Series("Apple");
                s.Bind(company2, "Mounth", "Money");
                c.AddSeries(s);

                // Insert chart into document
                document.InsertParagraph("Diagram").FontSize(20);
                document.InsertChart(c);
                document.Save();
            }
            Console.WriteLine("\tCreated: docs\\PieChart.docx\n");
        }

        private static void LineChart()
        {
            Console.WriteLine("\tLineChart()");
            // Create new document. 
            using (DocX document = DocX.Create(@"docs\LineChart.docx"))
            {
                // Create chart.
                LineChart c = new LineChart();
                c.AddLegend(ChartLegendPosition.Bottom, false);

                // Create data.
                List<ChartData> company1 = ChartData.CreateCompanyList1();
                List<ChartData> company2 = ChartData.CreateCompanyList2();

                // Create and add series
                Series s1 = new Series("Microsoft");
                s1.Color = WindowsColor.GreenYellow;
                s1.Bind(company1, "Mounth", "Money");
                c.AddSeries(s1);
                Series s2 = new Series("Apple");
                s2.Bind(company2, "Mounth", "Money");
                c.AddSeries(s2);

                // Insert chart into document
                document.InsertParagraph("Diagram").FontSize(20);
                document.InsertChart(c);
                document.Save();
            }
            Console.WriteLine("\tCreated: docs\\LineChart.docx\n");
        }

        private static void Chart3D()
        {
            Console.WriteLine("\tChart3D()");
            // Create new document. 
            using (DocX document = DocX.Create(@"docs\3DChart.docx"))
            {
                // Create chart.
                BarChart c = new BarChart();
                c.View3D = true;

                // Create data.
                List<ChartData> company1 = ChartData.CreateCompanyList1();

                // Create and add series
                Series s = new Series("Microsoft");
                s.Color = WindowsColor.GreenYellow;
                s.Bind(company1, "Mounth", "Money");
                c.AddSeries(s);

                // Insert chart into document
                document.InsertParagraph("3D Diagram").FontSize(20);
                document.InsertChart(c);
                document.Save();
            }
            Console.WriteLine("\tCreated: docs\\3DChart.docx\n");
        }

        #endregion

        /// <summary>
        /// Load a document and set content controls.
        /// </summary>
        private static void Contents()
        {
            Console.WriteLine("\tContent()");

            // Load a document.
            using (DocX document = DocX.Load(@"docs\Content.docx"))
            {
                foreach (var c in document.Contents)
                {
                    Console.WriteLine(String.Format("Name : {0}, Tag : {1}", c.Name, c.Tag));
                }

                (from d in document.Contents
                 where d.Name == "Name"
                 select d).First().SetText("NewerText");

                document.SaveAs(@"docs\ContentSetSingle.docx");

                XElement el = new XElement("Root",
                        new XElement("Name", "Claudia"),
                        new XElement("Address", "17 Liberty St"),
                        new XElement("Total", "123.45")

                        );

                XDocument doc = new XDocument(el);
                document.SetContent(el);
                document.SaveAs(@"docs\ContentSetWithElement.docx");


                doc.Save(@"docs\elements.xml");

                document.SetContent(@"docs\elements.xml");
                document.SaveAs(@"docs\ContentSetWithPath.docx");


            }
        }
        /// <summary>
        /// Create a document wit(h two equations.
        /// </summary>
        private static void Equations()
        {
            Console.WriteLine("\tEquations()");

            // Create a new document.
            using (DocX document = DocX.Create(@"docs\Equations.docx"))
            {
                // Insert first Equation in this document.
                Paragraph pEquation1 = document.InsertEquation("x = y+z");

                // Insert second Equation in this document and add formatting.
                Paragraph pEquation2 = document.InsertEquation("x = (y+z)/t").FontSize(18).Color(WindowsColor.Blue);

                // Save this document to disk.
                document.Save();
                Console.WriteLine("\tCreated: docs\\Equations.docx\n");
            }
        }
        public static void DocumentHeading()
        {
            Console.WriteLine("\tDocumentHeading()");
            using (DocX document = DocX.Create(@"docs\DocumentHeading.docx"))
            {

                foreach (HeadingType heading in (HeadingType[])Enum.GetValues(typeof(HeadingType)))
                {
                    string text = string.Format("{0} - The quick brown fox jumps over the lazy dog", heading.EnumDescription());

                    Paragraph p = document.InsertParagraph();
                    p.AppendLine(text).Heading(heading);
                }


                document.Save();
                Console.WriteLine("\tCreated: docs\\DocumentHeading.docx\n");
            }
        }

        /// <summary>
        /// Loads a document having a table with a given line as template.
        /// It avoids extra manipulation regarding style 
        /// </summary>
        private static void CreateTableRowsFromTemplate()
        {
            Console.WriteLine("\tCreateTableFromTemplate()");

            using (DocX docX = DocX.Load(@"docs\DocumentWithTemplateTable.docx"))
            {
                //look for one specific table here
                Table orderTable = docX.Tables.First(t => t.TableCaption == "ORDER_TABLE");
                if (orderTable != null)
                {
                    //Row 0 and 1 are Headers
                    //Row 2 is pattern
                    if (orderTable.RowCount >= 2)
                    {
                        //get the Pattern row for duplication
                        Row orderRowPattern = orderTable.Rows[2];
                        //Add 5 lines of product
                        for (int i = 0; i < 5; i++)
                        {
                            //InsertRow performs a copy, so we get markup in new line ready for replacements
                            Row newOrderRow = orderTable.InsertRow(orderRowPattern, 2 + i);
                            newOrderRow.ReplaceText("%PRODUCT_NAME%", "Product_" + i);
                            newOrderRow.ReplaceText("%PRODUCT_PRICE1%", "$ " + i * new Random().Next(1, 50));
                            newOrderRow.ReplaceText("%PRODUCT_PRICE2%", "$ " + i * new Random().Next(1, 50));
                        }

                        //pattern row is at the end now, can be removed from table
                        orderRowPattern.Remove();

                    }
                    docX.SaveAs(@"docs\CreateTableFromTemplate.docx");

                }
                else
                {
                    Console.WriteLine("\tError, couldn't find table with caption ORDER_TABLE in document");
                }


            }
            Console.WriteLine("\tCreated: docs\\CreateTableFromTemplate.docx");
        }

        private static void Bookmarks()
        {
            Console.WriteLine("\tBookmarks()");

            using (var document = DocX.Create(@"docs\Bookmarks.docx"))
            {
                var paragraph = document.InsertBookmark("firstBookmark");

                var paragraph2 = document.InsertParagraph("This is a paragraph which contains a ");
                paragraph2.AppendBookmark("secondBookmark");
                paragraph2.Append("bookmark");

                paragraph2.InsertAtBookmark("handy ", "secondBookmark");

                document.Save();
                Console.WriteLine("\tCreated: docs\\Bookmarks.docx\n");

            }
        }
        /// <summary>
        /// Loads a document 'DocumentWithBookmarks.docx' and changes text inside bookmark keeping formatting the same.
        /// This code creates the file 'BookmarksReplaceTextOfBookmarkKeepingFormat.docx'.
        /// </summary>
        private static void BookmarksReplaceTextOfBookmarkKeepingFormat()
        {
            Console.WriteLine("\tBookmarksReplaceTextOfBookmarkKeepingFormat()");

            using (DocX docX = DocX.Load(@"docs\DocumentWithBookmarks.docx"))
            {
                foreach (Bookmark bookmark in docX.Bookmarks)
                    Console.WriteLine("\t\tFound bookmark {0}", bookmark.Name);

                // Replace bookmars content
                docX.Bookmarks["bmkNoContent"].SetText("Here there was a bookmark");
                docX.Bookmarks["bmkContent"].SetText("Here there was a bookmark with a previous content");
                docX.Bookmarks["bmkFormattedContent"].SetText("Here there was a formatted bookmark");

                docX.SaveAs(@"docs\BookmarksReplaceTextOfBookmarkKeepingFormat.docx");
            }
            Console.WriteLine("\tCreated: docs\\BookmarksReplaceTextOfBookmarkKeepingFormat.docx");
        }

        /// <summary>
        /// Create a document with a Paragraph whos first line is indented.
        /// </summary>
        private static void Indentation()
        {
            Console.WriteLine("\tIndentation()");

            // Create a new document.
            using (DocX document = DocX.Create(@"docs\Indentation.docx"))
            {
                // Create a new Paragraph.
                Paragraph p = document.InsertParagraph("Line 1\nLine 2\nLine 3");

                // Indent only the first line of the Paragraph.
                p.IndentationFirstLine = 1.0f;


                // Save all changes made to this document.
                document.Save();
                Console.WriteLine("\tCreated: docs\\Indentation.docx\n");
            }
        }

        /// <summary>
        /// Create a document that with RightToLeft text flow.
        /// </summary>
        private static void RightToLeft()
        {
            Console.WriteLine("\tRightToLeft()");

            // Create a new document.
            using (DocX document = DocX.Create(@"docs\RightToLeft.docx"))
            {
                // Create a new Paragraph with the text "Hello World".
                Paragraph p = document.InsertParagraph("Hello World.");

                // Make this Paragraph flow right to left. Default is left to right.
                p.Direction = Direction.RightToLeft;

                // You don't need to manually set the text direction foreach Paragraph, you can just call this function.
                document.SetDirection(Direction.RightToLeft);

                // Save all changes made to this document.
                document.Save();
                Console.WriteLine("\tCreated: docs\\RightToLeft.docx\n");
            }
        }

        /// <summary>
        /// Creates a document with a Hyperlink, an Image and a Table.
        /// </summary>
        private static void HyperlinksImagesTables()
        {
            Console.WriteLine("\tHyperlinksImagesTables()");

            // Create a document.
            using (DocX document = DocX.Create(@"docs\HyperlinksImagesTables.docx"))
            {
                // Add a hyperlink into the document.
                Hyperlink link = document.AddHyperlink("link", new Uri("http://www.google.com"));

                // Add a Table into the document.
                Table table = document.AddTable(2, 2);
                table.Design = TableDesign.ColorfulGridAccent2;
                table.Alignment = Alignment.center;
                table.Rows[0].Cells[0].Paragraphs[0].Append("1");
                table.Rows[0].Cells[1].Paragraphs[0].Append("2");
                table.Rows[1].Cells[0].Paragraphs[0].Append("3");
                table.Rows[1].Cells[1].Paragraphs[0].Append("4");

                Row newRow = table.InsertRow(table.Rows[1]);
                newRow.ReplaceText("4", "5");

                // Add an image into the document.    
                RelativeDirectory rd = new RelativeDirectory(); // prepares the files for testing
                rd.Up(2);
                Image image = document.AddImage(rd.Path + @"\images\logo_template.png");

                // Create a picture (A custom view of an Image).
                Picture picture = image.CreatePicture();
                picture.Rotation = 10;
                picture.SetPictureShape(BasicShapes.cube);

                // Insert a new Paragraph into the document.
                Paragraph title = document.InsertParagraph().Append("Test").FontSize(20).Font(new Font("Comic Sans MS"));
                title.Alignment = Alignment.center;

                // Insert a new Paragraph into the document.
                Paragraph p1 = document.InsertParagraph();

                // Append content to the Paragraph
                p1.AppendLine("This line contains a ").Append("bold").Bold().Append(" word.");
                p1.AppendLine("Here is a cool ").AppendHyperlink(link).Append(".");
                p1.AppendLine();
                p1.AppendLine("Check out this picture ").AppendPicture(picture).Append(" its funky don't you think?");
                p1.AppendLine();
                p1.AppendLine("Can you check this Table of figures for me?");
                p1.AppendLine();

                // Insert the Table after Paragraph 1.
                p1.InsertTableAfterSelf(table);

                // Insert a new Paragraph into the document.
                Paragraph p2 = document.InsertParagraph();

                // Append content to the Paragraph.
                p2.AppendLine("Is it correct?");

                // Save this document.
                document.Save();

                Console.WriteLine("\tCreated: docs\\HyperlinksImagesTables.docx\n");
            }
        }
        private static void HyperlinksImagesTablesWithLists()
        {
            Console.WriteLine("\tHyperlinksImagesTablesWithLists()");

            // Create a document.
            using (DocX document = DocX.Create(@"docs\HyperlinksImagesTablesWithLists.docx"))
            {
                // Add a hyperlink into the document.
                Hyperlink link = document.AddHyperlink("link", new Uri("http://www.google.com"));

                // created numbered lists 
                var numberedList = document.AddList("First List Item.", 0, ListItemType.Numbered, 1);
                document.AddListItem(numberedList, "First sub list item", 1);
                document.AddListItem(numberedList, "Second List Item.");
                document.AddListItem(numberedList, "Third list item.");
                document.AddListItem(numberedList, "Nested item.", 1);
                document.AddListItem(numberedList, "Second nested item.", 1);

                // created bulleted lists

                var bulletedList = document.AddList("First Bulleted Item.", 0, ListItemType.Bulleted);
                document.AddListItem(bulletedList, "Second bullet item");
                document.AddListItem(bulletedList, "Sub bullet item", 1);
                document.AddListItem(bulletedList, "Second sub bullet item", 1);
                document.AddListItem(bulletedList, "Third bullet item");


                // Add a Table into the document.
                Table table = document.AddTable(2, 2);
                table.Design = TableDesign.ColorfulGridAccent2;
                table.Alignment = Alignment.center;
                table.Rows[0].Cells[0].Paragraphs[0].Append("1");
                table.Rows[0].Cells[1].Paragraphs[0].Append("2");
                table.Rows[1].Cells[0].Paragraphs[0].Append("3");
                table.Rows[1].Cells[1].Paragraphs[0].Append("4");

                Row newRow = table.InsertRow(table.Rows[1]);
                newRow.ReplaceText("4", "5");

                // Add an image into the document.    
                RelativeDirectory rd = new RelativeDirectory(); // prepares the files for testing
                rd.Up(2);
                Image image = document.AddImage(rd.Path + @"\images\logo_template.png");

                // Create a picture (A custom view of an Image).
                Picture picture = image.CreatePicture();
                picture.Rotation = 10;
                picture.SetPictureShape(BasicShapes.cube);

                // Insert a new Paragraph into the document.
                Paragraph title = document.InsertParagraph().Append("Test").FontSize(20).Font(new Font("Comic Sans MS"));
                title.Alignment = Alignment.center;



                // Insert a new Paragraph into the document.
                Paragraph p1 = document.InsertParagraph();

                // Append content to the Paragraph
                p1.AppendLine("This line contains a ").Append("bold").Bold().Append(" word.");
                p1.AppendLine("Here is a cool ").AppendHyperlink(link).Append(".");
                p1.AppendLine();
                p1.AppendLine("Check out this picture ").AppendPicture(picture).Append(" its funky don't you think?");
                p1.AppendLine();
                p1.AppendLine("Can you check this Table of figures for me?");
                p1.AppendLine();

                // Insert the Table after Paragraph 1.
                p1.InsertTableAfterSelf(table);

                // Insert a new Paragraph into the document.
                Paragraph p2 = document.InsertParagraph();

                // Append content to the Paragraph.
                p2.AppendLine("Is it correct?");
                p2.AppendLine();
                p2.AppendLine("Adding bullet list below: ");

                document.InsertList(bulletedList);

                // Adding another paragraph to add table and bullet list after it
                Paragraph p3 = document.InsertParagraph();
                p3.AppendLine();
                p3.AppendLine("Adding another table...");

                // Adding another table
                Table table1 = document.AddTable(2, 2);
                table1.Design = TableDesign.ColorfulGridAccent2;
                table1.Alignment = Alignment.center;
                table1.Rows[0].Cells[0].Paragraphs[0].Append("1");
                table1.Rows[0].Cells[1].Paragraphs[0].Append("2");
                table1.Rows[1].Cells[0].Paragraphs[0].Append("3");
                table1.Rows[1].Cells[1].Paragraphs[0].Append("4");

                Paragraph p4 = document.InsertParagraph();
                p4.InsertTableBeforeSelf(table1);

                p4.AppendLine();


                // Insert numbered list after table
                Paragraph p5 = document.InsertParagraph();
                p5.AppendLine("Adding numbered list below: ");
                p5.AppendLine();
                document.InsertList(numberedList);




                // Save this document.
                document.Save();

                Console.WriteLine("\tCreated: docs\\HyperlinksImagesTablesWithLists.docx\n");
            }
        }

        private static void DocumentMargins()
        {
            Console.WriteLine("\tDocumentMargins()");

            // Create a document.
            using (DocX document = DocX.Create(@"docs\DocumentMargins.docx"))
            {

                // Create a float var that contains doc Margins properties.
                float leftMargin = document.MarginLeft;
                float rightMargin = document.MarginRight;
                float topMargin = document.MarginTop;
                float bottomMargin = document.MarginBottom;
                // Modify using your own vars.
                leftMargin = 95F;
                rightMargin = 45F;
                topMargin = 50F;
                bottomMargin = 180F;

                // Or simply work the margins by setting the property directly. 
                document.MarginLeft = leftMargin;
                document.MarginRight = rightMargin;
                document.MarginTop = topMargin;
                document.MarginBottom = bottomMargin;

                // created bulleted lists

                var bulletedList = document.AddList("First Bulleted Item.", 0, ListItemType.Bulleted);
                document.AddListItem(bulletedList, "Second bullet item");
                document.AddListItem(bulletedList, "Sub bullet item", 1);
                document.AddListItem(bulletedList, "Second sub bullet item", 1);
                document.AddListItem(bulletedList, "Third bullet item");


                document.InsertList(bulletedList);

                // Save this document.
                document.Save();

                Console.WriteLine("\tCreated: docs\\DocumentMargins.docx\n");
            }
        }

        private static void DocumentsWithListsFontChange()
        {
            Console.WriteLine("\tDocumentsWithListsFontChange()");

            // Create a document.
            using (DocX document = DocX.Create(@"docs\DocumentsWithListsFontChange.docx"))
            {
                foreach (var oneFontFamily in WindowsFontFamily.Families)
                {
                    var fontFamily = new Font(oneFontFamily.Name);
                    var fontSize = 15.0;

                    // created numbered lists 
                    var numberedList = document.AddList("First List Item.", 0, ListItemType.Numbered, 1);
                    document.AddListItem(numberedList, "First sub list item", 1);
                    document.AddListItem(numberedList, "Second List Item.");
                    document.AddListItem(numberedList, "Third list item.");
                    document.AddListItem(numberedList, "Nested item.", 1);
                    document.AddListItem(numberedList, "Second nested item.", 1);

                    // created bulleted lists

                    var bulletedList = document.AddList("First Bulleted Item.", 0, ListItemType.Bulleted);
                    document.AddListItem(bulletedList, "Second bullet item");
                    document.AddListItem(bulletedList, "Sub bullet item", 1);
                    document.AddListItem(bulletedList, "Second sub bullet item", 1);
                    document.AddListItem(bulletedList, "Third bullet item");

                    document.InsertList(bulletedList);
                    document.InsertList(numberedList, fontFamily, fontSize);
                }

                // Save this document.
                document.Save();

                Console.WriteLine("\tCreated: docs\\DocumentsWithListsFontChange.docx\n");
            }
        }

        private static void AddList()
        {
            Console.WriteLine("\tAddList()");

            using (var document = DocX.Create(@"docs\Lists.docx"))
            {
                var numberedList = document.AddList("First List Item.", 0, ListItemType.Numbered);
                //Add a numbered list starting at 2
                document.AddListItem(numberedList, "Second List Item.");
                document.AddListItem(numberedList, "Third list item.");
                document.AddListItem(numberedList, "First sub list item", 1);

                document.AddListItem(numberedList, "Nested item.", 2);
                document.AddListItem(numberedList, "Fourth nested item.");

                var bulletedList = document.AddList("First Bulleted Item.", 0, ListItemType.Bulleted);
                document.AddListItem(bulletedList, "Second bullet item");
                document.AddListItem(bulletedList, "Sub bullet item", 1);
                document.AddListItem(bulletedList, "Second sub bullet item", 2);
                document.AddListItem(bulletedList, "Third bullet item");

                document.InsertList(numberedList);
                document.InsertList(bulletedList);
                document.Save();
                Console.WriteLine("\tCreated: docs\\Lists.docx");
            }
        }

        private static void HeadersAndFooters()
        {
            Console.WriteLine("\tHeadersAndFooters()");

            // Create a new document.
            using (DocX document = DocX.Create(@"docs\HeadersAndFooters.docx"))
            {
                // Add Headers and Footers to this document.
                document.AddHeaders();
                document.AddFooters();

                // Force the first page to have a different Header and Footer.
                document.DifferentFirstPage = true;

                // Force odd & even pages to have different Headers and Footers.
                document.DifferentOddAndEvenPages = true;

                // Get the first, odd and even Headers for this document.
                Header header_first = document.Headers.first;
                Header header_odd = document.Headers.odd;
                Header header_even = document.Headers.even;

                // Get the first, odd and even Footer for this document.
                Footer footer_first = document.Footers.first;
                Footer footer_odd = document.Footers.odd;
                Footer footer_even = document.Footers.even;

                // Insert a Paragraph into the first Header.
                Paragraph p0 = header_first.InsertParagraph();
                p0.Append("Hello First Header.").Bold();

                // Insert a Paragraph into the odd Header.
                Paragraph p1 = header_odd.InsertParagraph();
                p1.Append("Hello Odd Header.").Bold();

                // Insert a Paragraph into the even Header.
                Paragraph p2 = header_even.InsertParagraph();
                p2.Append("Hello Even Header.").Bold();

                // Insert a Paragraph into the first Footer.
                Paragraph p3 = footer_first.InsertParagraph();
                p3.Append("Hello First Footer.").Bold();

                // Insert a Paragraph into the odd Footer.
                Paragraph p4 = footer_odd.InsertParagraph();
                p4.Append("Hello Odd Footer.").Bold();

                // Insert a Paragraph into the even Header.
                Paragraph p5 = footer_even.InsertParagraph();
                p5.Append("Hello Even Footer.").Bold();

                // Insert a Paragraph into the document.
                Paragraph p6 = document.InsertParagraph();
                p6.AppendLine("Hello First page.");

                // Create a second page to show that the first page has its own header and footer.
                p6.InsertPageBreakAfterSelf();

                // Insert a Paragraph after the page break.
                Paragraph p7 = document.InsertParagraph();
                p7.AppendLine("Hello Second page.");

                // Create a third page to show that even and odd pages have different headers and footers.
                p7.InsertPageBreakAfterSelf();

                // Insert a Paragraph after the page break.
                Paragraph p8 = document.InsertParagraph();
                p8.AppendLine("Hello Third page.");

                //Insert a next page break, which is a section break combined with a page break
                document.InsertSectionPageBreak();

                //Insert a paragraph after the "Next" page break
                Paragraph p9 = document.InsertParagraph();
                p9.Append("Next page section break.");

                //Insert a continuous section break
                document.InsertSection();

                //Create a paragraph in the new section
                var p10 = document.InsertParagraph();
                p10.Append("Continuous section paragraph.");

                // Save all changes to this document.
                document.Save();

                Console.WriteLine("\tCreated: docs\\HeadersAndFooters.docx\n");
            }// Release this document from memory.
        }
        private static void HeadersAndFootersWithImagesAndTables()
        {
            Console.WriteLine("\tHeadersAndFootersWithImagesAndTables()");

            // Create a new document.
            using (DocX document = DocX.Create(@"docs\HeadersAndFootersWithImagesAndTables.docx"))
            {
                // Add a template logo image to this document.
                RelativeDirectory rd = new RelativeDirectory(); // prepares the files for testing
                rd.Up(2);
                Image logo = document.AddImage(rd.Path + @"\images\logo_the_happy_builder.png");

                // Add Headers and Footers to this document.
                document.AddHeaders();
                document.AddFooters();

                // Force the first page to have a different Header and Footer.
                document.DifferentFirstPage = true;

                // Force odd & even pages to have different Headers and Footers.
                document.DifferentOddAndEvenPages = true;

                // Get the first, odd and even Headers for this document.
                Header header_first = document.Headers.first;
                Header header_odd = document.Headers.odd;
                Header header_even = document.Headers.even;

                // Get the first, odd and even Footer for this document.
                Footer footer_first = document.Footers.first;
                Footer footer_odd = document.Footers.odd;
                Footer footer_even = document.Footers.even;

                // Insert a Paragraph into the first Header.
                Paragraph p0 = header_first.InsertParagraph();
                p0.Append("Hello First Header.").Bold();



                // Insert a Paragraph into the odd Header.
                Paragraph p1 = header_odd.InsertParagraph();
                p1.Append("Hello Odd Header.").Bold();


                // Insert a Paragraph into the even Header.
                Paragraph p2 = header_even.InsertParagraph();
                p2.Append("Hello Even Header.").Bold();

                // Insert a Paragraph into the first Footer.
                Paragraph p3 = footer_first.InsertParagraph();
                p3.Append("Hello First Footer.").Bold();

                // Insert a Paragraph into the odd Footer.
                Paragraph p4 = footer_odd.InsertParagraph();
                p4.Append("Hello Odd Footer.").Bold();

                // Insert a Paragraph into the even Header.
                Paragraph p5 = footer_even.InsertParagraph();
                p5.Append("Hello Even Footer.").Bold();

                // Insert a Paragraph into the document.
                Paragraph p6 = document.InsertParagraph();
                p6.AppendLine("Hello First page.");

                // Create a second page to show that the first page has its own header and footer.
                p6.InsertPageBreakAfterSelf();

                // Insert a Paragraph after the page break.
                Paragraph p7 = document.InsertParagraph();
                p7.AppendLine("Hello Second page.");

                // Create a third page to show that even and odd pages have different headers and footers.
                p7.InsertPageBreakAfterSelf();

                // Insert a Paragraph after the page break.
                Paragraph p8 = document.InsertParagraph();
                p8.AppendLine("Hello Third page.");

                //Insert a next page break, which is a section break combined with a page break
                document.InsertSectionPageBreak();

                //Insert a paragraph after the "Next" page break
                Paragraph p9 = document.InsertParagraph();
                p9.Append("Next page section break.");

                //Insert a continuous section break
                document.InsertSection();

                //Create a paragraph in the new section
                var p10 = document.InsertParagraph();
                p10.Append("Continuous section paragraph.");


                // Inserting logo into footer and header into Tables

                #region Company Logo in Header in Table
                // Insert Table into First Header - Create a new Table with 2 columns and 1 rows.
                Table header_first_table = header_first.InsertTable(1, 2);
                header_first_table.Design = TableDesign.TableGrid;
                header_first_table.AutoFit = AutoFit.Window;
                // Get the upper right Paragraph in the layout_table.
                Paragraph upperRightParagraph = header_first.Tables[0].Rows[0].Cells[1].Paragraphs[0];
                // Insert this template logo into the upper right Paragraph of Table.
                upperRightParagraph.AppendPicture(logo.CreatePicture());
                upperRightParagraph.Alignment = Alignment.right;

                // Get the upper left Paragraph in the layout_table.
                Paragraph upperLeftParagraphFirstTable = header_first.Tables[0].Rows[0].Cells[0].Paragraphs[0];
                upperLeftParagraphFirstTable.Append("Company Name - DocX Corporation");
                #endregion


                #region Company Logo in Header in Invisible Table
                // Insert Table into First Header - Create a new Table with 2 columns and 1 rows.
                Table header_second_table = header_odd.InsertTable(1, 2);
                header_second_table.Design = TableDesign.None;
                header_second_table.AutoFit = AutoFit.Window;
                // Get the upper right Paragraph in the layout_table.
                Paragraph upperRightParagraphSecondTable = header_second_table.Rows[0].Cells[1].Paragraphs[0];
                // Insert this template logo into the upper right Paragraph of Table.
                upperRightParagraphSecondTable.AppendPicture(logo.CreatePicture());
                upperRightParagraphSecondTable.Alignment = Alignment.right;

                // Get the upper left Paragraph in the layout_table.
                Paragraph upperLeftParagraphSecondTable = header_second_table.Rows[0].Cells[0].Paragraphs[0];
                upperLeftParagraphSecondTable.Append("Company Name - DocX Corporation");
                #endregion

                #region Company Logo in Footer in Table
                // Insert Table into First Header - Create a new Table with 2 columns and 1 rows.
                Table footer_first_table = footer_first.InsertTable(1, 2);
                footer_first_table.Design = TableDesign.TableGrid;
                footer_first_table.AutoFit = AutoFit.Window;
                // Get the upper right Paragraph in the layout_table.
                Paragraph upperRightParagraphFooterParagraph = footer_first.Tables[0].Rows[0].Cells[1].Paragraphs[0];
                // Insert this template logo into the upper right Paragraph of Table.
                upperRightParagraphFooterParagraph.AppendPicture(logo.CreatePicture());
                upperRightParagraphFooterParagraph.Alignment = Alignment.right;

                // Get the upper left Paragraph in the layout_table.
                Paragraph upperLeftParagraphFirstTableFooter = footer_first.Tables[0].Rows[0].Cells[0].Paragraphs[0];
                upperLeftParagraphFirstTableFooter.Append("Company Name - DocX Corporation");
                #endregion



                #region Company Logo in Header in Invisible Table
                // Insert Table into First Header - Create a new Table with 2 columns and 1 rows.
                Table footer_second_table = footer_odd.InsertTable(1, 2);
                footer_second_table.Design = TableDesign.None;
                footer_second_table.AutoFit = AutoFit.Window;
                // Get the upper right Paragraph in the layout_table.
                Paragraph upperRightParagraphSecondTableFooter = footer_second_table.Rows[0].Cells[1].Paragraphs[0];
                // Insert this template logo into the upper right Paragraph of Table.
                upperRightParagraphSecondTableFooter.AppendPicture(logo.CreatePicture());
                upperRightParagraphSecondTableFooter.Alignment = Alignment.right;

                // Get the upper left Paragraph in the layout_table.
                Paragraph upperLeftParagraphSecondTableFooter = footer_second_table.Rows[0].Cells[0].Paragraphs[0];
                upperLeftParagraphSecondTableFooter.Append("Company Name - DocX Corporation");
                #endregion

                // Save all changes to this document.
                document.Save();

                Console.WriteLine("\tCreated: docs\\HeadersAndFootersWithImagesAndTables.docx\n");
            }// Release this document from memory.
        }
        private static void HeadersAndFootersWithImagesAndTablesUsingInsertPicture()
        {
            Console.WriteLine("\tHeadersAndFootersWithImagesAndTablesUsingInsertPicture()");

            // Create a new document.
            using (DocX document = DocX.Create(@"docs\HeadersAndFootersWithImagesAndTablesUsingInsertPicture.docx"))
            {
                // Add a template logo image to this document.
                RelativeDirectory rd = new RelativeDirectory(); // prepares the files for testing
                rd.Up(2);
                Image logo = document.AddImage(rd.Path + @"\images\logo_the_happy_builder.png");

                // Add Headers and Footers to this document.
                document.AddHeaders();
                document.AddFooters();

                // Force the first page to have a different Header and Footer.
                document.DifferentFirstPage = true;

                // Force odd & even pages to have different Headers and Footers.
                document.DifferentOddAndEvenPages = true;

                // Get the first, odd and even Headers for this document.
                Header header_first = document.Headers.first;
                Header header_odd = document.Headers.odd;
                Header header_even = document.Headers.even;

                // Get the first, odd and even Footer for this document.
                Footer footer_first = document.Footers.first;
                Footer footer_odd = document.Footers.odd;
                Footer footer_even = document.Footers.even;

                // Insert a Paragraph into the first Header.
                Paragraph p0 = header_first.InsertParagraph();
                p0.Append("Hello First Header.").Bold();



                // Insert a Paragraph into the odd Header.
                Paragraph p1 = header_odd.InsertParagraph();
                p1.Append("Hello Odd Header.").Bold();


                // Insert a Paragraph into the even Header.
                Paragraph p2 = header_even.InsertParagraph();
                p2.Append("Hello Even Header.").Bold();

                // Insert a Paragraph into the first Footer.
                Paragraph p3 = footer_first.InsertParagraph();
                p3.Append("Hello First Footer.").Bold();

                // Insert a Paragraph into the odd Footer.
                Paragraph p4 = footer_odd.InsertParagraph();
                p4.Append("Hello Odd Footer.").Bold();

                // Insert a Paragraph into the even Header.
                Paragraph p5 = footer_even.InsertParagraph();
                p5.Append("Hello Even Footer.").Bold();

                // Insert a Paragraph into the document.
                Paragraph p6 = document.InsertParagraph();
                p6.AppendLine("Hello First page.");

                // Create a second page to show that the first page has its own header and footer.
                p6.InsertPageBreakAfterSelf();

                // Insert a Paragraph after the page break.
                Paragraph p7 = document.InsertParagraph();
                p7.AppendLine("Hello Second page.");

                // Create a third page to show that even and odd pages have different headers and footers.
                p7.InsertPageBreakAfterSelf();

                // Insert a Paragraph after the page break.
                Paragraph p8 = document.InsertParagraph();
                p8.AppendLine("Hello Third page.");

                //Insert a next page break, which is a section break combined with a page break
                document.InsertSectionPageBreak();

                //Insert a paragraph after the "Next" page break
                Paragraph p9 = document.InsertParagraph();
                p9.Append("Next page section break.");

                //Insert a continuous section break
                document.InsertSection();

                //Create a paragraph in the new section
                var p10 = document.InsertParagraph();
                p10.Append("Continuous section paragraph.");


                // Inserting logo into footer and header into Tables

                #region Company Logo in Header in Table
                // Insert Table into First Header - Create a new Table with 2 columns and 1 rows.
                Table header_first_table = header_first.InsertTable(1, 2);
                header_first_table.Design = TableDesign.TableGrid;
                header_first_table.AutoFit = AutoFit.Window;
                // Get the upper right Paragraph in the layout_table.
                Paragraph upperRightParagraph = header_first.Tables[0].Rows[0].Cells[1].Paragraphs[0];
                // Insert this template logo into the upper right Paragraph of Table.
                upperRightParagraph.InsertPicture(logo.CreatePicture());
                upperRightParagraph.Alignment = Alignment.right;

                // Get the upper left Paragraph in the layout_table.
                Paragraph upperLeftParagraphFirstTable = header_first.Tables[0].Rows[0].Cells[0].Paragraphs[0];
                upperLeftParagraphFirstTable.Append("Company Name - DocX Corporation");
                #endregion


                #region Company Logo in Header in Invisible Table
                // Insert Table into First Header - Create a new Table with 2 columns and 1 rows.
                Table header_second_table = header_odd.InsertTable(1, 2);
                header_second_table.Design = TableDesign.None;
                header_second_table.AutoFit = AutoFit.Window;
                // Get the upper right Paragraph in the layout_table.
                Paragraph upperRightParagraphSecondTable = header_second_table.Rows[0].Cells[1].Paragraphs[0];
                // Insert this template logo into the upper right Paragraph of Table.
                upperRightParagraphSecondTable.InsertPicture(logo.CreatePicture());
                upperRightParagraphSecondTable.Alignment = Alignment.right;

                // Get the upper left Paragraph in the layout_table.
                Paragraph upperLeftParagraphSecondTable = header_second_table.Rows[0].Cells[0].Paragraphs[0];
                upperLeftParagraphSecondTable.Append("Company Name - DocX Corporation");
                #endregion

                #region Company Logo in Footer in Table
                // Insert Table into First Header - Create a new Table with 2 columns and 1 rows.
                Table footer_first_table = footer_first.InsertTable(1, 2);
                footer_first_table.Design = TableDesign.TableGrid;
                footer_first_table.AutoFit = AutoFit.Window;
                // Get the upper right Paragraph in the layout_table.
                Paragraph upperRightParagraphFooterParagraph = footer_first.Tables[0].Rows[0].Cells[1].Paragraphs[0];
                // Insert this template logo into the upper right Paragraph of Table.
                upperRightParagraphFooterParagraph.InsertPicture(logo.CreatePicture());
                upperRightParagraphFooterParagraph.Alignment = Alignment.right;

                // Get the upper left Paragraph in the layout_table.
                Paragraph upperLeftParagraphFirstTableFooter = footer_first.Tables[0].Rows[0].Cells[0].Paragraphs[0];
                upperLeftParagraphFirstTableFooter.Append("Company Name - DocX Corporation");
                #endregion



                #region Company Logo in Header in Invisible Table
                // Insert Table into First Header - Create a new Table with 2 columns and 1 rows.
                Table footer_second_table = footer_odd.InsertTable(1, 2);
                footer_second_table.Design = TableDesign.None;
                footer_second_table.AutoFit = AutoFit.Window;
                // Get the upper right Paragraph in the layout_table.
                Paragraph upperRightParagraphSecondTableFooter = footer_second_table.Rows[0].Cells[1].Paragraphs[0];
                // Insert this template logo into the upper right Paragraph of Table.
                upperRightParagraphSecondTableFooter.InsertPicture(logo.CreatePicture());
                upperRightParagraphSecondTableFooter.Alignment = Alignment.right;

                // Get the upper left Paragraph in the layout_table.
                Paragraph upperLeftParagraphSecondTableFooter = footer_second_table.Rows[0].Cells[0].Paragraphs[0];
                upperLeftParagraphSecondTableFooter.Append("Company Name - DocX Corporation");
                #endregion

                // Save all changes to this document.
                document.Save();

                Console.WriteLine("\tCreated: docs\\HeadersAndFootersWithImagesAndTablesUsingInsertPicture.docx\n");
            }// Release this document from memory.
        }

        private static void CreateInvoice()
        {
            Console.WriteLine("\tCreateInvoice()");
            DocX g_document;

            try
            {
                // Store a global reference to the loaded document.
                g_document = DocX.Load(@"docs\InvoiceTemplate.docx");

                /*
                 * The template 'InvoiceTemplate.docx' does exist, 
                 * so lets use it to create an invoice for a factitious company
                 * called "The Happy Builder" and store a global reference it.
                 */
                g_document = CreateInvoiceFromTemplate(DocX.Load(@"docs\InvoiceTemplate.docx"));

                // Save all changes made to this template as Invoice_The_Happy_Builder.docx (We don't want to replace InvoiceTemplate.docx).
                g_document.SaveAs(@"docs\Invoice_The_Happy_Builder.docx");
                Console.WriteLine("\tCreated: docs\\Invoice_The_Happy_Builder.docx\n");
            }

            // The template 'InvoiceTemplate.docx' does not exist, so create it.
            catch (FileNotFoundException)
            {
                // Create and store a global reference to the template 'InvoiceTemplate.docx'.
                g_document = CreateInvoiceTemplate();

                // Save the template 'InvoiceTemplate.docx'.
                g_document.Save();
                Console.WriteLine("\tCreated: docs\\InvoiceTemplate.docx");

                // The template exists now so re-call CreateInvoice().
                CreateInvoice();
            }
        }
        private static void CreateTableWithTextDirection()
        {
            Console.WriteLine("\tCreateTableWithTextDirection()");

            // Create a document.
            using (DocX document = DocX.Create(@"docs\\CeateTableWithTextDirection.docx"))
            {
                // Add a Table to this document.
                Table t = document.AddTable(2, 3);
                // Specify some properties for this Table.
                t.Alignment = Alignment.center;
                t.Design = TableDesign.MediumGrid1Accent2;
                // Add content to this Table.
                t.Rows[0].Cells[0].Paragraphs.First().Append("A");
                t.Rows[0].Cells[0].TextDirection = TextDirection.btLr;
                t.Rows[0].Cells[1].Paragraphs.First().Append("B");
                t.Rows[0].Cells[1].TextDirection = TextDirection.btLr;
                t.Rows[0].Cells[2].Paragraphs.First().Append("C");
                t.Rows[0].Cells[2].TextDirection = TextDirection.btLr;
                t.Rows[1].Cells[0].Paragraphs.First().Append("D");
                t.Rows[1].Cells[1].Paragraphs.First().Append("E");
                t.Rows[1].Cells[2].Paragraphs.First().Append("F");
                // Insert the Table into the document.
                document.InsertTable(t);
                document.Save();
            }// Release this document from memory.
            Console.WriteLine("\tCreated: docs\\CreateTableWithTextDirection.docx");
        }

        // Create an invoice for a factitious company called "The Happy Builder".
        private static DocX CreateInvoiceFromTemplate(DocX template)
        {
            #region Logo
            // A quick glance at the template shows us that the logo Paragraph is in row zero cell 1.
            Paragraph logo_paragraph = template.Tables[0].Rows[0].Cells[1].Paragraphs[0];
            // Remove the template Picture that is in this Paragraph.
            logo_paragraph.Pictures[0].Remove();

            // Add the Happy Builders logo to this document.
            RelativeDirectory rd = new RelativeDirectory(); // prepares the files for testing
            rd.Up(2);
            Image logo = template.AddImage(rd.Path + @"\images\logo_the_happy_builder.png");

            // Insert the Happy Builders logo into this Paragraph.
            logo_paragraph.InsertPicture(logo.CreatePicture());
            #endregion

            #region Set CustomProperty values
            // Set the value of the custom property 'company_name'.
            template.AddCustomProperty(new CustomProperty("company_name", "The Happy Builder"));

            // Set the value of the custom property 'company_slogan'.
            template.AddCustomProperty(new CustomProperty("company_slogan", "No job too small"));

            // Set the value of the custom properties 'hired_company_address_line_one', 'hired_company_address_line_two' and 'hired_company_address_line_three'.
            template.AddCustomProperty(new CustomProperty("hired_company_address_line_one", "The Crooked House,"));
            template.AddCustomProperty(new CustomProperty("hired_company_address_line_two", "Dublin,"));
            template.AddCustomProperty(new CustomProperty("hired_company_address_line_three", "12345"));

            // Set the value of the custom property 'invoice_date'.
            template.AddCustomProperty(new CustomProperty("invoice_date", DateTime.Today.Date.ToString("d")));

            // Set the value of the custom property 'invoice_number'.
            template.AddCustomProperty(new CustomProperty("invoice_number", 1));

            // Set the value of the custom property 'hired_company_details_line_one' and 'hired_company_details_line_two'.
            template.AddCustomProperty(new CustomProperty("hired_company_details_line_one", "Business Street, Dublin, 12345"));
            template.AddCustomProperty(new CustomProperty("hired_company_details_line_two", "Phone: 012-345-6789, Fax: 012-345-6789, e-mail: support@thehappybuilder.com"));
            #endregion

            /* 
             * InvoiceTemplate.docx contains a blank Table, 
             * we want to replace this with a new Table that
             * contains all of our invoice data.
             */
            Table t = template.Tables[1];
            Table invoice_table = CreateAndInsertInvoiceTableAfter(t, ref template);
            t.Remove();

            // Return the template now that it has been modified to hold all of our custom data.
            return template;
        }

        // Create an invoice template.
        private static DocX CreateInvoiceTemplate()
        {
            // Create a new document.
            DocX document = DocX.Create(@"docs\InvoiceTemplate.docx");

            // Create a table for layout purposes (This table will be invisible).
            Table layout_table = document.InsertTable(2, 2);
            layout_table.Design = TableDesign.TableNormal;
            layout_table.AutoFit = AutoFit.Window;

            // Dark formatting
            Formatting dark_formatting = new Formatting();
            dark_formatting.Bold = true;
            dark_formatting.Size = 12;
            dark_formatting.FontColor = WindowsColor.FromArgb(31, 73, 125);

            // Light formatting
            Formatting light_formatting = new Formatting();
            light_formatting.Italic = true;
            light_formatting.Size = 11;
            light_formatting.FontColor = WindowsColor.FromArgb(79, 129, 189);

            #region Company Name
            // Get the upper left Paragraph in the layout_table.
            Paragraph upper_left_paragraph = layout_table.Rows[0].Cells[0].Paragraphs[0];

            // Create a custom property called company_name
            CustomProperty company_name = new CustomProperty("company_name", "Company Name");

            // Insert a field of type doc property (This will display the custom property 'company_name')
            layout_table.Rows[0].Cells[0].Paragraphs[0].InsertDocProperty(company_name, f: dark_formatting);

            // Force the next text insert to be on a new line.
            upper_left_paragraph.InsertText("\n", false);
            #endregion

            #region Company Slogan
            // Create a custom property called company_slogan
            CustomProperty company_slogan = new CustomProperty("company_slogan", "Company slogan goes here.");

            // Insert a field of type doc property (This will display the custom property 'company_slogan')
            upper_left_paragraph.InsertDocProperty(company_slogan, f: light_formatting);
            #endregion

            #region Company Logo
            // Get the upper right Paragraph in the layout_table.
            Paragraph upper_right_paragraph = layout_table.Rows[0].Cells[1].Paragraphs[0];

            // Add a template logo image to this document.
            RelativeDirectory rd = new RelativeDirectory(); // prepares the files for testing
            rd.Up(2);
            Image logo = document.AddImage(rd.Path + @"\images\logo_template.png");

            // Insert this template logo into the upper right Paragraph.
            upper_right_paragraph.InsertPicture(logo.CreatePicture());

            upper_right_paragraph.Alignment = Alignment.right;
            #endregion

            // Custom properties cannot contain newlines, so the company address must be split into 3 custom properties.
            #region Hired Company Address
            // Create a custom property called company_address_line_one
            CustomProperty hired_company_address_line_one = new CustomProperty("hired_company_address_line_one", "Street Address,");

            // Get the lower left Paragraph in the layout_table. 
            Paragraph lower_left_paragraph = layout_table.Rows[1].Cells[0].Paragraphs[0];
            lower_left_paragraph.InsertText("TO:\n", false, dark_formatting);

            // Insert a field of type doc property (This will display the custom property 'hired_company_address_line_one')
            lower_left_paragraph.InsertDocProperty(hired_company_address_line_one, f: light_formatting);

            // Force the next text insert to be on a new line.
            lower_left_paragraph.InsertText("\n", false);

            // Create a custom property called company_address_line_two
            CustomProperty hired_company_address_line_two = new CustomProperty("hired_company_address_line_two", "City,");

            // Insert a field of type doc property (This will display the custom property 'hired_company_address_line_two')
            lower_left_paragraph.InsertDocProperty(hired_company_address_line_two, f: light_formatting);

            // Force the next text insert to be on a new line.
            lower_left_paragraph.InsertText("\n", false);

            // Create a custom property called company_address_line_two
            CustomProperty hired_company_address_line_three = new CustomProperty("hired_company_address_line_three", "Zip Code");

            // Insert a field of type doc property (This will display the custom property 'hired_company_address_line_three')
            lower_left_paragraph.InsertDocProperty(hired_company_address_line_three, f: light_formatting);
            #endregion

            #region Date & Invoice number
            // Get the lower right Paragraph from the layout table.
            Paragraph lower_right_paragraph = layout_table.Rows[1].Cells[1].Paragraphs[0];

            CustomProperty invoice_date = new CustomProperty("invoice_date", DateTime.Today.Date.ToString("d"));
            lower_right_paragraph.InsertText("Date: ", false, dark_formatting);
            lower_right_paragraph.InsertDocProperty(invoice_date, f: light_formatting);

            CustomProperty invoice_number = new CustomProperty("invoice_number", 1);
            lower_right_paragraph.InsertText("\nInvoice: ", false, dark_formatting);
            lower_right_paragraph.InsertText("#", false, light_formatting);
            lower_right_paragraph.InsertDocProperty(invoice_number, f: light_formatting);

            lower_right_paragraph.Alignment = Alignment.right;
            #endregion

            // Insert an empty Paragraph between two Tables, so that they do not touch.
            document.InsertParagraph(string.Empty, false);

            // This table will hold all of the invoice data.
            Table invoice_table = document.InsertTable(4, 4);
            invoice_table.Design = TableDesign.LightShadingAccent1;
            invoice_table.Alignment = Alignment.center;

            // A nice thank you Paragraph.
            Paragraph thankyou = document.InsertParagraph("\nThank you for your business, we hope to work with you again soon.", false, dark_formatting);
            thankyou.Alignment = Alignment.center;

            #region Hired company details
            CustomProperty hired_company_details_line_one = new CustomProperty("hired_company_details_line_one", "Street Address, City, ZIP Code");
            CustomProperty hired_company_details_line_two = new CustomProperty("hired_company_details_line_two", "Phone: 000-000-0000, Fax: 000-000-0000, e-mail: support@companyname.com");

            Paragraph companyDetails = document.InsertParagraph(string.Empty, false);
            companyDetails.InsertDocProperty(hired_company_details_line_one, f: light_formatting);
            companyDetails.InsertText("\n", false);
            companyDetails.InsertDocProperty(hired_company_details_line_two, f: light_formatting);
            companyDetails.Alignment = Alignment.center;
            #endregion

            // Return the document now that it has been created.
            return document;
        }

        private static Table CreateAndInsertInvoiceTableAfter(Table t, ref DocX document)
        {
            // Grab data from somewhere (Most likely a database)
            DataTable data = GetDataFromDatabase();

            /* 
             * The trick to replacing one Table with another,
             * is to insert the new Table after the old one, 
             * and then remove the old one.
             */
            Table invoice_table = t.InsertTableAfterSelf(data.Rows.Count + 1, data.Columns.Count);
            invoice_table.Design = TableDesign.LightShadingAccent1;

            #region Table title
            Formatting table_title = new Formatting();
            table_title.Bold = true;

            invoice_table.Rows[0].Cells[0].Paragraphs[0].InsertText("Description", false, table_title);
            invoice_table.Rows[0].Cells[0].Paragraphs[0].Alignment = Alignment.center;
            invoice_table.Rows[0].Cells[1].Paragraphs[0].InsertText("Hours", false, table_title);
            invoice_table.Rows[0].Cells[1].Paragraphs[0].Alignment = Alignment.center;
            invoice_table.Rows[0].Cells[2].Paragraphs[0].InsertText("Rate", false, table_title);
            invoice_table.Rows[0].Cells[2].Paragraphs[0].Alignment = Alignment.center;
            invoice_table.Rows[0].Cells[3].Paragraphs[0].InsertText("Amount", false, table_title);
            invoice_table.Rows[0].Cells[3].Paragraphs[0].Alignment = Alignment.center;
            #endregion

            // Loop through the rows in the Table and insert data from the data source.
            for (int row = 1; row < invoice_table.RowCount; row++)
            {
                for (int cell = 0; cell < invoice_table.Rows[row].Cells.Count; cell++)
                {
                    Paragraph cell_paragraph = invoice_table.Rows[row].Cells[cell].Paragraphs[0];
                    cell_paragraph.InsertText(data.Rows[row - 1].ItemArray[cell].ToString(), false);
                }
            }

            // We want to fill in the total by suming the values from the amount column.
            Row total = invoice_table.InsertRow();
            total.Cells[0].Paragraphs[0].InsertText("Total:", false);
            Paragraph total_paragraph = total.Cells[invoice_table.ColumnCount - 1].Paragraphs[0];

            /* 
             * Lots of people are scared of LINQ,
             * so I will walk you through this line by line.
             * 
             * invoice_table.Rows is an IEnumerable<Row> (i.e a collection of rows), with LINQ you can query collections.
             * .Where(condition) is a filter that you want to apply to the items of this collection. 
             * My condition is that the index of the row must be greater than 0 and less than RowCount.
             * .Select(something) lets you select something from each item in the filtered collection.
             * I am selecting the Text value from each row, for example €100, then I am remove the €, 
             * and then I am parsing the remaining string as a double. This will return a collection of doubles,
             * the final thing I do is call .Sum() on this collection which return one double the sum of all the doubles,
             * this is the total.
             */
            double totalCost =
            (
                invoice_table.Rows
                .Where((row, index) => index > 0 && index < invoice_table.RowCount - 1)
                .Select(row => double.Parse(row.Cells[row.Cells.Count() - 1].Paragraphs[0].Text.Remove(0, 1)))
            ).Sum();

            // Insert the total calculated above using LINQ into the total Paragraph.
            total_paragraph.InsertText(string.Format("€{0}", totalCost), false);

            // Let the tables columns expand to fit its contents.
            invoice_table.AutoFit = AutoFit.Contents;

            // Center the Table
            invoice_table.Alignment = Alignment.center;

            // Return the invloce table now that it has been created.
            return invoice_table;
        }

        // You need to rewrite this function to grab data from your data source.
        private static DataTable GetDataFromDatabase()
        {
            DataTable table = new DataTable();
            table.Columns.AddRange(new DataColumn[] { new DataColumn("Description"), new DataColumn("Hours"), new DataColumn("Rate"), new DataColumn("Amount") });

            table.Rows.Add
            (
                "Install wooden doors (Kitchen, Sitting room, Dining room & Bedrooms)",
                "5",
                "€25",
                string.Format("€{0}", 5 * 25)
            );

            table.Rows.Add
            (
                "Fit stairs",
                "20",
                "€30",
                string.Format("€{0}", 20 * 30)
            );

            table.Rows.Add
            (
                "Replace Sitting room window",
                "6",
                "€50",
                string.Format("€{0}", 6 * 50)
            );

            table.Rows.Add
            (
                "Build garden shed",
                "10",
                "€10",
                string.Format("€{0}", 10 * 10)
            );

            table.Rows.Add
             (
                 "Fit new lock on back door",
                 "0.5",
                 "€30",
                 string.Format("€{0}", 0.5 * 30)
             );

            table.Rows.Add
             (
                 "Tile Kitchen floor",
                 "24",
                 "€25",
                 string.Format("€{0}", 24 * 25)
             );

            return table;
        }

        /// <summary>
        /// Creates a simple document with the text Hello World.
        /// </summary>
        static void HelloWorld()
        {
            Console.WriteLine("\tHelloWorld()");

            // Create a new document.
            using (DocX document = DocX.Create(@"docs\HelloWorld.docx"))
            {
                // Insert a Paragraph into this document.
                Paragraph p = document.InsertParagraph();

                // Append some text and add formatting.
                p.Append("Hello World!^011Hello World!")
                .Font(new Font("Times New Roman"))
                .FontSize(32)
                .Color(WindowsColor.Blue)
                .Bold();



                // Save this document to disk.
                document.Save();
                Console.WriteLine("\tCreated: docs\\HelloWorld.docx\n");
            }
        }

        static void HelloWorldInsertHorizontalLine()
        {
            Console.WriteLine("\tHelloWorldInsertHorizontalLine()");

            // Create a new document.
            using (DocX document = DocX.Create(@"docs\HelloWorldInsertHorizontalLine.docx"))
            {
                // Insert a Paragraph into this document.
                Paragraph p = document.InsertParagraph();

                // Append some text and add formatting.
                p.Append("Hello World!^011Hello World!")
                .Font(new Font("Times New Roman"))
                .FontSize(32)
                .Color(WindowsColor.Blue)
                .Bold();
                p.InsertHorizontalLine("double", 6, 1, "auto");

                Paragraph p1 = document.InsertParagraph();
                p1.InsertHorizontalLine("double", 6, 1, "red");
                Paragraph p2 = document.InsertParagraph();
                p2.InsertHorizontalLine("single", 6, 1, "red");
                Paragraph p3 = document.InsertParagraph();
                p3.InsertHorizontalLine("triple", 6, 1, "blue");
                Paragraph p4 = document.InsertParagraph();
                p4.InsertHorizontalLine("double", 3, 10, "red");


                // Save this document to disk.
                document.Save();
                Console.WriteLine("\tCreated: docs\\HelloWorldInsertHorizontalLine.docx\n");
            }
        }

        static void HelloWorldProtectedDocument()
        {
            Console.WriteLine("\tHelloWorldPasswordProtected()");

            // Create a new document.
            using (DocX document = DocX.Create(@"docs\HelloWorldPasswordProtected.docx"))
            {
                // Insert a Paragraph into this document.
                Paragraph p = document.InsertParagraph();

                // Append some text and add formatting.
                p.Append("Hello World!^011Hello World!")
                .Font(new Font("Times New Roman"))
                .FontSize(32)
                .Color(WindowsColor.Blue)
                .Bold();


                // Save this document to disk with different options
                // Protected with password for Read Only
                EditRestrictions erReadOnly = EditRestrictions.readOnly;
                document.AddProtection(erReadOnly, "SomePassword");
                document.SaveAs(@"docs\\HelloWorldPasswordProtectedReadOnly.docx");
                Console.WriteLine("\tCreated: docs\\HelloWorldPasswordProtectedReadOnly.docx\n");

                // Protected with password for Comments
                EditRestrictions erComments = EditRestrictions.comments;
                document.AddProtection(erComments, "SomePassword");
                document.SaveAs(@"docs\\HelloWorldPasswordProtectedCommentsOnly.docx");
                Console.WriteLine("\tCreated: docs\\HelloWorldPasswordProtectedCommentsOnly.docx\n");

                // Protected with password for Forms
                EditRestrictions erForms = EditRestrictions.forms;
                document.AddProtection(erForms, "SomePassword");
                document.SaveAs(@"docs\\HelloWorldPasswordProtectedFormsOnly.docx");
                Console.WriteLine("\tCreated: docs\\HelloWorldPasswordProtectedFormsOnly.docx\n");

                // Protected with password for Tracked Changes
                EditRestrictions erTrackedChanges = EditRestrictions.trackedChanges;
                document.AddProtection(erTrackedChanges, "SomePassword");
                document.SaveAs(@"docs\\HelloWorldPasswordProtectedTrackedChangesOnly.docx");
                Console.WriteLine("\tCreated: docs\\HelloWorldPasswordProtectedTrackedChangesOnly.docx\n");

                // But it's also possible to add restrictions without protecting it with password.

                // Protected with password for Read Only
                document.AddProtection(erReadOnly);
                document.SaveAs(@"docs\\HelloWorldWithoutPasswordReadOnly.docx");
                Console.WriteLine("\tCreated: docs\\HelloWorldWithoutPasswordReadOnly.docx\n");

                // Protected with password for Comments
                document.AddProtection(erComments);
                document.SaveAs(@"docs\\HelloWorldWithoutPasswordCommentsOnly.docx");
                Console.WriteLine("\tCreated: docs\\HelloWorldWithoutPasswordCommentsOnly.docx\n");

                // Protected with password for Forms
                document.AddProtection(erForms);
                document.SaveAs(@"docs\\HelloWorldWithoutPasswordFormsOnly.docx");
                Console.WriteLine("\tCreated: docs\\HelloWorldWithoutPasswordFormsOnly.docx\n");

                // Protected with password for Tracked Changes
                document.AddProtection(erTrackedChanges);
                document.SaveAs(@"docs\\HelloWorldWithoutPasswordTrackedChangesOnly.docx");
                Console.WriteLine("\tCreated: docs\\HelloWorldWithoutPasswordTrackedChangesOnly.docx\n");
            }
        }

        static void HelloWorldAdvancedFormatting()
        {
            Console.WriteLine("\tHelloWorldAdvancedFormatting()");
            // Create a document.
            using (DocX document = DocX.Create(@"docs\HelloWorldAdvancedFormatting.docx"))
            {
                // Insert a new Paragraphs.
                Paragraph p = document.InsertParagraph();

                p.Append("I am ").Append("bold").Bold()
                .Append(" and I am ")
                .Append("italic").Italic().Append(".")
                .AppendLine("I am ")
                .Append("Arial Black")
                .Font(new Font("Arial Black"))
                .Append(" and I am not.")
                .AppendLine("I am ")
                .Append("BLUE").Color(WindowsColor.Blue)
                .Append(" and I am")
                .Append("Red").Color(WindowsColor.Red).Append(".");

                // Save this document.
                document.Save();
                Console.WriteLine("\tCreated: docs\\HelloWorldAdvancedFormatting.docx\n");
            }// Release this document from memory.
        }

        /// <summary>
        /// Loads a document 'Input.docx' and writes the text 'Hello World' into the first imbedded Image.
        /// This code creates the file 'Output.docx'.
        /// </summary>
        static void ProgrammaticallyManipulateImbeddedImage()
        {
            Console.WriteLine("\tProgrammaticallyManipulateImbeddedImage()");
            const string str = "Hello World";

            // Open the document Input.docx.
            using (DocX document = DocX.Load(@"docs\Input.docx"))
            {
                // Make sure this document has at least one Image.
                if (document.Images.Count() > 0)
                {
                    Image img = document.Images[0];

                    // Write "Hello World" into this Image.
                    var b = new WindowsBitmap(img.GetStream(FileMode.Open, FileAccess.ReadWrite));

                    /*
                     * Get the Graphics object for this Bitmap.
                     * The Graphics object provides functions for drawing.
                     */
                    var g = WindowsGraphics.FromImage(b);

                    // Draw the string "Hello World".
                    g.DrawString
                    (
                        str,
                        new WindowsFont("Tahoma", 20),
                        WindowsBrushes.Blue,
                        0.0f, 0.0f
                    );

                    // Save this Bitmap back into the document using a Create\Write stream.
                    b.Save(img.GetStream(FileMode.Create, FileAccess.Write), WindowsImageFormat.Png);
                }
                else
                    Console.WriteLine("The provided document contains no Images.");

                // Save this document as Output.docx.
                document.SaveAs(@"docs\Output.docx");
                Console.WriteLine("\tCreated: docs\\Output.docx\n");
            }
        }

        /// <summary>
        /// For each of the documents in the folder 'docs\',
        /// Replace the string a with the string b,
        /// Do this in Parrallel accross many CPU cores.
        /// </summary>
        static void ReplaceTextParallel()
        {
            Console.WriteLine("\tReplaceTextParallel()\n");
            const string a = "apple";
            const string b = "pear";

            // Directory containing many .docx documents.
            DirectoryInfo di = new DirectoryInfo(@"docs\");

            // Loop through each document in this specified direction.
            Parallel.ForEach
            (
                di.GetFiles(),
                currentFile =>
                {
                    // Load the document.
                    using (DocX document = DocX.Load(currentFile.FullName))
                    {
                        // Replace text in this document.
                        document.ReplaceText(a, b);

                        // Save changes made to this document.
                        document.Save();
                    } // Release this document from memory.
                }
            );
            Console.WriteLine("\tCreated: None\n");
        }

        static void AddToc()
        {
            Console.WriteLine("\tAddToc()");

            using (var document = DocX.Create(@"docs\Toc.docx"))
            {
                document.InsertTableOfContents("I can haz table of contentz", TableOfContentsSwitches.O | TableOfContentsSwitches.U | TableOfContentsSwitches.Z | TableOfContentsSwitches.H, "Heading2");
                var h1 = document.InsertParagraph("Heading 1");
                h1.StyleName = "Heading1";
                document.InsertParagraph("Some very interesting content here");
                var h2 = document.InsertParagraph("Heading 2");
                document.InsertSectionPageBreak();
                h2.StyleName = "Heading1";
                document.InsertParagraph("Some very interesting content here as well");
                var h3 = document.InsertParagraph("Heading 2.1");
                h3.StyleName = "Heading2";
                document.InsertParagraph("Not so very interesting....");

                document.Save();
            }
        }

        static void AddTocByReference()
        {
            Console.WriteLine("\tAddTocByReference()");

            using (var document = DocX.Create(@"docs\TocByReference.docx"))
            {
                var h1 = document.InsertParagraph("Heading 1");
                h1.StyleName = "Heading1";
                document.InsertParagraph("Some very interesting content here");
                var h2 = document.InsertParagraph("Heading 2");
                document.InsertSectionPageBreak();
                h2.StyleName = "Heading1";
                document.InsertParagraph("Some very interesting content here as well");
                var h3 = document.InsertParagraph("Heading 2.1");
                h3.StyleName = "Heading2";
                document.InsertParagraph("Not so very interesting....");

                document.InsertTableOfContents(h2, "I can haz table of contentz", TableOfContentsSwitches.O | TableOfContentsSwitches.U | TableOfContentsSwitches.Z | TableOfContentsSwitches.H, "Heading2");

                document.Save();
            }
        }

        static void HelloWorldKeepWithNext()
        {
            // Create a Paragraph that will stay on the same page as the paragraph that comes next
            Console.WriteLine("\tHelloWorldKeepWithNext()");
            // Create a new document.
            using (DocX document = DocX.Create("docs\\HelloWorldKeepWithNext.docx"))

            {
                // Create a new Paragraph with the text "Hello World".
                Paragraph p = document.InsertParagraph("Hello World.");
                p.KeepWithNext();
                document.InsertParagraph("Previous paragraph will appear on the same page as this paragraph");

                // Save all changes made to this document.
                document.Save();
                Console.WriteLine("\tCreated: docs\\HelloWorldKeepWithNext.docx\n");
            }
        }
        static void HelloWorldKeepLinesTogether()
        {
            // Create a Paragraph that will stay on the same page as the paragraph that comes next
            Console.WriteLine("\tHelloWorldKeepLinesTogether()");
            // Create a new document.
            using (DocX document = DocX.Create("docs\\HelloWorldKeepLinesTogether.docx"))
            {
                // Create a new Paragraph with the text "Hello World".
                Paragraph p = document.InsertParagraph("All lines of this paragraph will appear on the same page...\nLine 2\nLine 3\nLine 4\nLine 5\nLine 6...");
                p.KeepLinesTogether();
                // Save all changes made to this document.
                document.Save();
                Console.WriteLine("\tCreated: docs\\HelloWorldKeepLinesTogether.docx\n");
            }
        }

        static void LargeTable()
        {
            Console.WriteLine("\tLargeTable()");
            var _directoryWithFiles = "docs\\";
            using (var output = File.Open(_directoryWithFiles + "LargeTable.docx", FileMode.Create))
            {
                using (var doc = DocX.Create(output))
                {
                    var tbl = doc.InsertTable(1, 18);

                    var wholeWidth = doc.PageWidth - doc.MarginLeft - doc.MarginRight;
                    var colWidth = wholeWidth / tbl.ColumnCount;
                    var colWidths = new int[tbl.ColumnCount];
                    tbl.AutoFit = AutoFit.Contents;
                    var r = tbl.Rows[0];
                    var cx = 0;
                    foreach (var cell in r.Cells)
                    {
                        cell.Paragraphs.First().Append("Col " + cx);
                        //cell.Width = colWidth;
                        cell.MarginBottom = 0;
                        cell.MarginLeft = 0;
                        cell.MarginRight = 0;
                        cell.MarginTop = 0;

                        cx++;
                    }
                    tbl.SetBorder(TableBorderType.Bottom, BlankBorder);
                    tbl.SetBorder(TableBorderType.Left, BlankBorder);
                    tbl.SetBorder(TableBorderType.Right, BlankBorder);
                    tbl.SetBorder(TableBorderType.Top, BlankBorder);
                    tbl.SetBorder(TableBorderType.InsideV, BlankBorder);
                    tbl.SetBorder(TableBorderType.InsideH, BlankBorder);

                    doc.Save();
                }
            }
            Console.WriteLine("\tCreated: docs\\LargeTable.docx\n");
        }


        static void TableWithSpecifiedWidths()
        {
            Console.WriteLine("\tTableSpecifiedWidths()");
            var _directoryWithFiles = "docs\\";
            using (var output = File.Open(_directoryWithFiles + "TableSpecifiedWidths.docx", FileMode.Create))
            {
                using (var doc = DocX.Create(output))
                {
                    var widths = new float[] { 200f, 100f, 300f };
                    var tbl = doc.InsertTable(1, widths.Length);
                    tbl.SetWidths(widths);
                    var wholeWidth = doc.PageWidth - doc.MarginLeft - doc.MarginRight;
                    tbl.AutoFit = AutoFit.Contents;
                    var r = tbl.Rows[0];
                    var cx = 0;
                    foreach (var cell in r.Cells)
                    {
                        cell.Paragraphs.First().Append("Col " + cx);
                        //cell.Width = colWidth;
                        cell.MarginBottom = 0;
                        cell.MarginLeft = 0;
                        cell.MarginRight = 0;
                        cell.MarginTop = 0;

                        cx++;
                    }
                    //add new rows 
                    for (var x = 0; x < 5; x++)
                    {
                        r = tbl.InsertRow();
                        cx = 0;
                        foreach (var cell in r.Cells)
                        {
                            cell.Paragraphs.First().Append("Col " + cx);
                            //cell.Width = colWidth;
                            cell.MarginBottom = 0;
                            cell.MarginLeft = 0;
                            cell.MarginRight = 0;
                            cell.MarginTop = 0;

                            cx++;
                        }
                    }
                    tbl.SetBorder(TableBorderType.Bottom, BlankBorder);
                    tbl.SetBorder(TableBorderType.Left, BlankBorder);
                    tbl.SetBorder(TableBorderType.Right, BlankBorder);
                    tbl.SetBorder(TableBorderType.Top, BlankBorder);
                    tbl.SetBorder(TableBorderType.InsideV, BlankBorder);
                    tbl.SetBorder(TableBorderType.InsideH, BlankBorder);

                    doc.Save();
                }
            }
            Console.WriteLine("\tCreated: docs\\TableSpecifiedWidths.docx\n");

        }

        /// <summary>
        /// Create a document with two pictures. One picture is inserted normal way, the other one with rotation
        /// </summary>
        static void HelloWorldAddPictureToWord()
        {
            Console.WriteLine("\tHelloWorldAddPictureToWord()");

            // Create a document.
            using (DocX document = DocX.Create(@"docs\HelloWorldAddPictureToWord.docx"))
            {
                // Add an image into the document.    
                RelativeDirectory rd = new RelativeDirectory(); // prepares the files for testing
                rd.Up(2);
                Image image = document.AddImage(rd.Path + @"\images\logo_template.png");

                // Create a picture (A custom view of an Image).
                Picture picture = image.CreatePicture();
                picture.Rotation = 10;
                picture.SetPictureShape(BasicShapes.cube);

                // Insert a new Paragraph into the document.
                Paragraph title = document.InsertParagraph().Append("This is a test for a picture").FontSize(20).Font(new Font("Comic Sans MS"));
                title.Alignment = Alignment.center;

                // Insert a new Paragraph into the document.
                Paragraph p1 = document.InsertParagraph();

                // Append content to the Paragraph
                p1.AppendLine("Just below there should be a picture ").Append("picture").Bold().Append(" inserted in a non-conventional way.");
                p1.AppendLine();
                p1.AppendLine("Check out this picture ").AppendPicture(picture).Append(" its funky don't you think?");
                p1.AppendLine();

                // Insert a new Paragraph into the document.
                Paragraph p2 = document.InsertParagraph();
                // Append content to the Paragraph.

                p2.AppendLine("Is it correct?");
                p2.AppendLine();

                // Lets add another picture (without the fancy stuff)
                Picture pictureNormal = image.CreatePicture();

                Paragraph p3 = document.InsertParagraph();
                p3.AppendLine("Lets add another picture (without the fancy  rotation stuff)");
                p3.AppendLine();
                p3.AppendPicture(pictureNormal);

                // Save this document.
                document.Save();
                Console.WriteLine("\tCreated: docs\\HelloWorldAddPictureToWord.docx\n");
            }
        }


    }
}
