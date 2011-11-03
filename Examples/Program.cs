using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Novacode;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;
using System.Threading.Tasks;
using System.Data;

namespace Examples
{
    class Program
    {
        static void Main(string[] args)
        {
            // Easy
            Console.WriteLine("\nRunning Easy Examples");
            HelloWorld();
            RightToLeft();
            Indentation();

            HeadersAndFooters();
            HyperlinksImagesTables();

            Equations();

            // Intermediate
            Console.WriteLine("\nRunning Intermediate Examples");
            CreateInvoice();

            // Advanced
            Console.WriteLine("\nRunning Advanced Examples");
            ProgrammaticallyManipulateImbeddedImage();
            ReplaceTextParallel();

            Console.WriteLine("\nPress any key to exit.");
            Console.ReadKey();
        }

        /// <summary>
        /// Create a document with two equations.
        /// </summary>
        private static void Equations()
        {
            Console.WriteLine("\nEquations()");

            // Create a new document.
            using (DocX document = DocX.Create(@"docs\Equations.docx"))
            {
                // Insert first Equation in this document.
                Paragraph pEquation1 = document.InsertEquation("x = y+z");

                // Insert second Equation in this document and add formatting.
                Paragraph pEquation2 = document.InsertEquation("x = (y+z)/t").FontSize(18).Color(Color.Blue);                

                // Save this document to disk.
                document.Save();
                Console.WriteLine("\tCreated: docs\\Equations.docx\n");
            }
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
                table.Rows[0].Cells[0].Paragraphs[0].Append("3");
                table.Rows[0].Cells[1].Paragraphs[0].Append("1");
                table.Rows[1].Cells[0].Paragraphs[0].Append("4");
                table.Rows[1].Cells[1].Paragraphs[0].Append("1");

                // Add an image into the document.    
                Novacode.Image image = document.AddImage(@"images\logo_template.png");

                // Create a picture (A custom view of an Image).
                Picture picture = image.CreatePicture();
                picture.Rotation = 10;
                picture.SetPictureShape(BasicShapes.cube);

                // Insert a new Paragraph into the document.
                Paragraph title = document.InsertParagraph().Append("Test").FontSize(20).Font(new FontFamily("Comic Sans MS"));
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

                // Save all changes to this document.
                document.Save();

                Console.WriteLine("\tCreated: docs\\HeadersAndFooters.docx\n");
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

        // Create an invoice for a factitious company called "The Happy Builder".
        private static DocX CreateInvoiceFromTemplate(DocX template)
        {
            #region Logo
            // A quick glance at the template shows us that the logo Paragraph is in row zero cell 1.
            Paragraph logo_paragraph = template.Tables[0].Rows[0].Cells[1].Paragraphs[0];
            // Remove the template Picture that is in this Paragraph.
            logo_paragraph.Pictures[0].Remove();

            // Add the Happy Builders logo to this document.
            Novacode.Image logo = template.AddImage(@"images\logo_the_happy_builder.png");

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
            dark_formatting.FontColor = Color.FromArgb(31, 73, 125);

            // Light formatting
            Formatting light_formatting = new Formatting();
            light_formatting.Italic = true;
            light_formatting.Size = 11;
            light_formatting.FontColor = Color.FromArgb(79, 129, 189);

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
            Novacode.Image logo = document.AddImage(@"images\logo_template.png");

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

            // We want to fill in the total by suming the values from the amount coloumn.
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

            // Let the tables coloumns expand to fit its contents.
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
            using (DocX document = DocX.Create(@"docs\Hello World.docx"))
            {
                // Insert a Paragraph into this document.
                Paragraph p = document.InsertParagraph();

                // Append some text and add formatting.
                p.Append("Hello World")
                .Font(new FontFamily("Times New Roman"))
                .FontSize(32)
                .Color(Color.Blue)
                .Bold();

                // Save this document to disk.
                document.Save();
                Console.WriteLine("\tCreated: docs\\Hello World.docx\n");
            }
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
            using (DocX document = DocX.Load(@"Input.docx"))
            {
                // Make sure this document has at least one Image.
                if (document.Images.Count() > 0)
                {
                    Novacode.Image img = document.Images[0];

                    // Write "Hello World" into this Image.
                    Bitmap b = new Bitmap(img.GetStream(FileMode.Open, FileAccess.ReadWrite));

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
                    b.Save(img.GetStream(FileMode.Create, FileAccess.Write), ImageFormat.Png);
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
    }
}
