using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml;
using System.IO;
using System.Text.RegularExpressions;
using System.IO.Packaging;

namespace Novacode
{
    /// <summary>
    /// Represents a document.
    /// </summary>
    public class DocX: IDisposable
    {
        #region Namespaces
        static internal XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        static internal XNamespace customPropertiesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        static internal XNamespace customVTypesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        #endregion

        #region Private instance variables defined foreach DocX object
        // The collection of Paragraphs in this document.
        private List<Paragraph> paragraphs = new List<Paragraph>();
        // A dictionary of CustomProperties in this document.
        private Dictionary<string, CustomProperty> customProperties;
        // A list of Images in this document.
        private List<Image> images;
        // A collection of Tables in this Paragraph
        private List<Table> tables;
        #endregion

        #region Internal variables defined foreach DocX object
        // Object representation of the .docx
        internal Package package;
        // The mainDocument is loaded into a XDocument object for easy querying and editing
        internal XDocument mainDoc; 
        // A lookup for the Paragraphs in this document.
        internal Dictionary<int, Paragraph> paragraphLookup = new Dictionary<int, Paragraph>();
        // Every document is stored in a MemoryStream, all edits made to a document are done in memory.
        internal MemoryStream memoryStream;
        // The filename that this document was loaded from
        internal string filename;
        // The stream that this document was loaded from
        internal Stream stream;
        #endregion

        internal DocX()
        {      
        }

        /// <summary>
        /// Returns a list of Paragraphs in this document.
        /// </summary>
        /// <example>
        /// Write to Console the Text from each Paragraph in this document.
        /// <code>
        /// // Load a document
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        ///
        /// // Loop through each Paragraph in this document.
        /// foreach (Paragraph p in document.Paragraphs)
        /// {
        ///     // Write this Paragraphs Text to Console.
        ///     Console.WriteLine(p.Text);
        /// }
        ///
        /// // Wait for the user to press a key before closing the console window.
        /// Console.ReadKey();
        /// </code>
        /// </example>
        /// <seealso cref="Paragraph.InsertText(string, bool)"/>
        /// <seealso cref="Paragraph.InsertText(int, string, bool)"/>
        /// <seealso cref="Paragraph.InsertText(string, bool, Formatting)"/>
        /// <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
        /// <seealso cref="Paragraph.RemoveText(int, bool)"/>
        /// <seealso cref="Paragraph.RemoveText(int, int, bool)"/>
        /// <seealso cref="Paragraph.ReplaceText(string, string, bool)"/>
        /// <seealso cref="Paragraph.ReplaceText(string, string, bool, RegexOptions)"/>
        /// <seealso cref="Paragraph.InsertPicture"/>
        public List<Paragraph> Paragraphs
        {
            get { return paragraphs; }
        }


        /// <summary>
        /// Returns a list of Tables in this Paragraph.
        /// </summary>
        public List<Table> Tables 
        { 
            get { return tables; } 
        }

        /// <summary>
        /// Returns a list of Images in this document.
        /// </summary>
        /// <example>
        /// Get the unique Id of every Image in this document.
        /// <code>
        /// // Load a document.
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        ///
        /// // Loop through each Image in this document.
        /// foreach (Novacode.Image i in document.Images)
        /// {
        ///     // Get the unique Id which identifies this Image.
        ///     string uniqueId = i.Id;
        /// }
        ///
        /// </code>
        /// </example>
        /// <seealso cref="AddImage(string)"/>
        /// <seealso cref="AddImage(Stream)"/>
        /// <seealso cref="Paragraph.Pictures"/>
        /// <seealso cref="Paragraph.InsertPicture"/>
        public List<Image> Images
        {
            get { return images; }
        }

        /// <summary>
        /// Returns a list of custom properties in this document.
        /// </summary>
        /// <example>
        /// Method 1: Get the name, type and value of each CustomProperty in this document.
        /// <code>
        /// // Load Example.docx
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        ///
        /// /*
        ///  * No two custom properties can have the same name,
        ///  * so a Dictionary is the perfect data structure to store them in.
        ///  * Each custom property can be accessed using its name.
        ///  */
        /// foreach (string name in document.CustomProperties.Keys)
        /// {
        ///     // Grab a custom property using its name.
        ///     CustomProperty cp = document.CustomProperties[name];
        ///
        ///     // Write this custom properties details to Console.
        ///     Console.WriteLine(string.Format("Name: '{0}', Value: {1}", cp.Name, cp.Value));
        /// }
        ///
        /// Console.WriteLine("Press any key...");
        ///
        /// // Wait for the user to press a key before closing the Console.
        /// Console.ReadKey();
        /// </code>
        /// </example>
        /// <example>
        /// Method 2: Get the name, type and value of each CustomProperty in this document.
        /// <code>
        /// // Load Example.docx
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        /// 
        /// /*
        ///  * No two custom properties can have the same name,
        ///  * so a Dictionary is the perfect data structure to store them in.
        ///  * The values of this Dictionary are CustomProperties.
        ///  */
        /// foreach (CustomProperty cp in document.CustomProperties.Values)
        /// {
        ///     // Write this custom properties details to Console.
        ///     Console.WriteLine(string.Format("Name: '{0}', Value: {1}", cp.Name, cp.Value));
        /// }
        ///
        /// Console.WriteLine("Press any key...");
        ///
        /// // Wait for the user to press a key before closing the Console.
        /// Console.ReadKey();
        /// </code>
        /// </example>
        /// <seealso cref="AddCustomProperty"/>
        public Dictionary<string, CustomProperty> CustomProperties
        {
            get { return customProperties; }
        }

        static internal void RebuildParagraphs(DocX document)
        {
            document.paragraphLookup.Clear();
            document.paragraphs.Clear();

            // Get the runs in this paragraph
            IEnumerable<XElement> paras = document.mainDoc.Descendants(XName.Get("p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"));

            int startIndex = 0;

            // Loop through each run in this paragraph
            foreach (XElement par in paras)
            {
                Paragraph xp = new Paragraph(document, startIndex, par);

                // Add to paragraph list
                document.paragraphs.Add(xp);

                // Only add runs which contain text
                if (Paragraph.GetElementTextLength(par) > 0)
                {
                    document.paragraphLookup.Add(xp.endIndex, xp);
                    startIndex = xp.endIndex;
                }
            }
        }

        /// <summary>
        /// Insert a new Paragraph at the end of this document.
        /// </summary>
        /// <param name="text">The text of this Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <returns>A new Paragraph.</returns>
        /// <example>
        /// Inserting a new Paragraph at the end of a document with text formatting.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Insert a new Paragraph at the end of this document.
        ///     document.InsertParagraph("New text", false);
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory
        /// </code>
        /// </example>
        public Paragraph InsertParagraph(string text, bool trackChanges)
        {
            int index = 0;
            if (paragraphLookup.Keys.Count() > 0)
                index = paragraphLookup.Last().Key;

            return InsertParagraph(index, text, trackChanges, null);
        }

        /// <summary>
        /// Insert a new Paragraph at the end of a document with text formatting.
        /// </summary>
        /// <param name="text">The text of this Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <param name="formatting">The formatting for the text of this Paragraph.</param>
        /// <returns>A new Paragraph.</returns>
        /// <example>
        /// Inserting a new Paragraph at the end of a document with text formatting.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Create a Formatting object
        ///     Formatting formatting = new Formatting();
        ///     formatting.Bold = true;
        ///     formatting.FontColor = Color.Red;
        ///     formatting.Size = 30;
        ///
        ///     // Insert a new Paragraph at the end of this document with text formatting.
        ///     document.InsertParagraph("New text", false, formatting);
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory
        /// </code>
        /// </example>
        public Paragraph InsertParagraph(string text, bool trackChanges, Formatting formatting)
        {
            int index = 0;
            if (paragraphLookup.Keys.Count() > 0)
                index = paragraphLookup.Last().Key;

            return InsertParagraph(index, text, trackChanges, formatting);
        }

        /// <summary>
        /// Get the Text of this document.
        /// </summary>
        /// <example>
        /// Write to Console the Text from this document.
        /// <code>
        /// // Load a document
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        ///
        /// // Get the text of this document.
        /// string text = document.Text;
        ///
        /// // Write the text of this document to Console.
        /// Console.Write(text);
        ///
        /// // Wait for the user to press a key before closing the console window.
        /// Console.ReadKey();
        /// </code>
        /// </example>
        public string Text
        {
            get
            {
                string text = string.Empty;
                foreach (Paragraph p in paragraphs)
                {
                    text += p.Text + "\n";
                }
                return text;
            }
        }
        /// <summary>
        /// Insert a new Paragraph into this document at a specified index.
        /// </summary>
        /// <param name="index">The character index to insert this document at.</param>
        /// <param name="text">The text of this Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <returns>A new Paragraph.</returns>
        /// <example>
        /// Insert a new Paragraph into the middle of a document.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Find the middle character index of this document.
        ///     int index = document.Text.Length / 2;
        ///
        ///     // Insert a new Paragraph at the middle of this document.
        ///     document.InsertParagraph(index, "New text", false);
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        ///}// Release this document from memory
        /// </code>
        /// </example>
        public Paragraph InsertParagraph(int index, string text, bool trackChanges)
        {
            return InsertParagraph(index, text, trackChanges, null);
        }
        
        /// <summary>
        /// Insert a Paragraph into this document, this Paragraph may have come from the same or another document.
        /// </summary>
        /// <param name="index">The index to insert this Paragragraph at.</param>
        /// <param name="p">The Paragraph to insert.</param>
        /// <returns>The Paragraph now associated with this document.</returns>
        /// <example>
        /// Take a Paragraph from document a, and insert it into document b at a specified position.
        /// <code>
        /// // Place holder for a Paragraph.
        /// Paragraph p;
        ///
        /// // Load document a.
        /// using (DocX documentA = DocX.Load(@"C:\Example\a.docx"))
        /// {
        ///     // Get the first paragraph from this document.
        ///     p = documentA.Paragraphs[0];
        /// }
        ///
        /// // Load document b.
        /// using (DocX documentB = DocX.Load(@"C:\Example\b.docx"))
        /// {
        ///     /* 
        ///      * Insert the Paragraph that was extracted from document a, into docment b. 
        ///      * This creates a new Paragraph that is now associated with document b.
        ///      */ 
        ///      Paragraph newParagraph = documentB.InsertParagraph(0, p);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Paragraph InsertParagraph(int index, Paragraph p)
        {
            XElement newXElement = new XElement(p.xml);
            p.xml = newXElement;

            Paragraph paragraph = GetFirstParagraphEffectedByInsert(this, index);
            
            if (paragraph == null)
                mainDoc.Descendants(XName.Get("body", DocX.w.NamespaceName)).First().Add(p.xml);
            else
            {
                XElement[] split = SplitParagraph(paragraph, index - paragraph.startIndex);

                paragraph.xml.ReplaceWith
                (
                    split[0],
                    newXElement,
                    split[1]
                );
            }

            RebuildParagraphs(this);
            return p;
        }

        /// <summary>
        /// Insert a Paragraph into this document, this Paragraph may have come from the same or another document.
        /// </summary>
        /// <param name="p">The Paragraph to insert.</param>
        /// <returns>The Paragraph now associated with this document.</returns>
        /// <example>
        /// Take a Paragraph from document a, and insert it into the end of document b.
        /// <code>
        /// // Place holder for a Paragraph.
        /// Paragraph p;
        ///
        /// // Load document a.
        /// using (DocX documentA = DocX.Load(@"C:\Example\a.docx"))
        /// {
        ///     // Get the first paragraph from this document.
        ///     p = documentA.Paragraphs[0];
        /// }
        ///
        /// // Load document b.
        /// using (DocX documentB = DocX.Load(@"C:\Example\b.docx"))
        /// {
        ///     /* 
        ///      * Insert the Paragraph that was extracted from document a, into docment b. 
        ///      * This creates a new Paragraph that is now associated with document b.
        ///      */ 
        ///      Paragraph newParagraph = documentB.InsertParagraph(p);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code> 
        /// </example>
        public Paragraph InsertParagraph(Paragraph p)
        {
            XElement newXElement = new XElement(p.xml);

            mainDoc.Descendants(XName.Get("body", DocX.w.NamespaceName)).First().Add(newXElement);
            int index = 0;
            if (paragraphLookup.Keys.Count() > 0)
                index = paragraphLookup.Last().Key + paragraphLookup.Last().Value.Text.Length;

            Paragraph newParagraph = new Paragraph(this, index, newXElement);
            paragraphLookup.Add(index, newParagraph);
            return newParagraph;
        }

        /// <summary>
        /// Insert a new Table at the end of this document.
        /// </summary>
        /// <param name="coloumnCount">The number of coloumns to create.</param>
        /// <param name="rowCount">The number of rows to create.</param>
        /// <returns>A new Table.</returns>
        /// <example>
        /// Insert a new Table with 2 coloumns and 3 rows, at the end of a document.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"C:\Example\Test.docx"))
        /// {
        ///     // Create a new Table with 2 coloumns and 3 rows.
        ///     Table newTable = document.InsertTable(2, 3);
        ///
        ///     // Set the design of this Table.
        ///     newTable.Design = TableDesign.LightShadingAccent2;
        ///
        ///     // Set the coloumn names.
        ///     newTable.Rows[0].Cells[0].Paragraph.InsertText("Ice Cream", false);
        ///     newTable.Rows[0].Cells[1].Paragraph.InsertText("Price", false);
        ///
        ///     // Fill row 1
        ///     newTable.Rows[1].Cells[0].Paragraph.InsertText("Chocolate", false);
        ///     newTable.Rows[1].Cells[1].Paragraph.InsertText("€3:50", false);
        ///
        ///     // Fill row 2
        ///     newTable.Rows[2].Cells[0].Paragraph.InsertText("Vanilla", false);
        ///     newTable.Rows[2].Cells[1].Paragraph.InsertText("€3:00", false);
        ///
        ///     // Save all changes made to document b.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Table InsertTable(int coloumnCount, int rowCount)
        {
            XElement newTable = CreateTable(rowCount, coloumnCount);
            mainDoc.Descendants(XName.Get("body", DocX.w.NamespaceName)).First().Add(newTable);

            RebuildTables();
            RebuildParagraphs(this);
            return new Table(this, newTable);
        }

        internal static XElement CreateTable(int rowCount, int coloumnCount)
        {
            XElement newTable =
            new XElement
            (
                XName.Get("tbl", DocX.w.NamespaceName),
                new XElement
                (
                    XName.Get("tblPr", DocX.w.NamespaceName),
                        new XElement(XName.Get("tblStyle", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "TableGrid")),
                        new XElement(XName.Get("tblW", DocX.w.NamespaceName), new XAttribute(XName.Get("w", DocX.w.NamespaceName), "0"), new XAttribute(XName.Get("type", DocX.w.NamespaceName), "auto")),
                        new XElement(XName.Get("tblLook", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "04A0"))
                )
            );

            XElement tableGrid = new XElement(XName.Get("tblGrid", DocX.w.NamespaceName));
            for (int i = 0; i < coloumnCount; i++)
                tableGrid.Add(new XElement(XName.Get("gridCol", DocX.w.NamespaceName), new XAttribute(XName.Get("w", DocX.w.NamespaceName), "2310")));

            newTable.Add(tableGrid);

            for (int i = 0; i < rowCount; i++)
            {
                XElement row = new XElement(XName.Get("tr", DocX.w.NamespaceName));

                for (int j = 0; j < coloumnCount; j++)
                {
                    XElement cell =
                    new XElement
                    (
                        XName.Get("tc", DocX.w.NamespaceName),
                            new XElement(XName.Get("tcPr", DocX.w.NamespaceName),
                                new XElement(XName.Get("tcW", DocX.w.NamespaceName), new XAttribute(XName.Get("w", DocX.w.NamespaceName), "2310"), new XAttribute(XName.Get("type", DocX.w.NamespaceName), "dxa"))),
                            new XElement(XName.Get("p", DocX.w.NamespaceName))
                    );

                    row.Add(cell);
                }

                newTable.Add(row);
            }
            return newTable;
        }

        /// <summary>
        /// Insert a Table into this document. The Table's source can be a completely different document.
        /// </summary>
        /// <param name="t">The Table to insert.</param>
        /// <param name="index">The index to insert this Table at.</param>
        /// <returns>The Table now associated with this document.</returns>
        /// <example>
        /// Extract a Table from document a and insert it into document b, at index 10.
        /// <code>
        /// // Place holder for a Table.
        /// Table t;
        ///
        /// // Load document a.
        /// using (DocX documentA = DocX.Load(@"C:\Example\a.docx"))
        /// {
        ///     // Get the first Table from this document.
        ///     t = documentA.Tables[0];
        /// }
        ///
        /// // Load document b.
        /// using (DocX documentB = DocX.Load(@"C:\Example\b.docx"))
        /// {
        ///     /* 
        ///      * Insert the Table that was extracted from document a, into document b. 
        ///      * This creates a new Table that is now associated with document b.
        ///      */
        ///     Table newTable = documentB.InsertTable(10, t);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Table InsertTable(int index, Table t)
        {
            Paragraph p = GetFirstParagraphEffectedByInsert(this, index);

            XElement[] split = SplitParagraph(p, index - p.startIndex);
            XElement newXElement = new XElement(t.xml);
            p.xml.ReplaceWith
            (
                split[0],
                newXElement,
                split[1]
            );

            Table newTable = new Table(this, newXElement);
            newTable.Design = t.Design;

            RebuildTables();
            RebuildParagraphs(this);
            return newTable;
        }

        /// <summary>
        /// Insert a Table into this document. The Table's source can be a completely different document.
        /// </summary>
        /// <param name="t">The Table to insert.</param>
        /// <returns>The Table now associated with this document.</returns>
        /// <example>
        /// Extract a Table from document a and insert it at the end of document b.
        /// <code>
        /// // Place holder for a Table.
        /// Table t;
        ///
        /// // Load document a.
        /// using (DocX documentA = DocX.Load(@"C:\Example\a.docx"))
        /// {
        ///     // Get the first Table from this document.
        ///     t = documentA.Tables[0];
        /// }
        ///
        /// // Load document b.
        /// using (DocX documentB = DocX.Load(@"C:\Example\b.docx"))
        /// {
        ///     /* 
        ///      * Insert the Table that was extracted from document a, into document b. 
        ///      * This creates a new Table that is now associated with document b.
        ///      */
        ///     Table newTable = documentB.InsertTable(t);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Table InsertTable(Table t)
        {
            XElement newXElement = new XElement(t.xml);
            mainDoc.Descendants(XName.Get("body", DocX.w.NamespaceName)).First().Add(newXElement);

            Table newTable = new Table(this, newXElement);
            newTable.Design = t.Design;

            tables.Add(newTable);
            return newTable;
        }

        /// <summary>
        /// Insert a new Table at the end of this document.
        /// </summary>
        /// <param name="coloumnCount">The number of coloumns to create.</param>
        /// <param name="rowCount">The number of rows to create.</param>
        /// <param name="index">The index to insert this Table at.</param>
        /// <returns>A new Table.</returns>
        /// <example>
        /// Insert a new Table with 2 coloumns and 3 rows, at index 37 in this document.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Create a new Table with 2 coloumns and 3 rows. Insert this Table at index 37.
        ///     Table newTable = document.InsertTable(37, 2, 3);
        ///
        ///     // Set the design of this Table.
        ///     newTable.Design = TableDesign.LightShadingAccent3;
        ///
        ///     // Set the coloumn names.
        ///     newTable.Rows[0].Cells[0].Paragraph.InsertText("Ice Cream", false);
        ///     newTable.Rows[0].Cells[1].Paragraph.InsertText("Price", false);
        ///
        ///     // Fill row 1
        ///     newTable.Rows[1].Cells[0].Paragraph.InsertText("Chocolate", false);
        ///     newTable.Rows[1].Cells[1].Paragraph.InsertText("€3:50", false);
        ///
        ///     // Fill row 2
        ///     newTable.Rows[2].Cells[0].Paragraph.InsertText("Vanilla", false);
        ///     newTable.Rows[2].Cells[1].Paragraph.InsertText("€3:00", false);
        ///
        ///     // Save all changes made to document b.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Table InsertTable(int index, int coloumnCount, int rowCount)
        {
            XElement newTable = CreateTable(rowCount, coloumnCount);

            Paragraph p = GetFirstParagraphEffectedByInsert(this, index);

            if (p == null)
                mainDoc.Descendants(XName.Get("body", DocX.w.NamespaceName)).First().AddFirst(newTable);

            else
            {
                XElement[] split = SplitParagraph(p, index - p.startIndex);

                p.xml.ReplaceWith
                (
                    split[0],
                    newTable,
                    split[1]
                );
            }

            RebuildParagraphs(this);
            RebuildTables();
            return new Table(this, newTable);
        }

        internal void RebuildTables()
        {
            tables =
            (
                from t in mainDoc.Descendants(XName.Get("tbl", DocX.w.NamespaceName))
                select new Table(this, t)
            ).ToList();
        }

        /// <summary>
        /// Insert a new Paragraph into this document at a specified index with text formatting.
        /// </summary>
        /// <param name="index">The character index to insert this document at.</param>
        /// <param name="text">The text of this Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <param name="formatting">The formatting for the text of this Paragraph.</param>
        /// <returns>A new Paragraph.</returns>
        /// /// <example>
        /// Insert a new Paragraph into the middle of a document with text formatting.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Create a Formatting object
        ///     Formatting formatting = new Formatting();
        ///     formatting.Bold = true;
        ///     formatting.FontColor = Color.Red;
        ///     formatting.Size = 30;
        ///
        ///     //  Middle character index of this document.
        ///     int index = document.Text.Length / 2;
        ///
        ///     // Insert a new Paragraph in the middle of this document.
        ///     document.InsertParagraph(index, "New text", false, formatting);
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory
        /// </code>
        /// <remarks>You must add a reference to System.Drawing in order to use Color.Red.</remarks>
        /// </example>
        public Paragraph InsertParagraph(int index, string text, bool trackChanges, Formatting formatting)
        {
            Paragraph newParagraph = new Paragraph(this, index, new XElement(w + "p"));
            newParagraph.InsertText(0, text, trackChanges, formatting);

            Paragraph firstPar = GetFirstParagraphEffectedByInsert(this, index);

            if (firstPar != null)
            {
                XElement[] splitParagraph = SplitParagraph(firstPar, index - firstPar.startIndex);

                firstPar.xml.ReplaceWith
                (
                    splitParagraph[0],
                    newParagraph.xml,
                    splitParagraph[1]
                );
            }

            else
                mainDoc.Descendants(XName.Get("body", DocX.w.NamespaceName)).First().Add(newParagraph.xml);

            DocX.RebuildParagraphs(this);
            return newParagraph;
        }

        static internal Paragraph GetFirstParagraphEffectedByInsert(DocX document, int index)
        {
            // This document contains no Paragraphs and insertion is at index 0
            if (document.paragraphLookup.Keys.Count() == 0 && index == 0)
                return null;

            foreach (int paragraphEndIndex in document.paragraphLookup.Keys)
            {
                if (paragraphEndIndex >= index)
                    return document.paragraphLookup[paragraphEndIndex];
            }

            throw new ArgumentOutOfRangeException();
        }

        internal XElement[] SplitParagraph(Paragraph p, int index)
        {
            Run r = p.GetFirstRunEffectedByInsert(index);

            XElement[] split;
            XElement before, after;

            if (r.xml.Parent.Name.LocalName == "ins")
            {
                split = p.SplitEdit(r.xml.Parent, index, EditType.ins);
                before = new XElement(p.xml.Name, p.xml.Attributes(), r.xml.Parent.ElementsBeforeSelf(), split[0]);
                after = new XElement(p.xml.Name, p.xml.Attributes(), r.xml.Parent.ElementsAfterSelf(), split[1]);
            }

            else if (r.xml.Parent.Name.LocalName == "del")
            {
                split = p.SplitEdit(r.xml.Parent, index, EditType.del);

                before = new XElement(p.xml.Name, p.xml.Attributes(), r.xml.Parent.ElementsBeforeSelf(), split[0]);
                after = new XElement(p.xml.Name, p.xml.Attributes(), r.xml.Parent.ElementsAfterSelf(), split[1]);
            }

            else
            {
                split = Run.SplitRun(r, index);

                before = new XElement(p.xml.Name, p.xml.Attributes(), r.xml.ElementsBeforeSelf(), split[0]);
                after = new XElement(p.xml.Name, p.xml.Attributes(), r.xml.ElementsAfterSelf(), split[1]);
            }

            if (before.Elements().Count() == 0)
                before = null;

            if (after.Elements().Count() == 0)
                after = null;

            return new XElement[] { before, after };
        }

        /// <summary>
        /// Creates a document using a Stream.
        /// </summary>
        /// <param name="stream">The Stream to create the document from.</param>
        /// <returns>Returns a DocX object which represents the document.</returns>
        /// <example>
        /// Creating a document from a FileStream.
        /// <code>
        /// // Use a FileStream fs to create a new document.
        /// using(FileStream fs = new FileStream(@"C:\Example\Test.docx", FileMode.Create))
        /// {
        ///     // Load the document using fs
        ///     using (DocX document = DocX.Create(fs))
        ///     {
        ///         // Do something with the document here.
        ///
        ///         // Save all changes made to this document.
        ///         document.Save();
        ///     }// Release this document from memory.
        /// }
        /// </code>
        /// </example>
        /// <example>
        /// Creating a document in a SharePoint site.
        /// <code>
        /// using(SPSite mySite = new SPSite("http://server/sites/site"))
        /// {
        ///     // Open a connection to the SharePoint site
        ///     using(SPWeb myWeb = mySite.OpenWeb())
        ///     {
        ///         // Create a MemoryStream ms.
        ///         using (MemoryStream ms = new MemoryStream())
        ///         {
        ///             // Create a document using ms.
        ///             using (DocX document = DocX.Create(ms))
        ///             {
        ///                 // Do something with the document here.
        ///
        ///                 // Save all changes made to this document.
        ///                 document.Save();
        ///             }// Release this document from memory
        ///
        ///             // Add the document to the SharePoint site
        ///             web.Files.Add("filename", ms.ToArray(), true);
        ///         }
        ///     }
        /// }
        /// </code>
        /// </example>
        /// <seealso cref="DocX.Create(string)"/>
        /// <seealso cref="DocX.Load(System.IO.Stream)"/>
        /// <seealso cref="DocX.Load(string)"/>
        /// <seealso cref="DocX.Save()"/>
        public static DocX Create(Stream stream)
        {
            // Store this document in memory
            MemoryStream ms = new MemoryStream();

            // Create the docx package
            Package package = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite);

            PostCreation(ref package);
            DocX document = DocX.Load(ms);
            document.stream = stream;
            return document;
        }

        /// <summary>
        /// Creates a document using a fully qualified or relative filename.
        /// </summary>
        /// <param name="filename">The fully qualified or relative filename.</param>
        /// <returns>Returns a DocX object which represents the document.</returns>
        /// <example>
        /// <code>
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Create(@"..\Test.docx"))
        /// {
        ///     // Do something with the document here.
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory
        /// </code>
        /// <code>
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Create(@"..\Test.docx"))
        /// {
        ///     // Do something with the document here.
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory
        /// </code>
        /// <seealso cref="DocX.Create(System.IO.Stream)"/>
        /// <seealso cref="DocX.Load(System.IO.Stream)"/>
        /// <seealso cref="DocX.Load(string)"/>
        /// <seealso cref="DocX.Save()"/>
        /// </example>
        public static DocX Create(string filename)
        {
            // Store this document in memory
            MemoryStream ms = new MemoryStream();

            // Create the docx package
            //WordprocessingDocument wdDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            Package package = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite);

            PostCreation(ref package);
            DocX document = DocX.Load(ms);
            document.filename = filename;
            return document;
        }

        private static void PostCreation(ref Package package)
        {
            XDocument mainDoc, stylesDoc;

            #region MainDocumentPart
            // Create the main document part for this package
            PackagePart mainDocumentPart = package.CreatePart(new Uri("/word/document.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
            package.CreateRelationship(mainDocumentPart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
            
            // Load the document part into a XDocument object
            using (TextReader tr = new StreamReader(mainDocumentPart.GetStream(FileMode.Create, FileAccess.ReadWrite)))
            {
                mainDoc = XDocument.Parse
                (@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                   <w:document xmlns:ve=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"">
                   <w:body>
                    <w:sectPr w:rsidR=""003E25F4"" w:rsidSect=""00FC3028"">
                        <w:pgSz w:w=""11906"" w:h=""16838""/>
                        <w:pgMar w:top=""1440"" w:right=""1440"" w:bottom=""1440"" w:left=""1440"" w:header=""708"" w:footer=""708"" w:gutter=""0""/>
                        <w:cols w:space=""708""/>
                        <w:docGrid w:linePitch=""360""/>
                    </w:sectPr>
                   </w:body>
                   </w:document>"
                );
            }

            // Save the main document
            using (TextWriter tw = new StreamWriter(mainDocumentPart.GetStream(FileMode.Create, FileAccess.Write)))
                mainDoc.Save(tw, SaveOptions.DisableFormatting);
            #endregion

            #region MainDocumentPart
            // Create the main document part for this package
            PackagePart word_styles = package.CreatePart(new Uri("/word/styles.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");
           
            stylesDoc =
            XDocument.Parse
            (
                @"<?xml version='1.0' encoding='utf-8' standalone='yes'?>
                  <w:styles xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                  </w:styles>"
            );

            // Save the main document
            using (TextWriter tw = new StreamWriter(word_styles.GetStream(FileMode.Create, FileAccess.Write)))
                stylesDoc.Save(tw, SaveOptions.DisableFormatting);

            mainDocumentPart.CreateRelationship(word_styles.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
            #endregion

            package.Close();
        }

        private static DocX PostLoad(ref Package package)
        {
            DocX document = new DocX();
            document.package = package;

            #region MainDocumentPart
            // Load the document part into a XDocument object
            PackagePart word_document = package.GetPart(new Uri("/word/document.xml", UriKind.Relative));
            using (TextReader tr = new StreamReader(word_document.GetStream(FileMode.Open, FileAccess.Read)))
                document.mainDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);

            RebuildParagraphs(document);

            document.images = new List<Image>();
            PackageRelationshipCollection imageRelationships = package.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
            if (imageRelationships.Count() > 0)
            {
                document.images =
                (
                    from i in imageRelationships
                    select new Image(document, i)
                ).ToList();
            }

            document.RebuildTables();
            #endregion

            #region CustomFilePropertiesPart
            ExtractCustomProperties(document);
            #endregion

            return document;
        }

        private static void ExtractCustomProperties(DocX document)
        {
            if(document.package.PartExists(new Uri("/docProps/custom.xml", UriKind.Relative)))
            {
                PackagePart docProps_custom = document.package.GetPart(new Uri("/docProps/custom.xml", UriKind.Relative));
                XDocument customPropDoc;
                using (TextReader tr = new StreamReader(docProps_custom.GetStream(FileMode.Open, FileAccess.Read)))
                    customPropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);

                // Get all of the custom properties in this document
                document.customProperties =
                (
                    from p in customPropDoc.Descendants(XName.Get("property", customPropertiesSchema.NamespaceName))
                    let Name = p.Attribute(XName.Get("name")).Value
                    let Type = p.Descendants().Single().Name.LocalName
                    let Value = p.Descendants().Single().Value
                    select new CustomProperty(Name, Type, Value)
                ).ToDictionary(p => p.Name, StringComparer.CurrentCultureIgnoreCase);
            }
        }

        /// <summary>
        /// Loads a document into a DocX object using a Stream.
        /// </summary>
        /// <param name="stream">The Stream to load the document from.</param>
        /// <returns>
        /// Returns a DocX object which represents the document.
        /// </returns>
        /// <example>
        /// Loading a document from a FileStream.
        /// <code>
        /// // Open a FileStream fs to a document.
        /// using (FileStream fs = new FileStream(@"C:\Example\Test.docx", FileMode.Open))
        /// {
        ///     // Load the document using fs.
        ///     using (DocX document = DocX.Load(fs))
        ///     {
        ///         // Do something with the document here.
        ///            
        ///         // Save all changes made to the document.
        ///         document.Save();
        ///     }// Release this document from memory.
        /// }
        /// </code>
        /// </example>
        /// <example>
        /// Loading a document from a SharePoint site.
        /// <code>
        /// // Get the SharePoint site that you want to access.
        /// using (SPSite mySite = new SPSite("http://server/sites/site"))
        /// {
        ///     // Open a connection to the SharePoint site
        ///     using (SPWeb myWeb = mySite.OpenWeb())
        ///     {
        ///         // Grab a document stored on this site.
        ///         SPFile file = web.GetFile("Source_Folder_Name/Source_File");
        ///
        ///         // DocX.Load requires a Stream, so open a Stream to this document.
        ///         Stream str = new MemoryStream(file.OpenBinary());
        ///
        ///         // Load the file using the Stream str.
        ///         using (DocX document = DocX.Load(str))
        ///         {
        ///             // Do something with the document here.
        ///
        ///             // Save all changes made to the document.
        ///             document.Save();
        ///         }// Release this document from memory.
        ///     }
        /// }
        /// </code>
        /// </example>
        /// <seealso cref="DocX.Load(string)"/>
        /// <seealso cref="DocX.Create(System.IO.Stream)"/>
        /// <seealso cref="DocX.Create(string)"/>
        /// <seealso cref="DocX.Save()"/>
        public static DocX Load(Stream stream)
        {
            MemoryStream ms = new MemoryStream();

            stream.Position = 0;
            byte[] data = new byte[stream.Length];
            stream.Read(data, 0, (int)stream.Length);
            ms.Write(data, 0, (int)stream.Length);

            // Open the docx package
            Package package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite);

            DocX document = PostLoad(ref package);
            document.package = package;
            document.memoryStream = ms;
            return document;
        }

        /// <summary>
        /// Loads a document into a DocX object using a fully qualified or relative filename.
        /// </summary>
        /// <param name="filename">The fully qualified or relative filename.</param>
        /// <returns>
        /// Returns a DocX object which represents the document.
        /// </returns>
        /// <example>
        /// <code>
        /// // Load a document using its fully qualified filename
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Do something with the document here
        ///
        ///     // Save all changes made to document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// <code>
        /// // Load a document using its relative filename.
        /// using(DocX document = DocX.Load(@"..\..\Test.docx"))
        /// { 
        ///     // Do something with the document here.
        ///                
        ///     // Save all changes made to document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// <seealso cref="DocX.Load(System.IO.Stream)"/>
        /// <seealso cref="DocX.Create(System.IO.Stream)"/>
        /// <seealso cref="DocX.Create(string)"/>
        /// <seealso cref="DocX.Save()"/>
        /// </example>
        public static DocX Load(string filename)
        {
            if (!File.Exists(filename))
                throw new FileNotFoundException(string.Format("File could not be found {0}", filename));

            MemoryStream ms = new MemoryStream();
            
            using (FileStream fs = new FileStream(filename, FileMode.Open))
            {
                byte[] data = new byte[fs.Length];
                fs.Read(data, 0, (int)fs.Length);
                ms.Write(data, 0, (int)fs.Length);
            }

            // Open the docx package
            Package package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite);

            DocX document = PostLoad(ref package);
            document.package = package;
            document.filename = filename;
            document.memoryStream = ms;

            return document;
        }

        /// <summary>
        /// Add an Image into this document from a fully qualified or relative filename.
        /// </summary>
        /// <param name="filename">The fully qualified or relative filename.</param>
        /// <returns>An Image file.</returns>
        /// <example>
        /// Add an Image into this document from a fully qualified filename.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Add an Image from a file.
        ///     document.AddImage(@"C:\Example\Image.png");
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <seealso cref="AddImage(System.IO.Stream)"/>
        /// <seealso cref="Paragraph.InsertPicture"/>
        public Image AddImage(string filename)
        {
            return AddImage(filename as object);
        }

        private bool IsSameFile(Stream streamOne, Stream streamTwo)
        {
            int file1byte, file2byte;

            if (streamOne.Length != streamOne.Length)
            {
                // Close the files
                streamOne.Close();
                streamTwo.Close();

                // Return false to indicate files are different
                return false;
            }

            // Read and compare a byte from each file until either a
            // non-matching set of bytes is found or until the end of
            // file1 is reached.
            do
            {
                // Read one byte from each file.
                file1byte = streamOne.ReadByte();
                file2byte = streamTwo.ReadByte();
            }
            while ((file1byte == file2byte) && (file1byte != -1));

            // Close the files.
            streamOne.Close();
            streamTwo.Close();

            // Return the success of the comparison. "file1byte" is 
            // equal to "file2byte" at this point only if the files are 
            // the same.
            return ((file1byte - file2byte) == 0);
        }

        /// <summary>
        /// Add an Image into this document from a Stream.
        /// </summary>
        /// <param name="stream">A Stream stream.</param>
        /// <returns>An Image file.</returns>
        /// <example>
        /// Add an Image into a document using a Stream. 
        /// <code>
        /// // Open a FileStream fs to an Image.
        /// using (FileStream fs = new FileStream(@"C:\Example\Image.jpg", FileMode.Open))
        /// {
        ///     // Load a document.
        ///     using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        ///     {
        ///         // Add an Image from a filestream fs.
        ///         document.AddImage(fs);
        ///
        ///         // Save all changes made to this document.
        ///         document.Save();
        ///     }// Release this document from memory.
        /// }
        /// </code>
        /// </example>
        /// <seealso cref="AddImage(string)"/>
        /// <seealso cref="Paragraph.InsertPicture"/>
        public Image AddImage(Stream stream)
        {
            return AddImage(stream as object);
        }

        internal Image AddImage(object o)
        {
            PackagePart word_document = package.GetPart(new Uri("/word/document.xml", UriKind.Relative));
            var imageParts = word_document.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image").Select(ir => package.GetPart(ir.TargetUri));

            foreach (PackagePart pp in imageParts)
            {
                Stream s;
                if (o is string)
                    s = new FileStream(o as string, FileMode.Open);
                else
                    s = o as Stream;

                if (IsSameFile(pp.GetStream(), s))
                {
                    string id = word_document.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
                    .Where(r => r.TargetUri == pp.Uri)
                    .Select(r => r.Id).First();

                    return images.Where(i => i.Id == id).First();
                }
            }

            int max = 0;
            var values =
            (
                from ip in imageParts
                let Name = Path.GetFileNameWithoutExtension(ip.Uri.ToString())
                let Number = Regex.Match(Name, @"\d+$").Value
                select Number != string.Empty ? int.Parse(Number) : 0
            );
            if (values.Count() > 0)
                max = Math.Max(max, values.Max());

            PackagePart img = package.CreatePart(new Uri(string.Format("/word/media/image{0}.jpeg", max + 1), UriKind.Relative), System.Net.Mime.MediaTypeNames.Image.Jpeg);
            PackageRelationship rel = word_document.CreateRelationship(img.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");

            using (Stream stream = img.GetStream(FileMode.Create, FileAccess.Write))
            {
                Stream s;
                if (o is string)
                    s = new FileStream(o as string, FileMode.Open);
                else
                    s = o as Stream;

                using (s)
                {
                    byte[] bytes = new byte[s.Length];
                    s.Read(bytes, 0, (int)s.Length);
                    stream.Write(bytes, 0, (int)s.Length);
                }
            }
            
            Image newImg = new Image(this, rel);
            images.Add(newImg);
            return newImg;
        }

        /// <summary>
        /// Save this document back to the location it was loaded from.
        /// </summary>
        /// <example>
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Add an Image from a file.
        ///     document.AddImage(@"C:\Example\Image.jpg");
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <seealso cref="DocX.SaveAs(string)"/>
        /// <seealso cref="DocX.Create(System.IO.Stream)"/>
        /// <seealso cref="DocX.Create(string)"/>
        /// <seealso cref="DocX.Load(System.IO.Stream)"/>
        /// <seealso cref="DocX.Load(string)"/>
        public void Save()
        {
            if (package.PartExists(new Uri("/word/document.xml", UriKind.Relative)))
            {
                // Save the main document
                using (TextWriter tw = new StreamWriter(package.GetPart(new Uri("/word/document.xml", UriKind.Relative)).GetStream(FileMode.Create, FileAccess.Write)))
                    mainDoc.Save(tw, SaveOptions.DisableFormatting);
            }

            // Close the document so that it can be saved.
            Dispose();

            #region Save this document back to a file or stream, that was specified by the user at save time.
            if (filename != null)
            {
                using (FileStream fs = new FileStream(filename, FileMode.Create))
                    fs.Write(memoryStream.ToArray(), 0, (int)memoryStream.Length);
            }

            else
            {
                // Set the length of this stream to 0
                stream.SetLength(0);

                // Write to the beginning of the stream
                stream.Position = 0;

                memoryStream.WriteTo(stream);
                stream.Close();
            }
            #endregion

            // Re-open the document
            package = Package.Open(memoryStream, FileMode.Open, FileAccess.ReadWrite);
        }

        /// <summary>
        /// Save this document to a file.
        /// </summary>
        /// <param name="filename">The filename to save this document as.</param>
        /// <example>
        /// Load a document from one file and save it to another.
        /// <code>
        /// // Load a document using its fully qualified filename.
        /// DocX document = DocX.Load(@"C:\Example\Test1.docx");
        ///
        /// // Insert a new Paragraph
        /// document.InsertParagraph("Hello world!", false);
        ///
        /// // Save the document to a new location.
        /// document.SaveAs(@"C:\Example\Test2.docx");
        /// </code>
        /// </example>
        /// <example>
        /// Load a document from a Stream and save it to a file.
        /// <code>
        /// DocX document;
        /// using (FileStream fs1 = new FileStream(@"C:\Example\Test1.docx", FileMode.Open))
        /// {
        ///     // Load a document using a stream.
        ///     document = DocX.Load(fs1);
        ///
        ///     // Insert a new Paragraph
        ///     document.InsertParagraph("Hello world again!", false);
        /// }
        ///    
        /// // Save the document to a new location.
        /// document.SaveAs(@"C:\Example\Test2.docx");
        /// </code>
        /// </example>
        /// <seealso cref="DocX.Save()"/>
        /// <seealso cref="DocX.Create(System.IO.Stream)"/>
        /// <seealso cref="DocX.Create(string)"/>
        /// <seealso cref="DocX.Load(System.IO.Stream)"/>
        /// <seealso cref="DocX.Load(string)"/>
        public void SaveAs(string filename)
        {
            this.filename = filename;
            this.stream = null;
            Save();
        }

        /// <summary>
        /// Save this document to a Stream.
        /// </summary>
        /// <param name="stream">The Stream to save this document to.</param>
        /// <example>
        /// Load a document from a file and save it to a Stream.
        /// <code>
        /// // Place holder for a document.
        /// DocX document;
        ///
        /// using (FileStream fs1 = new FileStream(@"C:\Example\Test1.docx", FileMode.Open))
        /// {
        ///     // Load a document using a stream.
        ///     document = DocX.Load(fs1);
        ///
        ///     // Insert a new Paragraph
        ///     document.InsertParagraph("Hello world again!", false);
        /// }
        ///
        /// using (FileStream fs2 = new FileStream(@"C:\Example\Test2.docx", FileMode.Create))
        /// {
        ///     // Save the document to a different stream.
        ///     document.SaveAs(fs2);
        /// }
        ///
        /// // Release this document from memory.
        /// document.Dispose();
        /// </code>
        /// </example>
        /// <example>
        /// Load a document from one Stream and save it to another.
        /// <code>
        /// DocX document;
        /// using (FileStream fs1 = new FileStream(@"C:\Example\Test1.docx", FileMode.Open))
        /// {
        ///     // Load a document using a stream.
        ///     document = DocX.Load(fs1);
        ///
        ///     // Insert a new Paragraph
        ///     document.InsertParagraph("Hello world again!", false);
        /// }
        /// 
        /// using (FileStream fs2 = new FileStream(@"C:\Example\Test2.docx", FileMode.Create))
        /// {
        ///     // Save the document to a different stream.
        ///     document.SaveAs(fs2);
        /// }
        /// </code>
        /// </example>
        /// <seealso cref="DocX.Save()"/>
        /// <seealso cref="DocX.Create(System.IO.Stream)"/>
        /// <seealso cref="DocX.Create(string)"/>
        /// <seealso cref="DocX.Load(System.IO.Stream)"/>
        /// <seealso cref="DocX.Load(string)"/>
        public void SaveAs(Stream stream)
        {
            this.filename = null;
            this.stream = stream;
            Save();
        }

        /// <summary>
        /// Add a custom property to this document. If a custom property already exists with the same name it will be replace. CustomProperty names are case insensitive.
        /// </summary>
        /// <param name="cp">The CustomProperty to add to this document.</param>
        /// <example>
        /// Add a custom properties of each type to a document.
        /// <code>
        /// // Load Example.docx
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // A CustomProperty called forename which stores a string.
        ///     CustomProperty forename;
        ///
        ///     // If this document does not contain a custom property called 'forename', create one.
        ///     if (!document.CustomProperties.ContainsKey("forename"))
        ///     {
        ///         // Create a new custom property called 'forename' and set its value.
        ///         document.AddCustomProperty(new CustomProperty("forename", "Cathal"));
        ///     }
        ///
        ///     // Get this documents custom property called 'forename'.
        ///     forename = document.CustomProperties["forename"];
        ///
        ///     // Print all of the information about this CustomProperty to Console.
        ///     Console.WriteLine(string.Format("Name: '{0}', Value: '{1}'\nPress any key...", forename.Name, forename.Value));
        ///     
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// } // Release this document from memory.
        ///
        /// // Wait for the user to press a key before exiting.
        /// Console.ReadKey();
        /// </code>
        /// </example>
        /// <example>
        /// Extract a CustomProperty from a document called 'forname'. If it doesn't exist, create it. Finally print this custom properties details to Console.
        /// <code>
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // A CustomProperty called forename which stores a string.
        ///     CustomProperty forename;
        ///
        ///     // If this document does not contain a custom property called 'forename', create one.
        ///     if (!document.CustomProperties.ContainsKey("forename"))
        ///     {
        ///         // Create a new custom property called 'forename' and set its value.
        ///         document.AddCustomProperty(new CustomProperty("forename", "Cathal"));
        ///     }
        ///
        ///     // Get this documents custom property called 'forename'.
        ///     forename = document.CustomProperties["forename"];
        ///
        ///     // Print all of the information about this CustomProperty to Console.
        ///     Console.WriteLine(string.Format("Name: '{0}', Value: '{1}'\nPress any key...", forename.Name, forename.Value));
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        ///    
        /// // Wait for the user to press a key before exiting.
        /// Console.ReadKey();
        /// </code>
        /// </example>
        /// <seealso cref="CustomProperty"/>
        /// <seealso cref="CustomProperties"/>
        public void AddCustomProperty(CustomProperty cp)
        {
            // If this document does not contain a customFilePropertyPart create one.
            if(!package.PartExists(new Uri("/docProps/custom.xml", UriKind.Relative)))
                CreateCustomPropertiesPart(this);

            XDocument customPropDoc;
            PackagePart customPropPart = package.GetPart(new Uri("/docProps/custom.xml", UriKind.Relative));
            using (TextReader tr = new StreamReader(customPropPart.GetStream(FileMode.Open, FileAccess.Read)))
                customPropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);

            // Each custom property has a PID, get the highest PID in this document.
            IEnumerable<int> pids =
            (
                from d in customPropDoc.Descendants()
                where d.Name.LocalName == "property"
                select int.Parse(d.Attribute(XName.Get("pid")).Value)
            );

            int pid = 1;
            if (pids.Count() > 0)
                pid = pids.Max();

            // Check if a custom property already exists with this name
            var customProperty =
            (
                from d in customPropDoc.Descendants()
                where (d.Name.LocalName == "property") && (d.Attribute(XName.Get("name")).Value == cp.Name)
                select d
            ).SingleOrDefault();

            // If a custom property with this name already exists remove it.
            if (customProperty != null)
                customProperty.Remove();

            XElement propertiesElement = customPropDoc.Element(XName.Get("Properties", customPropertiesSchema.NamespaceName));
            propertiesElement.Add
            (
                new XElement
                (
                    XName.Get("property", customPropertiesSchema.NamespaceName),
                    new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
                    new XAttribute("pid", pid + 1),
                    new XAttribute("name", cp.Name),
                        new XElement(customVTypesSchema + cp.Type, cp.Value)
                )
            );

            // Save the custom properties
            using (TextWriter tw = new StreamWriter(customPropPart.GetStream(FileMode.Create, FileAccess.Write)))
                customPropDoc.Save(tw, SaveOptions.DisableFormatting);

            // Refresh all fields in this document which display this custom property.
            UpdateCustomPropertyValue(this, cp.Name, cp.Value.ToString());

            // Get all of the custom properties in this document
            customProperties =
            (
                from p in customPropDoc.Descendants(XName.Get("property", customPropertiesSchema.NamespaceName))
                let Name = p.Attribute(XName.Get("name")).Value
                let Type = p.Descendants().Single().Name.LocalName
                let Value = p.Descendants().Single().Value
                select new CustomProperty(Name, Type, Value)
            ).ToDictionary(p => p.Name, StringComparer.CurrentCultureIgnoreCase);
        }

        internal static void CreateCustomPropertiesPart(DocX document)
        {
            PackagePart customPropertiesPart = document.package.CreatePart(new Uri("/docProps/custom.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.custom-properties+xml");

            XDocument customPropDoc = new XDocument
            (
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement
                (
                    XName.Get("Properties", customPropertiesSchema.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "vt", customVTypesSchema)
                )
            );

            using (TextWriter tw = new StreamWriter(customPropertiesPart.GetStream(FileMode.Create, FileAccess.Write)))
                customPropDoc.Save(tw, SaveOptions.DisableFormatting);

            document.package.CreateRelationship(customPropertiesPart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties");
        }

        internal static void UpdateCustomPropertyValue(DocX document, string customPropertyName, string customPropertyValue)
        {
            foreach (XElement e in document.mainDoc.Descendants(XName.Get("fldSimple", w.NamespaceName)))
            {
                if (e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Equals(string.Format(@" DOCPROPERTY  {0}  \* MERGEFORMAT ", customPropertyName), StringComparison.CurrentCultureIgnoreCase))
                {
                    XElement firstRun = e.Element(w + "r");

                    // Delete everything and insert updated text value
                    e.RemoveNodes();

                    XElement t = new XElement(w + "t", customPropertyValue);
                    Novacode.Text.PreserveSpace(t);
                    e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t));
                }
            }
        }

        internal static void RenumberIDs(DocX document)
        {
            IEnumerable<XAttribute> trackerIDs =
                            (from d in document.mainDoc.Descendants()
                             where d.Name.LocalName == "ins" || d.Name.LocalName == "del"
                             select d.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")));

            for (int i = 0; i < trackerIDs.Count(); i++)
                trackerIDs.ElementAt(i).Value = i.ToString();
        }

        /// <summary>
        /// Replace text in this document, not case sensetive.
        /// </summary>
        /// <example>
        /// Replace every instance of "old" in this document with "new".
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Replace every instance of "old" in this document with "new".
        ///     document.ReplaceText("old", "new", false, RegexOptions.IgnoreCase);
        ///                
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <param name="oldValue">The text to replace.</param>
        /// <param name="newValue">The new text to insert.</param>
        /// <param name="trackChanges">Should this change be tracked?</param>
        /// <param name="options">RegexOptions to use for this text replace.</param>
        public void ReplaceText(string oldValue, string newValue, bool trackChanges, RegexOptions options)
        {
            foreach (Paragraph p in paragraphs)
                p.ReplaceText(oldValue, newValue, trackChanges, options);
        }

        /// <summary>
        /// Replace text in this document, case sensetive.
        /// </summary>
        /// <example>
        /// Replace every instance of "old" in this document with "new".
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Replace every instance of "old" in this document with "new".
        ///     document.ReplaceText("old", "new", false);
        ///                
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <param name="oldValue">The text to replace.</param>
        /// <param name="newValue">The new text to insert.</param>
        /// <param name="trackChanges">Should this change be tracked?</param>
        /// <param name="options">RegexOptions to use for this text replace.</param>
        public void ReplaceText(string oldValue, string newValue, bool trackChanges)
        {
            ReplaceText(oldValue, newValue, trackChanges, RegexOptions.None);
        }

        #region IDisposable Members

        /// <summary>
        /// Releases all resources used by this document.
        /// </summary>
        /// <example>
        /// If you take advantage of the using keyword, Dispose() is automatically called for you.
        /// <code>
        /// // Load document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///      // The document is only in memory while in this scope.
        ///
        /// }// Dispose() is automatically called at this point.
        /// </code>
        /// </example>
        /// <example>
        /// This example is equilivant to the one above example.
        /// <code>
        /// // Load document.
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        /// 
        /// // Do something with the document here.
        ///
        /// // Dispose of the document.
        /// document.Dispose();
        /// </code>
        /// </example>
        public void Dispose()
        {
            package.Close();
        }

        #endregion
    }
}