using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using System.Security.Principal;
using System.Collections;
using System.IO.Packaging;
using System.IO;

namespace Novacode
{
    /// <summary>
    /// Represents a document paragraph.
    /// </summary>
    public class Paragraph
    {        
        // This paragraphs text alignment
        private Alignment alignment;

        // A lookup for the runs in this paragraph
        Dictionary<int, Run> runLookup = new Dictionary<int, Run>();

        // The underlying XElement which this Paragraph wraps
        internal XElement xml;
        internal int startIndex, endIndex;

        // A collection of Images in this Paragraph
        private List<Picture> pictures;

        /// <summary>
        /// Returns a list of Pictures in this Paragraph.
        /// </summary>
        public List<Picture> Pictures { get { return pictures; } }

        // A collection of field type DocProperty.
        private List<DocProperty> docProperties;

        internal List<XElement> styles = new List<XElement>();

        /// <summary>
        /// Returns a list of field type DocProperty in this document.
        /// </summary>
        public List<DocProperty> DocumentProperties
        {
            get { return docProperties; }
        }

        internal DocX document;
        internal Paragraph(DocX document, int startIndex, XElement p)
        {           
            this.document = document;
            this.startIndex = startIndex;
            this.endIndex = startIndex + GetElementTextLength(p);
            this.xml = p;

            BuildRunLookup(p);

            // Get all of the images in this document
            pictures = (from i in p.Descendants(XName.Get("drawing", DocX.w.NamespaceName))
                        select new Picture(i)).ToList();

            RebuildDocProperties();

            #region It's possible that a Paragraph may have pStyle references
            // Check if this Paragraph references any pStyle elements.
            var stylesElements = xml.Descendants(XName.Get("pStyle", DocX.w.NamespaceName));

            // If one or more pStyles are referenced.
            if (stylesElements.Count() > 0)
            {
                Uri style_package_uri = new Uri("/word/styles.xml", UriKind.Relative);
                PackagePart styles_document = document.package.GetPart(style_package_uri);
                
                using (TextReader tr = new StreamReader(styles_document.GetStream()))
                {
                    XDocument style_document = XDocument.Load(tr);
                    XElement styles_element = style_document.Element(XName.Get("styles", DocX.w.NamespaceName));

                    var styles_element_ids = stylesElements.Select(e => e.Attribute(XName.Get("val", DocX.w.NamespaceName)).Value);
                    
                    foreach(string id in styles_element_ids)
                    {
                        var style = 
                        (
                            from d in styles_element.Descendants()
                            let styleId = d.Attribute(XName.Get("styleId", DocX.w.NamespaceName))
                            where styleId != null && styleId.Value == id
                            select d
                        ).Single();

                        styles.Add(style);
                    } 
                }
            } 
            #endregion

            #region Pictures
		    // Check if this Paragraph contains any Pictures
            List<string> pictureElementIDs = 
            (
                from d in xml.Descendants()
                let embed = d.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))
                where embed != null
                select embed.Value
            ).ToList();
	        #endregion
        }

        /// <summary>
        /// Insert a new Table before this Paragraph, this Table can be from this document or another document.
        /// </summary>
        /// <param name="t">The Table t to be inserted.</param>
        /// <returns>A new Table inserted before this Paragraph.</returns>
        /// <example>
        /// Insert a new Table before this Paragraph.
        /// <code>
        /// // Place holder for a Table.
        /// Table t;
        ///
        /// // Load document a.
        /// using (DocX documentA = DocX.Load(@"a.docx"))
        /// {
        ///     // Get the first Table from this document.
        ///     t = documentA.Tables[0];
        /// }
        ///
        /// // Load document b.
        /// using (DocX documentB = DocX.Load(@"b.docx"))
        /// {
        ///     // Get the first Paragraph in document b.
        ///     Paragraph p2 = documentB.Paragraphs[0];
        ///
        ///     // Insert the Table from document a before this Paragraph.
        ///     Table newTable = p2.InsertTableBeforeSelf(t);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Table InsertTableBeforeSelf(Table t)
        {
            xml.AddBeforeSelf(t.xml);
            XElement newlyInserted = xml.ElementsBeforeSelf().First();

            t.xml = newlyInserted;
            DocX.RebuildTables(document);
            DocX.RebuildParagraphs(document);

            return t;
        }

        /// <summary>
        /// Insert a new Table into this document before this Paragraph.
        /// </summary>
        /// <param name="rowCount">The number of rows this Table should have.</param>
        /// <param name="coloumnCount">The number of coloumns this Table should have.</param>
        /// <returns>A new Table inserted before this Paragraph.</returns>
        /// <example>
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     //Insert a Paragraph into this document.
        ///     Paragraph p = document.InsertParagraph("Hello World", false);
        ///
        ///     // Insert a new Table before this Paragraph.
        ///     Table newTable = p.InsertTableBeforeSelf(2, 2);
        ///     newTable.Design = TableDesign.LightShadingAccent2;
        ///     newTable.Alignment = Alignment.center;
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Table InsertTableBeforeSelf(int rowCount, int coloumnCount)
        {
            XElement newTable = DocX.CreateTable(rowCount, coloumnCount);
            xml.AddBeforeSelf(newTable);
            XElement newlyInserted = xml.ElementsBeforeSelf().First();

            DocX.RebuildTables(document);
            DocX.RebuildParagraphs(document);
            return new Table(document, newlyInserted);
        }

        /// <summary>
        /// Insert a new Table after this Paragraph.
        /// </summary>
        /// <param name="t">The Table t to be inserted.</param>
        /// <returns>A new Table inserted after this Paragraph.</returns>
        /// <example>
        /// Insert a new Table after this Paragraph.
        /// <code>
        /// // Place holder for a Table.
        /// Table t;
        ///
        /// // Load document a.
        /// using (DocX documentA = DocX.Load(@"a.docx"))
        /// {
        ///     // Get the first Table from this document.
        ///     t = documentA.Tables[0];
        /// }
        ///
        /// // Load document b.
        /// using (DocX documentB = DocX.Load(@"b.docx"))
        /// {
        ///     // Get the first Paragraph in document b.
        ///     Paragraph p2 = documentB.Paragraphs[0];
        ///
        ///     // Insert the Table from document a after this Paragraph.
        ///     Table newTable = p2.InsertTableAfterSelf(t);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Table InsertTableAfterSelf(Table t)
        {
            xml.AddAfterSelf(t.xml);
            XElement newlyInserted = xml.ElementsAfterSelf().First();

            t.xml = newlyInserted;
            DocX.RebuildTables(document);
            DocX.RebuildParagraphs(document);

            return t;
        }

        /// <summary>
        /// Insert a new Table into this document after this Paragraph.
        /// </summary>
        /// <param name="rowCount">The number of rows this Table should have.</param>
        /// <param name="coloumnCount">The number of coloumns this Table should have.</param>
        /// <returns>A new Table inserted after this Paragraph.</returns>
        /// <example>
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     //Insert a Paragraph into this document.
        ///     Paragraph p = document.InsertParagraph("Hello World", false);
        ///
        ///     // Insert a new Table after this Paragraph.
        ///     Table newTable = p.InsertTableAfterSelf(2, 2);
        ///     newTable.Design = TableDesign.LightShadingAccent2;
        ///     newTable.Alignment = Alignment.center;
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Table InsertTableAfterSelf(int rowCount, int coloumnCount)
        {
            XElement newTable = DocX.CreateTable(rowCount, coloumnCount);
            xml.AddAfterSelf(newTable);
            XElement newlyInserted = xml.ElementsAfterSelf().First();

            DocX.RebuildTables(document);
            DocX.RebuildParagraphs(document);
            return new Table(document, newlyInserted);
        }

        /// <summary>
        /// Insert a Paragraph before this Paragraph, this Paragraph may have come from the same or another document.
        /// </summary>
        /// <param name="p">The Paragraph to insert.</param>
        /// <returns>The Paragraph now associated with this document.</returns>
        /// <example>
        /// Take a Paragraph from document a, and insert it into document b before this Paragraph.
        /// <code>
        /// // Place holder for a Paragraph.
        /// Paragraph p;
        ///
        /// // Load document a.
        /// using (DocX documentA = DocX.Load(@"a.docx"))
        /// {
        ///     // Get the first paragraph from this document.
        ///     p = documentA.Paragraphs[0];
        /// }
        ///
        /// // Load document b.
        /// using (DocX documentB = DocX.Load(@"b.docx"))
        /// {
        ///     // Get the first Paragraph in document b.
        ///     Paragraph p2 = documentB.Paragraphs[0];
        ///
        ///     // Insert the Paragraph from document a before this Paragraph.
        ///     Paragraph newParagraph = p2.InsertParagraphBeforeSelf(p);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code> 
        /// </example>
        public Paragraph InsertParagraphBeforeSelf(Paragraph p)
        {
            xml.AddBeforeSelf(p.xml);
            XElement newlyInserted = xml.ElementsBeforeSelf().First();

            p.xml = newlyInserted;
            DocX.RebuildParagraphs(document);

            return p;
        }

        /// <summary>
        /// Insert a new Paragraph before this Paragraph.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <returns>A new Paragraph inserted before this Paragraph.</returns>
        /// <example>
        /// Insert a new paragraph before the first Paragraph in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Paragraph into this document.
        ///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
        ///
        ///     p.InsertParagraphBeforeSelf("I was inserted before the next Paragraph.");
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public Paragraph InsertParagraphBeforeSelf(string text)
        {
            return InsertParagraphBeforeSelf(text, false, new Formatting());
        }

        /// <summary>
        /// Insert a new Paragraph before this Paragraph.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <returns>A new Paragraph inserted before this Paragraph.</returns>
        /// <example>
        /// Insert a new paragraph before the first Paragraph in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Paragraph into this document.
        ///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
        ///
        ///     p.InsertParagraphBeforeSelf("I was inserted before the next Paragraph.", false);
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges)
        {
            return InsertParagraphBeforeSelf(text, trackChanges, new Formatting());
        }

        /// <summary>
        /// Insert a new Paragraph before this Paragraph.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <param name="formatting">The formatting to apply to this insertion.</param>
        /// <returns>A new Paragraph inserted before this Paragraph.</returns>
        /// <example>
        /// Insert a new paragraph before the first Paragraph in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Paragraph into this document.
        ///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
        ///
        ///     Formatting boldFormatting = new Formatting();
        ///     boldFormatting.Bold = true;
        ///
        ///     p.InsertParagraphBeforeSelf("I was inserted before the next Paragraph.", false, boldFormatting);
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges, Formatting formatting)
        {
            XElement newParagraph = new XElement
            (
                XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName)), DocX.FormatInput(text, formatting.Xml)
            );

            if (trackChanges)
                newParagraph = CreateEdit(EditType.ins, DateTime.Now, newParagraph);

            xml.AddBeforeSelf(newParagraph);
            XElement newlyInserted = xml.ElementsBeforeSelf().First();

            Paragraph p = new Paragraph(document, -1, newlyInserted);
            DocX.RebuildParagraphs(document);

            return p;
        }

        /// <summary>
        /// Insert a Paragraph after this Paragraph, this Paragraph may have come from the same or another document.
        /// </summary>
        /// <param name="p">The Paragraph to insert.</param>
        /// <returns>The Paragraph now associated with this document.</returns>
        /// <example>
        /// Take a Paragraph from document a, and insert it into document b after this Paragraph.
        /// <code>
        /// // Place holder for a Paragraph.
        /// Paragraph p;
        ///
        /// // Load document a.
        /// using (DocX documentA = DocX.Load(@"a.docx"))
        /// {
        ///     // Get the first paragraph from this document.
        ///     p = documentA.Paragraphs[0];
        /// }
        ///
        /// // Load document b.
        /// using (DocX documentB = DocX.Load(@"b.docx"))
        /// {
        ///     // Get the first Paragraph in document b.
        ///     Paragraph p2 = documentB.Paragraphs[0];
        ///
        ///     // Insert the Paragraph from document a after this Paragraph.
        ///     Paragraph newParagraph = p2.InsertParagraphAfterSelf(p);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code> 
        /// </example>
        public Paragraph InsertParagraphAfterSelf(Paragraph p)
        {
            xml.AddAfterSelf(p.xml);
            XElement newlyInserted = xml.ElementsAfterSelf().First();

            p.xml = newlyInserted;
            DocX.RebuildParagraphs(document);

            return p;
        }

        /// <summary>
        /// Insert a new Paragraph after this Paragraph.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <param name="formatting">The formatting to apply to this insertion.</param>
        /// <returns>A new Paragraph inserted after this Paragraph.</returns>
        /// <example>
        /// Insert a new paragraph after the first Paragraph in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Paragraph into this document.
        ///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
        ///
        ///     Formatting boldFormatting = new Formatting();
        ///     boldFormatting.Bold = true;
        ///
        ///     p.InsertParagraphAfterSelf("I was inserted after the previous Paragraph.", false, boldFormatting);
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public Paragraph InsertParagraphAfterSelf(string text, bool trackChanges, Formatting formatting)
        {
            XElement newParagraph = new XElement
            (
                XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName)), DocX.FormatInput(text, formatting.Xml)
            );

            if (trackChanges)
                newParagraph = CreateEdit(EditType.ins, DateTime.Now, newParagraph);

            xml.AddAfterSelf(newParagraph);
            XElement newlyInserted = xml.ElementsAfterSelf().First();

            Paragraph p = new Paragraph(document, -1, newlyInserted);
            DocX.RebuildParagraphs(document);

            return p;
        }

        /// <summary>
        /// Insert a new Paragraph after this Paragraph.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <returns>A new Paragraph inserted after this Paragraph.</returns>
        /// <example>
        /// Insert a new paragraph after the first Paragraph in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Paragraph into this document.
        ///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
        ///
        ///     p.InsertParagraphAfterSelf("I was inserted after the previous Paragraph.", false);
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public Paragraph InsertParagraphAfterSelf(string text, bool trackChanges)
        {
            return InsertParagraphAfterSelf(text, trackChanges, new Formatting());
        }

        /// <summary>
        /// Insert a new Paragraph after this Paragraph.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <returns>A new Paragraph inserted after this Paragraph.</returns>
        /// <example>
        /// Insert a new paragraph after the first Paragraph in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Paragraph into this document.
        ///     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
        ///
        ///     p.InsertParagraphAfterSelf("I was inserted after the previous Paragraph.");
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public Paragraph InsertParagraphAfterSelf(string text)
        {
            return InsertParagraphAfterSelf(text, false, new Formatting());
        }

        private void RebuildDocProperties()
        {
            docProperties =
            (
                from dp in xml.Descendants(XName.Get("fldSimple", DocX.w.NamespaceName))
                select new DocProperty(dp)
            ).ToList();
        }

        /// <summary>
        /// Gets or set this Paragraphs text alignment.
        /// </summary>
        public Alignment Alignment 
        { 
            get { return alignment; }

            set 
            {
                alignment = value;

                XElement pPr = xml.Element(XName.Get("pPr", DocX.w.NamespaceName));

                if (alignment != Novacode.Alignment.left)
                {
                    if (pPr == null)
                        xml.Add(new XElement(XName.Get("pPr", DocX.w.NamespaceName)));
                    
                    pPr = xml.Element(XName.Get("pPr", DocX.w.NamespaceName));

                    XElement jc = pPr.Element(XName.Get("jc", DocX.w.NamespaceName));

                    if(jc == null)
                        pPr.Add(new XElement(XName.Get("jc", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), alignment.ToString())));
                    else
                        jc.Attribute(XName.Get("val", DocX.w.NamespaceName)).Value = alignment.ToString();
                }

                else
                {
                    if (pPr != null)
                    {
                        XElement jc = pPr.Element(XName.Get("jc", DocX.w.NamespaceName));

                        if (jc != null)
                            jc.Remove();
                    }
                }
            } 
        }

        /// <summary>
        /// Remove this Paragraph from the document.
        /// </summary>
        /// <param name="trackChanges">Should this remove be tracked as a change?</param>
        /// <example>
        /// Remove a Paragraph from a document and track it as a change.
        /// <code>
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Create and Insert a new Paragraph into this document.
        ///     Paragraph p = document.InsertParagraph("Hello", false);
        ///
        ///     // Remove the Paragraph and track this as a change.
        ///     p.Remove(true);
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public void Remove(bool trackChanges)
        {
            if (trackChanges)
            {
                DateTime now = DateTime.Now.ToUniversalTime();

                List<XElement> elements = xml.Elements().ToList();
                List<XElement> temp = new List<XElement>();
                for (int i = 0; i < elements.Count(); i++ )
                {
                    XElement e = elements[i];

                    if (e.Name.LocalName != "del")
                    {
                        temp.Add(e);
                        e.Remove();
                    }

                    else
                    {
                        if (temp.Count() > 0)
                        {
                            e.AddBeforeSelf(CreateEdit(EditType.del, now, temp.Elements()));
                            temp.Clear();
                        }
                    }
                }

                if (temp.Count() > 0)
                    xml.Add(CreateEdit(EditType.del, now, temp));                   
            }

            else
            {
                runLookup.Clear();

                if (xml.Parent.Name.LocalName == "tc")
                    xml.Value = string.Empty;

                else
                {
                    // Remove this paragraph from the document
                    xml.Remove();
                    xml = null;

                    runLookup = null;
                }
            }

            DocX.RebuildParagraphs(document);
        }

        internal void BuildRunLookup(XElement p)
        {
            // Get the runs in this paragraph
            IEnumerable<XElement> runs = p.Descendants(XName.Get("r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"));

            int startIndex = 0;

            // Loop through each run in this paragraph
            foreach (XElement run in runs)
            {
                // Only add runs which contain text
                if (GetElementTextLength(run) > 0)
                {
                    Run r = new Run(startIndex, run);
                    runLookup.Add(r.EndIndex, r);
                    startIndex = r.EndIndex;
                }
            }
        }

        /// <summary>
        /// Gets the text value of this Paragraph.
        /// </summary>
        public string Text
        {
            // Returns the underlying XElement's Value property.
            get
            {
                StringBuilder sb = new StringBuilder();

                // Loop through each run in this paragraph
                foreach (XElement r in xml.Descendants(XName.Get("r", DocX.w.NamespaceName)))
                {
                    // Loop through each text item in this run
                    foreach (XElement descendant in r.Descendants())
                    {
                        switch (descendant.Name.LocalName)
                        {
                            case "tab":
                                sb.Append("\t");
                                break;
                            case "br":
                                sb.Append("\n");
                                break;
                            case "t":
                                goto case "delText";
                            case "delText":
                                sb.Append(descendant.Value);
                                break;
                            default: break;
                        }
                    }
                }

                return sb.ToString();
            }
        }

        //public Picture InsertPicture(Picture picture)
        //{
        //    Picture newPicture = picture;
        //    newPicture.i = new XElement(picture.i);

        //    xml.Add(newPicture.i);
        //    pictures.Add(newPicture);
        //    return newPicture;  
        //}

        /// <summary>
        /// Insert a Picture at the end of this paragraph.
        /// </summary>
        /// <param name="description">A string to describe this Picture.</param>
        /// <param name="imageID">The unique id that identifies the Image this Picture represents.</param>
        /// <param name="name">The name of this image.</param>
        /// <returns>A Picture.</returns>
        /// <example>
        /// <code>
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Add a new Paragraph to this document.
        ///     Paragraph p = document.InsertParagraph("Here is Picture 1", false);
        ///
        ///     // Add an Image to this document.
        ///     Novacode.Image img = document.AddImage(@"Image.jpg");
        ///
        ///     // Insert pic at the end of Paragraph p.
        ///     Picture pic = p.InsertPicture(img.Id, "Photo 31415", "A pie I baked.");
        ///
        ///     // Rotate the Picture clockwise by 30 degrees. 
        ///     pic.Rotation = 30;
        ///
        ///     // Resize the Picture.
        ///     pic.Width = 400;
        ///     pic.Height = 300;
        ///
        ///     // Set the shape of this Picture to be a cube.
        ///     pic.SetPictureShape(BasicShapes.cube);
        ///
        ///     // Flip the Picture Horizontally.
        ///     pic.FlipHorizontal = true;
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Picture InsertPicture(string imageID, string name, string description)
        {
            Picture p = new Picture(document, imageID, name, description);
            xml.Add(p.xml);
            pictures.Add(p);
            return p;
        }

        public Picture InsertPicture(string imageID)
        {
            return InsertPicture(imageID, string.Empty, string.Empty);
        }

        //public Picture InsertPicture(int index, Picture picture)
        //{
        //    Picture p = picture;
        //    p.i = new XElement(picture.i);

        //    Run run = GetFirstRunEffectedByEdit(index);

        //    if (run == null)
        //        xml.Add(p.i);
        //    else
        //    {
        //        // Split this run at the point you want to insert
        //        XElement[] splitRun = Run.SplitRun(run, index);

        //        // Replace the origional run
        //        run.xml.ReplaceWith
        //        (
        //            splitRun[0],
        //            p.i,
        //            splitRun[1]
        //        );
        //    }

        //    // Rebuild the run lookup for this paragraph
        //    runLookup.Clear();
        //    BuildRunLookup(xml);
        //    DocX.RenumberIDs(document);
        //    return p;
        //}

        /// <summary>
        /// Insert a Picture into this Paragraph at a specified index.
        /// </summary>
        /// <param name="description">A string to describe this Picture.</param>
        /// <param name="imageID">The unique id that identifies the Image this Picture represents.</param>
        /// <param name="name">The name of this image.</param>
        /// <param name="index">The index to insert this Picture at.</param>
        /// <returns>A Picture.</returns>
        /// <example>
        /// <code>
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Add a new Paragraph to this document.
        ///     Paragraph p = document.InsertParagraph("Here is Picture 1", false);
        ///
        ///     // Add an Image to this document.
        ///     Novacode.Image img = document.AddImage(@"Image.jpg");
        ///
        ///     // Insert pic at the start of Paragraph p.
        ///     Picture pic = p.InsertPicture(0, img.Id, "Photo 31415", "A pie I baked.");
        ///
        ///     // Rotate the Picture clockwise by 30 degrees. 
        ///     pic.Rotation = 30;
        ///
        ///     // Resize the Picture.
        ///     pic.Width = 400;
        ///     pic.Height = 300;
        ///
        ///     // Set the shape of this Picture to be a cube.
        ///     pic.SetPictureShape(BasicShapes.cube);
        ///
        ///     // Flip the Picture Horizontally.
        ///     pic.FlipHorizontal = true;
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Picture InsertPicture(int index, string imageID, string name, string description)
        {
            Picture picture = new Picture(document, imageID, name, description);
            
            Run run = GetFirstRunEffectedByEdit(index);

            if (run == null)
                xml.Add(picture.xml);
            else
            {
                // Split this run at the point you want to insert
                XElement[] splitRun = Run.SplitRun(run, index);

                // Replace the origional run
                run.xml.ReplaceWith
                (
                    splitRun[0],
                    picture.xml,
                    splitRun[1]
                );
            }

            // Rebuild the run lookup for this paragraph
            runLookup.Clear();
            BuildRunLookup(xml);
            DocX.RenumberIDs(document);
            return picture;
        }

        public Picture InsertPicture(int index, string imageID)
        {
            return InsertPicture(index, imageID, string.Empty, string.Empty);
        }

        /// <summary>
        /// Creates an Edit either a ins or a del with the specified content and date
        /// </summary>
        /// <param name="t">The type of this edit (ins or del)</param>
        /// <param name="edit_time">The time stamp to use for this edit</param>
        /// <param name="content">The initial content of this edit</param>
        /// <returns></returns>
        internal static XElement CreateEdit(EditType t, DateTime edit_time, object content)
        {
            if (t == EditType.del)
            {
                foreach (object o in (IEnumerable<XElement>)content)
                {
                    if (o is XElement)
                    {
                       XElement e = (o as XElement);
                       IEnumerable<XElement> ts = e.DescendantsAndSelf(XName.Get("t", DocX.w.NamespaceName));
                       
                       for(int i = 0; i < ts.Count(); i ++)
                       {
                           XElement text = ts.ElementAt(i);
                           text.ReplaceWith(new XElement(DocX.w + "delText", text.Attributes(), text.Value));  
                       }
                    }
                }
            }

            return
            (
                new XElement(DocX.w + t.ToString(),
                    new XAttribute(DocX.w + "id", 0),
                    new XAttribute(DocX.w + "author", WindowsIdentity.GetCurrent().Name),
                    new XAttribute(DocX.w + "date", edit_time),
                content)
            );
        }

        internal Run GetFirstRunEffectedByEdit(int index)
        {
            foreach (int runEndIndex in runLookup.Keys)
            {
                if (runEndIndex > index)
                    return runLookup[runEndIndex];
            }

            if (runLookup.Last().Value.EndIndex == index)
                return runLookup.Last().Value;

            throw new ArgumentOutOfRangeException();
        }

        internal Run GetFirstRunEffectedByInsert(int index)
        {
            // This paragraph contains no Runs and insertion is at index 0
            if (runLookup.Keys.Count() == 0 && index == 0)
                return null;

            foreach (int runEndIndex in runLookup.Keys)
            {
                if (runEndIndex >= index)
                    return runLookup[runEndIndex];
            }

            throw new ArgumentOutOfRangeException();
        }

        /// <!-- 
        /// Bug found and fixed by krugs525 on August 12 2009.
        /// Use TFS compare to see exact code change.
        /// -->
        static internal int GetElementTextLength(XElement run)
        {
            int count = 0;

            if (run == null)
                return count;

            foreach (var d in run.Descendants())
            {
                switch (d.Name.LocalName)
                {
                    case "tab":
                        if (d.Parent.Name.LocalName != "tabs")
                            goto case "br"; break;
                    case "br": count++; break;
                    case "t": goto case "delText";
                    case "delText": count += d.Value.Length; break;
                    default: break;
                }
            }
            return count;
        }

        internal XElement[] SplitEdit(XElement edit, int index, EditType type)
        {
            Run run;
            if(type == EditType.del)
                run = GetFirstRunEffectedByEdit(index);
            else
                run = GetFirstRunEffectedByInsert(index);

            XElement[] splitRun = Run.SplitRun(run, index);
            
            XElement splitLeft = new XElement(edit.Name, edit.Attributes(), run.xml.ElementsBeforeSelf(), splitRun[0]);
            if (GetElementTextLength(splitLeft) == 0)
                splitLeft = null;

            XElement splitRight = new XElement(edit.Name, edit.Attributes(), splitRun[1], run.xml.ElementsAfterSelf());
            if (GetElementTextLength(splitRight) == 0)
                splitRight = null;

            return
            (
                new XElement[]
                {
                    splitLeft,
                    splitRight
                }
            );
        }

        /// <summary>
        /// Inserts a specified instance of System.String into a Novacode.DocX.Paragraph at a specified index position.
        /// </summary>
        /// <example>
        /// <code> 
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Iterate through the Paragraphs in this document.
        ///     foreach (Paragraph p in document.Paragraphs)
        ///     {
        ///         // Insert the string "Start: " at the begining of every Paragraph and flag it as a change.
        ///         p.InsertText(0, "Start: ", true);
        ///     }
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <example>
        /// Inserting tabs using the \t switch.
        /// <code>  
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Iterate through the paragraphs in this document.
        ///     foreach (Paragraph p in document.Paragraphs)
        ///     {
        ///         // Insert the string "\tStart:\t" at the begining of every paragraph and flag it as a change.
        ///         p.InsertText(0, "\tStart:\t", true);
        ///     }
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <seealso cref="Paragraph.RemoveText(int, bool)"/>
        /// <seealso cref="Paragraph.RemoveText(int, int, bool)"/>
        /// <seealso cref="Paragraph.ReplaceText(string, string, bool)"/>
        /// <seealso cref="Paragraph.ReplaceText(string, string, bool, RegexOptions)"/>
        /// <param name="index">The index position of the insertion.</param>
        /// <param name="value">The System.String to insert.</param>
        /// <param name="trackChanges">Flag this insert as a change.</param>
        public void InsertText(int index, string value, bool trackChanges)
        {
            InsertText(index, value, trackChanges, null);
        }

        /// <summary>
        /// Inserts a specified instance of System.String into a Novacode.DocX.Paragraph at a specified index position.
        /// </summary>
        /// <example>
        /// <code> 
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Iterate through the Paragraphs in this document.
        ///     foreach (Paragraph p in document.Paragraphs)
        ///     {
        ///         // Insert the string "End: " at the end of every Paragraph and flag it as a change.
        ///         p.InsertText("End: ", true);
        ///     }
        ///       
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <example>
        /// Inserting tabs using the \t switch.
        /// <code>  
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Iterate through the paragraphs in this document.
        ///     foreach (Paragraph p in document.Paragraphs)
        ///     {
        ///         // Insert the string "\tEnd" at the end of every paragraph and flag it as a change.
        ///         p.InsertText("\tEnd", true);
        ///     }
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <seealso cref="Paragraph.RemoveText(int, bool)"/>
        /// <seealso cref="Paragraph.RemoveText(int, int, bool)"/>
        /// <seealso cref="Paragraph.ReplaceText(string, string, bool)"/>
        /// <seealso cref="Paragraph.ReplaceText(string, string, bool, RegexOptions)"/>
        /// <param name="value">The System.String to insert.</param>
        /// <param name="trackChanges">Flag this insert as a change.</param>
        public void InsertText(string value, bool trackChanges)
        {
            List<XElement> newRuns = DocX.FormatInput(value, null);
            xml.Add(newRuns);

            runLookup.Clear();
            BuildRunLookup(xml);
            DocX.RenumberIDs(document);
        }

        /// <summary>
        /// Inserts a specified instance of System.String into a Novacode.DocX.Paragraph at a specified index position.
        /// </summary>
        /// <example>
        /// <code> 
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Create a text formatting.
        ///     Formatting f = new Formatting();
        ///     f.FontColor = Color.Red;
        ///     f.Size = 30;
        ///
        ///     // Iterate through the Paragraphs in this document.
        ///     foreach (Paragraph p in document.Paragraphs)
        ///     {
        ///         // Insert the string "Start: " at the begining of every Paragraph and flag it as a change.
        ///         p.InsertText("Start: ", true, f);
        ///     }
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <example>
        /// Inserting tabs using the \t switch.
        /// <code>  
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///      // Create a text formatting.
        ///      Formatting f = new Formatting();
        ///      f.FontColor = Color.Red;
        ///      f.Size = 30;
        ///        
        ///      // Iterate through the paragraphs in this document.
        ///      foreach (Paragraph p in document.Paragraphs)
        ///      {
        ///          // Insert the string "\tEnd" at the end of every paragraph and flag it as a change.
        ///          p.InsertText("\tEnd", true, f);
        ///      }
        ///       
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <seealso cref="Paragraph.RemoveText(int, bool)"/>
        /// <seealso cref="Paragraph.RemoveText(int, int, bool)"/>
        /// <seealso cref="Paragraph.ReplaceText(string, string, bool)"/>
        /// <seealso cref="Paragraph.ReplaceText(string, string, bool, RegexOptions)"/>
        /// <param name="value">The System.String to insert.</param>
        /// <param name="trackChanges">Flag this insert as a change.</param>
        /// <param name="formatting">The text formatting.</param>
        public void InsertText(string value, bool trackChanges, Formatting formatting)
        {
            List<XElement> newRuns = DocX.FormatInput(value, formatting.Xml);
            xml.Add(newRuns);

            runLookup.Clear();
            BuildRunLookup(xml);
            DocX.RenumberIDs(document);
            DocX.RebuildParagraphs(document);
        }

        /// <summary>
        /// Inserts a specified instance of System.String into a Novacode.DocX.Paragraph at a specified index position.
        /// </summary>
        /// <example>
        /// <code> 
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Create a text formatting.
        ///     Formatting f = new Formatting();
        ///     f.FontColor = Color.Red;
        ///     f.Size = 30;
        ///
        ///     // Iterate through the Paragraphs in this document.
        ///     foreach (Paragraph p in document.Paragraphs)
        ///     {
        ///         // Insert the string "Start: " at the begining of every Paragraph and flag it as a change.
        ///         p.InsertText(0, "Start: ", true, f);
        ///     }
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <example>
        /// Inserting tabs using the \t switch.
        /// <code>  
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Create a text formatting.
        ///     Formatting f = new Formatting();
        ///     f.FontColor = Color.Red;
        ///     f.Size = 30;
        ///
        ///     // Iterate through the paragraphs in this document.
        ///     foreach (Paragraph p in document.Paragraphs)
        ///     {
        ///         // Insert the string "\tStart:\t" at the begining of every paragraph and flag it as a change.
        ///         p.InsertText(0, "\tStart:\t", true, f);
        ///     }
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <seealso cref="Paragraph.RemoveText(int, bool)"/>
        /// <seealso cref="Paragraph.RemoveText(int, int, bool)"/>
        /// <seealso cref="Paragraph.ReplaceText(string, string, bool)"/>
        /// <seealso cref="Paragraph.ReplaceText(string, string, bool, RegexOptions)"/>
        /// <param name="index">The index position of the insertion.</param>
        /// <param name="value">The System.String to insert.</param>
        /// <param name="trackChanges">Flag this insert as a change.</param>
        /// <param name="formatting">The text formatting.</param>
        public void InsertText(int index, string value, bool trackChanges, Formatting formatting)
        {
            // Timestamp to mark the start of insert
            DateTime now = DateTime.Now;
            DateTime insert_datetime = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);

            // Get the first run effected by this Insert
            Run run = GetFirstRunEffectedByInsert(index);

            if (run == null)
            {
                object insert;
                if (formatting != null)
                    insert = DocX.FormatInput(value, formatting.Xml);
                else
                    insert = DocX.FormatInput(value, null);
                
                if (trackChanges)
                    insert = CreateEdit(EditType.ins, insert_datetime, insert);
                xml.Add(insert);
            }

            else
            {
                object newRuns;
                if (formatting != null)
                    newRuns = DocX.FormatInput(value, formatting.Xml);
                else
                    newRuns = DocX.FormatInput(value, run.xml.Element(XName.Get("rPr", DocX.w.NamespaceName)));

                // The parent of this Run
                XElement parentElement = run.xml.Parent;
                switch (parentElement.Name.LocalName)
                {
                    case "ins":
                        {
                            // The datetime that this ins was created
                            DateTime parent_ins_date = DateTime.Parse(parentElement.Attribute(XName.Get("date", DocX.w.NamespaceName)).Value);

                            /* 
                             * Special case: You want to track changes,
                             * and the first Run effected by this insert
                             * has a datetime stamp equal to now.
                            */
                            if (trackChanges && parent_ins_date.CompareTo(insert_datetime) == 0)
                            {
                                /*
                                 * Inserting into a non edit and this special case, is the same procedure.
                                */
                                goto default;
                            }

                            /*
                             * If not the special case above, 
                             * then inserting into an ins or a del, is the same procedure.
                            */
                            goto case "del";
                        }

                    case "del":
                        {
                            object insert = newRuns;
                            if (trackChanges)
                                insert = CreateEdit(EditType.ins, insert_datetime, newRuns);

                            // Split this Edit at the point you want to insert
                            XElement[] splitEdit = SplitEdit(parentElement, index, EditType.ins);

                            // Replace the origional run
                            parentElement.ReplaceWith
                            (
                                splitEdit[0],
                                insert,
                                splitEdit[1]
                            );

                            break;
                        }

                    default:
                        {
                            object insert = newRuns;
                            if (trackChanges && !parentElement.Name.LocalName.Equals("ins"))
                                insert = CreateEdit(EditType.ins, insert_datetime, newRuns);

                            // Split this run at the point you want to insert
                            XElement[] splitRun = Run.SplitRun(run, index);

                            // Replace the origional run
                            run.xml.ReplaceWith
                            (
                                splitRun[0],
                                insert,
                                splitRun[1]
                            );

                            break;
                        }
                }
            }

            // Rebuild the run lookup for this paragraph
            runLookup.Clear();
            BuildRunLookup(xml);
            DocX.RenumberIDs(document);
        }
        /// <summary>
        /// Insert a field of type document property, this field will display the custom property cp, at the end of this paragraph.
        /// </summary>
        /// <param name="cp">The custom property to display.</param>
        /// <param name="f">The formatting to use for this text.</param>
        /// <example>
        /// Create, add and display a custom property in a document.
        /// <code>
        /// // Load a document
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Create a custom property.
        ///     CustomProperty name = new CustomProperty("name", "Cathal Coffey");
        ///        
        ///     // Add this custom property to this document.
        ///     document.AddCustomProperty(name);
        ///
        ///     // Create a text formatting.
        ///     Formatting f = new Formatting();
        ///     f.Bold = true;
        ///     f.Size = 14;
        ///     f.StrikeThrough = StrickThrough.strike;
        ///
        ///     // Insert a new paragraph.
        ///     Paragraph p = document.InsertParagraph("Author: ", false, f);
        ///
        ///     // Insert a field of type document property to display the custom property name
        ///     p.InsertDocProperty(name, f);
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public void InsertDocProperty(CustomProperty cp, Formatting f)
        {
            XElement e = new XElement
            (
                XName.Get("fldSimple", DocX.w.NamespaceName),
                new XAttribute(XName.Get("instr", DocX.w.NamespaceName), string.Format(@"DOCPROPERTY {0} \* MERGEFORMAT", cp.Name)),
                    new XElement(XName.Get("r", DocX.w.NamespaceName),
                        new XElement(XName.Get("t", DocX.w.NamespaceName), f.Xml, cp.Value))
            );

            xml.Add(e);                    
        }

        /// <summary>
        /// Insert a field of type document property, this field will display the custom property cp, at the end of this paragraph.
        /// </summary>
        /// <param name="cp">The custom property to display.</param>
        /// <example>
        /// Create, add and display a custom property in a document.
        /// <code>
        /// // Load a document
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Create a custom property.
        ///     CustomProperty name = new CustomProperty("name", "Cathal Coffey");
        ///        
        ///     // Add this custom property to this document.
        ///     document.AddCustomProperty(name);
        ///
        ///     // Insert a new paragraph.
        ///     Paragraph p = document.InsertParagraph("Author: ", false);
        ///        
        ///     // Insert a field of type document property to display the custom property name
        ///     p.InsertDocProperty(name);
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public void InsertDocProperty(CustomProperty cp)
        {
            InsertDocProperty(cp, new Formatting());
        }

        /// <summary>
        /// Removes characters from a Novacode.DocX.Paragraph.
        /// </summary>
        /// <example>
        /// <code>
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Iterate through the paragraphs
        ///     foreach (Paragraph p in document.Paragraphs)
        ///     {
        ///         // Remove the first two characters from every paragraph
        ///         p.RemoveText(0, 2, false);
        ///     }
        ///        
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <seealso cref="Paragraph.ReplaceText(string, string, bool)"/>
        /// <seealso cref="Paragraph.ReplaceText(string, string, bool, RegexOptions)"/>
        /// <seealso cref="Paragraph.InsertText(string, bool)"/>
        /// <seealso cref="Paragraph.InsertText(int, string, bool)"/>
        /// <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
        /// <seealso cref="Paragraph.InsertText(string, bool, Formatting)"/>
        /// <param name="index">The position to begin deleting characters.</param>
        /// <param name="count">The number of characters to delete</param>
        /// <param name="trackChanges">Track changes</param>
        public void RemoveText(int index, int count, bool trackChanges)
        {
            // Timestamp to mark the start of insert
            DateTime now = DateTime.Now;
            DateTime remove_datetime = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);

            // The number of characters processed so far
            int processed = 0;

            do
            {
                // Get the first run effected by this Remove
                Run run = GetFirstRunEffectedByEdit(index + processed);

                // The parent of this Run
                XElement parentElement = run.xml.Parent;
                switch (parentElement.Name.LocalName)
                {
                    case "ins":
                        {
                            XElement[] splitEditBefore = SplitEdit(parentElement, index + processed, EditType.del);
                            int min = Math.Min(count - processed, run.xml.ElementsAfterSelf().Sum(e => GetElementTextLength(e)));
                            XElement[] splitEditAfter = SplitEdit(parentElement, index + processed + min, EditType.del);

                            XElement temp = SplitEdit(splitEditBefore[1], index + processed + min, EditType.del)[0];
                            object middle = CreateEdit(EditType.del, remove_datetime, temp.Elements());
                            processed += GetElementTextLength(middle as XElement);
                            
                            if (!trackChanges)
                                middle = null;
                                
                            parentElement.ReplaceWith
                            (
                                splitEditBefore[0],
                                middle,
                                splitEditAfter[1]
                            );

                            processed += GetElementTextLength(middle as XElement);
                            break;
                        }

                    case "del":
                        {
                            if (trackChanges)
                            {
                                // You cannot delete from a deletion, advance processed to the end of this del
                                processed += GetElementTextLength(parentElement);
                            }

                            else
                                goto case "ins";

                            break;
                        }

                    default:
                        {
                            XElement[] splitRunBefore = Run.SplitRun(run, index + processed);
                            int min = Math.Min(index + processed + (count - processed), run.EndIndex);
                            XElement[] splitRunAfter = Run.SplitRun(run, min);

                            object middle = CreateEdit(EditType.del, remove_datetime, new List<XElement>() { Run.SplitRun(new Run(run.StartIndex + GetElementTextLength(splitRunBefore[0]), splitRunBefore[1]), min)[0] });
                            processed += GetElementTextLength(middle as XElement);
                            
                            if (!trackChanges)
                                middle = null;

                            run.xml.ReplaceWith
                            (
                                splitRunBefore[0],
                                middle,
                                splitRunAfter[1]
                            );

                            break;
                        }
                }

                // If after this remove the parent element is empty, remove it.
                if (GetElementTextLength(parentElement) == 0)
                {
                    if (parentElement.Parent != null && parentElement.Parent.Name.LocalName != "tc")
                        parentElement.Remove();
                }
            }
            while (processed < count);

            // Rebuild the run lookup
            runLookup.Clear();
            BuildRunLookup(xml);
            DocX.RenumberIDs(document);
        }


        /// <summary>
        /// Removes characters from a Novacode.DocX.Paragraph.
        /// </summary>
        /// <example>
        /// <code>
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Iterate through the paragraphs
        ///     foreach (Paragraph p in document.Paragraphs)
        ///     {
        ///         // Remove all but the first 2 characters from this Paragraph.
        ///         p.RemoveText(2, false);
        ///     }
        ///        
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <seealso cref="Paragraph.ReplaceText(string, string, bool)"/>
        /// <seealso cref="Paragraph.ReplaceText(string, string, bool, RegexOptions)"/>
        /// <seealso cref="Paragraph.InsertText(string, bool)"/>
        /// <seealso cref="Paragraph.InsertText(int, string, bool)"/>
        /// <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
        /// <seealso cref="Paragraph.InsertText(string, bool, Formatting)"/>
        /// <param name="index">The position to begin deleting characters.</param>
        /// <param name="trackChanges">Track changes</param>
        public void RemoveText(int index, bool trackChanges)
        {
            RemoveText(index, Text.Length - index, trackChanges);
        }

        /// <summary>
        /// Replaces all occurrences of a specified System.String in this instance, with another specified System.String.
        /// </summary>
        /// <example>
        /// <code>
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Iterate through the paragraphs in this document.
        ///     foreach (Paragraph p in document.Paragraphs)
        ///     {
        ///         // Replace all instances of the string "wrong" with the string "right" and ignore case.
        ///         p.ReplaceText("wrong", "right", false, RegexOptions.IgnoreCase);
        ///     }
        ///        
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <seealso cref="Paragraph.RemoveText(int, int, bool)"/>
        /// <seealso cref="Paragraph.RemoveText(int, bool)"/>
        /// <seealso cref="Paragraph.InsertText(string, bool)"/>
        /// <seealso cref="Paragraph.InsertText(int, string, bool)"/>
        /// <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
        /// <seealso cref="Paragraph.InsertText(string, bool, Formatting)"/>
        /// <param name="newValue">A System.String to replace all occurances of oldValue.</param>
        /// <param name="oldValue">A System.String to be replaced.</param>
        /// <param name="options">A bitwise OR combination of RegexOption enumeration options.</param>
        /// <param name="trackChanges">Track changes</param>
        public void ReplaceText(string oldValue, string newValue, bool trackChanges, RegexOptions options)
        {
            MatchCollection mc = Regex.Matches(this.Text, Regex.Escape(oldValue), options);
            
            // Loop through the matches in reverse order
            foreach (Match m in mc.Cast<Match>().Reverse())
            {
                InsertText(m.Index + oldValue.Length, newValue, trackChanges);
                RemoveText(m.Index, m.Length, trackChanges);
            }
        }

        /// <summary>
        /// Find all instances of a string in this paragraph and return their indexes in a List.
        /// </summary>
        /// <param name="str">The string to find</param>
        /// <returns>A list of indexes.</returns>
        /// <example>
        /// Find all instances of Hello in this document and insert 'don't' in frount of them.
        /// <code>
        /// // Load a document
        /// using (DocX document = DocX.Load(@"Test.docx"))
        /// {
        ///     // Loop through the paragraphs in this document.
        ///     foreach(Paragraph p in document.Paragraphs)
        ///     {
        ///         // Find all instances of 'go' in this paragraph.
        ///         List&lt;int&gt; gos = document.FindAll("go");
        ///
        ///         /* 
        ///          * Insert 'don't' in frount of every instance of 'go' in this document to produce 'don't go'.
        ///          * An important trick here is to do the inserting in reverse document order. If you inserted 
        ///          * in document order, every insert would shift the index of the remaining matches.
        ///          */
        ///         gos.Reverse();
        ///         foreach (int index in gos)
        ///         {
        ///             p.InsertText(index, "don't ", false);
        ///         }
        ///     }
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public List<int> FindAll(string str)
        {
            return FindAll(str, RegexOptions.None);
        }

        /// <summary>
        /// Find all instances of a string in this paragraph and return their indexes in a List.
        /// </summary>
        /// <param name="str">The string to find</param>
        /// <param name="options">The options to use when finding a string match.</param>
        /// <returns>A list of indexes.</returns>
        /// <example>
        /// Find all instances of Hello in this document and insert 'don't' in frount of them.
        /// <code>
        /// // Load a document
        /// using (DocX document = DocX.Load(@"Test.docx"))
        /// {
        ///     // Loop through the paragraphs in this document.
        ///     foreach(Paragraph p in document.Paragraphs)
        ///     {
        ///         // Find all instances of 'go' in this paragraph (Ignore case).
        ///         List&lt;int&gt; gos = document.FindAll("go", RegexOptions.IgnoreCase);
        ///
        ///         /* 
        ///          * Insert 'don't' in frount of every instance of 'go' in this document to produce 'don't go'.
        ///          * An important trick here is to do the inserting in reverse document order. If you inserted 
        ///          * in document order, every insert would shift the index of the remaining matches.
        ///          */
        ///         gos.Reverse();
        ///         foreach (int index in gos)
        ///         {
        ///             p.InsertText(index, "don't ", false);
        ///         }
        ///     }
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public List<int> FindAll(string str, RegexOptions options)
        {
            MatchCollection mc = Regex.Matches(this.Text, Regex.Escape(str), options);

            var query =
            (
                from m in mc.Cast<Match>()
                select m.Index
            ).ToList();

            return query;
        }

        /// <summary>
        /// Replaces all occurrences of a specified System.String in this instance, with another specified System.String.
        /// </summary>
        /// <example>
        /// <code>
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Iterate through the paragraphs in this document.
        ///     foreach (Paragraph p in document.Paragraphs)
        ///     {
        ///         // Replace all instances of the string "wrong" with the string "right".
        ///         p.ReplaceText("wrong", "right", false);
        ///     }
        ///       
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <seealso cref="Paragraph.RemoveText(int, int, bool)"/>
        /// <seealso cref="Paragraph.RemoveText(int, bool)"/>
        /// <seealso cref="Paragraph.InsertText(string, bool)"/>
        /// <seealso cref="Paragraph.InsertText(int, string, bool)"/>
        /// <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
        /// <seealso cref="Paragraph.InsertText(string, bool, Formatting)"/>
        /// <param name="newValue">A System.String to replace all occurances of oldValue.</param>
        /// <param name="oldValue">A System.String to be replaced.</param>
        /// <param name="trackChanges">Track changes</param>
        public void ReplaceText(string oldValue, string newValue, bool trackChanges)
        {
            ReplaceText(oldValue, newValue, trackChanges, RegexOptions.None);
        }
    }
}
