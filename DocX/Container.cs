using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using System.IO.Packaging;
using System.IO;

namespace Novacode
{
    public abstract class Container : DocXElement
    {
        /// <summary>
        /// Returns a list of all Paragraphs inside this container.
        /// </summary>
        /// <example>
        /// <code>
        ///  Load a document.
        /// using (DocX document = DocX.Load(@"Test.docx"))
        /// {
        ///    // All Paragraphs in this document.
        ///    List<Paragraph> documentParagraphs = document.Paragraphs;
        ///    
        ///    // Make sure this document contains at least one Table.
        ///    if (document.Tables.Count() > 0)
        ///    {
        ///        // Get the first Table in this document.
        ///        Table t = document.Tables[0];
        ///
        ///        // All Paragraphs in this Table.
        ///        List<Paragraph> tableParagraphs = t.Paragraphs;
        ///    
        ///        // Make sure this Table contains at least one Row.
        ///        if (t.Rows.Count() > 0)
        ///        {
        ///            // Get the first Row in this document.
        ///            Row r = t.Rows[0];
        ///
        ///            // All Paragraphs in this Row.
        ///            List<Paragraph> rowParagraphs = r.Paragraphs;
        ///
        ///            // Make sure this Row contains at least one Cell.
        ///            if (r.Cells.Count() > 0)
        ///            {
        ///                // Get the first Cell in this document.
        ///                Cell c = r.Cells[0];
        ///
        ///                // All Paragraphs in this Cell.
        ///                List<Paragraph> cellParagraphs = c.Paragraphs;
        ///            }
        ///        }
        ///    }
        ///
        ///    // Save all changes to this document.
        ///    document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public virtual List<Paragraph> Paragraphs 
        {
            get 
            {
                List<Paragraph> paragraphs = GetParagraphs();

                foreach (var p in paragraphs)
                {
                    if ((p.Xml.ElementsAfterSelf().FirstOrDefault() != null) && (p.Xml.ElementsAfterSelf().First().Name.Equals(DocX.w + "tbl")))
                        p.FollowingTable = new Table(this.Document, p.Xml.ElementsAfterSelf().First());
                }
                
                return paragraphs;
            }
        }

        internal List<Paragraph> GetParagraphs()
        {
            // Need some memory that can be updated by the recursive search.
            int index = 0;
            List<Paragraph> paragraphs = new List<Paragraph>();

            GetParagraphsRecursive(Xml, ref index, ref paragraphs);

            return paragraphs;
        }

        internal void GetParagraphsRecursive(XElement Xml, ref int index, ref List<Paragraph> paragraphs)
        {
            if (Xml.Name.LocalName == "p")
            {
                paragraphs.Add(new Paragraph(Document, Xml, index));

                index += HelperFunctions.GetText(Xml).Length;
            }

            else
            {
                if (Xml.HasElements)
                    foreach (XElement e in Xml.Elements())
                        GetParagraphsRecursive(e, ref index, ref paragraphs);
            }
        }

        public virtual List<Table> Tables
        {
            get
            {
                List<Table> tables =
                (
                    from t in Xml.Descendants(DocX.w + "tbl")
                    select new Table(Document, t)
                ).ToList();

                return tables;
            }
        }

        public virtual List<Hyperlink> Hyperlinks
        {
            get
            {
                List<Hyperlink> hyperlinks = new List<Hyperlink>();

                foreach (Paragraph p in Paragraphs)
                    hyperlinks.AddRange(p.Hyperlinks);

                return hyperlinks;
            }
        }

        public virtual List<Picture> Pictures
        {
            get
            {
                List<Picture> pictures = new List<Picture>();

                foreach (Paragraph p in Paragraphs)
                    pictures.AddRange(p.Pictures);

                return pictures;
            }
        }

        /// <summary>
        /// Sets the Direction of content.
        /// </summary>
        /// <param name="direction">Direction either LeftToRight or RightToLeft</param>
        /// <example>
        /// Set the Direction of content in a Paragraph to RightToLeft.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"Test.docx"))
        /// {
        ///    // Get the first Paragraph from this document.
        ///    Paragraph p = document.InsertParagraph();
        ///
        ///    // Set the Direction of this Paragraph.
        ///    p.Direction = Direction.RightToLeft;
        ///
        ///    // Make sure the document contains at lest one Table.
        ///    if (document.Tables.Count() > 0)
        ///    {
        ///        // Get the first Table from this document.
        ///        Table t = document.Tables[0];
        ///
        ///        /* 
        ///         * Set the direction of the entire Table.
        ///         * Note: The same function is available at the Row and Cell level.
        ///         */
        ///        t.SetDirection(Direction.RightToLeft);
        ///    }
        ///
        ///    // Save all changes to this document.
        ///    document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public virtual void SetDirection(Direction direction)
        {
            foreach (Paragraph p in Paragraphs)
                p.Direction = direction;
        }

        public virtual List<int> FindAll(string str)
        {
            return FindAll(str, RegexOptions.None);
        }

        public virtual List<int> FindAll(string str, RegexOptions options)
        {
            List<int> list = new List<int>();

            foreach (Paragraph p in Paragraphs)
            {
                List<int> indexes = p.FindAll(str, options);

                for (int i = 0; i < indexes.Count(); i++)
                    indexes[0] += p.startIndex;

                list.AddRange(indexes);
            }

            return list;
        }

        public virtual void ReplaceText(string oldValue, string newValue, bool trackChanges = false, RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch)
        {
            // ReplaceText in Headers of the document.
            Headers headers = Document.Headers;
            List<Header> headerList = new List<Header> { headers.first, headers.even, headers.odd };
            foreach (Header h in headerList)
                if (h != null)
                    foreach (Paragraph p in h.Paragraphs)
                        p.ReplaceText(oldValue, newValue, trackChanges, options, newFormatting, matchFormatting, fo);

            // ReplaceText int main body of document.
            foreach (Paragraph p in Paragraphs)
                p.ReplaceText(oldValue, newValue, trackChanges, options, newFormatting, matchFormatting, fo);

            // ReplaceText in Footers of the document.
            Footers footers = Document.Footers;
            List<Footer> footerList = new List<Footer> { footers.first, footers.even, footers.odd };
            foreach (Footer f in footerList)
                if (f != null)
                    foreach (Paragraph p in f.Paragraphs)
                        p.ReplaceText(oldValue, newValue, trackChanges, options, newFormatting, matchFormatting, fo);
        }

        public virtual Paragraph InsertParagraph(int index, string text, bool trackChanges)
        {
            return InsertParagraph(index, text, trackChanges, null);
        }

        public virtual Paragraph InsertParagraph()
        {
            return InsertParagraph(string.Empty, false);
        }

        public virtual Paragraph InsertParagraph(int index, Paragraph p)
        {
            XElement newXElement = new XElement(p.Xml);
            p.Xml = newXElement;

            Paragraph paragraph = HelperFunctions.GetFirstParagraphEffectedByInsert(Document, index);

            if (paragraph == null)
                Xml.Add(p.Xml);
            else
            {
                XElement[] split = HelperFunctions.SplitParagraph(paragraph, index - paragraph.startIndex);

                paragraph.Xml.ReplaceWith
                (
                    split[0],
                    newXElement,
                    split[1]
                );
            }

            return p;
        }

        public virtual Paragraph InsertParagraph(Paragraph p)
        {
            #region Styles
            XDocument style_document;

            if (p.styles.Count() > 0)
            {
                Uri style_package_uri = new Uri("/word/styles.xml", UriKind.Relative);
                if (!Document.package.PartExists(style_package_uri))
                {
                    PackagePart style_package = Document.package.CreatePart(style_package_uri, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");
                    using (TextWriter tw = new StreamWriter(style_package.GetStream()))
                    {
                        style_document = new XDocument
                        (
                            new XDeclaration("1.0", "UTF-8", "yes"),
                            new XElement(XName.Get("styles", DocX.w.NamespaceName))
                        );

                        style_document.Save(tw);
                    }
                }

                PackagePart styles_document = Document.package.GetPart(style_package_uri);
                using (TextReader tr = new StreamReader(styles_document.GetStream()))
                {
                    style_document = XDocument.Load(tr);
                    XElement styles_element = style_document.Element(XName.Get("styles", DocX.w.NamespaceName));

                    var ids = from d in styles_element.Descendants(XName.Get("style", DocX.w.NamespaceName))
                              let a = d.Attribute(XName.Get("styleId", DocX.w.NamespaceName))
                              where a != null
                              select a.Value;

                    foreach (XElement style in p.styles)
                    {
                        // If styles_element does not contain this element, then add it.

                        if (!ids.Contains(style.Attribute(XName.Get("styleId", DocX.w.NamespaceName)).Value))
                            styles_element.Add(style);
                    }
                }

                using (TextWriter tw = new StreamWriter(styles_document.GetStream()))
                    style_document.Save(tw);
            }
            #endregion

            XElement newXElement = new XElement(p.Xml);

            Xml.Add(newXElement);

            int index = 0;
            if (Document.paragraphLookup.Keys.Count() > 0)
            {
                index = Document.paragraphLookup.Last().Key;

                if (Document.paragraphLookup.Last().Value.Text.Length == 0)
                    index++;
                else
                    index += Document.paragraphLookup.Last().Value.Text.Length;
            }

            Paragraph newParagraph = new Paragraph(Document, newXElement, index);
            Document.paragraphLookup.Add(index, newParagraph);
            return newParagraph;
        }

        public virtual Paragraph InsertParagraph(int index, string text, bool trackChanges, Formatting formatting)
        {
            Paragraph newParagraph = new Paragraph(Document, new XElement(DocX.w + "p"), index);
            newParagraph.InsertText(0, text, trackChanges, formatting);

            Paragraph firstPar = HelperFunctions.GetFirstParagraphEffectedByInsert(Document, index);

            if (firstPar != null)
            {
                XElement[] splitParagraph = HelperFunctions.SplitParagraph(firstPar, index - firstPar.startIndex);

                firstPar.Xml.ReplaceWith
                (
                    splitParagraph[0],
                    newParagraph.Xml,
                    splitParagraph[1]
                );
            }

            else
                Xml.Add(newParagraph);

            return newParagraph;
        }

        public virtual Paragraph InsertParagraph(string text)
        {
            return InsertParagraph(text, false, new Formatting());
        }

        public virtual Paragraph InsertParagraph(string text, bool trackChanges)
        {
            return InsertParagraph(text, trackChanges, new Formatting());
        }

        public virtual Paragraph InsertParagraph(string text, bool trackChanges, Formatting formatting)
        {
            XElement newParagraph = new XElement
            (
                XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName)), HelperFunctions.FormatInput(text, formatting.Xml)
            );

            if (trackChanges)
                newParagraph = HelperFunctions.CreateEdit(EditType.ins, DateTime.Now, newParagraph);

            Xml.Add(newParagraph);

            return Paragraphs.Last();
        }

        public Table InsertTable(int coloumnCount, int rowCount)
        {
            XElement newTable = HelperFunctions.CreateTable(rowCount, coloumnCount);
            Xml.Elements().First().Add(newTable);

            return new Table(Document, newTable);
        }

        public Table InsertTable(int index, int coloumnCount, int rowCount)
        {
            XElement newTable = HelperFunctions.CreateTable(rowCount, coloumnCount);

            Paragraph p = HelperFunctions.GetFirstParagraphEffectedByInsert(Document, index);

            if (p == null)
                Xml.Elements().First().AddFirst(newTable);

            else
            {
                XElement[] split = HelperFunctions.SplitParagraph(p, index - p.startIndex);

                p.Xml.ReplaceWith
                (
                    split[0],
                    newTable,
                    split[1]
                );
            }


            return new Table(Document, newTable);
        }

        public Table InsertTable(Table t)
        {
            XElement newXElement = new XElement(t.Xml);
            Xml.Add(newXElement);

            Table newTable = new Table(Document, newXElement);
            newTable.Design = t.Design;

            return newTable;
        }

        public Table InsertTable(int index, Table t)
        {
            Paragraph p = HelperFunctions.GetFirstParagraphEffectedByInsert(Document, index);

            XElement[] split = HelperFunctions.SplitParagraph(p, index - p.startIndex);
            XElement newXElement = new XElement(t.Xml);
            p.Xml.ReplaceWith
            (
                split[0],
                newXElement,
                split[1]
            );

            Table newTable = new Table(Document, newXElement);
            newTable.Design = t.Design;

            return newTable;
        }
        internal Container(DocX document, XElement xml)
            : base(document, xml)
        {
            
        }
    }
}
