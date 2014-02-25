using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.IO.Packaging;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Collections.ObjectModel;

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
        public virtual ReadOnlyCollection<Paragraph> Paragraphs
        {
            get
            {
                List<Paragraph> paragraphs = GetParagraphs();

                foreach (var p in paragraphs)
                {
                    if ((p.Xml.ElementsAfterSelf().FirstOrDefault() != null) && (p.Xml.ElementsAfterSelf().First().Name.Equals(DocX.w + "tbl")))
                        p.FollowingTable = new Table(this.Document, p.Xml.ElementsAfterSelf().First());

                    p.ParentContainer = GetParentFromXmlName(p.Xml.Ancestors().First().Name.LocalName);

                    if (p.IsListItem)
                    {
                        GetListItemType(p);
                    }
                }

                return paragraphs.AsReadOnly();
            }
        }
        // <summary>
        /// Removes paragraph at specified position
        /// </summary>
        /// <param name="index">Index of paragraph to remove</param>
        /// <returns>True if removed</returns>
        public bool RemoveParagraphAt(int index)
        {
            int i = 0;
            foreach (var paragraph in Xml.Descendants(DocX.w + "p"))
            {
                if (i == index)
                {
                    paragraph.Remove();
                    return true;
                }
                ++i;
 
            }
 
            return false;
        }

        /// <summary>
        /// Removes paragraph
        /// </summary>
        /// <param name="paragraph">Paragraph to remove</param>
        /// <returns>True if removed</returns>
        public bool RemoveParagraph(Paragraph p)
        {
            foreach (var paragraph in Xml.Descendants(DocX.w + "p"))
            {
                if (paragraph.Equals(p.Xml))
                {
                    paragraph.Remove();
                    return true;
                }
            }

            return false;
        }


        public virtual List<Section> Sections
        {
            get
            {
                var allParas = Paragraphs;

                var parasInASection = new List<Paragraph>();
                var sections = new List<Section>();

                foreach (var para in allParas)
                {

                    var sectionInPara = para.Xml.Descendants().FirstOrDefault(s => s.Name.LocalName == "sectPr");

                    if (sectionInPara == null)
                    {
                        parasInASection.Add(para);
                    }
                    else
                    {
                        parasInASection.Add(para);
                        var section = new Section(Document, sectionInPara) { SectionParagraphs = parasInASection };
                        sections.Add(section);
                        parasInASection = new List<Paragraph>();
                    }

                }

                XElement body = Xml.Element(XName.Get("body", DocX.w.NamespaceName));
                XElement baseSectionXml = body.Element(XName.Get("sectPr", DocX.w.NamespaceName));
                var baseSection = new Section(Document, baseSectionXml) { SectionParagraphs = parasInASection };
                sections.Add(baseSection);

                return sections;
            }
        }


        private void GetListItemType(Paragraph p)
        {
            var ilvlNode = p.ParagraphNumberProperties.Descendants().FirstOrDefault(el => el.Name.LocalName == "ilvl");
            var ilvlValue = ilvlNode.Attribute(DocX.w + "val").Value;

            var numIdNode = p.ParagraphNumberProperties.Descendants().FirstOrDefault(el => el.Name.LocalName == "numId");
            var numIdValue = numIdNode.Attribute(DocX.w + "val").Value;

            //find num node in numbering 
            var numNodes = Document.numbering.Descendants().Where(n => n.Name.LocalName == "num");
            XElement numNode = numNodes.FirstOrDefault(node => node.Attribute(DocX.w + "numId").Value.Equals(numIdValue));
           
	        if (numNode != null)
            {
               //Get abstractNumId node and its value from numNode
	            var abstractNumIdNode = numNode.Descendants().First(n => n.Name.LocalName == "abstractNumId");
	            var abstractNumNodeValue = abstractNumIdNode.Attribute(DocX.w + "val").Value;
	
	            var abstractNumNodes = Document.numbering.Descendants().Where(n => n.Name.LocalName == "abstractNum");
	            XElement abstractNumNode =
	              abstractNumNodes.FirstOrDefault(node => node.Attribute(DocX.w + "abstractNumId").Value.Equals(abstractNumNodeValue));
	
	            //Find lvl node
	            var lvlNodes = abstractNumNode.Descendants().Where(n => n.Name.LocalName == "lvl");
	            XElement lvlNode = null;
	            foreach (XElement node in lvlNodes)
	            {
	                if (node.Attribute(DocX.w + "ilvl").Value.Equals(ilvlValue))
	                {
	                    lvlNode = node;
	                    break;
	                }
	            }
	           
	           	var numFmtNode = lvlNode.Descendants().First(n => n.Name.LocalName == "numFmt");
	          		p.ListItemType = GetListItemType(numFmtNode.Attribute(DocX.w + "val").Value);
            }         

        }


        public ContainerType ParentContainer;


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
            // sdtContent are for PageNumbers inside Headers or Footers, don't go any deeper.
            //if (Xml.Name.LocalName == "sdtContent")
            //    return;

            if (Xml.Name.LocalName == "p")
            {
                paragraphs.Add(new Paragraph(Document, Xml, index));

                index += HelperFunctions.GetText(Xml).Length;
            }

            else
            {
                if (Xml.HasElements)
                {
                    foreach (XElement e in Xml.Elements())
                    {
                        GetParagraphsRecursive(e, ref index, ref paragraphs);
                    }
                }
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

        public virtual List<List> Lists
        {
            get
            {
                var lists = new List<List>();
                var list = new List(Document, Xml);

                foreach (var paragraph in Paragraphs)
                {
                    if (paragraph.IsListItem)
                    {
                        if (list.CanAddListItem(paragraph))
                        {
                            list.AddItem(paragraph);
                        }
                        else
                        {
                            lists.Add(list);
                            list = new List(Document, Xml);
                            list.AddItem(paragraph);
                        }
                    }
                }

                lists.Add(list);

                return lists;
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

        /// <summary>
        /// Find all unique instances of the given Regex Pattern,
        /// returning the list of the unique strings found
        /// </summary>
        /// <param name="str"></param>
        /// <param name="options"></param>
        /// <returns></returns>
        public virtual List<string> FindUniqueByPattern(string pattern, RegexOptions options)
        {
            List<string> rawResults = new List<string>();

            foreach (Paragraph p in Paragraphs)
            {   // accumulate the search results from all paragraphs
                List<string> partials = p.FindAllByPattern(pattern, options);
                rawResults.AddRange(partials);
            }

            // this dictionary is used to collect results and test for uniqueness
            Dictionary<string, int> uniqueResults = new Dictionary<string, int>();

            foreach (string currValue in rawResults)
            {
                if (!uniqueResults.ContainsKey(currValue))
                {   // if the dictionary doesn't have it, add it
                    uniqueResults.Add(currValue, 0);
                }
            }

            return uniqueResults.Keys.ToList();  // return the unique list of results
        }

        public virtual void ReplaceText(string oldValue, string newValue, bool trackChanges = false, RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch)
        {
            if (oldValue == null || oldValue.Length == 0)
                throw new ArgumentException("oldValue cannot be null or empty", "oldValue");

            if (newValue == null)
                throw new ArgumentException("newValue cannot be null or empty", "newValue");
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

        /// <summary>
        /// Removes all items with required formatting
        /// </summary>
        /// <returns>Numer of texts removed</returns>
        public int RemoveTextInGivenFormat(Formatting matchFormatting, MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch)
        {
            var deletedCount = 0;
            foreach (var x in Xml.Elements())
            {
                deletedCount += RemoveTextWithFormatRecursive(x, matchFormatting, fo);
            }

            return deletedCount;
        }

        internal int RemoveTextWithFormatRecursive(XElement element, Formatting matchFormatting, MatchFormattingOptions fo)
        {
            var deletedCount = 0;
            foreach (var x in element.Elements())
            {
                if ("rPr".Equals(x.Name.LocalName))
                {
                    if (HelperFunctions.ContainsEveryChildOf(matchFormatting.Xml, x, fo))
                    {
                        x.Parent.Remove();
                        ++deletedCount;
                    }
                }

                deletedCount += RemoveTextWithFormatRecursive(x, matchFormatting, fo);
            }

            return deletedCount;
        }

        public virtual void InsertAtBookmark(string toInsert, string bookmarkName)
        {
            if (bookmarkName.IsNullOrWhiteSpace())
                throw new ArgumentException("bookmark cannot be null or empty", "bookmarkName");

            var headerCollection = Document.Headers;
            var headers = new List<Header> { headerCollection.first, headerCollection.even, headerCollection.odd };
            foreach (var header in headers.Where(x => x != null))
                foreach (var paragraph in header.Paragraphs)
                    paragraph.InsertAtBookmark(toInsert, bookmarkName);

            foreach (var paragraph in Paragraphs)
                paragraph.InsertAtBookmark(toInsert, bookmarkName);

            var footerCollection = Document.Footers;
            var footers = new List<Footer> { footerCollection.first, footerCollection.even, footerCollection.odd };
            foreach (var footer in footers.Where(x => x != null))
                foreach (var paragraph in footer.Paragraphs)
                    paragraph.InsertAtBookmark(toInsert, bookmarkName);
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

            GetParent(p);

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
                    PackagePart style_package = Document.package.CreatePart(style_package_uri, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", CompressionOption.Maximum);
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

            GetParent(newParagraph);

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

            GetParent(newParagraph);

            return newParagraph;
        }


        private ContainerType GetParentFromXmlName(string xmlName)
        {
            ContainerType parent;

            switch (xmlName)
            {
                case "body":
                    parent = ContainerType.Body;
                    break;
                case "p":
                    parent = ContainerType.Paragraph;
                    break;
                case "tbl":
                    parent = ContainerType.Table;
                    break;
                case "sectPr":
                    parent = ContainerType.Section;
                    break;
                case "tc":
                    parent = ContainerType.Cell;
                    break;
                default:
                    parent = ContainerType.None;
                    break;
            }
            return parent;
        }

        private void GetParent(Paragraph newParagraph)
        {
            var containerType = GetType();

            switch (containerType.Name)
            {

                case "Body":
                    newParagraph.ParentContainer = ContainerType.Body;
                    break;
                case "Table":
                    newParagraph.ParentContainer = ContainerType.Table;
                    break;
                case "TOC":
                    newParagraph.ParentContainer = ContainerType.TOC;
                    break;
                case "Section":
                    newParagraph.ParentContainer = ContainerType.Section;
                    break;
                case "Cell":
                    newParagraph.ParentContainer = ContainerType.Cell;
                    break;
                case "Header":
                    newParagraph.ParentContainer = ContainerType.Header;
                    break;
                case "Footer":
                    newParagraph.ParentContainer = ContainerType.Footer;
                    break;
                case "Paragraph":
                    newParagraph.ParentContainer = ContainerType.Paragraph;
                    break;
            }
        }


        private ListItemType GetListItemType(string styleName)
        {
            ListItemType listItemType;

            switch (styleName)
            {
                case "bullet":
                    listItemType = ListItemType.Bulleted;
                    break;
                default:
                    listItemType = ListItemType.Numbered;
                    break;
            }

            return listItemType;
        }



        public virtual void InsertSection()
        {

            InsertSection(false);
        }

        public virtual void InsertSection(bool trackChanges)
        {
            var newParagraphSection = new XElement
            (
                XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName), new XElement(XName.Get("sectPr", DocX.w.NamespaceName), new XElement(XName.Get("type", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", "continuous"))))
            );

            if (trackChanges)
                newParagraphSection = HelperFunctions.CreateEdit(EditType.ins, DateTime.Now, newParagraphSection);

            Xml.Add(newParagraphSection);
        }

        public virtual void InsertSectionPageBreak(bool trackChanges = false)
        {
            var newParagraphSection = new XElement
            (
                XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName), new XElement(XName.Get("sectPr", DocX.w.NamespaceName)))
            );

            if (trackChanges)
                newParagraphSection = HelperFunctions.CreateEdit(EditType.ins, DateTime.Now, newParagraphSection);

            Xml.Add(newParagraphSection);
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

            var paragraphAdded = Paragraphs.Last();

            GetParent(paragraphAdded);

            return paragraphAdded;
        }

        public virtual Paragraph InsertEquation(string equation)
        {
            Paragraph p = InsertParagraph();
            p.AppendEquation(equation);
            return p;
        }

        public virtual Paragraph InsertBookmark(String bookmarkName)
        {
            var p = InsertParagraph();
            p.AppendBookmark(bookmarkName);
            return p;
        }

        public virtual Table InsertTable(int rowCount, int columnCount) //Dmitchern, changed to virtual, and overrided in Table.Cell
        {
            XElement newTable = HelperFunctions.CreateTable(rowCount, columnCount);
            Xml.Add(newTable);

            return new Table(Document, newTable);
        }

        public Table InsertTable(int index, int rowCount, int columnCount)
        {
            XElement newTable = HelperFunctions.CreateTable(rowCount, columnCount);

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

        public List InsertList(List list)
        {
            foreach (var item in list.Items)
            {
              //  item.Font(System.Drawing.FontFamily fontFamily)

                Xml.Add(item.Xml);
            }

            return list;
        }
        public List InsertList(List list, double fontSize)
        {
            foreach (var item in list.Items)
            {
                item.FontSize(fontSize);
                Xml.Add(item.Xml);
            }
            return list;
        }

        public List InsertList(List list, System.Drawing.FontFamily fontFamily, double fontSize)
        {
            foreach (var item in list.Items)
            {
                item.Font(fontFamily);
                item.FontSize(fontSize);
                Xml.Add(item.Xml);
            }
            return list;
        }

        public List InsertList(int index, List list)
        {
            Paragraph p = HelperFunctions.GetFirstParagraphEffectedByInsert(Document, index);

            XElement[] split = HelperFunctions.SplitParagraph(p, index - p.startIndex);
            var elements = new List<XElement> { split[0] };
            elements.AddRange(list.Items.Select(i => new XElement(i.Xml)));
            elements.Add(split[1]);
            p.Xml.ReplaceWith(elements.ToArray());

            return list;
        }
    }
}
