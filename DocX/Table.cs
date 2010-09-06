using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Novacode;
using System.IO.Packaging;
using System.IO;
using System.Reflection;
using System.Drawing;

namespace Novacode
{   
    /// <summary>
    /// Represents a Table in a document.
    /// </summary>
    public class Table : InsertBeforeOrAfter
    {
        private Alignment alignment;
        private AutoFit autofit;
        private List<Row> rows;
        private int rowCount, columnCount;

        /// <summary>
        /// Returns a list of all Paragraphs inside this container.
        /// </summary>
        /// 
        public virtual List<Paragraph> Paragraphs
        {
            get
            {
                List<Paragraph> paragraphs = new List<Paragraph>();

                foreach (Row r in Rows)
                    paragraphs.AddRange(r.Paragraphs);

                return paragraphs;
            }
        }

        /// <summary>
        /// Returns a list of all Pictures in a Table.
        /// </summary>
        /// <example>
        /// Returns a list of all Pictures in a Table.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Load(@"Test.docx"))
        /// {
        ///     // Get the first Table in a document.
        ///     Table t = document.Tables[0];
        ///
        ///     // Get all of the Pictures in this Table.
        ///     List<Picture> pictures = t.Pictures;
        ///
        ///     // Save this document.
        ///     document.Save();
        /// }
        /// </code>
        /// </example>
        public List<Picture> Pictures
        {
            get
            {
                List<Picture> pictures = new List<Picture>();

                foreach (Row r in Rows)
                    pictures.AddRange(r.Pictures);

                return pictures;
            }
        }

        /// <summary>
        /// Set the direction of all content in this Table.
        /// </summary>
        /// <param name="direction">(Left to Right) or (Right to Left)</param>
        /// <example>
        /// Set the content direction for all content in a table to RightToLeft.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"Test.docx"))
        /// {
        ///     // Get the first table in a document.
        ///     Table table = document.Tables[0];
        ///
        ///     // Set the content direction for all content in this table to RightToLeft.
        ///     table.SetDirection(Direction.RightToLeft);
        ///    
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }
        /// </code>
        /// </example>
        public void SetDirection(Direction direction)
        {
            XElement tblPr = GetOrCreate_tblPr();
            tblPr.Add(new XElement(DocX.w + "bidiVisual"));
            
            foreach (Row r in Rows)
                r.SetDirection(direction);
        }

        /// <summary>
        /// Get all of the Hyperlinks in this Table.
        /// </summary>
        /// <example>
        /// Get all of the Hyperlinks in this Table.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Load(@"Test.docx"))
        /// {
        ///     // Get the first Table in this document.
        ///     Table t = document.Tables[0];
        ///
        ///     // Get a list of all Hyperlinks in this Table.
        ///     List<Hyperlink> hyperlinks = t.Hyperlinks;
        ///
        ///     // Save this document.
        ///     document.Save();
        /// }
        /// </code>
        /// </example>
        public List<Hyperlink> Hyperlinks
        {
            get
            {
                List<Hyperlink> hyperlinks = new List<Hyperlink>();

                foreach (Row r in Rows)
                    hyperlinks.AddRange(r.Hyperlinks);

                return hyperlinks;
            }
        }

        /// <summary>
        /// If the tblPr element doesent exist it is created, either way it is returned by this function.
        /// </summary>
        /// <returns>The tblPr element for this Table.</returns>
        internal XElement GetOrCreate_tblPr()
        {
            // Get the element.
            XElement tblPr = Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));

            // If it dosen't exist, create it.
            if (tblPr == null)
            {
                Xml.AddFirst(new XElement(XName.Get("tblPr", DocX.w.NamespaceName)));
                tblPr = Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
            }

            // Return the pPr element for this Paragraph.
            return tblPr;
        }

        /// <summary>
        /// Returns the number of rows in this table.
        /// </summary>
        public int RowCount { get { return rowCount; } }

        /// <summary>
        /// Returns the number of coloumns in this table.
        /// </summary>
        public int ColumnCount { get { return columnCount; } }

        /// <summary>
        /// Returns a list of rows in this table.
        /// </summary>
        public List<Row> Rows { get { return rows; } }
        private TableDesign design;

        internal PackagePart mainPart;

        internal Table(DocX document, XElement xml):base(document, xml)
        {
            autofit = AutoFit.ColoumnWidth;
            this.Xml = xml;

            XElement properties = xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
            
            rows = (from r in xml.Elements(XName.Get("tr", DocX.w.NamespaceName))
                       select new Row(this, document, r)).ToList();

            rowCount = rows.Count;

            if (rows.Count > 0)
                if (rows[0].Cells.Count > 0)
                    columnCount = rows[0].Cells.Count;

            XElement style = properties.Element(XName.Get("tblStyle", DocX.w.NamespaceName));
            if (style != null)
            {
                XAttribute val = style.Attribute(XName.Get("val", DocX.w.NamespaceName));

                if (val != null)
                {
                    try
                    {
                        design = (TableDesign)Enum.Parse(typeof(TableDesign), val.Value.Replace("-", string.Empty));
                    }

                    catch (Exception e)
                    {
                        design = TableDesign.Custom;
                    }
                }
                else
                    design = TableDesign.None;
            }

            else
                design = TableDesign.None;
        }

        public Alignment Alignment
        {
            get { return alignment; }
            set
            {
                string alignmentString = string.Empty;
                switch (value)
                {
                    case Alignment.left:
                    {
                        alignmentString = "left";
                        break;
                    }

                    case Alignment.both:
                    {
                        alignmentString = "both";
                        break;
                    }


                    case Alignment.right:
                    {
                        alignmentString = "right";
                        break;
                    }

                    case Alignment.center:
                    {
                        alignmentString = "center";
                        break;
                    }
                }

                XElement tblPr = Xml.Descendants(XName.Get("tblPr", DocX.w.NamespaceName)).First();
                XElement jc = tblPr.Descendants(XName.Get("jc", DocX.w.NamespaceName)).FirstOrDefault();

                if(jc != null)
                    jc.Remove();

                jc = new XElement(XName.Get("jc", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), alignmentString));
                tblPr.Add(jc);
                alignment = value;
            }
        }

        /// <summary>
        /// Auto size this table according to some rule.
        /// </summary>
        public AutoFit AutoFit
        {
            get{return autofit;}
            
            set
            {
                string attributeValue = string.Empty;
                switch(value)
                {
                    case AutoFit.ColoumnWidth:
                    {
                        attributeValue = "dxa";
                        break;
                    }

                    case AutoFit.Contents:
                    {
                        attributeValue = "auto";
                        break;
                    }

                    case AutoFit.Window:
                    {
                        attributeValue = "pct";
                        break;
                    }
                }

                var query = from d in Xml.Descendants()
                            let type = d.Attribute(XName.Get("type", DocX.w.NamespaceName))
                            where (d.Name.LocalName == "tcW" || d.Name.LocalName == "tblW") && type != null
                            select type;

                foreach (XAttribute type in query)
                    type.Value = attributeValue;

                autofit = value;
            }
        }
        /// <summary>
        /// The design\style to apply to this table.
        /// </summary>
        public TableDesign Design 
        {
            get { return design; }
            set
            {
                XElement tblPr = Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
                XElement style = tblPr.Element(XName.Get("tblStyle", DocX.w.NamespaceName));
                if (style == null)
                {
                    tblPr.Add(new XElement(XName.Get("tblStyle", DocX.w.NamespaceName)));
                    style = tblPr.Element(XName.Get("tblStyle", DocX.w.NamespaceName));
                }

                XAttribute val = style.Attribute(XName.Get("val", DocX.w.NamespaceName));
                if(val == null)
                {
                    style.Add(new XAttribute(XName.Get("val", DocX.w.NamespaceName), ""));
                    val = style.Attribute(XName.Get("val", DocX.w.NamespaceName));
                }

                design = value;

                if (design == TableDesign.None)
                {
                    if (style != null)
                        style.Remove();
                }

                switch (design)
                {
                    case TableDesign.TableNormal: val.Value = "TableNormal"; break;
                    case TableDesign.TableGrid: val.Value = "TableGrid"; break;
                    case TableDesign.LightShading: val.Value = "LightShading"; break;
                    case TableDesign.LightShadingAccent1: val.Value = "LightShading-Accent1"; break;
                    case TableDesign.LightShadingAccent2: val.Value = "LightShading-Accent2"; break;
                    case TableDesign.LightShadingAccent3: val.Value = "LightShading-Accent3"; break;
                    case TableDesign.LightShadingAccent4: val.Value = "LightShading-Accent4"; break;
                    case TableDesign.LightShadingAccent5: val.Value = "LightShading-Accent5"; break;
                    case TableDesign.LightShadingAccent6: val.Value = "LightShading-Accent6"; break;
                    case TableDesign.LightList: val.Value = "LightList"; break;
                    case TableDesign.LightListAccent1: val.Value = "LightList-Accent1"; break;
                    case TableDesign.LightListAccent2: val.Value = "LightList-Accent2"; break;
                    case TableDesign.LightListAccent3: val.Value = "LightList-Accent3"; break;
                    case TableDesign.LightListAccent4: val.Value = "LightList-Accent4"; break;
                    case TableDesign.LightListAccent5: val.Value = "LightList-Accent5"; break;
                    case TableDesign.LightListAccent6: val.Value = "LightList-Accent6"; break;
                    case TableDesign.LightGrid: val.Value = "LightGrid"; break;
                    case TableDesign.LightGridAccent1: val.Value = "LightGrid-Accent1"; break;
                    case TableDesign.LightGridAccent2: val.Value = "LightGrid-Accent2"; break;
                    case TableDesign.LightGridAccent3: val.Value = "LightGrid-Accent3"; break;
                    case TableDesign.LightGridAccent4: val.Value = "LightGrid-Accent4"; break;
                    case TableDesign.LightGridAccent5: val.Value = "LightGrid-Accent5"; break;
                    case TableDesign.LightGridAccent6: val.Value = "LightGrid-Accent6"; break;
                    case TableDesign.MediumShading1: val.Value = "MediumShading1"; break;
                    case TableDesign.MediumShading1Accent1: val.Value = "MediumShading1-Accent1"; break;
                    case TableDesign.MediumShading1Accent2: val.Value = "MediumShading1-Accent2"; break;
                    case TableDesign.MediumShading1Accent3: val.Value = "MediumShading1-Accent3"; break;
                    case TableDesign.MediumShading1Accent4: val.Value = "MediumShading1-Accent4"; break;
                    case TableDesign.MediumShading1Accent5: val.Value = "MediumShading1-Accent5"; break;
                    case TableDesign.MediumShading1Accent6: val.Value = "MediumShading1-Accent6"; break;
                    case TableDesign.MediumShading2: val.Value = "MediumShading2"; break;
                    case TableDesign.MediumShading2Accent1: val.Value = "MediumShading2-Accent1"; break;
                    case TableDesign.MediumShading2Accent2: val.Value = "MediumShading2-Accent2"; break;
                    case TableDesign.MediumShading2Accent3: val.Value = "MediumShading2-Accent3"; break;
                    case TableDesign.MediumShading2Accent4: val.Value = "MediumShading2-Accent4"; break;
                    case TableDesign.MediumShading2Accent5: val.Value = "MediumShading2-Accent5"; break;
                    case TableDesign.MediumShading2Accent6: val.Value = "MediumShading2-Accent6"; break;
                    case TableDesign.MediumList1: val.Value = "MediumList1"; break;
                    case TableDesign.MediumList1Accent1: val.Value = "MediumList1-Accent1"; break;
                    case TableDesign.MediumList1Accent2: val.Value = "MediumList1-Accent2"; break;
                    case TableDesign.MediumList1Accent3: val.Value = "MediumList1-Accent3"; break;
                    case TableDesign.MediumList1Accent4: val.Value = "MediumList1-Accent4"; break;
                    case TableDesign.MediumList1Accent5: val.Value = "MediumList1-Accent5"; break;
                    case TableDesign.MediumList1Accent6: val.Value = "MediumList1-Accent6"; break;
                    case TableDesign.MediumList2: val.Value = "MediumList2"; break;
                    case TableDesign.MediumList2Accent1: val.Value = "MediumList2-Accent1"; break;
                    case TableDesign.MediumList2Accent2: val.Value = "MediumList2-Accent2"; break;
                    case TableDesign.MediumList2Accent3: val.Value = "MediumList2-Accent3"; break;
                    case TableDesign.MediumList2Accent4: val.Value = "MediumList2-Accent4"; break;
                    case TableDesign.MediumList2Accent5: val.Value = "MediumList2-Accent5"; break;
                    case TableDesign.MediumList2Accent6: val.Value = "MediumList2-Accent6"; break;
                    case TableDesign.MediumGrid1: val.Value = "MediumGrid1"; break;
                    case TableDesign.MediumGrid1Accent1: val.Value = "MediumGrid1-Accent1"; break;
                    case TableDesign.MediumGrid1Accent2: val.Value = "MediumGrid1-Accent2"; break;
                    case TableDesign.MediumGrid1Accent3: val.Value = "MediumGrid1-Accent3"; break;
                    case TableDesign.MediumGrid1Accent4: val.Value = "MediumGrid1-Accent4"; break;
                    case TableDesign.MediumGrid1Accent5: val.Value = "MediumGrid1-Accent5"; break;
                    case TableDesign.MediumGrid1Accent6: val.Value = "MediumGrid1-Accent6"; break;
                    case TableDesign.MediumGrid2: val.Value = "MediumGrid2"; break;
                    case TableDesign.MediumGrid2Accent1: val.Value = "MediumGrid2-Accent1"; break;
                    case TableDesign.MediumGrid2Accent2: val.Value = "MediumGrid2-Accent2"; break;
                    case TableDesign.MediumGrid2Accent3: val.Value = "MediumGrid2-Accent3"; break;
                    case TableDesign.MediumGrid2Accent4: val.Value = "MediumGrid2-Accent4"; break;
                    case TableDesign.MediumGrid2Accent5: val.Value = "MediumGrid2-Accent5"; break;
                    case TableDesign.MediumGrid2Accent6: val.Value = "MediumGrid2-Accent6"; break;
                    case TableDesign.MediumGrid3: val.Value = "MediumGrid3"; break;
                    case TableDesign.MediumGrid3Accent1: val.Value = "MediumGrid3-Accent1"; break;
                    case TableDesign.MediumGrid3Accent2: val.Value = "MediumGrid3-Accent2"; break;
                    case TableDesign.MediumGrid3Accent3: val.Value = "MediumGrid3-Accent3"; break;
                    case TableDesign.MediumGrid3Accent4: val.Value = "MediumGrid3-Accent4"; break;
                    case TableDesign.MediumGrid3Accent5: val.Value = "MediumGrid3-Accent5"; break;
                    case TableDesign.MediumGrid3Accent6: val.Value = "MediumGrid3-Accent6"; break;

                    case TableDesign.DarkList: val.Value = "DarkList"; break;
                    case TableDesign.DarkListAccent1: val.Value = "DarkList-Accent1"; break;
                    case TableDesign.DarkListAccent2: val.Value = "DarkList-Accent2"; break;
                    case TableDesign.DarkListAccent3: val.Value = "DarkList-Accent3"; break;
                    case TableDesign.DarkListAccent4: val.Value = "DarkList-Accent4"; break;
                    case TableDesign.DarkListAccent5: val.Value = "DarkList-Accent5"; break;
                    case TableDesign.DarkListAccent6: val.Value = "DarkList-Accent6"; break;

                    case TableDesign.ColorfulShading: val.Value = "ColorfulShading"; break;
                    case TableDesign.ColorfulShadingAccent1: val.Value = "ColorfulShading-Accent1"; break;
                    case TableDesign.ColorfulShadingAccent2: val.Value = "ColorfulShading-Accent2"; break;
                    case TableDesign.ColorfulShadingAccent3: val.Value = "ColorfulShading-Accent3"; break;
                    case TableDesign.ColorfulShadingAccent4: val.Value = "ColorfulShading-Accent4"; break;
                    case TableDesign.ColorfulShadingAccent5: val.Value = "ColorfulShading-Accent5"; break;
                    case TableDesign.ColorfulShadingAccent6: val.Value = "ColorfulShading-Accent6"; break;

                    case TableDesign.ColorfulList: val.Value = "ColorfulList"; break;
                    case TableDesign.ColorfulListAccent1: val.Value = "ColorfulList-Accent1"; break;
                    case TableDesign.ColorfulListAccent2: val.Value = "ColorfulList-Accent2"; break;
                    case TableDesign.ColorfulListAccent3: val.Value = "ColorfulList-Accent3"; break;
                    case TableDesign.ColorfulListAccent4: val.Value = "ColorfulList-Accent4"; break;
                    case TableDesign.ColorfulListAccent5: val.Value = "ColorfulList-Accent5"; break;
                    case TableDesign.ColorfulListAccent6: val.Value = "ColorfulList-Accent6"; break;

                    case TableDesign.ColorfulGrid: val.Value = "ColorfulGrid"; break;
                    case TableDesign.ColorfulGridAccent1: val.Value = "ColorfulGrid-Accent1"; break;
                    case TableDesign.ColorfulGridAccent2: val.Value = "ColorfulGrid-Accent2"; break;
                    case TableDesign.ColorfulGridAccent3: val.Value = "ColorfulGrid-Accent3"; break;
                    case TableDesign.ColorfulGridAccent4: val.Value = "ColorfulGrid-Accent4"; break;
                    case TableDesign.ColorfulGridAccent5: val.Value = "ColorfulGrid-Accent5"; break;
                    case TableDesign.ColorfulGridAccent6: val.Value = "ColorfulGrid-Accent6"; break;

                    default: break;
                }

                XDocument style_doc;
                PackagePart word_styles = Document.package.GetPart(new Uri("/word/styles.xml", UriKind.Relative));
                using (TextReader tr = new StreamReader(word_styles.GetStream()))
                    style_doc = XDocument.Load(tr);

                var tableStyle =
                (
                    from e in style_doc.Descendants()
                    let styleId = e.Attribute(XName.Get("styleId", DocX.w.NamespaceName))
                    where (styleId != null && styleId.Value == val.Value)
                    select e
                ).FirstOrDefault();

                if (tableStyle == null)
                {
                    XDocument external_style_doc = HelperFunctions.DecompressXMLResource("Novacode.Resources.styles.xml.gz");

                    var styleElement =
                    (
                        from e in external_style_doc.Descendants()
                        let styleId = e.Attribute(XName.Get("styleId", DocX.w.NamespaceName))
                        where (styleId != null && styleId.Value == val.Value)
                        select e
                    ).First();

                    style_doc.Element(XName.Get("styles", DocX.w.NamespaceName)).Add(styleElement);

                    using (TextWriter tw = new StreamWriter(word_styles.GetStream(FileMode.Create)))
                        style_doc.Save(tw, SaveOptions.None);
                }
            }
        }

        /// <summary>
        /// Insert a row at the end of this table.
        /// </summary>
        /// <example>
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Get the first table in this document.
        ///     Table table = document.Tables[0];
        ///        
        ///     // Insert a new row at the end of this table.
        ///     Row row = table.InsertRow();
        ///
        ///     // Loop through each cell in this new row.
        ///     foreach (Cell c in row.Cells)
        ///     {
        ///         // Set the text of each new cell to "Hello".
        ///         c.Paragraphs[0].InsertText("Hello", false);
        ///     }
        ///
        ///     // Save the document to a new file.
        ///     document.SaveAs(@"C:\Example\Test2.docx");
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <returns>A new row.</returns>
        public Row InsertRow()
        {
            return InsertRow(rows.Count);
        }

        /// <summary>
        /// Returns the index of this Table.
        /// </summary>
        /// <example>
        /// Replace the first table in this document with a new Table.
        /// <code>
        /// // Load a document into memory.
        /// using (DocX document = DocX.Load(@"Test.docx"))
        /// {
        ///     // Get the first Table in this document.
        ///     Table t = document.Tables[0];
        ///
        ///     // Get the character index of Table t in this document.
        ///     int index = t.Index;
        ///
        ///     // Remove Table t.
        ///     t.Remove();
        ///
        ///     // Insert a new Table at the original index of Table t.
        ///     Table newTable = document.InsertTable(index, 4, 4);
        ///
        ///     // Set the design of this new Table, so that we can see it.
        ///     newTable.Design = TableDesign.LightShadingAccent1;
        ///
        ///     // Save all changes made to the document.
        ///     document.Save();
        /// } // Release this document from memory.
        /// </code>
        /// </example>
        public int Index
        {
            get
            {
                int index = 0;
                IEnumerable<XElement> previous = Xml.ElementsBeforeSelf();

                foreach (XElement e in previous)
                    index += Paragraph.GetElementTextLength(e);

                return index;
            }
        }

        /// <summary>
        /// Remove this Table from this document.
        /// </summary>
        /// <example>
        /// Remove the first Table from this document.
        /// <code>
        /// // Load a document into memory.
        /// using (DocX document = DocX.Load(@"Test.docx"))
        /// {
        ///     // Get the first Table in this document.
        ///     Table t = d.Tables[0];
        ///        
        ///     // Remove this Table.
        ///     t.Remove();
        ///
        ///     // Save all changes made to the document.
        ///     document.Save();
        /// } // Release this document from memory.
        /// </code>
        /// </example>
        public void Remove()
        {
            Xml.Remove();
        }

        /// <summary>
        /// Insert a column to the right of a Table.
        /// </summary>
        /// <example>
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Get the first Table in this document.
        ///     Table table = document.Tables[0];
        ///
        ///     // Insert a new column to this right of this table.
        ///     table.InsertColumn();
        ///
        ///     // Set the new coloumns text to "Row no."
        ///     table.Rows[0].Cells[table.ColumnCount - 1].Paragraph.InsertText("Row no.", false);
        ///
        ///     // Loop through each row in the table.
        ///     for (int i = 1; i &lt; table.Rows.Count; i++)
        ///     {
        ///         // The current row.
        ///         Row row = table.Rows[i];
        ///
        ///         // The cell in this row that belongs to the new coloumn.
        ///         Cell cell = row.Cells[table.ColumnCount - 1];
        ///
        ///         // The first Paragraph that this cell houses.
        ///         Paragraph p = cell.Paragraphs[0];
        ///
        ///         // Insert this rows index.
        ///         p.InsertText(i.ToString(), false);
        ///     }
        ///
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public void InsertColumn()
        {
            InsertColumn(columnCount);
        }

        /// <summary>
        /// Remove the last row from this Table.
        /// </summary>
        /// <example>
        /// Remove the last row from a Table.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Get the first table in this document.
        ///     Table table = document.Tables[0];
        ///
        ///     // Remove the last row from this table.
        ///     table.RemoveRow();
        ///
        ///     // Save the document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public void RemoveRow()
        {
            RemoveRow(rowCount - 1);
        }

        /// <summary>
        /// Remove a row from this Table.
        /// </summary>
        /// <param name="index">The row to remove.</param>
        /// <example>
        /// Remove the first row from a Table.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Get the first table in this document.
        ///     Table table = document.Tables[0];
        ///
        ///     // Remove the first row from this table.
        ///     table.RemoveRow(0);
        ///
        ///     // Save the document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public void RemoveRow(int index)
        {
            if (index < 0 || index > rows.Count)
                throw new IndexOutOfRangeException();

            rows[index].Xml.Remove();
        }

        /// <summary>
        /// Remove the last column for this Table.
        /// </summary>
        /// <example>
        /// Remove the last column from a Table.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Get the first table in this document.
        ///     Table table = document.Tables[0];
        ///
        ///     // Remove the last column from this table.
        ///     table.RemoveColumn();
        ///
        ///     // Save the document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public void RemoveColumn()
        {
            RemoveColumn(columnCount - 1);
        }

        /// <summary>
        /// Remove a coloumn from this Table.
        /// </summary>
        /// <param name="index">The coloumn to remove.</param>
        /// <example>
        /// Remove the first column from a Table.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Get the first table in this document.
        ///     Table table = document.Tables[0];
        ///
        ///     // Remove the first column from this table.
        ///     table.RemoveColumn(0);
        ///
        ///     // Save the document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public void RemoveColumn(int index)
        {
            if (index < 0 || index > columnCount - 1)
                throw new IndexOutOfRangeException();

            foreach (Row r in rows)
                r.Cells[index].Xml.Remove();
        }

        /// <summary>
        /// Insert a row into this table.
        /// </summary>
        /// <example>
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Get the first table in this document.
        ///     Table table = document.Tables[0];
        ///        
        ///     // Insert a new row at index 1 in this table.
        ///     Row row = table.InsertRow(1);
        ///
        ///     // Loop through each cell in this new row.
        ///     foreach (Cell c in row.Cells)
        ///     {
        ///         // Set the text of each new cell to "Hello".
        ///         c.Paragraphs[0].InsertText("Hello", false);
        ///     }
        ///
        ///     // Save the document to a new file.
        ///     document.SaveAs(@"C:\Example\Test2.docx");
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <param name="index">Index to insert row at.</param>
        /// <returns>A new Row</returns>
        public Row InsertRow(int index)
        {
            if (index < 0 || index > rows.Count)
                throw new IndexOutOfRangeException();

            List<XElement> content = new List<XElement>();
                        
            foreach (Cell c in rows[0].Cells)
                content.Add(new XElement(XName.Get("tc", DocX.w.NamespaceName), new XElement(XName.Get("p", DocX.w.NamespaceName))));

            XElement e = new XElement(XName.Get("tr", DocX.w.NamespaceName), content);
            Row newRow = new Row(this, Document, e);

            XElement rowXml;
            if (index == rows.Count)
            {
                rowXml = rows.Last().Xml;
                rowXml.AddAfterSelf(newRow.Xml);
            }
            
            else
            {
                rowXml = rows[index].Xml;
                rowXml.AddBeforeSelf(newRow.Xml);
            }

            rows.Insert(index, newRow);
            rowCount = rows.Count;
            return newRow;
        }

        /// <summary>
        /// Insert a column into a table.
        /// </summary>
        /// <param name="index">The index to insert the column at.</param>
        /// <example>
        /// Insert a column to the left of a table.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Get the first Table in this document.
        ///     Table table = document.Tables[0];
        ///
        ///     // Insert a new column to this left of this table.
        ///     table.InsertColumn(0);
        ///
        ///     // Set the new coloumns text to "Row no."
        ///     table.Rows[0].Cells[table.ColumnCount - 1].Paragraph.InsertText("Row no.", false);
        ///
        ///     // Loop through each row in the table.
        ///     for (int i = 1; i &lt; table.Rows.Count; i++)
        ///     {
        ///         // The current row.
        ///         Row row = table.Rows[i];
        ///
        ///         // The cell in this row that belongs to the new coloumn.
        ///         Cell cell = row.Cells[table.ColumnCount - 1];
        ///
        ///         // The first Paragraph that this cell houses.
        ///         Paragraph p = cell.Paragraphs[0];
        ///
        ///         // Insert this rows index.
        ///         p.InsertText(i.ToString(), false);
        ///     }
        ///
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public void InsertColumn(int index)
        {
            if (rows.Count > 0)
            {
                foreach (Row r in rows)
                {
                    if(columnCount == index)
                        r.Cells[index - 1].Xml.AddAfterSelf(new XElement(XName.Get("tc", DocX.w.NamespaceName), new XElement(XName.Get("p", DocX.w.NamespaceName))));
                    else
                        r.Cells[index].Xml.AddBeforeSelf(new XElement(XName.Get("tc", DocX.w.NamespaceName), new XElement(XName.Get("p", DocX.w.NamespaceName))));
                }

                rows = (from r in Xml.Elements(XName.Get("tr", DocX.w.NamespaceName))
                        select new Row(this, Document, r)).ToList();

                rowCount = rows.Count;

                if (rows.Count > 0)
                    if (rows[0].Cells.Count > 0)
                        columnCount = rows[0].Cells.Count;
            }
        }

        /// <summary>
        /// Insert a page break before a Table.
        /// </summary>
        /// <example>
        /// Insert a Table and a Paragraph into a document with a page break between them.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {              
        ///     // Insert a new Paragraph.
        ///     Paragraph p1 = document.InsertParagraph("Paragraph", false);
        ///
        ///     // Insert a new Table.
        ///     Table t1 = document.InsertTable(2, 2);
        ///     t1.Design = TableDesign.LightShadingAccent1;
        ///     
        ///     // Insert a page break before this Table.
        ///     t1.InsertPageBreakBeforeSelf();
        ///     
        ///     // Save this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public override void InsertPageBreakBeforeSelf()
        {
            base.InsertPageBreakBeforeSelf();
        }


        /// <summary>
        /// Insert a page break after a Table.
        /// </summary>
        /// <example>
        /// Insert a Table and a Paragraph into a document with a page break between them.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a new Table.
        ///     Table t1 = document.InsertTable(2, 2);
        ///     t1.Design = TableDesign.LightShadingAccent1;
        ///        
        ///     // Insert a page break after this Table.
        ///     t1.InsertPageBreakAfterSelf();
        ///        
        ///     // Insert a new Paragraph.
        ///     Paragraph p1 = document.InsertParagraph("Paragraph", false);
        ///
        ///     // Save this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public override void InsertPageBreakAfterSelf()
        {
            base.InsertPageBreakAfterSelf();
        }

        /// <summary>
        /// Insert a new Table before this Table, this Table can be from this document or another document.
        /// </summary>
        /// <param name="t">The Table t to be inserted</param>
        /// <returns>A new Table inserted before this Table.</returns>
        /// <example>
        /// Insert a new Table before this Table.
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
        ///     // Get the first Table in document b.
        ///     Table t2 = documentB.Tables[0];
        ///
        ///     // Insert the Table from document a before this Table.
        ///     Table newTable = t2.InsertTableBeforeSelf(t);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public override Table InsertTableBeforeSelf(Table t)
        {
            return base.InsertTableBeforeSelf(t);
        }

        /// <summary>
        /// Insert a new Table into this document before this Table.
        /// </summary>
        /// <param name="rowCount">The number of rows this Table should have.</param>
        /// <param name="coloumnCount">The number of coloumns this Table should have.</param>
        /// <returns>A new Table inserted before this Table.</returns>
        /// <example>
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     //Insert a Table into this document.
        ///     Table t = document.InsertTable(2, 2);
        ///     t.Design = TableDesign.LightShadingAccent1;
        ///     t.Alignment = Alignment.center;
        ///     
        ///     // Insert a new Table before this Table.
        ///     Table newTable = t.InsertTableBeforeSelf(2, 2);
        ///     newTable.Design = TableDesign.LightShadingAccent2;
        ///     newTable.Alignment = Alignment.center;
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public override Table InsertTableBeforeSelf(int rowCount, int coloumnCount)
        {
            return base.InsertTableBeforeSelf(rowCount, coloumnCount);
        }

        /// <summary>
        /// Insert a new Table after this Table, this Table can be from this document or another document.
        /// </summary>
        /// <param name="t">The Table t to be inserted</param>
        /// <returns>A new Table inserted after this Table.</returns>
        /// <example>
        /// Insert a new Table after this Table.
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
        ///     // Get the first Table in document b.
        ///     Table t2 = documentB.Tables[0];
        ///
        ///     // Insert the Table from document a after this Table.
        ///     Table newTable = t2.InsertTableAfterSelf(t);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public override Table InsertTableAfterSelf(Table t)
        {
            return base.InsertTableAfterSelf(t);
        }

        /// <summary>
        /// Insert a new Table into this document after this Table.
        /// </summary>
        /// <param name="rowCount">The number of rows this Table should have.</param>
        /// <param name="coloumnCount">The number of coloumns this Table should have.</param>
        /// <returns>A new Table inserted before this Table.</returns>
        /// <example>
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     //Insert a Table into this document.
        ///     Table t = document.InsertTable(2, 2);
        ///     t.Design = TableDesign.LightShadingAccent1;
        ///     t.Alignment = Alignment.center;
        ///     
        ///     // Insert a new Table after this Table.
        ///     Table newTable = t.InsertTableAfterSelf(2, 2);
        ///     newTable.Design = TableDesign.LightShadingAccent2;
        ///     newTable.Alignment = Alignment.center;
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public override Table InsertTableAfterSelf(int rowCount, int coloumnCount)
        {
            return base.InsertTableAfterSelf(rowCount, coloumnCount);
        }

        /// <summary>
        /// Insert a Paragraph before this Table, this Paragraph may have come from the same or another document.
        /// </summary>
        /// <param name="p">The Paragraph to insert.</param>
        /// <returns>The Paragraph now associated with this document.</returns>
        /// <example>
        /// Take a Paragraph from document a, and insert it into document b before this Table.
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
        ///     // Get the first Table in document b.
        ///     Table t = documentB.Tables[0];
        ///
        ///     // Insert the Paragraph from document a before this Table.
        ///     Paragraph newParagraph = t.InsertParagraphBeforeSelf(p);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code> 
        /// </example>
        public override Paragraph InsertParagraphBeforeSelf(Paragraph p)
        {
            return base.InsertParagraphBeforeSelf(p);
        }

        /// <summary>
        /// Insert a new Paragraph before this Table.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <returns>A new Paragraph inserted before this Table.</returns>
        /// <example>
        /// Insert a new Paragraph before the first Table in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Table into this document.
        ///     Table t = document.InsertTable(2, 2);
        ///
        ///     t.InsertParagraphBeforeSelf("I was inserted before the next Table.");
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public override Paragraph InsertParagraphBeforeSelf(string text)
        {
            return base.InsertParagraphBeforeSelf(text);
        }

        /// <summary>
        /// Insert a new Paragraph before this Table.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <returns>A new Paragraph inserted before this Table.</returns>
        /// <example>
        /// Insert a new paragraph before the first Table in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Table into this document.
        ///     Table t = document.InsertTable(2, 2);
        ///
        ///     t.InsertParagraphBeforeSelf("I was inserted before the next Table.", false);
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public override Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges)
        {
            return base.InsertParagraphBeforeSelf(text, trackChanges);
        }

        /// <summary>
        /// Insert a new Paragraph before this Table.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <param name="formatting">The formatting to apply to this insertion.</param>
        /// <returns>A new Paragraph inserted before this Table.</returns>
        /// <example>
        /// Insert a new paragraph before the first Table in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Table into this document.
        ///     Table t = document.InsertTable(2, 2);
        ///
        ///     Formatting boldFormatting = new Formatting();
        ///     boldFormatting.Bold = true;
        ///
        ///     t.InsertParagraphBeforeSelf("I was inserted before the next Table.", false, boldFormatting);
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public override Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges, Formatting formatting)
        {
            return base.InsertParagraphBeforeSelf(text, trackChanges, formatting);
        }

        /// <summary>
        /// Insert a Paragraph after this Table, this Paragraph may have come from the same or another document.
        /// </summary>
        /// <param name="p">The Paragraph to insert.</param>
        /// <returns>The Paragraph now associated with this document.</returns>
        /// <example>
        /// Take a Paragraph from document a, and insert it into document b after this Table.
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
        ///     // Get the first Table in document b.
        ///     Table t = documentB.Tables[0];
        ///
        ///     // Insert the Paragraph from document a after this Table.
        ///     Paragraph newParagraph = t.InsertParagraphAfterSelf(p);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code> 
        /// </example>
        public override Paragraph InsertParagraphAfterSelf(Paragraph p)
        {
            return base.InsertParagraphAfterSelf(p);
        }

        /// <summary>
        /// Insert a new Paragraph after this Table.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <param name="formatting">The formatting to apply to this insertion.</param>
        /// <returns>A new Paragraph inserted after this Table.</returns>
        /// <example>
        /// Insert a new paragraph after the first Table in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Table into this document.
        ///     Table t = document.InsertTable(2, 2);
        ///
        ///     Formatting boldFormatting = new Formatting();
        ///     boldFormatting.Bold = true;
        ///
        ///     t.InsertParagraphAfterSelf("I was inserted after the previous Table.", false, boldFormatting);
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public override Paragraph InsertParagraphAfterSelf(string text, bool trackChanges, Formatting formatting)
        {
            return base.InsertParagraphAfterSelf(text, trackChanges, formatting);
        }

        /// <summary>
        /// Insert a new Paragraph after this Table.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <returns>A new Paragraph inserted after this Table.</returns>
        /// <example>
        /// Insert a new paragraph after the first Table in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Table into this document.
        ///     Table t = document.InsertTable(2, 2);
        ///
        ///     t.InsertParagraphAfterSelf("I was inserted after the previous Table.", false);
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public override Paragraph InsertParagraphAfterSelf(string text, bool trackChanges)
        {
            return base.InsertParagraphAfterSelf(text, trackChanges);
        }

        /// <summary>
        /// Insert a new Paragraph after this Table.
        /// </summary>
        /// <param name="text">The initial text for this new Paragraph.</param>
        /// <returns>A new Paragraph inserted after this Table.</returns>
        /// <example>
        /// Insert a new Paragraph after the first Table in this document.
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a Table into this document.
        ///     Table t = document.InsertTable(2, 2);
        ///
        ///     t.InsertParagraphAfterSelf("I was inserted after the previous Table.");
        ///
        ///     // Save all changes made to this new document.
        ///     document.Save();
        ///    }// Release this new document form memory.
        /// </code>
        /// </example>
        public override Paragraph InsertParagraphAfterSelf(string text)
        {
            return base.InsertParagraphAfterSelf(text);
        }
    }

    /// <summary>
    /// Represents a single row in a Table.
    /// </summary>
    public class Row : Container
    {
        /// <summary>
        /// A list of Cells in this Row.
        /// </summary>
        public List<Cell> Cells 
        { 
            get 
            {
                List<Cell> cells = 
                (
                    from c in Xml.Elements(XName.Get("tc", DocX.w.NamespaceName))
                    select new Cell(this, Document, c)
                ).ToList();

                return cells;
            } 
        }

        public override List<Paragraph> Paragraphs
        {
            get
            {
                List<Paragraph> paragraphs =
                (
                    from p in Xml.Descendants(DocX.w + "p")
                    select new Paragraph(Document, p, 0)
                ).ToList();

                foreach (Paragraph p in paragraphs)
                    p.PackagePart = table.mainPart;

                return paragraphs;
            }
        }

        internal Table table;
        internal Row(Table table, DocX document, XElement xml):base(document, xml)
        {
            this.table = table;
        }

        /// <summary>
        /// Height in pixels. // Added by Joel, refactored by Cathal.
        /// </summary>
        public double Height
        {
            get
            {
                /*
                * Get the trPr (table row properties) element for this Row,
                * null will be return if no such element exists.
                */
                XElement trPr = Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));

                // If trPr is null, this row contains no height information.
                if(trPr == null)
                    return double.NaN;

                /*
                 * Get the trHeight element for this Row,
                 * null will be return if no such element exists.
                 */
                XElement trHeight = trPr.Element(XName.Get("trHeight", DocX.w.NamespaceName));
               
                // If trHeight is null, this row contains no height information.
                if (trHeight == null)
                    return double.NaN;

                // Get the val attribute for this trHeight element.
                XAttribute val = trHeight.Attribute(XName.Get("val", DocX.w.NamespaceName));

                // If w is null, this cell contains no width information.
                if (val == null)
                    return double.NaN;

                // If val is not a double, something is wrong with this attributes value, so remove it and return double.NaN;
                double heightInWordUnits;
                if (!double.TryParse(val.Value, out heightInWordUnits))
                {
                    val.Remove();
                    return double.NaN;
                }

                // 15 "word units" in one pixel
                return (heightInWordUnits / 15);
            }

            set
            {
                /*
                 * Get the trPr (table row properties) element for this Row,
                 * null will be return if no such element exists.
                 */
                XElement trPr = Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
                if (trPr == null)
                {
                    Xml.SetElementValue(XName.Get("trPr", DocX.w.NamespaceName), string.Empty);
                    trPr = Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
                }

                /*
                 * Get the trHeight element for this Row,
                 * null will be return if no such element exists.
                 */
                XElement trHeight = trPr.Element(XName.Get("trHeight", DocX.w.NamespaceName));
                if (trHeight == null)
                {
                    trPr.SetElementValue(XName.Get("trHeight", DocX.w.NamespaceName), string.Empty);
                    trHeight = trPr.Element(XName.Get("trHeight", DocX.w.NamespaceName));
                }

                // The hRule attribute needs to be set to exact.
                trHeight.SetAttributeValue(XName.Get("hRule", DocX.w.NamespaceName), "exact");

                // 15 "word units" is equal to one pixel. 
                trHeight.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), (value * 15).ToString());
            }
        }

        /// <summary>
        /// Merge cells starting with startIndex and ending with endIndex.
        /// </summary>
        public void MergeCells(int startIndex, int endIndex)
        {
            // Check for valid start and end indexes.
            if (startIndex < 0 || endIndex <= startIndex || endIndex > Cells.Count + 1)
                throw new IndexOutOfRangeException();

            // The sum of all merged gridSpans.
            int gridSpanSum = 0;
            
            // Foreach each Cell between startIndex and endIndex inclusive.
            foreach (Cell c in Cells.Where((z, i) => i > startIndex && i <= endIndex))
            {
                XElement tcPr = c.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
                if (tcPr != null)
                {
                    XElement gridSpan = tcPr.Element(XName.Get("gridSpan", DocX.w.NamespaceName));
                    if (gridSpan != null)
                    {
                        XAttribute val = gridSpan.Attribute(XName.Get("val", DocX.w.NamespaceName));

                        int value = 0;
                        if (val != null)
                            if (int.TryParse(val.Value, out value))
                                gridSpanSum += value - 1;
                    }
                }

                // Add this cells Pragraph to the merge start Cell.
                Cells[startIndex].Xml.Add(c.Xml.Elements(XName.Get("p", DocX.w.NamespaceName)));
                
                // Remove this Cell.
                c.Xml.Remove();
            }

            /* 
             * Get the tcPr (table cell properties) element for the first cell in this merge,
             * null will be returned if no such element exists.
             */
            XElement start_tcPr = Cells[startIndex].Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
            if (start_tcPr == null)
            {
                Cells[startIndex].Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
                start_tcPr = Cells[startIndex].Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
            }

            /* 
             * Get the gridSpan element of this row,
             * null will be returned if no such element exists.
             */
            XElement start_gridSpan = start_tcPr.Element(XName.Get("gridSpan", DocX.w.NamespaceName));
            if (start_gridSpan == null)
            {
                start_tcPr.SetElementValue(XName.Get("gridSpan", DocX.w.NamespaceName), string.Empty);
                start_gridSpan = start_tcPr.Element(XName.Get("gridSpan", DocX.w.NamespaceName));
            }

            /* 
             * Get the val attribute of this row,
             * null will be returned if no such element exists.
             */
            XAttribute start_val = start_gridSpan.Attribute(XName.Get("val", DocX.w.NamespaceName));

            int start_value = 0;
            if (start_val != null)
                if (int.TryParse(start_val.Value, out start_value))
                    gridSpanSum += start_value - 1;

            // Set the val attribute to the number of merged cells.
            start_gridSpan.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), (gridSpanSum + (endIndex - startIndex + 1)).ToString());
        }
    }

    public class Cell:Container
    {
        internal Row row;
        internal Cell(Row row, DocX document, XElement xml):base(document, xml)
        {
            this.row = row;
        }

        public override List<Paragraph> Paragraphs
        {
            get
            {
                List<Paragraph> paragraphs = base.Paragraphs;

                foreach (Paragraph p in paragraphs)
                    p.PackagePart = row.table.mainPart;

                return paragraphs;
            }
        }

        public Color Shading
        {
            get
            {
                /*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
                XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));

                // If tcPr is null, this cell contains no Color information.
                if (tcPr == null)
                    return Color.White;

                /*
                 * Get the shd (table shade) element for this Cell,
                 * null will be return if no such element exists.
                 */
                XElement shd = tcPr.Element(XName.Get("shd", DocX.w.NamespaceName));

                // If shd is null, this cell contains no Color information.
                if (shd == null)
                    return Color.White;
             
                // Get the w attribute of the tcW element.
                XAttribute fill = shd.Attribute(XName.Get("fill", DocX.w.NamespaceName));

                // If fill is null, this cell contains no Color information.
                if (fill == null)
                    return Color.White;

               return ColorTranslator.FromHtml(string.Format("#{0}", fill.Value));
            }

            set
            {
                /*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
                XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
                if (tcPr == null)
                {
                    Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
                    tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
                }

                /*
                 * Get the shd (table shade) element for this Cell,
                 * null will be return if no such element exists.
                 */
                XElement shd = tcPr.Element(XName.Get("shd", DocX.w.NamespaceName));
                if (shd == null)
                {
                    tcPr.SetElementValue(XName.Get("shd", DocX.w.NamespaceName), string.Empty);
                    shd = tcPr.Element(XName.Get("shd", DocX.w.NamespaceName));
                }

                // The val attribute needs to be set to clear
                shd.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), "clear");

                // The color attribute needs to be set to auto
                shd.SetAttributeValue(XName.Get("color", DocX.w.NamespaceName), "auto");

                // The fill attribute needs to be set to the hex for this Color.
                shd.SetAttributeValue(XName.Get("fill", DocX.w.NamespaceName), value.ToHex());
            }
        }

        /// <summary>
        /// Width in pixels. // Added by Joel, refactored by Cathal
        /// </summary>
        public double Width
        {
            get
            {
                /*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
                XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));

                // If tcPr is null, this cell contains no width information.
                if (tcPr == null)
                    return double.NaN;

                /*
                 * Get the tcW (table cell width) element for this Cell,
                 * null will be return if no such element exists.
                 */
                XElement tcW = tcPr.Element(XName.Get("tcW", DocX.w.NamespaceName));

                // If tcW is null, this cell contains no width information.
                if (tcW == null)
                    return double.NaN;
             
                // Get the w attribute of the tcW element.
                XAttribute w = tcW.Attribute(XName.Get("w", DocX.w.NamespaceName));

                // If w is null, this cell contains no width information.
                if (w == null)
                    return double.NaN;

                // If w is not a double, something is wrong with this attributes value, so remove it and return double.NaN;
                double widthInWordUnits;
                if (!double.TryParse(w.Value, out widthInWordUnits))
                {
                    w.Remove();
                    return double.NaN;
                }

                // 15 "word units" is equal to one pixel.
                return (widthInWordUnits / 15);
            }

            set
            {
                /*
                 * Get the tcPr (table cell properties) element for this Cell,
                 * null will be return if no such element exists.
                 */
                XElement tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
                if (tcPr == null)
                {
                    Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
                    tcPr = Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
                }

                /*
                 * Get the tcW (table cell width) element for this Cell,
                 * null will be return if no such element exists.
                 */
                XElement tcW = tcPr.Element(XName.Get("tcW", DocX.w.NamespaceName));
                if (tcW == null)
                {
                    tcPr.SetElementValue(XName.Get("tcW", DocX.w.NamespaceName), string.Empty);
                    tcW = tcPr.Element(XName.Get("tcW", DocX.w.NamespaceName));
                }

                // The type attribute needs to be set to dxa which represents "twips" or twentieths of a point. In other words, 1/1440th of an inch.
                tcW.SetAttributeValue(XName.Get("type", DocX.w.NamespaceName), "dxa");

                // 15 "word units" is equal to one pixel. 
                tcW.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), (value * 15).ToString());
            }
        }
    }
}
