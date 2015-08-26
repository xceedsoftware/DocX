using System;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace Novacode
{
    /// <summary>
    /// All DocX types are derived from DocXElement. 
    /// This class contains properties which every element of a DocX must contain.
    /// </summary>
    public abstract class DocXElement
    {
        internal PackagePart mainPart;
        public PackagePart PackagePart { get { return mainPart; } set { mainPart = value; } }

        /// <summary>
        /// This is the actual Xml that gives this element substance. 
        /// For example, a Paragraphs Xml might look something like the following
        /// <p>
        ///     <r>
        ///         <t>Hello World!</t>
        ///     </r>
        /// </p>
        /// </summary>
        
        public XElement Xml { get; set; }
        /// <summary>
        /// This is a reference to the DocX object that this element belongs to.
        /// Every DocX element is connected to a document.
        /// </summary>
        internal DocX Document { get; set; }
        /// <summary>
        /// Store both the document and xml so that they can be accessed by derived types.
        /// </summary>
        /// <param name="document">The document that this element belongs to.</param>
        /// <param name="xml">The Xml that gives this element substance</param>
        public DocXElement(DocX document, XElement xml)
        {
            this.Document = document;
            this.Xml = xml;
        }
    }

    /// <summary>
    /// This class provides functions for inserting new DocXElements before or after the current DocXElement.
    /// Only certain DocXElements can support these functions without creating invalid documents, at the moment these are Paragraphs and Table.
    /// </summary>
    public abstract class InsertBeforeOrAfter : DocXElement
    {
        public InsertBeforeOrAfter(DocX document, XElement xml) : base(document, xml) { }

        public virtual void InsertPageBreakBeforeSelf()
        {
            XElement p = new XElement
            (
                XName.Get("p", DocX.w.NamespaceName),
                    new XElement
                    (
                        XName.Get("r", DocX.w.NamespaceName),
                            new XElement
                            (
                                XName.Get("br", DocX.w.NamespaceName),
                                new XAttribute(XName.Get("type", DocX.w.NamespaceName), "page")
                            )
                    )
            );

            Xml.AddBeforeSelf(p);
        }

        public virtual void InsertPageBreakAfterSelf()
        {
            XElement p = new XElement
            (
                XName.Get("p", DocX.w.NamespaceName),
                    new XElement
                    (
                        XName.Get("r", DocX.w.NamespaceName),
                            new XElement
                            (
                                XName.Get("br", DocX.w.NamespaceName),
                                new XAttribute(XName.Get("type", DocX.w.NamespaceName), "page")
                            )
                    )
            );

            Xml.AddAfterSelf(p);
        }

        public virtual Paragraph InsertParagraphBeforeSelf(Paragraph p)
        {
            Xml.AddBeforeSelf(p.Xml);
            XElement newlyInserted = Xml.ElementsBeforeSelf().First();

            p.Xml = newlyInserted;

            return p;
        }

        public virtual Paragraph InsertParagraphAfterSelf(Paragraph p)
        {
            Xml.AddAfterSelf(p.Xml);
            XElement newlyInserted = Xml.ElementsAfterSelf().First();

            //Dmitchern
            if (this as Paragraph != null)
            {
                return new Paragraph(Document, newlyInserted, (this as Paragraph).endIndex);
            }
            else
            {
                p.Xml = newlyInserted; //IMPORTANT: I think we have return new paragraph in any case, but I dont know what to put as startIndex parameter into Paragraph constructor
                return p;
            }
        }

        public virtual Paragraph InsertParagraphBeforeSelf(string text)
        {
            return InsertParagraphBeforeSelf(text, false, new Formatting());
        }

        public virtual Paragraph InsertParagraphAfterSelf(string text)
        {
            return InsertParagraphAfterSelf(text, false, new Formatting());
        }

        public virtual Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges)
        {
            return InsertParagraphBeforeSelf(text, trackChanges, new Formatting());
        }

        public virtual Paragraph InsertParagraphAfterSelf(string text, bool trackChanges)
        {
            return InsertParagraphAfterSelf(text, trackChanges, new Formatting());
        }

        public virtual Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges, Formatting formatting)
        {
            XElement newParagraph = new XElement
            (
                XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName)), HelperFunctions.FormatInput(text, formatting.Xml)
            );

            if (trackChanges)
                newParagraph = Paragraph.CreateEdit(EditType.ins, DateTime.Now, newParagraph);

            Xml.AddBeforeSelf(newParagraph);
            XElement newlyInserted = Xml.ElementsBeforeSelf().Last();

            Paragraph p = new Paragraph(Document, newlyInserted, -1);

            return p;
        }

        public virtual Paragraph InsertParagraphAfterSelf(string text, bool trackChanges, Formatting formatting)
        {
            XElement newParagraph = new XElement
            (
                XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName)), HelperFunctions.FormatInput(text, formatting.Xml)
            );

            if (trackChanges)
                newParagraph = Paragraph.CreateEdit(EditType.ins, DateTime.Now, newParagraph);

            Xml.AddAfterSelf(newParagraph);
            XElement newlyInserted = Xml.ElementsAfterSelf().First();

            Paragraph p = new Paragraph(Document, newlyInserted, -1);

            return p;
        }

        public virtual Table InsertTableAfterSelf(int rowCount, int columnCount)
        {
            XElement newTable = HelperFunctions.CreateTable(rowCount, columnCount);
            Xml.AddAfterSelf(newTable);
            XElement newlyInserted = Xml.ElementsAfterSelf().First();

            return new Table(Document, newlyInserted) { mainPart = mainPart };
        }

        public virtual Table InsertTableAfterSelf(Table t)
        {
            Xml.AddAfterSelf(t.Xml);
            XElement newlyInserted = Xml.ElementsAfterSelf().First();
            //Dmitchern
            return new Table(Document, newlyInserted) { mainPart = mainPart }; //return new table, dont affect parameter table

            //t.Xml = newlyInserted;
            //return t;
        }

        public virtual Table InsertTableBeforeSelf(int rowCount, int columnCount)
        {
            XElement newTable = HelperFunctions.CreateTable(rowCount, columnCount);
            Xml.AddBeforeSelf(newTable);
            XElement newlyInserted = Xml.ElementsBeforeSelf().Last();

            return new Table(Document, newlyInserted) { mainPart = mainPart };
        }

        public virtual Table InsertTableBeforeSelf(Table t)
        {
            Xml.AddBeforeSelf(t.Xml);
            XElement newlyInserted = Xml.ElementsBeforeSelf().Last();

            //Dmitchern
            return new Table(Document, newlyInserted) { mainPart=mainPart}; //return new table, dont affect parameter table

            //t.Xml = newlyInserted;
            //return t;
        }
    }

    public static class XmlTemplateBases
    {
        #region TocXml
        public const string TocXmlBase = @"<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                  <w:sdtPr>
                    <w:docPartObj>
                      <w:docPartGallery w:val='Table of Contents'/>
                      <w:docPartUnique/>
                    </w:docPartObj>\
                  </w:sdtPr>
                  <w:sdtEndPr>
                    <w:rPr>
                      <w:rFonts w:asciiTheme='minorHAnsi' w:cstheme='minorBidi' w:eastAsiaTheme='minorHAnsi' w:hAnsiTheme='minorHAnsi'/>
                      <w:color w:val='auto'/>
                      <w:sz w:val='22'/>
                      <w:szCs w:val='22'/>
                      <w:lang w:eastAsia='en-US'/>
                    </w:rPr>
                  </w:sdtEndPr>
                  <w:sdtContent>
                    <w:p>
                      <w:pPr>
                        <w:pStyle w:val='{0}'/>
                      </w:pPr>
                      <w:r>
                        <w:t>{1}</w:t>
                      </w:r>
                    </w:p>
                    <w:p>
                      <w:pPr>
                        <w:pStyle w:val='TOC1'/>
                        <w:tabs>
                          <w:tab w:val='right' w:leader='dot' w:pos='{2}'/>
                        </w:tabs>
                        <w:rPr>
                          <w:noProof/>
                        </w:rPr>
                      </w:pPr>
                      <w:r>
                        <w:fldChar w:fldCharType='begin' w:dirty='true'/>
                      </w:r>
                      <w:r>
                        <w:instrText xml:space='preserve'> {3} </w:instrText>
                      </w:r>
                      <w:r>
                        <w:fldChar w:fldCharType='separate'/>
                      </w:r>
                    </w:p>
                    <w:p>
                      <w:r>
                        <w:rPr>
                          <w:b/>
                          <w:bCs/>
                          <w:noProof/>
                        </w:rPr>
                        <w:fldChar w:fldCharType='end'/>
                      </w:r>
                    </w:p>
                  </w:sdtContent>
                </w:sdt>
            ";
        public const string TocHeadingStyleBase = @"<w:style w:type='paragraph' w:styleId='{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
            <w:name w:val='TOC Heading'/>
            <w:basedOn w:val='Heading1'/>
            <w:next w:val='Normal'/>
            <w:uiPriority w:val='39'/>
            <w:semiHidden/>
            <w:unhideWhenUsed/>
            <w:qFormat/>
            <w:rsid w:val='00E67AA6'/>
            <w:pPr>
              <w:outlineLvl w:val='9'/>
            </w:pPr>
            <w:rPr>
              <w:lang w:eastAsia='nb-NO'/>
            </w:rPr>
          </w:style>
        ";
        public const string TocElementStyleBase = @"  <w:style w:type='paragraph' w:styleId='{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
            <w:name w:val='{1}' />
            <w:basedOn w:val='Normal' />
            <w:next w:val='Normal' />
            <w:autoRedefine />
            <w:uiPriority w:val='39' />
            <w:unhideWhenUsed />
            <w:pPr>
              <w:spacing w:after='100' />
              <w:ind w:left='440' />
            </w:pPr>
          </w:style>
        ";
        public const string TocHyperLinkStyleBase = @"  <w:style w:type='character' w:styleId='Hyperlink' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
            <w:name w:val='Hyperlink' />
            <w:basedOn w:val='Normal' />
            <w:uiPriority w:val='99' />
            <w:unhideWhenUsed />
            <w:rPr>
              <w:color w:val='0000FF' w:themeColor='hyperlink' />
              <w:u w:val='single' />
            </w:rPr>
          </w:style>
        ";
        #endregion
    }
}
