using System;
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

            return new Table(Document, newlyInserted);
        }

        public virtual Table InsertTableAfterSelf(Table t)
        {
            Xml.AddAfterSelf(t.Xml);
            XElement newlyInserted = Xml.ElementsAfterSelf().First();
            //Dmitchern
            return new Table(Document, newlyInserted); //return new table, dont affect parameter table

            //t.Xml = newlyInserted;
            //return t;
        }

        public virtual Table InsertTableBeforeSelf(int rowCount, int columnCount)
        {
            XElement newTable = HelperFunctions.CreateTable(rowCount, columnCount);
            Xml.AddBeforeSelf(newTable);
            XElement newlyInserted = Xml.ElementsBeforeSelf().Last();

            return new Table(Document, newlyInserted);
        }

        public virtual Table InsertTableBeforeSelf(Table t)
        {
            Xml.AddBeforeSelf(t.Xml);
            XElement newlyInserted = Xml.ElementsBeforeSelf().Last();

            //Dmitchern
            return new Table(Document, newlyInserted); //return new table, dont affect parameter table

            //t.Xml = newlyInserted;
            //return t;
        }
    }
}
