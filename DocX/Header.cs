using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.IO.Packaging;
using System.Collections.ObjectModel;

namespace Novacode
{
    public class Header : Container
    {
        public bool PageNumbers
        {
            get
            {
                return false;
            }

            set
            {
                XElement e = XElement.Parse
                (@"<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                    <w:sdtPr>
                      <w:id w:val='157571950' />
                      <w:docPartObj>
                        <w:docPartGallery w:val='Page Numbers (Top of Page)' />
                        <w:docPartUnique />
                      </w:docPartObj>
                    </w:sdtPr>
                    <w:sdtContent>
                      <w:p w:rsidR='008D2BFB' w:rsidRDefault='008D2BFB'>
                        <w:pPr>
                          <w:pStyle w:val='Header' />
                          <w:jc w:val='center' />
                        </w:pPr>
                        <w:fldSimple w:instr=' PAGE \* MERGEFORMAT'>
                          <w:r>
                            <w:rPr>
                              <w:noProof />
                            </w:rPr>
                            <w:t>1</w:t>
                          </w:r>
                        </w:fldSimple>
                      </w:p>
                    </w:sdtContent>
                  </w:sdt>"
               );
               
               Xml.AddFirst(e);
               
               PageNumberParagraph = new Paragraph(Document, e.Descendants(XName.Get("p", DocX.w.NamespaceName)).SingleOrDefault(), 0); 
            }
        }

        public Paragraph PageNumberParagraph;

        internal PackagePart mainPart;
        internal Header(DocX document, XElement xml, PackagePart mainPart):base(document, xml)
        {
            this.mainPart = mainPart;
        }

        public override Paragraph InsertParagraph()
        {
            Paragraph p = base.InsertParagraph();
            p.PackagePart = mainPart;
            return p;
        }

        public override Paragraph InsertParagraph(int index, string text, bool trackChanges)
        {
            Paragraph p = base.InsertParagraph(index, text, trackChanges);
            p.PackagePart = mainPart;
            return p;
        }

        public override Paragraph InsertParagraph(Paragraph p)
        {
            p.PackagePart = mainPart;
            return base.InsertParagraph(p);
        }

        public override Paragraph InsertParagraph(int index, Paragraph p)
        {
            p.PackagePart = mainPart;
            return base.InsertParagraph(index, p);
        }

        public override Paragraph InsertParagraph(int index, string text, bool trackChanges, Formatting formatting)
        {
            Paragraph p = base.InsertParagraph(index, text, trackChanges, formatting);
            p.PackagePart = mainPart;
            return p;
        }

        public override Paragraph InsertParagraph(string text)
        {
            Paragraph p = base.InsertParagraph(text);
            p.PackagePart = mainPart;
            return p;
        }

        public override Paragraph InsertParagraph(string text, bool trackChanges)
        {
            Paragraph p = base.InsertParagraph(text, trackChanges);
            p.PackagePart = mainPart;
            return p;
        }

        public override Paragraph InsertParagraph(string text, bool trackChanges, Formatting formatting)
        {
            Paragraph p = base.InsertParagraph(text, trackChanges, formatting);
            p.PackagePart = mainPart;

            return p;
        }

        public override Paragraph InsertEquation(String equation)
        {
            Paragraph p = base.InsertEquation(equation);
            p.PackagePart = mainPart;
            return p;
        }


        public override ReadOnlyCollection<Paragraph> Paragraphs
        {
            get
            {
                ReadOnlyCollection<Paragraph> l = base.Paragraphs;
                foreach (var paragraph in l)
                {
                    paragraph.mainPart = mainPart;
                }
                return l;
            }
        }

        public override List<Table> Tables
        {
            get
            {
                List<Table> l = base.Tables;
                l.ForEach(x => x.mainPart = mainPart);
                return l;
            }
        }

        public List<Image> Images
        {
            get
            {
                PackageRelationshipCollection imageRelationships = mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                if (imageRelationships.Count() > 0)
                {
                    return
                    (
                        from i in imageRelationships
                        select new Image(Document, i)
                    ).ToList();
                }

                return new List<Image>();
            }
        }
        public new Table InsertTable(int rowCount, int columnCount)
        {
            if (rowCount < 1 || columnCount < 1)
                throw new ArgumentOutOfRangeException("Row and Column count must be greater than zero.");

            Table t = base.InsertTable(rowCount, columnCount);
            t.mainPart = mainPart;
            return t;
        }
        public new Table InsertTable(int index, Table t)
        {
            Table t2 = base.InsertTable(index, t);
            t2.mainPart = mainPart;
            return t2;
        }
        public new Table InsertTable(Table t)
        {
            t = base.InsertTable(t);
            t.mainPart = mainPart;
            return t;
        }
        public new Table InsertTable(int index, int rowCount, int columnCount)
        {
            if (rowCount < 1 || columnCount < 1)
                throw new ArgumentOutOfRangeException("Row and Column count must be greater than zero.");

            Table t = base.InsertTable(index, rowCount, columnCount);
            t.mainPart = mainPart;
            return t;
        }

    }
}
