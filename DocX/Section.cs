using System.Collections.Generic;
using System.IO.Packaging;
using System.Xml.Linq;

namespace Novacode
{
  public class Section : Container
  {

    public SectionBreakType SectionBreakType;

    internal Section(DocX document, XElement xml) : base(document, xml)
    {
    }

    public List<Paragraph> SectionParagraphs { get; set; }
  }
}