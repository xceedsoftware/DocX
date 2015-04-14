using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace Novacode
{
    /// <summary>
    /// Represents a table of contents in the document
    /// </summary>
    public class TableOfContents : DocXElement
    {
        #region TocBaseValues

        private const string HeaderStyle = "TOCHeading";
        private const int RightTabPos = 9350;
        #endregion

        private TableOfContents(DocX document, XElement xml, string headerStyle) : base(document, xml)
        {
            AssureUpdateField(document);
            AssureStyles(document, headerStyle);
        }        

        internal static TableOfContents CreateTableOfContents(DocX document, string title, TableOfContentsSwitches switches, string headerStyle = null, int lastIncludeLevel = 3, int? rightTabPos = null)
        {
            var reader = XmlReader.Create(new StringReader(string.Format(XmlTemplateBases.TocXmlBase, headerStyle ?? HeaderStyle, title, rightTabPos ?? RightTabPos, BuildSwitchString(switches, lastIncludeLevel))));
            var xml = XElement.Load(reader);
            return new TableOfContents(document, xml, headerStyle);
        }

        private void AssureUpdateField(DocX document)
        {
            if (document.settings.Descendants().Any(x => x.Name.Equals(DocX.w + "updateFields"))) return;
            
            var element = new XElement(XName.Get("updateFields", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", true));
            document.settings.Root.Add(element);
        }

        private void AssureStyles(DocX document, string headerStyle)
        {
            if (!HasStyle(document, headerStyle, "paragraph"))
            {
                var reader = XmlReader.Create(new StringReader(string.Format(XmlTemplateBases.TocHeadingStyleBase, headerStyle ?? HeaderStyle)));
                var xml = XElement.Load(reader);
                document.styles.Root.Add(xml);
            }
            if (!HasStyle(document, "TOC1", "paragraph"))
            {
                var reader = XmlReader.Create(new StringReader(string.Format(XmlTemplateBases.TocElementStyleBase, "TOC1", "toc 1")));
                var xml = XElement.Load(reader);
                document.styles.Root.Add(xml);
            }
            if (!HasStyle(document, "TOC2", "paragraph"))
            {
                var reader = XmlReader.Create(new StringReader(string.Format(XmlTemplateBases.TocElementStyleBase, "TOC2", "toc 2")));
                var xml = XElement.Load(reader);
                document.styles.Root.Add(xml);
            }
            if (!HasStyle(document, "TOC3", "paragraph"))
            {
                var reader = XmlReader.Create(new StringReader(string.Format(XmlTemplateBases.TocElementStyleBase, "TOC3", "toc 3")));
                var xml = XElement.Load(reader);
                document.styles.Root.Add(xml);
            }
            if (!HasStyle(document, "TOC4", "paragraph"))
            {
                var reader = XmlReader.Create(new StringReader(string.Format(XmlTemplateBases.TocElementStyleBase, "TOC4", "toc 4")));
                var xml = XElement.Load(reader);
                document.styles.Root.Add(xml);
            }
            if (!HasStyle(document, "Hyperlink", "character"))
            {
                var reader = XmlReader.Create(new StringReader(string.Format(XmlTemplateBases.TocHyperLinkStyleBase)));
                var xml = XElement.Load(reader);
                document.styles.Root.Add(xml);
            }
        }

        private bool HasStyle(DocX document, string value, string type)
        {
            return document.styles.Descendants().Any(x => x.Name.Equals(DocX.w + "style")&& (x.Attribute(DocX.w + "type") == null || x.Attribute(DocX.w + "type").Value.Equals(type)) && x.Attribute(DocX.w + "styleId") != null && x.Attribute(DocX.w + "styleId").Value.Equals(value));
        }

        private static string BuildSwitchString(TableOfContentsSwitches switches, int lastIncludeLevel)
        {
            var allSwitches = Enum.GetValues(typeof (TableOfContentsSwitches)).Cast<TableOfContentsSwitches>();
            var switchString = "TOC";
            foreach (var s in allSwitches.Where(s => s != TableOfContentsSwitches.None && switches.HasFlag(s)))
            {
                switchString += " " + s.EnumDescription();
                if (s == TableOfContentsSwitches.O)
                {
                    switchString += string.Format(" '{0}-{1}'", 1, lastIncludeLevel);
                }
            }

            return switchString;
        }

    }
}