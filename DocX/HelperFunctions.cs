using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using System.Text;
using System.Xml.Linq;
using System.Xml;

namespace Novacode
{
    internal static class HelperFunctions
    {
        public const string DOCUMENT_DOCUMENTTYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
        public const string TEMPLATE_DOCUMENTTYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";

            public static bool IsNullOrWhiteSpace(this string value)
            {
                if (value == null) return true;
                return string.IsNullOrEmpty(value.Trim());
            }

        /// <summary>
        /// Checks whether 'toCheck' has all children that 'desired' has and values of 'val' attributes are the same
        /// </summary>
        /// <param name="desired"></param>
        /// <param name="toCheck"></param>
        /// <param name="fo">Matching options whether check if desired attributes are inder a, or a has exactly and only these attributes as b has.</param>
        /// <returns></returns>
        internal static bool ContainsEveryChildOf(XElement desired, XElement toCheck, MatchFormattingOptions fo)
        {
            foreach (XElement e in desired.Elements())
            {
                // If a formatting property has the same name and 'val' attribute's value, its considered to be equivalent.
                if (!toCheck.Elements(e.Name).Where(bElement => bElement.GetAttribute(XName.Get("val", DocX.w.NamespaceName)) == e.GetAttribute(XName.Get("val", DocX.w.NamespaceName))).Any())
                    return false;
            }

            // If the formatting has to be exact, no additionaly formatting must exist.
            if (fo == MatchFormattingOptions.ExactMatch)
                return desired.Elements().Count() == toCheck.Elements().Count();

            return true;
        }
        internal static void CreateRelsPackagePart(DocX Document, Uri uri)
        {
            PackagePart pp = Document.package.CreatePart(uri, "application/vnd.openxmlformats-package.relationships+xml", CompressionOption.Maximum);
            using (TextWriter tw = new StreamWriter(pp.GetStream()))
            {
                XDocument d = new XDocument
                (
                    new XDeclaration("1.0", "UTF-8", "yes"),
                    new XElement(XName.Get("Relationships", DocX.rel.NamespaceName))
                );
                var root = d.Root;
                d.Save(tw);
            }
        }

        internal static int GetSize(XElement Xml)
        {
            switch (Xml.Name.LocalName)
            {
                case "tab":
                    return 1;
                case "br":
                    return 1;
                case "t":
                    goto case "delText";
                case "delText":
                    return Xml.Value.Length;
                case "tr":
                    goto case "br";
                case "tc":
                    goto case "br";
                default:
                    return 0;
            }
        }

        internal static string GetText(XElement e)
        {
            StringBuilder sb = new StringBuilder();
            GetTextRecursive(e, ref sb);
            return sb.ToString();
        }

        internal static void GetTextRecursive(XElement Xml, ref StringBuilder sb)
        {
            sb.Append(ToText(Xml));

            if (Xml.HasElements)
                foreach (XElement e in Xml.Elements())
                    GetTextRecursive(e, ref sb);
        }

        internal static List<FormattedText> GetFormattedText(XElement e)
        {
            List<FormattedText> alist = new List<FormattedText>();
            GetFormattedTextRecursive(e, ref alist);
            return alist;
        }

        internal static void GetFormattedTextRecursive(XElement Xml, ref List<FormattedText> alist)
        {
            FormattedText ft = ToFormattedText(Xml);
            FormattedText last = null;

            if (ft != null)
            {
                if (alist.Count() > 0)
                    last = alist.Last();

                if (last != null && last.CompareTo(ft) == 0)
                {
                    // Update text of last entry.
                    last.text += ft.text;
                }
                else
                {
                    if (last != null)
                        ft.index = last.index + last.text.Length;

                    alist.Add(ft);
                }
            }

            if (Xml.HasElements)
                foreach (XElement e in Xml.Elements())
                    GetFormattedTextRecursive(e, ref alist);
        }

        internal static FormattedText ToFormattedText(XElement e)
        {
            // The text representation of e.
            String text = ToText(e);
            if (text == String.Empty)
                return null;

            // e is a w:t element, it must exist inside a w:r element, lets climb until we find it.
            while (!e.Name.Equals(XName.Get("r", DocX.w.NamespaceName)))
                e = e.Parent;

            // e is a w:r element, lets find the rPr element.
            XElement rPr = e.Element(XName.Get("rPr", DocX.w.NamespaceName));

            FormattedText ft = new FormattedText();
            ft.text = text;
            ft.index = 0;
            ft.formatting = null;

            // Return text with formatting.
            if (rPr != null)
                ft.formatting = Formatting.Parse(rPr);

            return ft;
        }

        internal static string ToText(XElement e)
        {
            switch (e.Name.LocalName)
            {
                case "tab":
                    return "\t";
                case "br":
                    return "\n";
                case "t":
                    goto case "delText";
                case "delText":
                    {
                        if (e.Parent != null && e.Parent.Name.LocalName == "r")
                        {
                            XElement run = e.Parent;
                            var rPr = run.Elements().FirstOrDefault(a => a.Name.LocalName == "rPr");
                            if (rPr != null)
                            {
                                var caps = rPr.Elements().FirstOrDefault(a => a.Name.LocalName == "caps");

                                if (caps != null)
                                    return e.Value.ToUpper();
                            }
                        }

                        return e.Value;
                    }
                case "tr":
                    goto case "br";
                case "tc":
                    goto case "tab";
                default: return "";
            }
        }

        internal static XElement CloneElement(XElement element)
        {
            return new XElement
            (
                element.Name,
                element.Attributes(),
                element.Nodes().Select
                (
                    n =>
                    {
                        XElement e = n as XElement;
                        if (e != null)
                            return CloneElement(e);
                        return n;
                    }
                )
            );
        }

        internal static PackagePart CreateOrGetSettingsPart(Package package)
        {
            PackagePart settingsPart;

            Uri settingsUri = new Uri("/word/settings.xml", UriKind.Relative);
            if (!package.PartExists(settingsUri))
            {
                settingsPart = package.CreatePart(settingsUri, "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml", CompressionOption.Maximum);

                PackagePart mainDocumentPart = package.GetParts().Single(p => p.ContentType.Equals(DOCUMENT_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase) ||
                                                                              p.ContentType.Equals(TEMPLATE_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase));

                mainDocumentPart.CreateRelationship(settingsUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings");

                XDocument settings = XDocument.Parse
                (@"<?xml version='1.0' encoding='utf-8' standalone='yes'?>
                <w:settings xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math' xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w10='urn:schemas-microsoft-com:office:word' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:sl='http://schemas.openxmlformats.org/schemaLibrary/2006/main'>
                  <w:zoom w:percent='100' />
                  <w:defaultTabStop w:val='720' />
                  <w:characterSpacingControl w:val='doNotCompress' />
                  <w:compat />
                  <w:rsids>
                    <w:rsidRoot w:val='00217F62' />
                    <w:rsid w:val='001915A3' />
                    <w:rsid w:val='00217F62' />
                    <w:rsid w:val='00A906D8' />
                    <w:rsid w:val='00AB5A74' />
                    <w:rsid w:val='00F071AE' />
                  </w:rsids>
                  <m:mathPr>
                    <m:mathFont m:val='Cambria Math' />
                    <m:brkBin m:val='before' />
                    <m:brkBinSub m:val='--' />
                    <m:smallFrac m:val='off' />
                    <m:dispDef />
                    <m:lMargin m:val='0' />
                    <m:rMargin m:val='0' />
                    <m:defJc m:val='centerGroup' />
                    <m:wrapIndent m:val='1440' />
                    <m:intLim m:val='subSup' />
                    <m:naryLim m:val='undOvr' />
                  </m:mathPr>
                  <w:themeFontLang w:val='en-IE' w:bidi='ar-SA' />
                  <w:clrSchemeMapping w:bg1='light1' w:t1='dark1' w:bg2='light2' w:t2='dark2' w:accent1='accent1' w:accent2='accent2' w:accent3='accent3' w:accent4='accent4' w:accent5='accent5' w:accent6='accent6' w:hyperlink='hyperlink' w:followedHyperlink='followedHyperlink' />
                  <w:shapeDefaults>
                    <o:shapedefaults v:ext='edit' spidmax='2050' />
                    <o:shapelayout v:ext='edit'>
                      <o:idmap v:ext='edit' data='1' />
                    </o:shapelayout>
                  </w:shapeDefaults>
                  <w:decimalSymbol w:val='.' />
                  <w:listSeparator w:val=',' />
                </w:settings>"
                );

                XElement themeFontLang = settings.Root.Element(XName.Get("themeFontLang", DocX.w.NamespaceName));
                themeFontLang.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), CultureInfo.CurrentCulture);

                // Save the settings document.
                using (TextWriter tw = new StreamWriter(settingsPart.GetStream()))
                    settings.Save(tw);
            }
            else
                settingsPart = package.GetPart(settingsUri);
            return settingsPart;
        }

        internal static void CreateCustomPropertiesPart(DocX document)
        {
            PackagePart customPropertiesPart = document.package.CreatePart(new Uri("/docProps/custom.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.custom-properties+xml", CompressionOption.Maximum);

            XDocument customPropDoc = new XDocument
            (
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement
                (
                    XName.Get("Properties", DocX.customPropertiesSchema.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "vt", DocX.customVTypesSchema)
                )
            );

            using (TextWriter tw = new StreamWriter(customPropertiesPart.GetStream(FileMode.Create, FileAccess.Write)))
                customPropDoc.Save(tw, SaveOptions.None);

            document.package.CreateRelationship(customPropertiesPart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties");
        }

        internal static XDocument DecompressXMLResource(string manifest_resource_name)
        {
            // XDocument to load the compressed Xml resource into.
            XDocument document;

            // Get a reference to the executing assembly.
            Assembly assembly = Assembly.GetExecutingAssembly();

            // Open a Stream to the embedded resource.
            Stream stream = assembly.GetManifestResourceStream(manifest_resource_name);

            // Decompress the embedded resource.
            using (GZipStream zip = new GZipStream(stream, CompressionMode.Decompress))
            {
                // Load this decompressed embedded resource into an XDocument using a TextReader.
                using (TextReader sr = new StreamReader(zip))
                {
                    document = XDocument.Load(sr);
                }
            }

            // Return the decompressed Xml as an XDocument.
            return document;
        }


        /// <summary>
        /// If this document does not contain a /word/numbering.xml add the default one generated by Microsoft Word 
        /// when the default bullet, numbered and multilevel lists are added to a blank document
        /// </summary>
        /// <param name="package"></param>
        /// <param name="mainDocumentPart"></param>
        /// <returns></returns>
        internal static XDocument AddDefaultNumberingXml(Package package)
        {
            XDocument numberingDoc;
            // Create the main document part for this package
            PackagePart wordNumbering = package.CreatePart(new Uri("/word/numbering.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", CompressionOption.Maximum);

            numberingDoc = DecompressXMLResource("Novacode.Resources.numbering.xml.gz");

            // Save /word/numbering.xml
            using (TextWriter tw = new StreamWriter(wordNumbering.GetStream(FileMode.Create, FileAccess.Write)))
                numberingDoc.Save(tw, SaveOptions.None);

            PackagePart mainDocumentPart = package.GetParts().Single(p => p.ContentType.Equals(DOCUMENT_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase) ||
                                                                          p.ContentType.Equals(TEMPLATE_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase));

            mainDocumentPart.CreateRelationship(wordNumbering.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering");
            return numberingDoc;
        }



        /// <summary>
        /// If this document does not contain a /word/styles.xml add the default one generated by Microsoft Word.
        /// </summary>
        /// <param name="package"></param>
        /// <param name="mainDocumentPart"></param>
        /// <returns></returns>
        internal static XDocument AddDefaultStylesXml(Package package)
        {
            XDocument stylesDoc;
            // Create the main document part for this package
            PackagePart word_styles = package.CreatePart(new Uri("/word/styles.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", CompressionOption.Maximum);

            stylesDoc = HelperFunctions.DecompressXMLResource("Novacode.Resources.default_styles.xml.gz");
            XElement lang = stylesDoc.Root.Element(XName.Get("docDefaults", DocX.w.NamespaceName)).Element(XName.Get("rPrDefault", DocX.w.NamespaceName)).Element(XName.Get("rPr", DocX.w.NamespaceName)).Element(XName.Get("lang", DocX.w.NamespaceName));
            lang.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), CultureInfo.CurrentCulture);

            // Save /word/styles.xml
            using (TextWriter tw = new StreamWriter(word_styles.GetStream(FileMode.Create, FileAccess.Write)))
                stylesDoc.Save(tw, SaveOptions.None);

            PackagePart mainDocumentPart = package.GetParts().Where
            (
                p => p.ContentType.Equals(DOCUMENT_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase)||p.ContentType.Equals(TEMPLATE_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase)
            ).Single();

            mainDocumentPart.CreateRelationship(word_styles.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
            return stylesDoc;
        }

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

                        for (int i = 0; i < ts.Count(); i++)
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

        internal static XElement CreateTable(int rowCount, int columnCount)
		{
			int[] columnWidths = new int[columnCount];
			for (int i = 0; i < columnCount; i++)
			{
				columnWidths[i] = 2310;
			}
			return CreateTable(rowCount, columnWidths);
		}

		internal static XElement CreateTable(int rowCount, int[] columnWidths)
        {
            XElement newTable =
            new XElement
            (
                XName.Get("tbl", DocX.w.NamespaceName),
                new XElement
                (
                    XName.Get("tblPr", DocX.w.NamespaceName),
                        new XElement(XName.Get("tblStyle", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "TableGrid")),
                        new XElement(XName.Get("tblW", DocX.w.NamespaceName), new XAttribute(XName.Get("w", DocX.w.NamespaceName), "5000"), new XAttribute(XName.Get("type", DocX.w.NamespaceName), "auto")),
                        new XElement(XName.Get("tblLook", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "04A0"))
                )
            );

            XElement tableGrid = new XElement(XName.Get("tblGrid", DocX.w.NamespaceName));
            for (int i = 0; i < columnWidths.Length; i++)
                tableGrid.Add(new XElement(XName.Get("gridCol", DocX.w.NamespaceName), new XAttribute(XName.Get("w", DocX.w.NamespaceName), XmlConvert.ToString(columnWidths[i]))));

            newTable.Add(tableGrid);

            for (int i = 0; i < rowCount; i++)
            {
                XElement row = new XElement(XName.Get("tr", DocX.w.NamespaceName));

                for (int j = 0; j < columnWidths.Length; j++)
                {
                    XElement cell = CreateTableCell();
                    row.Add(cell);
                }

                newTable.Add(row);
            }
            return newTable;
        }

        /// <summary>
        /// Create and return a cell of a table        
        /// </summary>        
        internal static XElement CreateTableCell()
        {
            return new XElement
                    (
                        XName.Get("tc", DocX.w.NamespaceName),
                            new XElement(XName.Get("tcPr", DocX.w.NamespaceName),
                            new XElement(XName.Get("tcW", DocX.w.NamespaceName),
                                    new XAttribute(XName.Get("w", DocX.w.NamespaceName), "2310"),
                                    new XAttribute(XName.Get("type", DocX.w.NamespaceName), "dxa"))),
                            new XElement(XName.Get("p", DocX.w.NamespaceName),
                                new XElement(XName.Get("pPr", DocX.w.NamespaceName)))
                    );
        }

        internal static List CreateItemInList(List list, string listText, int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = null, bool trackChanges = false)
        {
            if (list.NumId == 0)
            {
                list.CreateNewNumberingNumId(level, listType);
            }

            if (!string.IsNullOrEmpty(listText))
            {
                var newParagraphSection = new XElement
                    (
                    XName.Get("p", DocX.w.NamespaceName),
                    new XElement(XName.Get("pPr", DocX.w.NamespaceName),
                                 new XElement(XName.Get("numPr", DocX.w.NamespaceName),
                                              new XElement(XName.Get("ilvl", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", level)),
                                              new XElement(XName.Get("numId", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", list.NumId)))),
                    new XElement(XName.Get("r", DocX.w.NamespaceName), new XElement(XName.Get("t", DocX.w.NamespaceName), listText))
                    );

                if (trackChanges)
                    newParagraphSection = CreateEdit(EditType.ins, DateTime.Now, newParagraphSection);

                if (startNumber == null)
                {
                    list.AddItem(new Paragraph(list.Document, newParagraphSection, 0, ContainerType.Paragraph));
                }
                else
                {
                    list.AddItemWithStartValue(new Paragraph(list.Document, newParagraphSection, 0, ContainerType.Paragraph), (int)startNumber);
                }
            }

            return list;
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

        internal static Paragraph GetFirstParagraphEffectedByInsert(DocX document, int index)
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

        internal static List<XElement> FormatInput(string text, XElement rPr)
        {
            List<XElement> newRuns = new List<XElement>();
            XElement tabRun = new XElement(DocX.w + "tab");
            XElement breakRun = new XElement(DocX.w + "br");

            StringBuilder sb = new StringBuilder();

            if (string.IsNullOrEmpty(text))
            {
                return newRuns; //I dont wanna get an exception if text == null, so just return empy list
            }

            foreach (char c in text)
            {
                switch (c)
                {
                    case '\t':
                        if (sb.Length > 0)
                        {
                            XElement t = new XElement(DocX.w + "t", sb.ToString());
                            Novacode.Text.PreserveSpace(t);
                            newRuns.Add(new XElement(DocX.w + "r", rPr, t));
                            sb = new StringBuilder();
                        }
                        newRuns.Add(new XElement(DocX.w + "r", rPr, tabRun));
                        break;
                    case '\n':
                        if (sb.Length > 0)
                        {
                            XElement t = new XElement(DocX.w + "t", sb.ToString());
                            Novacode.Text.PreserveSpace(t);
                            newRuns.Add(new XElement(DocX.w + "r", rPr, t));
                            sb = new StringBuilder();
                        }
                        newRuns.Add(new XElement(DocX.w + "r", rPr, breakRun));
                        break;

                    default:
                        sb.Append(c);
                        break;
                }
            }

            if (sb.Length > 0)
            {
                XElement t = new XElement(DocX.w + "t", sb.ToString());
                Novacode.Text.PreserveSpace(t);
                newRuns.Add(new XElement(DocX.w + "r", rPr, t));
            }

            return newRuns;
        }

        internal static XElement[] SplitParagraph(Paragraph p, int index)
        {
            // In this case edit dosent really matter, you have a choice.
            Run r = p.GetFirstRunEffectedByEdit(index, EditType.ins);

            XElement[] split;
            XElement before, after;

            if (r.Xml.Parent.Name.LocalName == "ins")
            {
                split = p.SplitEdit(r.Xml.Parent, index, EditType.ins);
                before = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsBeforeSelf(), split[0]);
                after = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsAfterSelf(), split[1]);
            }
            else if (r.Xml.Parent.Name.LocalName == "del")
            {
                split = p.SplitEdit(r.Xml.Parent, index, EditType.del);

                before = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsBeforeSelf(), split[0]);
                after = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsAfterSelf(), split[1]);
            }
            else
            {
                split = Run.SplitRun(r, index);

                before = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.ElementsBeforeSelf(), split[0]);
                after = new XElement(p.Xml.Name, p.Xml.Attributes(), split[1], r.Xml.ElementsAfterSelf());
            }

            if (before.Elements().Count() == 0)
                before = null;

            if (after.Elements().Count() == 0)
                after = null;

            return new XElement[] { before, after };
        }

        /// <!-- 
        /// Bug found and fixed by trnilse. To see the change, 
        /// please compare this release to the previous release using TFS compare.
        /// -->
        internal static bool IsSameFile(Stream streamOne, Stream streamTwo)
        {
            int file1byte, file2byte;

            if (streamOne.Length != streamOne.Length)
            {
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

            // Return the success of the comparison. "file1byte" is 
            // equal to "file2byte" at this point only if the files are 
            // the same.

            streamOne.Position = 0;
            streamTwo.Position = 0;

            return ((file1byte - file2byte) == 0);
        }

      internal static UnderlineStyle GetUnderlineStyle(string underlineStyle)
      {
        switch (underlineStyle)
        {
          case "single":
            return UnderlineStyle.singleLine;
          case "double": 
            return UnderlineStyle.doubleLine;
          case "thick":
            return UnderlineStyle.thick;
          case "dotted":
            return UnderlineStyle.dotted;
          case "dottedHeavy":
            return UnderlineStyle.dottedHeavy;
          case "dash":
            return UnderlineStyle.dash;
          case "dashedHeavy":
            return UnderlineStyle.dashedHeavy;
          case "dashLong":
            return UnderlineStyle.dashLong;
          case "dashLongHeavy":
            return UnderlineStyle.dashLongHeavy;
          case "dotDash":
            return UnderlineStyle.dotDash;
          case "dashDotHeavy":
            return UnderlineStyle.dashDotHeavy;
          case "dotDotDash":
            return UnderlineStyle.dotDotDash;
          case "dashDotDotHeavy":
            return UnderlineStyle.dashDotDotHeavy;
          case "wave":
            return UnderlineStyle.wave;
          case "wavyHeavy":
            return UnderlineStyle.wavyHeavy;
          case "wavyDouble":
            return UnderlineStyle.wavyDouble;
          case "words":
            return UnderlineStyle.words;
          default: 
            return UnderlineStyle.none;
        }
      }



    }
}
