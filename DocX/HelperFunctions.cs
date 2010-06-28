using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Packaging;
using System.Xml.Linq;
using System.IO;
using System.Reflection;
using System.IO.Compression;
using System.Security.Principal;

namespace Novacode
{
    internal static class HelperFunctions
    {
        internal static PackagePart CreateOrGetSettingsPart(Package package)
        {
            PackagePart settingsPart;

            Uri settingsUri = new Uri("/word/settings.xml", UriKind.Relative);
            if (!package.PartExists(settingsUri))
            {
                settingsPart = package.CreatePart(settingsUri, "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml");

                PackagePart mainDocumentPart = package.GetParts().Where
                (
                    p => p.ContentType.Equals
                    (
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                        StringComparison.CurrentCultureIgnoreCase
                    )
                ).Single();

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
            PackagePart customPropertiesPart = document.package.CreatePart(new Uri("/docProps/custom.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.custom-properties+xml");

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
                customPropDoc.Save(tw, SaveOptions.DisableFormatting);

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
        /// If this document does not contain a /word/styles.xml add the default one generated by Microsoft Word.
        /// </summary>
        /// <param name="package"></param>
        /// <param name="mainDocumentPart"></param>
        /// <returns></returns>
        internal static XDocument AddDefaultStylesXml(Package package)
        {
            XDocument stylesDoc;
            // Create the main document part for this package
            PackagePart word_styles = package.CreatePart(new Uri("/word/styles.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");

            stylesDoc = HelperFunctions.DecompressXMLResource("Novacode.Resources.default_styles.xml.gz");

            // Save /word/styles.xml
            using (TextWriter tw = new StreamWriter(word_styles.GetStream(FileMode.Create, FileAccess.Write)))
                stylesDoc.Save(tw, SaveOptions.DisableFormatting);

            PackagePart mainDocumentPart = package.GetParts().Where
            (
                p => p.ContentType.Equals
                (
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                    StringComparison.CurrentCultureIgnoreCase
                )
            ).Single();

            mainDocumentPart.CreateRelationship(word_styles.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
            return stylesDoc;
        }

        static internal XElement CreateEdit(EditType t, DateTime edit_time, object content)
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

        internal static XElement CreateTable(int rowCount, int coloumnCount)
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
            for (int i = 0; i < coloumnCount; i++)
                tableGrid.Add(new XElement(XName.Get("gridCol", DocX.w.NamespaceName), new XAttribute(XName.Get("w", DocX.w.NamespaceName), "2310")));

            newTable.Add(tableGrid);

            for (int i = 0; i < rowCount; i++)
            {
                XElement row = new XElement(XName.Get("tr", DocX.w.NamespaceName));

                for (int j = 0; j < coloumnCount; j++)
                {
                    XElement cell =
                    new XElement
                    (
                        XName.Get("tc", DocX.w.NamespaceName),
                            new XElement(XName.Get("tcPr", DocX.w.NamespaceName),
                                new XElement(XName.Get("tcW", DocX.w.NamespaceName), new XAttribute(XName.Get("w", DocX.w.NamespaceName), "2310"), new XAttribute(XName.Get("type", DocX.w.NamespaceName), "dxa"))),
                            new XElement(XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName)))
                    );

                    row.Add(cell);
                }

                newTable.Add(row);
            }
            return newTable;
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

        static internal Paragraph GetFirstParagraphEffectedByInsert(DocX document, int index)
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

        static internal XElement[] SplitParagraph(Paragraph p, int index)
        {
            Run r = p.GetFirstRunEffectedByInsert(index);

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

        static internal void RebuildParagraphs(DocX document)
        {
            document.paragraphLookup.Clear();
            document.paragraphs.Clear();

            // Get the runs in this paragraph
            IEnumerable<XElement> paras = document.mainDoc.Descendants(XName.Get("p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"));

            int startIndex = 0;

            // Loop through each run in this paragraph
            foreach (XElement par in paras)
            {
                Paragraph xp = new Paragraph(document, par, startIndex);

                // Add to paragraph list
                document.paragraphs.Add(xp);

                // Only add runs which contain text
                if (Paragraph.GetElementTextLength(par) > 0)
                {
                    document.paragraphLookup.Add(xp.endIndex, xp);
                    startIndex = xp.endIndex;
                }
            }
        }

        /// <!-- 
        /// Bug found and fixed by trnilse. To see the change, 
        /// please compare this release to the previous release using TFS compare.
        /// -->
        static internal bool IsSameFile(Stream streamOne, Stream streamTwo)
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
    }
}
