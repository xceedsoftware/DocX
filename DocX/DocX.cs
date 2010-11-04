using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml;
using System.IO;
using System.Text.RegularExpressions;
using System.IO.Packaging;
using System.Security.Principal;
using System.Reflection;
using System.IO.Compression;

namespace Novacode
{
    /// <summary>
    /// Represents a document.
    /// </summary>
    public class DocX: Container, IDisposable
    {
        #region Namespaces
        static internal XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        static internal XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        static internal XNamespace customPropertiesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        static internal XNamespace customVTypesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        #endregion

        /// <summary>
        /// Returns a collection of Headers in this Document.
        /// A document typically contains three Headers.
        /// A default one (odd), one for the first page and one for even pages.
        /// </summary>
        /// <example>
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///    // Add header support to this document.
        ///    document.AddHeaders();
        ///
        ///    // Get a collection of all headers in this document.
        ///    Headers headers = document.Headers;
        ///
        ///    // The header used for the first page of this document.
        ///    Header first = headers.first;
        ///
        ///    // The header used for odd pages of this document.
        ///    Header odd = headers.odd;
        ///
        ///    // The header used for even pages of this document.
        ///    Header even = headers.even;
        /// }
        /// </code>
        /// </example>
        public Headers Headers 
        {
            get 
            {
                return headers;
            } 
        }
        private Headers headers;

        /// <summary>
        /// Returns a collection of Footers in this Document.
        /// A document typically contains three Footers.
        /// A default one (odd), one for the first page and one for even pages.
        /// </summary>
        /// <example>
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///    // Add footer support to this document.
        ///    document.AddFooters();
        ///
        ///    // Get a collection of all footers in this document.
        ///    Footers footers = document.Footers;
        ///
        ///    // The footer used for the first page of this document.
        ///    Footer first = footers.first;
        ///
        ///    // The footer used for odd pages of this document.
        ///    Footer odd = footers.odd;
        ///
        ///    // The footer used for even pages of this document.
        ///    Footer even = footers.even;
        /// }
        /// </code>
        /// </example>
        public Footers Footers
        {
            get
            {
                return footers;
            }
        }

        private Footers footers;

        /// <summary>
        /// Should the Document use different Headers and Footers for odd and even pages?
        /// </summary>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Add header support to this document.
        ///     document.AddHeaders();
        ///
        ///     // Get a collection of all headers in this document.
        ///     Headers headers = document.Headers;
        ///
        ///     // The header used for odd pages of this document.
        ///     Header odd = headers.odd;
        ///
        ///     // The header used for even pages of this document.
        ///     Header even = headers.even;
        ///
        ///     // Force the document to use a different header for odd and even pages.
        ///     document.DifferentOddAndEvenPages = true;
        ///
        ///     // Content can be added to the Headers in the same manor that it would be added to the main document.
        ///     Paragraph p1 = odd.InsertParagraph();
        ///     p1.Append("This is the odd pages header.");
        ///     
        ///     Paragraph p2 = even.InsertParagraph();
        ///     p2.Append("This is the even pages header.");
        ///
        ///     // Save all changes to this document.
        ///     document.Save();    
        /// }// Release this document from memory.
        /// </example>
        public bool DifferentOddAndEvenPages
        {
            get
            {
                XDocument settings;
                using (TextReader tr = new StreamReader(settingsPart.GetStream()))
                    settings = XDocument.Load(tr);

                XElement evenAndOddHeaders = settings.Root.Element(w + "evenAndOddHeaders");

                return evenAndOddHeaders != null;
            }

            set
            {
                XDocument settings;
                using (TextReader tr = new StreamReader(settingsPart.GetStream()))
                    settings = XDocument.Load(tr);

                XElement evenAndOddHeaders = settings.Root.Element(w + "evenAndOddHeaders");
                if (evenAndOddHeaders == null)
                {
                    if (value)
                        settings.Root.AddFirst(new XElement(w + "evenAndOddHeaders"));
                }

                else
                {
                    if (!value)
                        evenAndOddHeaders.Remove();
                }

                using (TextWriter tw = new StreamWriter(settingsPart.GetStream()))
                    settings.Save(tw);
            }
        }

        /// <summary>
        /// Should the Document use an independent Header and Footer for the first page?
        /// </summary>
        /// <example>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Add header support to this document.
        ///     document.AddHeaders();
        ///
        ///     // The header used for the first page of this document.
        ///     Header first = document.Headers.first;
        ///
        ///     // Force the document to use a different header for first page.
        ///     document.DifferentFirstPage = true;
        ///     
        ///     // Content can be added to the Headers in the same manor that it would be added to the main document.
        ///     Paragraph p = first.InsertParagraph();
        ///     p.Append("This is the first pages header.");
        ///
        ///     // Save all changes to this document.
        ///     document.Save();    
        /// }// Release this document from memory.
        /// </example>
        public bool DifferentFirstPage
        {
            get
            {
                XElement body = mainDoc.Root.Element(w + "body");
                XElement sectPr = body.Element(w + "sectPr");
                
                if (sectPr != null)
                {
                    XElement titlePg = sectPr.Element(w + "titlePg");
                    if (titlePg != null)
                        return true;
                }

                return false;
            }

            set
            {
                XElement body = mainDoc.Root.Element(w + "body");
                XElement sectPr = null;
                XElement titlePg = null;

                if (sectPr == null)
                    body.Add(new XElement(w + "sectPr", string.Empty));

                sectPr = body.Element(w + "sectPr");
                
                titlePg = sectPr.Element(w + "titlePg");
                if (titlePg == null)
                {
                    if (value)
                        sectPr.Add(new XElement(w + "titlePg", string.Empty));
                }

                else
                {
                    if (!value)
                        titlePg.Remove();
                }
            }
        }

        private Header GetHeaderByType(string type)
        {
            return (Header)GetHeaderOrFooterByType(type, true);
        }

        private Footer GetFooterByType(string type)
        {
            return (Footer)GetHeaderOrFooterByType(type, false);
        }

        private object GetHeaderOrFooterByType(string type, bool b)
        {
            string reference = "footerReference";
            if (b)
                reference = "headerReference";

            string Id =
            (
                from e in mainDoc.Descendants(XName.Get("body", DocX.w.NamespaceName)).Descendants()
                where (e.Name.LocalName == reference) && (e.Attribute(w + "type").Value == type)
                select e.Attribute(r + "id").Value
            ).FirstOrDefault();

            if (Id != null)
            {
                Uri partUri = mainPart.GetRelationship(Id).TargetUri;
                if (!partUri.OriginalString.StartsWith("/word/"))
                    partUri = new Uri("/word/" + partUri.OriginalString, UriKind.Relative);

                PackagePart part = package.GetPart(partUri);
                XDocument doc;
                using (TextReader tr = new StreamReader(part.GetStream()))
                {
                    doc = XDocument.Load(tr);
                    if(b)
                        return new Header(this, doc.Element(w + "hdr"), part);
                    else
                        return new Footer(this, doc.Element(w + "ftr"), part);
                }
            }

            return null;
        }

        // Get the word\document.xml part
        internal PackagePart mainPart;

        // Get the word\settings.xml part
        internal PackagePart settingsPart;

        #region Internal variables defined foreach DocX object
        // Object representation of the .docx
        internal Package package;
        // The mainDocument is loaded into a XDocument object for easy querying and editing
        internal XDocument mainDoc;
        internal XDocument header1;
        internal XDocument header2;
        internal XDocument header3;

        // A lookup for the Paragraphs in this document.
        internal Dictionary<int, Paragraph> paragraphLookup = new Dictionary<int, Paragraph>();
        // Every document is stored in a MemoryStream, all edits made to a document are done in memory.
        internal MemoryStream memoryStream;
        // The filename that this document was loaded from
        internal string filename;
        // The stream that this document was loaded from
        internal Stream stream;
        #endregion

        internal DocX(DocX document, XElement xml): base(document, xml)
        {      
        }

        /// <summary>
        /// Returns a list of Images in this document.
        /// </summary>
        /// <example>
        /// Get the unique Id of every Image in this document.
        /// <code>
        /// // Load a document.
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        ///
        /// // Loop through each Image in this document.
        /// foreach (Novacode.Image i in document.Images)
        /// {
        ///     // Get the unique Id which identifies this Image.
        ///     string uniqueId = i.Id;
        /// }
        ///
        /// </code>
        /// </example>
        /// <seealso cref="AddImage(string)"/>
        /// <seealso cref="AddImage(Stream)"/>
        /// <seealso cref="Paragraph.Pictures"/>
        /// <seealso cref="Paragraph.InsertPicture"/>
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
                        select new Image(this, i)
                    ).ToList();
                }

                return new List<Image>();
            }
        }

        /// <summary>
        /// Returns a list of custom properties in this document.
        /// </summary>
        /// <example>
        /// Method 1: Get the name, type and value of each CustomProperty in this document.
        /// <code>
        /// // Load Example.docx
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        ///
        /// /*
        ///  * No two custom properties can have the same name,
        ///  * so a Dictionary is the perfect data structure to store them in.
        ///  * Each custom property can be accessed using its name.
        ///  */
        /// foreach (string name in document.CustomProperties.Keys)
        /// {
        ///     // Grab a custom property using its name.
        ///     CustomProperty cp = document.CustomProperties[name];
        ///
        ///     // Write this custom properties details to Console.
        ///     Console.WriteLine(string.Format("Name: '{0}', Value: {1}", cp.Name, cp.Value));
        /// }
        ///
        /// Console.WriteLine("Press any key...");
        ///
        /// // Wait for the user to press a key before closing the Console.
        /// Console.ReadKey();
        /// </code>
        /// </example>
        /// <example>
        /// Method 2: Get the name, type and value of each CustomProperty in this document.
        /// <code>
        /// // Load Example.docx
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        /// 
        /// /*
        ///  * No two custom properties can have the same name,
        ///  * so a Dictionary is the perfect data structure to store them in.
        ///  * The values of this Dictionary are CustomProperties.
        ///  */
        /// foreach (CustomProperty cp in document.CustomProperties.Values)
        /// {
        ///     // Write this custom properties details to Console.
        ///     Console.WriteLine(string.Format("Name: '{0}', Value: {1}", cp.Name, cp.Value));
        /// }
        ///
        /// Console.WriteLine("Press any key...");
        ///
        /// // Wait for the user to press a key before closing the Console.
        /// Console.ReadKey();
        /// </code>
        /// </example>
        /// <seealso cref="AddCustomProperty"/>
        public Dictionary<string, CustomProperty> CustomProperties
        {
            get 
            {
                if (package.PartExists(new Uri("/docProps/custom.xml", UriKind.Relative)))
                {
                    PackagePart docProps_custom = package.GetPart(new Uri("/docProps/custom.xml", UriKind.Relative));
                    XDocument customPropDoc;
                    using (TextReader tr = new StreamReader(docProps_custom.GetStream(FileMode.Open, FileAccess.Read)))
                        customPropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);

                    // Get all of the custom properties in this document
                    return
                    (
                        from p in customPropDoc.Descendants(XName.Get("property", customPropertiesSchema.NamespaceName))
                        let Name = p.Attribute(XName.Get("name")).Value
                        let Type = p.Descendants().Single().Name.LocalName
                        let Value = p.Descendants().Single().Value
                        select new CustomProperty(Name, Type, Value)
                    ).ToDictionary(p => p.Name, StringComparer.CurrentCultureIgnoreCase);
                }

                return new Dictionary<string, CustomProperty>();
            }
        }

      ///<summary>
      /// Returns the list of document core properties with corresponding values.
      ///</summary>
      public Dictionary<string, string> CoreProperties
      {
        get
        {
          if (package.PartExists(new Uri("/docProps/core.xml", UriKind.Relative)))
          {
            PackagePart docProps_Core = package.GetPart(new Uri("/docProps/core.xml", UriKind.Relative));
            XDocument corePropDoc;
            using (TextReader tr = new StreamReader(docProps_Core.GetStream(FileMode.Open, FileAccess.Read)))
              corePropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);

            // Get all of the core properties in this document
            return (from docProperty in corePropDoc.Root.Elements()
                    select
                      new KeyValuePair<string, string>(
                      string.Format(
                        "{0}:{1}",
                        corePropDoc.Root.GetPrefixOfNamespace(docProperty.Name.Namespace),
                        docProperty.Name.LocalName),
                      docProperty.Value)).ToDictionary(p => p.Key, v => v.Value);
          }

          return new Dictionary<string, string>();
        }
      }

        /// <summary>
        /// Get the Text of this document.
        /// </summary>
        /// <example>
        /// Write to Console the Text from this document.
        /// <code>
        /// // Load a document
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        ///
        /// // Get the text of this document.
        /// string text = document.Text;
        ///
        /// // Write the text of this document to Console.
        /// Console.Write(text);
        ///
        /// // Wait for the user to press a key before closing the console window.
        /// Console.ReadKey();
        /// </code>
        /// </example>
        public string Text
        {
            get
            {
                StringBuilder sb = new StringBuilder();
                HelperFunctions.GetText(Xml, sb);
                return sb.ToString();
            }
        }

        internal string GetCollectiveText(List<PackagePart> list)
        {
            string text = string.Empty;

            foreach (var hp in list)
            {
                using (TextReader tr = new StreamReader(hp.GetStream()))
                {
                    XDocument d = XDocument.Load(tr);

                    StringBuilder sb = new StringBuilder();

                    // Loop through each text item in this run
                    foreach (XElement descendant in d.Descendants())
                    {
                        switch (descendant.Name.LocalName)
                        {
                            case "tab":
                                sb.Append("\t");
                                break;
                            case "br":
                                sb.Append("\n");
                                break;
                            case "t":
                                goto case "delText";
                            case "delText":
                                sb.Append(descendant.Value);
                                break;
                            default: break;
                        }
                    }

                    text += "\n" + sb.ToString();
                }
            }

            return text;
        }

        /// <summary>
        /// Insert the contents of another document at the end of this document. 
        /// </summary>
        /// <param name="document">The document to insert at the end of this document.</param>
        /// <example>
        /// Create a new document and insert an old document into it.
        /// <code>
        /// // Create a new document.
        /// using (DocX newDocument = DocX.Create(@"NewDocument.docx"))
        /// {
        ///     // Load an old document.
        ///     using (DocX oldDocument = DocX.Load(@"OldDocument.docx"))
        ///     {
        ///         // Insert the old document into the new document.
        ///         newDocument.InsertDocument(oldDocument);
        ///
        ///         // Save the new document.
        ///         newDocument.Save();
        ///     }// Release the old document from memory.
        /// }// Release the new document from memory.
        /// </code>
        /// <remarks>
        /// If the document being inserted contains Images, CustomProperties and or custom styles, these will be correctly inserted into the new document. In the case of Images, new ID's are generated for the Images being inserted to avoid ID conflicts. CustomProperties with the same name will be ignored not replaced.
        /// </remarks>
        /// </example>
        public void InsertDocument(DocX document)
        {
            #region /word/document.xml
            // Get the external elements that are going to be inserted.
            IEnumerable<XElement> external_elements = document.mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName)).Elements();

            // Get the body element of the internal document.
            XElement internal_body = mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName));

            // Insert the elements
            internal_body.Add(external_elements);

            // A moment of genius
            int count = external_elements.Count();
            external_elements = internal_body.Elements().Reverse().TakeWhile((i, j) => j < count);
            #endregion

            #region /word/styles.xml
            Uri word_styles_Uri = new Uri("/word/styles.xml", UriKind.Relative);

            // If the external document has a styles.xml, we need to insert its elements into the internal documents styles.xml.
            if (document.package.PartExists(word_styles_Uri))
            {
                // Load the external documents styles.xml into memory.
                XDocument external_word_styles;
                using (TextReader tr = new StreamReader(document.package.GetPart(word_styles_Uri).GetStream()))
                    external_word_styles = XDocument.Load(tr);

                // If the internal document contains no /word/styles.xml create one.
                if (!package.PartExists(word_styles_Uri))
                    HelperFunctions.AddDefaultStylesXml(package);

                // Load the internal documents styles.xml into memory.
                XDocument internal_word_styles;
                using (TextReader tr = new StreamReader(package.GetPart(word_styles_Uri).GetStream()))
                    internal_word_styles = XDocument.Load(tr);

                // Create a list of internal and external style elements for easy iteration.
                var internal_style_list = internal_word_styles.Root.Elements(XName.Get("style", DocX.w.NamespaceName));
                var external_style_list = external_word_styles.Root.Elements(XName.Get("style", DocX.w.NamespaceName));
                
                // Loop through the external style elements
                foreach (XElement style in external_style_list)
                {
                    // If the internal styles document does not contain this element, add it.
                    if (!internal_style_list.Contains(style))
                        internal_word_styles.Root.Add(style);
                }

                // Save the internal styles document.
                using (TextWriter tw = new StreamWriter(package.GetPart(word_styles_Uri).GetStream()))
                    internal_word_styles.Save(tw);
            }
            #endregion

            #region Images
            PackagePart internal_word_document = mainPart;
            PackagePart external_word_document = document.mainPart;

            // Get all Image relationships in the external document.
            var external_image_rels = external_word_document.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");

            // Get all Image relationships in the internal document.
            var internal_image_rels = internal_word_document.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");

            var internal_image_parts = internal_image_rels.Select(ir => package.GetParts().Where(p => p.Uri.ToString().EndsWith(ir.TargetUri.ToString())).First());

            int max = 0;
            var values =
            (
                from ip in internal_image_parts
                let Name = Path.GetFileNameWithoutExtension(ip.Uri.ToString())
                let Number = Regex.Match(Name, @"\d+$").Value
                select Number != string.Empty ? int.Parse(Number) : 0
            );

            if (values.Count() > 0)
                max = Math.Max(max, values.Max());

            // Foreach external image relationship
            foreach (var rel in external_image_rels)
            {
                string uri_string = rel.TargetUri.ToString();
                if (!uri_string.StartsWith("/"))
                    uri_string = "/" + uri_string;
                
                PackagePart external_image_part = rel.Package.GetPart(new Uri("/word" + uri_string, UriKind.RelativeOrAbsolute));
                PackagePart internal_image_part = package.CreatePart(new Uri(string.Format("/word/media/image{0}.jpeg", max + 1), UriKind.RelativeOrAbsolute), System.Net.Mime.MediaTypeNames.Image.Jpeg);

                PackageRelationship pr = internal_word_document.CreateRelationship(internal_image_part.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                
                var query = from e in external_elements.DescendantsAndSelf()
                            let embed = e.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))
                            where embed != null && embed.Value == rel.Id
                            select embed;

                foreach (XAttribute a in query)
                    a.Value = pr.Id;

                using (Stream stream = internal_image_part.GetStream(FileMode.Create, FileAccess.Write))
                {
                    using (Stream s = external_image_part.GetStream())
                    {
                        byte[] bytes = new byte[s.Length];
                        s.Read(bytes, 0, (int)s.Length);
                        stream.Write(bytes, 0, (int)s.Length);
                    }
                }

                max++;
            }
            #endregion

            #region CustomProperties
            
            // Check if the external document contains custom properties.
            if (document.package.PartExists(new Uri("/docProps/custom.xml", UriKind.Relative)))
            {
                PackagePart external_docProps_custom = document.package.GetPart(new Uri("/docProps/custom.xml", UriKind.Relative));
                XDocument external_customPropDoc;
                using (TextReader tr = new StreamReader(external_docProps_custom.GetStream(FileMode.Open, FileAccess.Read)))
                    external_customPropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);

                // Get all of the custom properties in this document.
                IEnumerable<XElement> external_customProperties =
                (
                    from cp in external_customPropDoc.Descendants(XName.Get("property", customPropertiesSchema.NamespaceName))
                    select cp
                );

                // If the internal document does not contain a customFilePropertyPart, create one.
                if (!package.PartExists(new Uri("/docProps/custom.xml", UriKind.Relative)))
                    HelperFunctions.CreateCustomPropertiesPart(this);

                
                PackagePart internal_docProps_custom = package.GetPart(new Uri("/docProps/custom.xml", UriKind.Relative));
                XDocument internal_customPropDoc;
                using (TextReader tr = new StreamReader(internal_docProps_custom.GetStream(FileMode.Open, FileAccess.Read)))
                    internal_customPropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);

                foreach (XElement cp in external_customProperties)
                {
                    // Does the internal document already have a custom property with this name?
                    XElement conflict = 
                    (
                        from d in internal_customPropDoc.Descendants(XName.Get("property", customPropertiesSchema.NamespaceName))
                        let ExternalName = d.Attribute(XName.Get("name", customPropertiesSchema.NamespaceName))
                        let InternalName = cp.Attribute(XName.Get("name", customPropertiesSchema.NamespaceName))
                        where ExternalName != null && InternalName != null && ExternalName == InternalName
                        select d
                    ).FirstOrDefault();

                    // Same name
                    if (conflict != null)
                    {

                    }

                    // There is no conflict, just add the Custom Property.
                    else
                        internal_customPropDoc.Root.Add(cp);
                }

                using (TextWriter tw = new StreamWriter(internal_docProps_custom.GetStream(FileMode.Open, FileAccess.Write)))
                    internal_customPropDoc.Save(tw);


            }
            #endregion

            // A document can only have one header and one footer.
            #region Remove external (header & footer) references
            var externalHeaderAndFooterReferences = 
            (
                from d in external_elements.Descendants()
                where d.Name.LocalName == "headerReference" || d.Name.LocalName == "footerReference"
                select d
            );

            foreach (var r in externalHeaderAndFooterReferences.ToList())
                r.Remove();
            #endregion
        }

        

        /// <summary>
        /// Insert a new Table at the end of this document.
        /// </summary>
        /// <param name="coloumnCount">The number of coloumns to create.</param>
        /// <param name="rowCount">The number of rows to create.</param>
        /// <returns>A new Table.</returns>
        /// <example>
        /// Insert a new Table with 2 coloumns and 3 rows, at the end of a document.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"C:\Example\Test.docx"))
        /// {
        ///     // Create a new Table with 2 coloumns and 3 rows.
        ///     Table newTable = document.InsertTable(2, 3);
        ///
        ///     // Set the design of this Table.
        ///     newTable.Design = TableDesign.LightShadingAccent2;
        ///
        ///     // Set the coloumn names.
        ///     newTable.Rows[0].Cells[0].Paragraph.InsertText("Ice Cream", false);
        ///     newTable.Rows[0].Cells[1].Paragraph.InsertText("Price", false);
        ///
        ///     // Fill row 1
        ///     newTable.Rows[1].Cells[0].Paragraph.InsertText("Chocolate", false);
        ///     newTable.Rows[1].Cells[1].Paragraph.InsertText("€3:50", false);
        ///
        ///     // Fill row 2
        ///     newTable.Rows[2].Cells[0].Paragraph.InsertText("Vanilla", false);
        ///     newTable.Rows[2].Cells[1].Paragraph.InsertText("€3:00", false);
        ///
        ///     // Save all changes made to document b.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Table InsertTable(int coloumnCount, int rowCount)
        {
            XElement newTable = HelperFunctions.CreateTable(rowCount, coloumnCount);
            mainDoc.Descendants(XName.Get("body", DocX.w.NamespaceName)).First().Add(newTable);

            return new Table(this, newTable);
        }

        public Table AddTable(int rowCount, int coloumnCount)
        {
            return (new Table(this, HelperFunctions.CreateTable(rowCount, coloumnCount)));
        }

        /// <summary>
        /// Insert a Table into this document. The Table's source can be a completely different document.
        /// </summary>
        /// <param name="t">The Table to insert.</param>
        /// <param name="index">The index to insert this Table at.</param>
        /// <returns>The Table now associated with this document.</returns>
        /// <example>
        /// Extract a Table from document a and insert it into document b, at index 10.
        /// <code>
        /// // Place holder for a Table.
        /// Table t;
        ///
        /// // Load document a.
        /// using (DocX documentA = DocX.Load(@"C:\Example\a.docx"))
        /// {
        ///     // Get the first Table from this document.
        ///     t = documentA.Tables[0];
        /// }
        ///
        /// // Load document b.
        /// using (DocX documentB = DocX.Load(@"C:\Example\b.docx"))
        /// {
        ///     /* 
        ///      * Insert the Table that was extracted from document a, into document b. 
        ///      * This creates a new Table that is now associated with document b.
        ///      */
        ///     Table newTable = documentB.InsertTable(10, t);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Table InsertTable(int index, Table t)
        {
            Paragraph p = HelperFunctions.GetFirstParagraphEffectedByInsert(this, index);

            XElement[] split = HelperFunctions.SplitParagraph(p, index - p.startIndex);
            XElement newXElement = new XElement(t.Xml);
            p.Xml.ReplaceWith
            (
                split[0],
                newXElement,
                split[1]
            );

            Table newTable = new Table(this, newXElement);
            newTable.Design = t.Design;

            return newTable;
        }

        /// <summary>
        /// Insert a Table into this document. The Table's source can be a completely different document.
        /// </summary>
        /// <param name="t">The Table to insert.</param>
        /// <returns>The Table now associated with this document.</returns>
        /// <example>
        /// Extract a Table from document a and insert it at the end of document b.
        /// <code>
        /// // Place holder for a Table.
        /// Table t;
        ///
        /// // Load document a.
        /// using (DocX documentA = DocX.Load(@"C:\Example\a.docx"))
        /// {
        ///     // Get the first Table from this document.
        ///     t = documentA.Tables[0];
        /// }
        ///
        /// // Load document b.
        /// using (DocX documentB = DocX.Load(@"C:\Example\b.docx"))
        /// {
        ///     /* 
        ///      * Insert the Table that was extracted from document a, into document b. 
        ///      * This creates a new Table that is now associated with document b.
        ///      */
        ///     Table newTable = documentB.InsertTable(t);
        ///
        ///     // Save all changes made to document b.
        ///     documentB.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Table InsertTable(Table t)
        {
            XElement newXElement = new XElement(t.Xml);
            mainDoc.Descendants(XName.Get("body", DocX.w.NamespaceName)).First().Add(newXElement);

            Table newTable = new Table(this, newXElement);
            newTable.Design = t.Design;

            return newTable;
        }

        /// <summary>
        /// Insert a new Table at the end of this document.
        /// </summary>
        /// <param name="coloumnCount">The number of coloumns to create.</param>
        /// <param name="rowCount">The number of rows to create.</param>
        /// <param name="index">The index to insert this Table at.</param>
        /// <returns>A new Table.</returns>
        /// <example>
        /// Insert a new Table with 2 coloumns and 3 rows, at index 37 in this document.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Create a new Table with 2 coloumns and 3 rows. Insert this Table at index 37.
        ///     Table newTable = document.InsertTable(37, 2, 3);
        ///
        ///     // Set the design of this Table.
        ///     newTable.Design = TableDesign.LightShadingAccent3;
        ///
        ///     // Set the coloumn names.
        ///     newTable.Rows[0].Cells[0].Paragraph.InsertText("Ice Cream", false);
        ///     newTable.Rows[0].Cells[1].Paragraph.InsertText("Price", false);
        ///
        ///     // Fill row 1
        ///     newTable.Rows[1].Cells[0].Paragraph.InsertText("Chocolate", false);
        ///     newTable.Rows[1].Cells[1].Paragraph.InsertText("€3:50", false);
        ///
        ///     // Fill row 2
        ///     newTable.Rows[2].Cells[0].Paragraph.InsertText("Vanilla", false);
        ///     newTable.Rows[2].Cells[1].Paragraph.InsertText("€3:00", false);
        ///
        ///     // Save all changes made to document b.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Table InsertTable(int index, int coloumnCount, int rowCount)
        {
            XElement newTable = HelperFunctions.CreateTable(rowCount, coloumnCount);

            Paragraph p = HelperFunctions.GetFirstParagraphEffectedByInsert(this, index);

            if (p == null)
                mainDoc.Descendants(XName.Get("body", DocX.w.NamespaceName)).First().AddFirst(newTable);

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


            return new Table(this, newTable);
        }
        
        /// <summary>
        /// Creates a document using a Stream.
        /// </summary>
        /// <param name="stream">The Stream to create the document from.</param>
        /// <returns>Returns a DocX object which represents the document.</returns>
        /// <example>
        /// Creating a document from a FileStream.
        /// <code>
        /// // Use a FileStream fs to create a new document.
        /// using(FileStream fs = new FileStream(@"C:\Example\Test.docx", FileMode.Create))
        /// {
        ///     // Load the document using fs
        ///     using (DocX document = DocX.Create(fs))
        ///     {
        ///         // Do something with the document here.
        ///
        ///         // Save all changes made to this document.
        ///         document.Save();
        ///     }// Release this document from memory.
        /// }
        /// </code>
        /// </example>
        /// <example>
        /// Creating a document in a SharePoint site.
        /// <code>
        /// using(SPSite mySite = new SPSite("http://server/sites/site"))
        /// {
        ///     // Open a connection to the SharePoint site
        ///     using(SPWeb myWeb = mySite.OpenWeb())
        ///     {
        ///         // Create a MemoryStream ms.
        ///         using (MemoryStream ms = new MemoryStream())
        ///         {
        ///             // Create a document using ms.
        ///             using (DocX document = DocX.Create(ms))
        ///             {
        ///                 // Do something with the document here.
        ///
        ///                 // Save all changes made to this document.
        ///                 document.Save();
        ///             }// Release this document from memory
        ///
        ///             // Add the document to the SharePoint site
        ///             web.Files.Add("filename", ms.ToArray(), true);
        ///         }
        ///     }
        /// }
        /// </code>
        /// </example>
        /// <seealso cref="DocX.Create(string)"/>
        /// <seealso cref="DocX.Load(System.IO.Stream)"/>
        /// <seealso cref="DocX.Load(string)"/>
        /// <seealso cref="DocX.Save()"/>
        public static DocX Create(Stream stream)
        {
            // Store this document in memory
            MemoryStream ms = new MemoryStream();

            // Create the docx package
            Package package = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite);

            PostCreation(ref package);
            DocX document = DocX.Load(ms);
            document.stream = stream;
            return document;
        }

        /// <summary>
        /// Creates a document using a fully qualified or relative filename.
        /// </summary>
        /// <param name="filename">The fully qualified or relative filename.</param>
        /// <returns>Returns a DocX object which represents the document.</returns>
        /// <example>
        /// <code>
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Create(@"..\Test.docx"))
        /// {
        ///     // Do something with the document here.
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory
        /// </code>
        /// <code>
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Create(@"..\Test.docx"))
        /// {
        ///     // Do something with the document here.
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory
        /// </code>
        /// <seealso cref="DocX.Create(System.IO.Stream)"/>
        /// <seealso cref="DocX.Load(System.IO.Stream)"/>
        /// <seealso cref="DocX.Load(string)"/>
        /// <seealso cref="DocX.Save()"/>
        /// </example>
        public static DocX Create(string filename)
        {
            // Store this document in memory
            MemoryStream ms = new MemoryStream();

            // Create the docx package
            //WordprocessingDocument wdDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            Package package = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite);

            PostCreation(ref package);
            DocX document = DocX.Load(ms);
            document.filename = filename;
            return document;
        }

        internal static void PostCreation(ref Package package)
        {
            XDocument mainDoc, stylesDoc;

            #region MainDocumentPart
            // Create the main document part for this package
            PackagePart mainDocumentPart = package.CreatePart(new Uri("/word/document.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
            package.CreateRelationship(mainDocumentPart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
            
            // Load the document part into a XDocument object
            using (TextReader tr = new StreamReader(mainDocumentPart.GetStream(FileMode.Create, FileAccess.ReadWrite)))
            {
                mainDoc = XDocument.Parse
                (@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                   <w:document xmlns:ve=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"">
                   <w:body>
                    <w:sectPr w:rsidR=""003E25F4"" w:rsidSect=""00FC3028"">
                        <w:pgSz w:w=""11906"" w:h=""16838""/>
                        <w:pgMar w:top=""1440"" w:right=""1440"" w:bottom=""1440"" w:left=""1440"" w:header=""708"" w:footer=""708"" w:gutter=""0""/>
                        <w:cols w:space=""708""/>
                        <w:docGrid w:linePitch=""360""/>
                    </w:sectPr>
                   </w:body>
                   </w:document>"
                );
            }

            // Save the main document
            using (TextWriter tw = new StreamWriter(mainDocumentPart.GetStream(FileMode.Create, FileAccess.Write)))
                mainDoc.Save(tw, SaveOptions.DisableFormatting);
            #endregion

            #region StylePart
            stylesDoc = HelperFunctions.AddDefaultStylesXml(package);
            #endregion

            package.Close();
        }

        internal static DocX PostLoad(ref Package package)
        {
            DocX document = new DocX(null, null);
            document.package = package;
            document.Document = document;

            #region MainDocumentPart
            document.mainPart = package.GetParts().Where
            (
                p => p.ContentType.Equals
                (
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                    StringComparison.CurrentCultureIgnoreCase
                )
            ).Single();

            using (TextReader tr = new StreamReader(document.mainPart.GetStream(FileMode.Open, FileAccess.Read)))
                document.mainDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);
            #endregion

            PopulateDocument(document, package);

          return document;
        }

      private static void PopulateDocument(DocX document, Package package)
      {
        Headers headers = new Headers();
        headers.odd = document.GetHeaderByType("default");
        headers.even = document.GetHeaderByType("even");
        headers.first = document.GetHeaderByType("first");

        Footers footers = new Footers();
        footers.odd = document.GetFooterByType("default");
        footers.even = document.GetFooterByType("even");
        footers.first = document.GetFooterByType("first");

        document.Xml = document.mainDoc.Root.Element(w + "body");
        document.headers = headers;
        document.footers = footers;
        document.settingsPart = HelperFunctions.CreateOrGetSettingsPart(package);
      }

      /// <summary>
        /// Loads a document into a DocX object using a Stream.
        /// </summary>
        /// <param name="stream">The Stream to load the document from.</param>
        /// <returns>
        /// Returns a DocX object which represents the document.
        /// </returns>
        /// <example>
        /// Loading a document from a FileStream.
        /// <code>
        /// // Open a FileStream fs to a document.
        /// using (FileStream fs = new FileStream(@"C:\Example\Test.docx", FileMode.Open))
        /// {
        ///     // Load the document using fs.
        ///     using (DocX document = DocX.Load(fs))
        ///     {
        ///         // Do something with the document here.
        ///            
        ///         // Save all changes made to the document.
        ///         document.Save();
        ///     }// Release this document from memory.
        /// }
        /// </code>
        /// </example>
        /// <example>
        /// Loading a document from a SharePoint site.
        /// <code>
        /// // Get the SharePoint site that you want to access.
        /// using (SPSite mySite = new SPSite("http://server/sites/site"))
        /// {
        ///     // Open a connection to the SharePoint site
        ///     using (SPWeb myWeb = mySite.OpenWeb())
        ///     {
        ///         // Grab a document stored on this site.
        ///         SPFile file = web.GetFile("Source_Folder_Name/Source_File");
        ///
        ///         // DocX.Load requires a Stream, so open a Stream to this document.
        ///         Stream str = new MemoryStream(file.OpenBinary());
        ///
        ///         // Load the file using the Stream str.
        ///         using (DocX document = DocX.Load(str))
        ///         {
        ///             // Do something with the document here.
        ///
        ///             // Save all changes made to the document.
        ///             document.Save();
        ///         }// Release this document from memory.
        ///     }
        /// }
        /// </code>
        /// </example>
        /// <seealso cref="DocX.Load(string)"/>
        /// <seealso cref="DocX.Create(System.IO.Stream)"/>
        /// <seealso cref="DocX.Create(string)"/>
        /// <seealso cref="DocX.Save()"/>
        public static DocX Load(Stream stream)
        {
            MemoryStream ms = new MemoryStream();

            stream.Position = 0;
            byte[] data = new byte[stream.Length];
            stream.Read(data, 0, (int)stream.Length);
            ms.Write(data, 0, (int)stream.Length);

            // Open the docx package
            Package package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite);

            DocX document = PostLoad(ref package);
            document.package = package;
            document.memoryStream = ms;
            document.stream = stream;
            return document;
        }

        /// <summary>
        /// Loads a document into a DocX object using a fully qualified or relative filename.
        /// </summary>
        /// <param name="filename">The fully qualified or relative filename.</param>
        /// <returns>
        /// Returns a DocX object which represents the document.
        /// </returns>
        /// <example>
        /// <code>
        /// // Load a document using its fully qualified filename
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Do something with the document here
        ///
        ///     // Save all changes made to document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// <code>
        /// // Load a document using its relative filename.
        /// using(DocX document = DocX.Load(@"..\..\Test.docx"))
        /// { 
        ///     // Do something with the document here.
        ///                
        ///     // Save all changes made to document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// <seealso cref="DocX.Load(System.IO.Stream)"/>
        /// <seealso cref="DocX.Create(System.IO.Stream)"/>
        /// <seealso cref="DocX.Create(string)"/>
        /// <seealso cref="DocX.Save()"/>
        /// </example>
        public static DocX Load(string filename)
        {
            if (!File.Exists(filename))
                throw new FileNotFoundException(string.Format("File could not be found {0}", filename));

            MemoryStream ms = new MemoryStream();
            
            using (FileStream fs = new FileStream(filename, FileMode.Open))
            {
                byte[] data = new byte[fs.Length];
                fs.Read(data, 0, (int)fs.Length);
                ms.Write(data, 0, (int)fs.Length);
            }

            // Open the docx package
            Package package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite);

            DocX document = PostLoad(ref package);
            document.package = package;
            document.filename = filename;
            document.memoryStream = ms;

            return document;
        }

        ///<summary>
        /// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
        ///</summary>
        ///<param name="templateFilePath">The path to the document template file.</param>
        ///<exception cref="FileNotFoundException">The document template file not found.</exception>
        public void ApplyTemplate(string templateFilePath)
        {
          ApplyTemplate(templateFilePath, true);
        }

        ///<summary>
        /// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
        ///</summary>
        ///<param name="templateFilePath">The path to the document template file.</param>
        ///<param name="includeContent">Whether to copy the document template text content to document.</param>
        ///<exception cref="FileNotFoundException">The document template file not found.</exception>
        public void ApplyTemplate(string templateFilePath, bool includeContent)
        {
          if (!File.Exists(templateFilePath))
          {
            throw new FileNotFoundException(string.Format("File could not be found {0}", templateFilePath));
          }
          using (FileStream packageStream = new FileStream(templateFilePath, FileMode.Open, FileAccess.Read))
          {
            ApplyTemplate(packageStream, includeContent);
          }
        }

        ///<summary>
        /// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
        ///</summary>
        ///<param name="templateStream">The stream of the document template file.</param>
        public void ApplyTemplate(Stream templateStream)
        {
          ApplyTemplate(templateStream, true);
        }

        ///<summary>
        /// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
        ///</summary>
        ///<param name="templateStream">The stream of the document template file.</param>
        ///<param name="includeContent">Whether to copy the document template text content to document.</param>
        public void ApplyTemplate(Stream templateStream, bool includeContent)
        {
          Package templatePackage = Package.Open(templateStream);
          try
          {
            PackagePart documentPart = null;
            XDocument documentDoc = null;
            foreach (PackagePart packagePart in templatePackage.GetParts())
            {
              switch (packagePart.Uri.ToString())
              {
                case "/word/document.xml":
                  documentPart = packagePart;
                  using (XmlReader xr = XmlReader.Create(packagePart.GetStream(FileMode.Open, FileAccess.Read)))
                  {
                    documentDoc = XDocument.Load(xr);
                  }
                  break;
                case "/_rels/.rels":
                  if (!this.package.PartExists(packagePart.Uri))
                  {
                    this.package.CreatePart(packagePart.Uri, packagePart.ContentType, packagePart.CompressionOption);
                  }
                  PackagePart globalRelsPart = this.package.GetPart(packagePart.Uri);
                  using (
                    StreamReader tr = new StreamReader(
                      packagePart.GetStream(FileMode.Open, FileAccess.Read), Encoding.UTF8))
                  {
                    using (
                      StreamWriter tw = new StreamWriter(
                        globalRelsPart.GetStream(FileMode.Create, FileAccess.Write), Encoding.UTF8))
                    {
                      tw.Write(tr.ReadToEnd());
                    }
                  }
                  break;
                case "/word/_rels/document.xml.rels":
                  break;
                default:
                  if (!this.package.PartExists(packagePart.Uri))
                  {
                    this.package.CreatePart(packagePart.Uri, packagePart.ContentType, packagePart.CompressionOption);
                  }
                  Encoding packagePartEncoding = Encoding.Default;
                  if (packagePart.Uri.ToString().EndsWith(".xml") || packagePart.Uri.ToString().EndsWith(".rels"))
                  {
                    packagePartEncoding = Encoding.UTF8;
                  }
                  PackagePart nativePart = this.package.GetPart(packagePart.Uri);
                  using (
                    StreamReader tr = new StreamReader(
                      packagePart.GetStream(FileMode.Open, FileAccess.Read), packagePartEncoding))
                  {
                    using (
                      StreamWriter tw = new StreamWriter(
                        nativePart.GetStream(FileMode.Create, FileAccess.Write), tr.CurrentEncoding))
                    {
                      tw.Write(tr.ReadToEnd());
                    }
                  }
                  break;
              }
            }
            if (documentPart != null)
            {
              string mainContentType = documentPart.ContentType.Replace("template.main", "document.main");
              if (this.package.PartExists(documentPart.Uri))
              {
                this.package.DeletePart(documentPart.Uri);
              }
              PackagePart documentNewPart = this.package.CreatePart(
                documentPart.Uri, mainContentType, documentPart.CompressionOption);
              using (XmlWriter xw = XmlWriter.Create(documentNewPart.GetStream(FileMode.Create, FileAccess.Write)))
              {
                documentDoc.WriteTo(xw);
              }
              foreach (PackageRelationship documentPartRel in documentPart.GetRelationships())
              {
                documentNewPart.CreateRelationship(
                  documentPartRel.TargetUri,
                  documentPartRel.TargetMode,
                  documentPartRel.RelationshipType,
                  documentPartRel.Id);
              }
              this.mainPart = documentNewPart;
              this.mainDoc = documentDoc;
              PopulateDocument(this, templatePackage);
            }
            if (!includeContent)
            {
              foreach (Paragraph paragraph in this.Paragraphs)
              {
                paragraph.Remove(false);
              }
            }
          }
          finally
          {
            this.package.Flush();
            var documentRelsPart = this.package.GetPart(new Uri("/word/_rels/document.xml.rels", UriKind.Relative));
            using (TextReader tr = new StreamReader(documentRelsPart.GetStream(FileMode.Open, FileAccess.Read)))
            {
              tr.Read();
            }
            templatePackage.Close();
          }
        }

        /// <summary>
        /// Add an Image into this document from a fully qualified or relative filename.
        /// </summary>
        /// <param name="filename">The fully qualified or relative filename.</param>
        /// <returns>An Image file.</returns>
        /// <example>
        /// Add an Image into this document from a fully qualified filename.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Add an Image from a file.
        ///     document.AddImage(@"C:\Example\Image.png");
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <seealso cref="AddImage(System.IO.Stream)"/>
        /// <seealso cref="Paragraph.InsertPicture"/>
        public Image AddImage(string filename)
        {
            return AddImage(filename as object);
        }

        /// <summary>
        /// Add an Image into this document from a Stream.
        /// </summary>
        /// <param name="stream">A Stream stream.</param>
        /// <returns>An Image file.</returns>
        /// <example>
        /// Add an Image into a document using a Stream. 
        /// <code>
        /// // Open a FileStream fs to an Image.
        /// using (FileStream fs = new FileStream(@"C:\Example\Image.jpg", FileMode.Open))
        /// {
        ///     // Load a document.
        ///     using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        ///     {
        ///         // Add an Image from a filestream fs.
        ///         document.AddImage(fs);
        ///
        ///         // Save all changes made to this document.
        ///         document.Save();
        ///     }// Release this document from memory.
        /// }
        /// </code>
        /// </example>
        /// <seealso cref="AddImage(string)"/>
        /// <seealso cref="Paragraph.InsertPicture"/>
        public Image AddImage(Stream stream)
        {
            return AddImage(stream as object);
        }

        /// <summary>
        /// Adds a hyperlink to a document and creates a Paragraph which uses it.
        /// </summary>
        /// <param name="text">The text as displayed by the hyperlink.</param>
        /// <param name="uri">The hyperlink itself.</param>
        /// <returns>Returns a hyperlink that can be inserted into a Paragraph.</returns>
        /// <example>
        /// Adds a hyperlink to a document and creates a Paragraph which uses it.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///    // Add a hyperlink to this document.
        ///    Hyperlink h = document.AddHyperlink("Google", new Uri("http://www.google.com"));
        ///    
        ///    // Add a new Paragraph to this document.
        ///    Paragraph p = document.InsertParagraph();
        ///    p.Append("My favourite search engine is ");
        ///    p.AppendHyperlink(h);
        ///    p.Append(", I think it's great.");
        ///
        ///    // Save all changes made to this document.
        ///    document.Save();
        /// }
        /// </code>
        /// </example>
        public Hyperlink AddHyperlink(string text, Uri uri)
        {
            XElement i = new XElement
            (
                XName.Get("hyperlink", DocX.w.NamespaceName),
                new XAttribute(r + "id", string.Empty),
                new XAttribute(w + "history", "1"),
                new XElement(XName.Get("r", DocX.w.NamespaceName),
                new XElement(XName.Get("rPr", DocX.w.NamespaceName),
                new XElement(XName.Get("rStyle", DocX.w.NamespaceName),
                new XAttribute(w + "val", "Hyperlink"))),
                new XElement(XName.Get("t", DocX.w.NamespaceName), text))
            );

            Hyperlink h = new Hyperlink(this, i);
            h.Text = text;
            h.Uri = uri;

            AddHyperlinkStyleIfNotPresent();

            return h;
        }

        internal void AddHyperlinkStyleIfNotPresent()
        {
            Uri word_styles_Uri = new Uri("/word/styles.xml", UriKind.Relative);

            // If the internal document contains no /word/styles.xml create one.
            if (!package.PartExists(word_styles_Uri))
                HelperFunctions.AddDefaultStylesXml(package);

            // Load the styles.xml into memory.
            XDocument word_styles;
            using (TextReader tr = new StreamReader(package.GetPart(word_styles_Uri).GetStream()))
                word_styles = XDocument.Load(tr);

            bool hyperlinkStyleExists = 
            (
                from s in word_styles.Element(w + "styles").Elements()
                let styleId = s.Attribute(XName.Get("styleId", w.NamespaceName))
                where (styleId != null && styleId.Value == "Hyperlink")
                select s
            ).Count() > 0;

            if (!hyperlinkStyleExists)
            {
                XElement style = new XElement
                (
                    w + "style",
                    new XAttribute(w + "type", "character"),
                    new XAttribute(w + "styleId", "Hyperlink"),
                        new XElement(w + "name",         new XAttribute(w + "val", "Hyperlink")),
                        new XElement(w + "basedOn",      new XAttribute(w + "val", "DefaultParagraphFont")),
                        new XElement(w + "uiPriority",   new XAttribute(w + "val", "99")),
                        new XElement(w + "unhideWhenUsed"),
                        new XElement(w + "rsid",         new XAttribute(w + "val", "0005416C")),
                        new XElement
                        (
                            w + "rPr",
                            new XElement(w + "color", new XAttribute(w + "val", "0000FF"), new XAttribute(w + "themeColor", "hyperlink")),
                            new XElement
                            (
                                w + "u",
                                new XAttribute(w + "val", "single")
                            )
                        )
                );
                word_styles.Element(w + "styles").Add(style);

                // Save the styles document.
                using (TextWriter tw = new StreamWriter(package.GetPart(word_styles_Uri).GetStream()))
                    word_styles.Save(tw);
            }
        }

        private string GetNextFreeRelationshipID()
        {
            string id =
            (
                from r in mainPart.GetRelationships()
                select r.Id
            ).Max();

            // The convension for ids is rid01, rid02, etc
            string newId = id.Replace("rId", "");
            int result;
            if (int.TryParse(newId, out result))
                return ("rId" + (result + 1));

            else
                return Guid.NewGuid().ToString();
        }

        /// <summary>
        /// Adds three new Headers to this document. One for the first page, one for odd pages and one for even pages.
        /// </summary>
        /// <example>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Add header support to this document.
        ///     document.AddHeaders();
        ///
        ///     // Get a collection of all headers in this document.
        ///     Headers headers = document.Headers;
        ///
        ///     // The header used for the first page of this document.
        ///     Header first = headers.first;
        ///
        ///     // The header used for odd pages of this document.
        ///     Header odd = headers.odd;
        ///
        ///     // The header used for even pages of this document.
        ///     Header even = headers.even;
        ///
        ///     // Force the document to use a different header for first, odd and even pages.
        ///     document.DifferentFirstPage = true;
        ///     document.DifferentOddAndEvenPages = true;
        ///
        ///     // Content can be added to the Headers in the same manor that it would be added to the main document.
        ///     Paragraph p = first.InsertParagraph();
        ///     p.Append("This is the first pages header.");
        ///
        ///     // Save all changes to this document.
        ///     document.Save();    
        /// }// Release this document from memory.
        /// </example>
        public void AddHeaders()
        {
            AddHeadersOrFooters(true);

            headers.odd = Document.GetHeaderByType("default");
            headers.even = Document.GetHeaderByType("even");
            headers.first = Document.GetHeaderByType("first");
        }

        /// <summary>
        /// Adds three new Footers to this document. One for the first page, one for odd pages and one for even pages.
        /// </summary>
        /// <example>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Add footer support to this document.
        ///     document.AddFooters();
        ///
        ///     // Get a collection of all footers in this document.
        ///     Footers footers = document.Footers;
        ///
        ///     // The footer used for the first page of this document.
        ///     Footer first = footers.first;
        ///
        ///     // The footer used for odd pages of this document.
        ///     Footer odd = footers.odd;
        ///
        ///     // The footer used for even pages of this document.
        ///     Footer even = footers.even;
        ///
        ///     // Force the document to use a different footer for first, odd and even pages.
        ///     document.DifferentFirstPage = true;
        ///     document.DifferentOddAndEvenPages = true;
        ///
        ///     // Content can be added to the Footers in the same manor that it would be added to the main document.
        ///     Paragraph p = first.InsertParagraph();
        ///     p.Append("This is the first pages footer.");
        ///
        ///     // Save all changes to this document.
        ///     document.Save();    
        /// }// Release this document from memory.
        /// </example>
        public void AddFooters()
        {
            AddHeadersOrFooters(false);

            footers.odd = Document.GetFooterByType("default");
            footers.even = Document.GetFooterByType("even");
            footers.first = Document.GetFooterByType("first");
        }

        /// <summary>
        /// Adds a Header to a document.
        /// If the document already contains a Header it will be replaced.
        /// </summary>
        /// <returns>The Header that was added to the document.</returns>
        internal void AddHeadersOrFooters(bool b)
        {
            string element = "ftr";
            string reference = "footer";
            if (b)
            {
                element = "hdr";
                reference = "header";
            }

            DeleteHeadersOrFooters(b);

            XElement sectPr = mainDoc.Root.Element(w + "body").Element(w + "sectPr");

            for (int i = 1; i < 4; i++)
            {
                string header_uri = string.Format("/word/{0}{1}.xml", reference, i);

                PackagePart headerPart = package.CreatePart(new Uri(header_uri, UriKind.Relative), string.Format("application/vnd.openxmlformats-officedocument.wordprocessingml.{0}+xml", reference));
                PackageRelationship headerRelationship = mainPart.CreateRelationship(headerPart.Uri, TargetMode.Internal, string.Format("http://schemas.openxmlformats.org/officeDocument/2006/relationships/{0}", reference));

                XDocument header;

                // Load the document part into a XDocument object
                using (TextReader tr = new StreamReader(headerPart.GetStream(FileMode.Create, FileAccess.ReadWrite)))
                {
                    header = XDocument.Parse
                    (string.Format(@"<?xml version=""1.0"" encoding=""utf-16"" standalone=""yes""?>
                       <w:{0} xmlns:ve=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"">
                         <w:p w:rsidR=""009D472B"" w:rsidRDefault=""009D472B"">
                           <w:pPr>
                             <w:pStyle w:val=""{1}"" />
                           </w:pPr>
                         </w:p>
                       </w:{0}>", element, reference)
                    );
                }

                // Save the main document
                using (TextWriter tw = new StreamWriter(headerPart.GetStream(FileMode.Create, FileAccess.Write)))
                    header.Save(tw, SaveOptions.DisableFormatting);

                string type;
                switch (i)
                {
                    case 1: type = "default"; break;
                    case 2: type = "even"; break;
                    case 3: type = "first"; break;
                    default: throw new ArgumentOutOfRangeException();
                }

                sectPr.Add
                (
                    new XElement
                    (
                        w + string.Format("{0}Reference", reference),
                        new XAttribute(w + "type", type),
                        new XAttribute(r + "id", headerRelationship.Id)
                    )
                );
            }
        }

        internal void DeleteHeadersOrFooters(bool b)
        {
            string reference = "footer";
            if (b)
                reference = "header";

            // Get all header Relationships in this document.
            var header_relationships = mainPart.GetRelationshipsByType(string.Format("http://schemas.openxmlformats.org/officeDocument/2006/relationships/{0}", reference));

            foreach (PackageRelationship header_relationship in header_relationships)
            {
                // Get the TargetUri for this Part.
                Uri header_uri = header_relationship.TargetUri;

                // Check to see if the document actually contains the Part.
                if (!header_uri.OriginalString.StartsWith("/word/"))
                    header_uri = new Uri("/word/" + header_uri.OriginalString, UriKind.Relative);

                if (package.PartExists(header_uri))
                {
                    // Delete the Part
                    package.DeletePart(header_uri);

                    // Get all references to this Relationship in the document.
                    var query =
                    (
                        from e in mainDoc.Descendants(XName.Get("body", DocX.w.NamespaceName)).Descendants()
                        where (e.Name.LocalName == string.Format("{0}Reference", reference)) && (e.Attribute(r + "id").Value == header_relationship.Id)
                        select e
                    );

                    // Remove all references to this Relationship in the document.
                    for (int i = 0; i < query.Count(); i++)
                        query.ElementAt(i).Remove();

                    // Delete the Relationship.
                    package.DeleteRelationship(header_relationship.Id);
                }
            }
        }

        internal Image AddImage(object o)
        {
            // Open a Stream to the new image being added.
            Stream newImageStream;
            if (o is string)
                newImageStream = new FileStream(o as string, FileMode.Open, FileAccess.Read);
            else
                newImageStream = o as Stream;

            // Get all image parts in word\document.xml
            List<PackagePart> imageParts = mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image").Select(ir => package.GetParts().Where(p => p.Uri.ToString().EndsWith(ir.TargetUri.ToString())).First()).ToList();
            foreach (PackagePart relsPart in package.GetParts().Where(part => part.Uri.ToString().Contains("/word/")).Where(part => part.ContentType.Equals("application/vnd.openxmlformats-package.relationships+xml")))
          {
            XDocument relsPartContent;
            using (TextReader tr = new StreamReader(relsPart.GetStream(FileMode.Open, FileAccess.Read)))
            {
              relsPartContent = XDocument.Load(tr);
            }
            IEnumerable<XElement> imageRelationships =
              relsPartContent.Root.Elements().Where(
                imageRel =>
                imageRel.Attribute(XName.Get("Type")).Value.Equals(
                  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"));
            foreach (XElement imageRelationship in imageRelationships)
            {
              if (imageRelationship.Attribute(XName.Get("Target")) != null)
              {
                string imagePartUri = Path.Combine(Path.GetDirectoryName(relsPart.Uri.ToString()), imageRelationship.Attribute(XName.Get("Target")).Value);
                imagePartUri = Path.GetFullPath(imagePartUri.Replace("\\_rels", string.Empty));
                imagePartUri = imagePartUri.Replace(Path.GetFullPath("\\"), string.Empty).Replace("\\", "/");
                if (!imagePartUri.StartsWith("/"))
                {
                  imagePartUri = "/" + imagePartUri;
                }
                PackagePart imagePart = package.GetPart(new Uri(imagePartUri, UriKind.Relative));
                imageParts.Add(imagePart);
              }
            }
          }
            
            // Loop through each image part in this document.
            foreach (PackagePart pp in imageParts)
            {
                // Open a tempory Stream to this image part.
                using (Stream tempStream = pp.GetStream(FileMode.Open, FileAccess.Read))
                {
                    // Compare this image to the new image being added.
                    if (HelperFunctions.IsSameFile(tempStream, newImageStream))
                    {
                        // Get the image object for this image part
                        string id = mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
                        .Where(r => r.TargetUri == pp.Uri)
                        .Select(r => r.Id).First();

                        // Return the Image object
                        return Images.Where(i => i.Id == id).First();
                    }
                }
            }

            /* 
             * This Image being added is infact a new Image,
             * we need to generate a unique name for this image of the format imageN.ext,
             * where n is an integer that has not been used before.
             * This could probabily be replace by a Guid.
             */
            int max = 0;
            var values =
            (
                from ip in imageParts
                let Name = Path.GetFileNameWithoutExtension(ip.Uri.ToString())
                let Number = Regex.Match(Name, @"\d+$").Value
                select Number != string.Empty ? int.Parse(Number) : 0
            );
            if (values.Count() > 0)
                max = Math.Max(max, values.Max());

            // Create a new image part.
          string imgPartUriPath = string.Format("/word/media/image{0}.jpeg", max + 1);
          if (package.PartExists(new Uri(imgPartUriPath, UriKind.Relative)))
          {
            package.DeletePart(new Uri(imgPartUriPath, UriKind.Relative));
          }
            PackagePart img = package.CreatePart(new Uri(imgPartUriPath, UriKind.Relative), System.Net.Mime.MediaTypeNames.Image.Jpeg);

            // Create a new image relationship
            PackageRelationship rel = mainPart.CreateRelationship(img.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");

            // Open a Stream to the newly created Image part.
            using (Stream stream = img.GetStream(FileMode.Create, FileAccess.Write))
            {
                // Using the Stream to the real image, copy this streams data into the newly create Image part.
                using (newImageStream)
                {
                    byte[] bytes = new byte[newImageStream.Length];
                    newImageStream.Read(bytes, 0, (int)newImageStream.Length);
                    stream.Write(bytes, 0, (int)newImageStream.Length);
                }// Close the Stream to the new image.
            }// Close the Stream to the new image part.

            return new Image(this, rel);
        }

        /// <summary>
        /// Save this document back to the location it was loaded from.
        /// </summary>
        /// <example>
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Add an Image from a file.
        ///     document.AddImage(@"C:\Example\Image.jpg");
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <seealso cref="DocX.SaveAs(string)"/>
        /// <seealso cref="DocX.Create(System.IO.Stream)"/>
        /// <seealso cref="DocX.Create(string)"/>
        /// <seealso cref="DocX.Load(System.IO.Stream)"/>
        /// <seealso cref="DocX.Load(string)"/> 
        /// <!-- 
        /// Bug found and fixed by krugs525 on August 12 2009.
        /// Use TFS compare to see exact code change.
        /// -->
        public void Save()
        {
            Headers headers = Headers;

            // Save the main document
            using (TextWriter tw = new StreamWriter(mainPart.GetStream(FileMode.Create, FileAccess.Write)))
                mainDoc.Save(tw, SaveOptions.DisableFormatting);

            XElement body = mainDoc.Root.Element(w + "body");
            XElement sectPr = body.Element(w + "sectPr");
            
            var evenHeaderRef = 
            (
                from e in sectPr.Elements(w + "headerReference")
                let type = e.Attribute(w + "type")
                where type != null && type.Value.Equals("even", StringComparison.CurrentCultureIgnoreCase)
                select e.Attribute(r + "id").Value
             ).SingleOrDefault();

            if(evenHeaderRef != null)
            {
                XElement even = headers.even.Xml;

                Uri target = PackUriHelper.ResolvePartUri
                (
                    mainPart.Uri,
                    mainPart.GetRelationship(evenHeaderRef).TargetUri
                );
                
                using (TextWriter tw = new StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write)))
                {
                    new XDocument
                    (
                        new XDeclaration("1.0", "UTF-8", "yes"),
                        even
                    ).Save(tw, SaveOptions.DisableFormatting);
                }
            }

            var oddHeaderRef = 
            (
                from e in sectPr.Elements(w + "headerReference")
                let type = e.Attribute(w + "type")
                where type != null && type.Value.Equals("default", StringComparison.CurrentCultureIgnoreCase)
                select e.Attribute(r + "id").Value
             ).SingleOrDefault();

            if(oddHeaderRef != null)
            {
                XElement odd = headers.odd.Xml;

                Uri target = PackUriHelper.ResolvePartUri
                (
                    mainPart.Uri,
                    mainPart.GetRelationship(oddHeaderRef).TargetUri
                );

                // Save header1
                using (TextWriter tw = new StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write)))
                {
                    new XDocument
                    (
                        new XDeclaration("1.0", "UTF-8", "yes"),
                        odd
                    ).Save(tw, SaveOptions.DisableFormatting);
                }
            }

            var firstHeaderRef =
            (
                from e in sectPr.Elements(w + "headerReference")
                let type = e.Attribute(w + "type")
                where type != null && type.Value.Equals("first", StringComparison.CurrentCultureIgnoreCase)
                select e.Attribute(r + "id").Value
             ).SingleOrDefault();

            if(firstHeaderRef != null)
            {
                XElement first = headers.first.Xml;
                Uri target = PackUriHelper.ResolvePartUri
                (
                    mainPart.Uri,
                    mainPart.GetRelationship(firstHeaderRef).TargetUri
                );
               
                // Save header3
                using (TextWriter tw = new StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write)))
                {
                    new XDocument
                    (
                        new XDeclaration("1.0", "UTF-8", "yes"),
                        first
                    ).Save(tw, SaveOptions.DisableFormatting);
                }
            }

            var oddFooterRef =
            (
                from e in sectPr.Elements(w + "footerReference")
                let type = e.Attribute(w + "type")
                where type != null && type.Value.Equals("default", StringComparison.CurrentCultureIgnoreCase)
                select e.Attribute(r + "id").Value
             ).SingleOrDefault();

            if(oddFooterRef != null)
            {
                XElement odd = footers.odd.Xml;
                Uri target = PackUriHelper.ResolvePartUri
                (
                    mainPart.Uri,
                    mainPart.GetRelationship(oddFooterRef).TargetUri
                );
             
                // Save header1
                using (TextWriter tw = new StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write)))
                {
                    new XDocument
                    (
                        new XDeclaration("1.0", "UTF-8", "yes"),
                        odd
                    ).Save(tw, SaveOptions.DisableFormatting);
                }
            }

            var evenFooterRef =
            (
                from e in sectPr.Elements(w + "footerReference")
                let type = e.Attribute(w + "type")
                where type != null && type.Value.Equals("even", StringComparison.CurrentCultureIgnoreCase)
                select e.Attribute(r + "id").Value
             ).SingleOrDefault();

            if (evenFooterRef != null)
            {
                XElement even = footers.even.Xml;
                Uri target = PackUriHelper.ResolvePartUri
                (
                    mainPart.Uri,
                    mainPart.GetRelationship(evenFooterRef).TargetUri
                );
             
                // Save header2
                using (TextWriter tw = new StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write)))
                {
                    new XDocument
                    (
                        new XDeclaration("1.0", "UTF-8", "yes"),
                        even
                    ).Save(tw, SaveOptions.DisableFormatting);
                }
            }

            var firstFooterRef =
            (
                 from e in sectPr.Elements(w + "footerReference")
                 let type = e.Attribute(w + "type")
                 where type != null && type.Value.Equals("first", StringComparison.CurrentCultureIgnoreCase)
                 select e.Attribute(r + "id").Value
            ).SingleOrDefault();

            if (firstFooterRef != null)
            {         
                XElement first = footers.first.Xml;
                Uri target = PackUriHelper.ResolvePartUri
                (
                    mainPart.Uri,
                    mainPart.GetRelationship(firstFooterRef).TargetUri
                );

                // Save header3
                using (TextWriter tw = new StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write)))
                {
                    new XDocument
                    (
                        new XDeclaration("1.0", "UTF-8", "yes"),
                        first
                    ).Save(tw, SaveOptions.DisableFormatting);
                }
            }

            // Close the document so that it can be saved.
            package.Flush();

            #region Save this document back to a file or stream, that was specified by the user at save time.
            if (filename != null)
            {
                using (FileStream fs = new FileStream(filename, FileMode.Create))
                    fs.Write(memoryStream.ToArray(), 0, (int)memoryStream.Length);
            }

            else
            {
                // Set the length of this stream to 0
                stream.SetLength(0);

                // Write to the beginning of the stream
                stream.Position = 0;

                memoryStream.WriteTo(stream);
            }
            #endregion
        }

        /// <summary>
        /// Save this document to a file.
        /// </summary>
        /// <param name="filename">The filename to save this document as.</param>
        /// <example>
        /// Load a document from one file and save it to another.
        /// <code>
        /// // Load a document using its fully qualified filename.
        /// DocX document = DocX.Load(@"C:\Example\Test1.docx");
        ///
        /// // Insert a new Paragraph
        /// document.InsertParagraph("Hello world!", false);
        ///
        /// // Save the document to a new location.
        /// document.SaveAs(@"C:\Example\Test2.docx");
        /// </code>
        /// </example>
        /// <example>
        /// Load a document from a Stream and save it to a file.
        /// <code>
        /// DocX document;
        /// using (FileStream fs1 = new FileStream(@"C:\Example\Test1.docx", FileMode.Open))
        /// {
        ///     // Load a document using a stream.
        ///     document = DocX.Load(fs1);
        ///
        ///     // Insert a new Paragraph
        ///     document.InsertParagraph("Hello world again!", false);
        /// }
        ///    
        /// // Save the document to a new location.
        /// document.SaveAs(@"C:\Example\Test2.docx");
        /// </code>
        /// </example>
        /// <seealso cref="DocX.Save()"/>
        /// <seealso cref="DocX.Create(System.IO.Stream)"/>
        /// <seealso cref="DocX.Create(string)"/>
        /// <seealso cref="DocX.Load(System.IO.Stream)"/>
        /// <seealso cref="DocX.Load(string)"/>
        public void SaveAs(string filename)
        {
            this.filename = filename;
            this.stream = null;
            Save();
        }

        /// <summary>
        /// Save this document to a Stream.
        /// </summary>
        /// <param name="stream">The Stream to save this document to.</param>
        /// <example>
        /// Load a document from a file and save it to a Stream.
        /// <code>
        /// // Place holder for a document.
        /// DocX document;
        ///
        /// using (FileStream fs1 = new FileStream(@"C:\Example\Test1.docx", FileMode.Open))
        /// {
        ///     // Load a document using a stream.
        ///     document = DocX.Load(fs1);
        ///
        ///     // Insert a new Paragraph
        ///     document.InsertParagraph("Hello world again!", false);
        /// }
        ///
        /// using (FileStream fs2 = new FileStream(@"C:\Example\Test2.docx", FileMode.Create))
        /// {
        ///     // Save the document to a different stream.
        ///     document.SaveAs(fs2);
        /// }
        ///
        /// // Release this document from memory.
        /// document.Dispose();
        /// </code>
        /// </example>
        /// <example>
        /// Load a document from one Stream and save it to another.
        /// <code>
        /// DocX document;
        /// using (FileStream fs1 = new FileStream(@"C:\Example\Test1.docx", FileMode.Open))
        /// {
        ///     // Load a document using a stream.
        ///     document = DocX.Load(fs1);
        ///
        ///     // Insert a new Paragraph
        ///     document.InsertParagraph("Hello world again!", false);
        /// }
        /// 
        /// using (FileStream fs2 = new FileStream(@"C:\Example\Test2.docx", FileMode.Create))
        /// {
        ///     // Save the document to a different stream.
        ///     document.SaveAs(fs2);
        /// }
        /// </code>
        /// </example>
        /// <seealso cref="DocX.Save()"/>
        /// <seealso cref="DocX.Create(System.IO.Stream)"/>
        /// <seealso cref="DocX.Create(string)"/>
        /// <seealso cref="DocX.Load(System.IO.Stream)"/>
        /// <seealso cref="DocX.Load(string)"/>
        public void SaveAs(Stream stream)
        {
            this.filename = null;
            this.stream = stream;
            Save();
        }
      
      /// <summary>
      /// Add a core property to this document. If a core property already exists with the same name it will be replaced. Core property names are case insensitive.
      /// </summary>
      ///<param name="propertyName">The property name.</param>
      ///<param name="propertyValue">The property value.</param>
      ///<example>
      /// Add a core properties of each type to a document.
      /// <code>
      /// // Load Example.docx
      /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
      /// {
      ///     // If this document does not contain a core property called 'forename', create one.
      ///     if (!document.CoreProperties.ContainsKey("forename"))
      ///     {
      ///         // Create a new core property called 'forename' and set its value.
      ///         document.AddCoreProperty("forename", "Cathal");
      ///     }
      ///
      ///     // Get this documents core property called 'forename'.
      ///     string forenameValue = document.CoreProperties["forename"];
      ///
      ///     // Print all of the information about this core property to Console.
      ///     Console.WriteLine(string.Format("Name: '{0}', Value: '{1}'\nPress any key...", "forename", forenameValue));
      ///     
      ///     // Save all changes made to this document.
      ///     document.Save();
      /// } // Release this document from memory.
      ///
      /// // Wait for the user to press a key before exiting.
      /// Console.ReadKey();
      /// </code>
      /// </example>
      /// <seealso cref="CoreProperties"/>
      /// <seealso cref="CustomProperty"/>
      /// <seealso cref="CustomProperties"/>
      public void AddCoreProperty(string propertyName, string propertyValue)
      {
        string propertyNamespacePrefix = propertyName.Contains(":") ? propertyName.Split(new[] { ':' })[0] : "cp";
        string propertyLocalName = propertyName.Contains(":") ? propertyName.Split(new[] { ':' })[1] : propertyName;

        // If this document does not contain a coreFilePropertyPart create one.)
        if (!package.PartExists(new Uri("/docProps/core.xml", UriKind.Relative)))
          throw new Exception("Core properties part doesn't exist.");

        XDocument corePropDoc;
        PackagePart corePropPart = package.GetPart(new Uri("/docProps/core.xml", UriKind.Relative));
        using (TextReader tr = new StreamReader(corePropPart.GetStream(FileMode.Open, FileAccess.Read)))
        {
          corePropDoc = XDocument.Load(tr);
        }

        XElement corePropElement =
          (from propElement in corePropDoc.Root.Elements()
           where (propElement.Name.LocalName.Equals(propertyLocalName))
           select propElement).SingleOrDefault();
        if (corePropElement != null)
        {
          corePropElement.SetValue(propertyValue);
        }
        else
        {
          var propertyNamespace = corePropDoc.Root.GetNamespaceOfPrefix(propertyNamespacePrefix);
          corePropDoc.Root.Add(new XElement(XName.Get(propertyLocalName, propertyNamespace.NamespaceName), propertyValue));
        }
        
        using(TextWriter tw = new StreamWriter(corePropPart.GetStream(FileMode.Create, FileAccess.Write)))
        {
          corePropDoc.Save(tw);
        }
        UpdateCorePropertyValue(this, propertyLocalName, propertyValue);
      }

      internal static void UpdateCorePropertyValue(DocX document, string corePropertyName, string corePropertyValue)
      {
        string matchPattern = string.Format(@"(DOCPROPERTY)?{0}\\\*MERGEFORMAT", corePropertyName).ToLower();
        foreach (XElement e in document.mainDoc.Descendants(XName.Get("fldSimple", w.NamespaceName)))
        {
          string attr_value = e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Replace(" ", string.Empty).Trim().ToLower();
          
          if (Regex.IsMatch(attr_value, matchPattern))
          {
            XElement firstRun = e.Element(w + "r");
            XElement firstText = firstRun.Element(w + "t");
            XElement rPr = firstText.Element(w + "rPr");

            // Delete everything and insert updated text value
            e.RemoveNodes();

            XElement t = new XElement(w + "t", rPr, corePropertyValue);
            Novacode.Text.PreserveSpace(t);
            e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t));
          }
        }

        #region Headers

        IEnumerable<PackagePart> headerParts = from headerPart in document.package.GetParts()
                                               where (Regex.IsMatch(headerPart.Uri.ToString(), @"/word/header\d?.xml"))
                                               select headerPart;
        foreach (PackagePart pp in headerParts)
        {
          XDocument header = XDocument.Load(new StreamReader(pp.GetStream()));

          foreach (XElement e in header.Descendants(XName.Get("fldSimple", w.NamespaceName)))
          {
            string attr_value = e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Replace(" ", string.Empty).Trim().ToLower();
            if (Regex.IsMatch(attr_value, matchPattern))
            {
              XElement firstRun = e.Element(w + "r");

              // Delete everything and insert updated text value
              e.RemoveNodes();

              XElement t = new XElement(w + "t", corePropertyValue);
              Novacode.Text.PreserveSpace(t);
              e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t));
            }
          }

          using (TextWriter tw = new StreamWriter(pp.GetStream(FileMode.Create, FileAccess.Write)))
            header.Save(tw);
        }
        #endregion

        #region Footers
        IEnumerable<PackagePart> footerParts = from footerPart in document.package.GetParts()
                                               where (Regex.IsMatch(footerPart.Uri.ToString(), @"/word/footer\d?.xml"))
                                               select footerPart;
        foreach (PackagePart pp in footerParts)
        {
          XDocument footer = XDocument.Load(new StreamReader(pp.GetStream()));

          foreach (XElement e in footer.Descendants(XName.Get("fldSimple", w.NamespaceName)))
          {
            string attr_value = e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Replace(" ", string.Empty).Trim().ToLower();
            if (Regex.IsMatch(attr_value, matchPattern))
            {
              XElement firstRun = e.Element(w + "r");

              // Delete everything and insert updated text value
              e.RemoveNodes();

              XElement t = new XElement(w + "t", corePropertyValue);
              Novacode.Text.PreserveSpace(t);
              e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t));
            }
          }

          using (TextWriter tw = new StreamWriter(pp.GetStream(FileMode.Create, FileAccess.Write)))
            footer.Save(tw);
        }
        #endregion
        PopulateDocument(document, document.package);
      }

        /// <summary>
        /// Add a custom property to this document. If a custom property already exists with the same name it will be replace. CustomProperty names are case insensitive.
        /// </summary>
        /// <param name="cp">The CustomProperty to add to this document.</param>
        /// <example>
        /// Add a custom properties of each type to a document.
        /// <code>
        /// // Load Example.docx
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // A CustomProperty called forename which stores a string.
        ///     CustomProperty forename;
        ///
        ///     // If this document does not contain a custom property called 'forename', create one.
        ///     if (!document.CustomProperties.ContainsKey("forename"))
        ///     {
        ///         // Create a new custom property called 'forename' and set its value.
        ///         document.AddCustomProperty(new CustomProperty("forename", "Cathal"));
        ///     }
        ///
        ///     // Get this documents custom property called 'forename'.
        ///     forename = document.CustomProperties["forename"];
        ///
        ///     // Print all of the information about this CustomProperty to Console.
        ///     Console.WriteLine(string.Format("Name: '{0}', Value: '{1}'\nPress any key...", forename.Name, forename.Value));
        ///     
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// } // Release this document from memory.
        ///
        /// // Wait for the user to press a key before exiting.
        /// Console.ReadKey();
        /// </code>
        /// </example>
        /// <seealso cref="CustomProperty"/>
        /// <seealso cref="CustomProperties"/>
        public void AddCustomProperty(CustomProperty cp)
        {
            // If this document does not contain a customFilePropertyPart create one.
            if(!package.PartExists(new Uri("/docProps/custom.xml", UriKind.Relative)))
                HelperFunctions.CreateCustomPropertiesPart(this);

            XDocument customPropDoc;
            PackagePart customPropPart = package.GetPart(new Uri("/docProps/custom.xml", UriKind.Relative));
            using (TextReader tr = new StreamReader(customPropPart.GetStream(FileMode.Open, FileAccess.Read)))
                customPropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);

            // Each custom property has a PID, get the highest PID in this document.
            IEnumerable<int> pids =
            (
                from d in customPropDoc.Descendants()
                where d.Name.LocalName == "property"
                select int.Parse(d.Attribute(XName.Get("pid")).Value)
            );

            int pid = 1;
            if (pids.Count() > 0)
                pid = pids.Max();

            // Check if a custom property already exists with this name
            var customProperty =
            (
                from d in customPropDoc.Descendants()
                where (d.Name.LocalName == "property") && (d.Attribute(XName.Get("name")).Value == cp.Name)
                select d
            ).SingleOrDefault();

            // If a custom property with this name already exists remove it.
            if (customProperty != null)
                customProperty.Remove();

            XElement propertiesElement = customPropDoc.Element(XName.Get("Properties", customPropertiesSchema.NamespaceName));
            propertiesElement.Add
            (
                new XElement
                (
                    XName.Get("property", customPropertiesSchema.NamespaceName),
                    new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
                    new XAttribute("pid", pid + 1),
                    new XAttribute("name", cp.Name),
                        new XElement(customVTypesSchema + cp.Type, cp.Value)
                )
            );

            // Save the custom properties
            using (TextWriter tw = new StreamWriter(customPropPart.GetStream(FileMode.Create, FileAccess.Write)))
                customPropDoc.Save(tw, SaveOptions.DisableFormatting);

            // Refresh all fields in this document which display this custom property.
            UpdateCustomPropertyValue(this, cp.Name, cp.Value.ToString());
        }

        

        internal static void UpdateCustomPropertyValue(DocX document, string customPropertyName, string customPropertyValue)
        {
            foreach (XElement e in document.mainDoc.Descendants(XName.Get("fldSimple", w.NamespaceName)))
            {
                string attr_value = e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Replace(" ", string.Empty).Trim();
                string match_value = string.Format(@"DOCPROPERTY  {0}  \* MERGEFORMAT", customPropertyName).Replace(" ", string.Empty);

                if (attr_value.Equals(match_value, StringComparison.CurrentCultureIgnoreCase))
                {
                    XElement firstRun = e.Element(w + "r");              
                    XElement firstText = firstRun.Element(w + "t");
                    XElement rPr = firstText.Element(w + "rPr");

                    // Delete everything and insert updated text value
                    e.RemoveNodes();

                    XElement t = new XElement(w + "t", rPr, customPropertyValue);
                    Novacode.Text.PreserveSpace(t);
                    e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t));
                }
            }

            //#region Headers
            //foreach(PackagePart pp in document.headers)
            //{
            //    XDocument header = XDocument.Load(new StreamReader(pp.GetStream()));

            //    foreach (XElement e in header.Descendants(XName.Get("fldSimple", w.NamespaceName)))
            //    {
            //        if (e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Trim().Equals(string.Format(@"DOCPROPERTY  {0}  \* MERGEFORMAT", customPropertyName), StringComparison.CurrentCultureIgnoreCase))
            //        {
            //            XElement firstRun = e.Element(w + "r");

            //            // Delete everything and insert updated text value
            //            e.RemoveNodes();

            //            XElement t = new XElement(w + "t", customPropertyValue);
            //            Novacode.Text.PreserveSpace(t);
            //            e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t));
            //        }
            //    }

            //    using (TextWriter tw = new StreamWriter(pp.GetStream()))
            //        header.Save(tw);
            //} 
            //#endregion

            //#region Footers
            //foreach (PackagePart pp in document.footers)
            //{
            //    XDocument footer = XDocument.Load(new StreamReader(pp.GetStream()));

            //    foreach (XElement e in footer.Descendants(XName.Get("fldSimple", w.NamespaceName)))
            //    {
            //        if (e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Trim().Equals(string.Format(@"DOCPROPERTY  {0}  \* MERGEFORMAT", customPropertyName), StringComparison.CurrentCultureIgnoreCase))
            //        {
            //            XElement firstRun = e.Element(w + "r");

            //            // Delete everything and insert updated text value
            //            e.RemoveNodes();

            //            XElement t = new XElement(w + "t", customPropertyValue);
            //            Novacode.Text.PreserveSpace(t);
            //            e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t));
            //        }
            //    }

            //    using (TextWriter tw = new StreamWriter(pp.GetStream()))
            //        footer.Save(tw);
            //}
            //#endregion
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

        public override List<Paragraph> Paragraphs
        {
            get
            {
                List<Paragraph> l = base.Paragraphs;
                l.ForEach(x => x.PackagePart = mainPart);
                return l;
            }
        }

        #region IDisposable Members

        /// <summary>
        /// Releases all resources used by this document.
        /// </summary>
        /// <example>
        /// If you take advantage of the using keyword, Dispose() is automatically called for you.
        /// <code>
        /// // Load document.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///      // The document is only in memory while in this scope.
        ///
        /// }// Dispose() is automatically called at this point.
        /// </code>
        /// </example>
        /// <example>
        /// This example is equilivant to the one above example.
        /// <code>
        /// // Load document.
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        /// 
        /// // Do something with the document here.
        ///
        /// // Dispose of the document.
        /// document.Dispose();
        /// </code>
        /// </example>
        public void Dispose()
        {
            package.Close();
        }

        #endregion
    }
}