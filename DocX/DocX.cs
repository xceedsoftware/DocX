using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;
using System.Xml;
using System.IO;
using System.Text.RegularExpressions;

namespace Novacode
{
    /// <summary>
    /// Represents a .docx file.
    /// </summary>
    public class DocX: IDisposable
    {
        static internal string editSessionID;

        static internal XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        /// <summary>
        /// Default constructor
        /// </summary>
        internal DocX()
        {
            editSessionID = Guid.NewGuid().ToString().Substring(0, 8);
        }

        #region Private static variable declarations
        // The URI of this .docx
        private static string uri;
        // This memory stream will hold the DocX file
        private static MemoryStream document;
        // Object representation of the .docx
        private static WordprocessingDocument wdDoc;
        // Object representation of the \word\document.xml part
        internal static MainDocumentPart mainDocumentPart;
        // Object representation of the \docProps\custom.xml part
        private static CustomFilePropertiesPart customPropertiesPart = null;
        // The mainDocument is loaded into a XDocument object for easy querying and editing
        private static XDocument mainDoc;
        // The customPropertyDocument is loaded into a XDocument object for easy querying and editing
        private static XDocument customPropDoc;
        // The collection of Paragraphs <w:p></w:p> in this .docx
        private static IEnumerable<Paragraph> paragraphs;
        // The collection of custom properties in this .docx
        private static IEnumerable<CustomProperty> customProperties;
        private static IEnumerable<Image> images;
        private static XNamespace customPropertiesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        private static XNamespace customVTypesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

        private Random random = new Random();
        #endregion
        
        /// <summary>
        /// Gets the paragraphs of the .docx file.
        /// </summary>
        /// <seealso cref="Paragraph.Remove"/>
        /// <seealso cref="Paragraph.Insert"/>
        /// <seealso cref="Paragraph.Replace"/>
        /// <example>
        /// <code>
        /// // Load Test.docx
        /// DocX dx = DocX.Load(@"C:\Example.docx");
        ///    
        /// // Iterate through the paragraphs
        /// foreach(Paragraph p in dx.Paragraphs)
        /// {
        ///     string paragraphText = p.Value;
        /// }
        /// </code>
        /// </example>
        public IEnumerable<Paragraph> Paragraphs
        {
            get { return paragraphs; }      
        }

        /// <summary>
        /// Gets the custom properties of the .docx file.
        /// </summary>
        /// <seealso cref="SetCustomProperty"/>
        /// <example>
        /// <code>
        /// // Load Example.docx
        /// DocX dx = DocX.Load(@"C:\Example.docx");
        /// 
        /// // Iterate through the custom properties
        /// foreach(CustomProperty cp in dx.CustomProperties)
        /// {
        ///     string customPropertyName = cp.Name;
        ///     CustomPropertyType customPropertyType = cp.Type;
        ///     object customPropertyValue = cp.Value;
        /// }
        /// </code>
        /// </example>
        public IEnumerable<CustomProperty> CustomProperties
        {
            get { return customProperties; }
        }

        static internal void RebuildParagraphs()
        {
            // Get all of the paragraphs in this document
            DocX.paragraphs = from p in mainDoc.Descendants(XName.Get("p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"))
                              select new Paragraph(p);
        }

        public Paragraph AddParagraph()
        {
            XElement newParagraph = new XElement(w + "p");
            mainDoc.Descendants(XName.Get("body", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")).Single().Add(newParagraph);

            RebuildParagraphs();

            return new Paragraph(newParagraph);
        }

        public static DocX Create(string uri)
        {
            FileInfo fi = new FileInfo(uri);

            if (fi.Extension != ".docx")
                throw new Exception(string.Format("The input file {0} is not a .docx file", fi.FullName));

            DocX.uri = uri;

            // Create the docx package
            wdDoc = WordprocessingDocument.Create(uri, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            #region MainDocumentPart
            // Create the main document part for this package
            mainDocumentPart = wdDoc.AddMainDocumentPart();  
                       
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

            // Get all of the paragraphs in this document
            DocX.paragraphs = from p in mainDoc.Descendants(XName.Get("p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"))
                              select new Paragraph(p);
            #endregion

            #region CustomFilePropertiesPart

            // Get the custom file properties part from the package
            customPropertiesPart = wdDoc.CustomFilePropertiesPart;

            // This docx contains a customFilePropertyPart
            if (customPropertiesPart != null)
            {
                // Load the customFilePropertyPart
                using (TextReader tr = new StreamReader(customPropertiesPart.GetStream(FileMode.Open, FileAccess.Read)))
                {
                    customPropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);
                }

                // Get all of the custom properties in this document
                DocX.customProperties = from cp in customPropDoc.Descendants(XName.Get("property", customPropertiesSchema.NamespaceName))
                                        select new CustomProperty(cp);
            }

            #endregion
            
            // Save the new docx file to disk and return
            DocX created = new DocX();
            created.Save();
            return created;
        }

        /// <summary>
        /// Loads a .docx file into a DocX object.
        /// </summary>
        /// <param name="uri">The fully qualified name of the .docx file, or the relative file name.</param>
        /// <returns>
        /// Returns a DocX object which represents the .docx file.
        /// </returns>
        /// <example>
        /// <code>
        /// // Load Example.docx
        /// DocX dx = DocX.Load(@"C:\Example.docx");
        /// </code>
        /// </example>
        public static DocX Load(string uri)
        {           
            FileInfo fi = new FileInfo(uri);

            if (fi.Extension != ".docx")
                throw new Exception(string.Format("The input file {0} is not a .docx file", fi.FullName));

            if (!fi.Exists)
                throw new Exception(string.Format("The input file {0} does not exist", fi.FullName));

            DocX.uri = uri;

            // Open the docx package
            wdDoc = WordprocessingDocument.Open(uri, true);

            #region MainDocumentPart
            // Get the main document part from the package
            mainDocumentPart = wdDoc.MainDocumentPart;

            // Load the document part into a XDocument object
            using (TextReader tr = new StreamReader(mainDocumentPart.GetStream(FileMode.Open, FileAccess.Read)))
            {
                mainDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);
            }

            // Get all of the paragraphs in this document
            DocX.paragraphs = from p in mainDoc.Descendants(XName.Get("p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"))
                              select new Paragraph(p);

            DocX.images = from i in mainDocumentPart.ImageParts
                          select new Image(i);
            #endregion

            #region CustomFilePropertiesPart
                    
            // Get the custom file properties part from the package
            customPropertiesPart = wdDoc.CustomFilePropertiesPart;
            
            // This docx contains a customFilePropertyPart
            if (customPropertiesPart != null)
            {
                // Load the customFilePropertyPart
                using (TextReader tr = new StreamReader(customPropertiesPart.GetStream(FileMode.Open, FileAccess.Read)))
                {
                    customPropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);
                }

                // Get all of the custom properties in this document
                DocX.customProperties = from cp in customPropDoc.Descendants(XName.Get("property", customPropertiesSchema.NamespaceName))
                                        select new CustomProperty(cp);
            }

            #endregion
            
            return new DocX();
        }

        public Image AddImage(string filename)
        {
            ImagePart ip = mainDocumentPart.AddImagePart(ImagePartType.Jpeg);

            using (Stream stream = ip.GetStream(FileMode.Create, FileAccess.Write))
            {
                using (Stream s = new FileStream(filename, FileMode.Open))
                {
                    byte[] bytes = new byte[s.Length];
                    s.Read(bytes, 0, (int)s.Length);
                    stream.Write(bytes, 0, (int)s.Length);
                }
            }

            return new Image(ip);
        }

        public Image AddImage(Stream s)
        {
            ImagePart ip = mainDocumentPart.AddImagePart(ImagePartType.Jpeg);

            using (Stream stream = ip.GetStream(FileMode.Create, FileAccess.Write))
            {
                byte[] bytes = new byte[s.Length];
                s.Read(bytes, 0, (int)s.Length);
                stream.Write(bytes, 0, (int)s.Length);
            }

            return new Image(ip);
        }

        /// <summary>
        /// Saves all changes made to the DocX object back to the .docx file it represents.
        /// </summary>
        /// <example>
        /// <code>
        /// // Load Example.docx
        /// DocX dx = DocX.Load(@"C:\Example.docx");
        /// 
        /// // Insert code to do something useful with dx
        /// 
        /// // Save changes to Example.docx
        /// dx.Save();
        /// </code>
        /// </example>
        public void Save()
        {
            // Save the main document part back to disk
            using (TextWriter tw = new StreamWriter(mainDocumentPart.GetStream(FileMode.Create, FileAccess.Write)))
            {
                mainDoc.Save(tw, SaveOptions.DisableFormatting);
            }

            if (customPropertiesPart != null)
            {
                // Save the custom properties part back to disk
                using (TextWriter tw = new StreamWriter(customPropertiesPart.GetStream(FileMode.Create, FileAccess.Write)))
                {
                    customPropDoc.Save(tw, SaveOptions.DisableFormatting);
                }
            }

            // Save and close the .docx file
            wdDoc.Close();

            // Reopen the .docx file
            Load(uri);
        }

        #region IDisposable Members
        /// <summary>
        /// Releases all resources used by Novacode.DocX object.
        /// </summary>
        public void Dispose()
        {
            wdDoc.Dispose();
        }

        #endregion

        /// <summary>
        /// Set the value of a custom property. If no custom property exists with the name customPropertyName, then a new custom property is created. 
        /// </summary>
        /// <param name="customPropertyName">The name of the custom property to set</param>
        /// <param name="customPropertyType">The type of the custom property</param>
        /// <param name="customPropertyValue">The value of the custom property</param>
        /// <example>
        /// <code>
        /// // Load Example.docx
        /// DocX dx = DocX.Load(@"C:\Example.docx");
        ///
        /// // Add custom property 'Name'
        /// dx.SetCustomProperty("Name", CustomPropertyType.Text, "Helium");
        ///
        /// // Add custom property 'Date discovered'
        /// dx.SetCustomProperty("Date discovered", CustomPropertyType.Date, new DateTime(1868, 08, 18));
        ///
        /// // Add the custom property 'Noble gas'
        /// dx.SetCustomProperty("Noble gas", CustomPropertyType.YesOrNo, true);
        ///
        /// // Add the custom property 'Atomic number'
        /// dx.SetCustomProperty("Atomic number", CustomPropertyType.NumberInteger, 2);
        ///
        /// // Add the custom property 'Boiling point'
        /// dx.SetCustomProperty("Boiling point", CustomPropertyType.NumberDecimal, -268.93);
        ///
        /// // Save changes to Example.docx
        /// dx.Save(); 
        /// </code>
        /// <code>
        /// // Load Example.docx
        /// DocX dx = DocX.Load(@"C:\Example.docx");
        /// 
        /// // Add the custom property 'LCID'
        /// dx.SetCustomProperty("LCID", CustomPropertyType.NumberInteger, 1036);
        ///
        /// // Save changes to Example.docx
        /// dx.Save(); 
        /// 
        /// // Update the custom property 'LCID' with a new value
        /// dx.SetCustomProperty("LCID", CustomPropertyType.NumberInteger, 1041);
        /// 
        /// // Save changes to Example.docx
        /// dx.Save(); 
        /// </code>
        /// </example>
        public void SetCustomProperty(string customPropertyName, CustomPropertyType customPropertyType, object customPropertyValue)
        {
            string type = "";
            string value = "";

            switch (customPropertyType)
            {
                case CustomPropertyType.Text:
                    {
                        if (!(customPropertyValue is string))
                            throw new Exception("Not a string");

                        type = "lpwstr";
                        value = customPropertyValue as string;
                        break;
                    }

                case CustomPropertyType.Date:
                    {
                        if (!(customPropertyValue is DateTime))
                            throw new Exception("Not a DateTime");

                        type = "filetime";
                        // Must be UTC time
                        value = (((DateTime)customPropertyValue).ToUniversalTime()).ToString();
                        break;
                    }

                case CustomPropertyType.YesOrNo:
                    {
                        if (!(customPropertyValue is bool))
                            throw new Exception("Not a Boolean");

                        type = "bool";
                        // Must be lower case either {true or false}
                        value = (((bool)customPropertyValue)).ToString().ToLower();
                        break;
                    }

                case CustomPropertyType.NumberInteger:
                    {
                        if (!(customPropertyValue is int))
                            throw new Exception("Not an int");

                        type = "i4";
                        value = customPropertyValue.ToString();
                        break;
                    }

                case CustomPropertyType.NumberDecimal:
                    {
                        if (!(customPropertyValue is double))
                            throw new Exception("Not a double");

                        type = "r8";
                        value = customPropertyValue.ToString();
                        break;
                    }
            }

            // This docx does not contain a customFilePropertyPart
            if (customPropertiesPart == null)
            {
                customPropertiesPart = wdDoc.AddCustomFilePropertiesPart();

                customPropDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
                    new XElement(XName.Get("Properties", customPropertiesSchema.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "vt", customVTypesSchema),
                        new XElement(XName.Get("property", customPropertiesSchema.NamespaceName),
                            new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
                            new XAttribute("pid", "2"),
                            new XAttribute("name", customPropertyName),
                                new XElement(customVTypesSchema + type, customPropertyValue)
                )));
            }

            // This docx contains a customFilePropertyPart
            else
            {
                // Get the highest PID
                int pid = (from d in customPropDoc.Descendants()
                           where d.Name.LocalName == "property"
                           select int.Parse(d.Attribute(XName.Get("pid")).Value)).Max<int>();

                // Get the custom property or null
                var customProperty = (from d in customPropDoc.Descendants()
                                      where (d.Name.LocalName == "property") && (d.Attribute(XName.Get("name")).Value == customPropertyName)
                                      select d).SingleOrDefault();

                if (customProperty != null)
                {
                    customProperty.Descendants(XName.Get(type, customVTypesSchema.NamespaceName)).SingleOrDefault().ReplaceWith(
                                new XElement(customVTypesSchema + type, customPropertyValue));
                }

                else
                {
                    XElement propertiesElement = customPropDoc.Element(XName.Get("Properties", customPropertiesSchema.NamespaceName));
                    propertiesElement.Add(new XElement(XName.Get("property", customPropertiesSchema.NamespaceName),
                            new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
                            new XAttribute("pid", pid + 1),
                            new XAttribute("name", customPropertyName),
                                new XElement(customVTypesSchema + type, customPropertyValue)
                            ));
                }
            }

            UpdateCustomPropertyValue(customPropertyName, customPropertyValue.ToString());
        }

        private static void UpdateCustomPropertyValue(string customPropertyName, string customPropertyValue)
        {
            foreach (XElement e in mainDoc.Descendants(XName.Get("fldSimple", w.NamespaceName)))
            {
                if (e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Equals(string.Format(@" DOCPROPERTY  {0}  \* MERGEFORMAT ", customPropertyName), StringComparison.CurrentCultureIgnoreCase))
                {
                    XElement firstRun = e.Element(w + "r");

                    // Delete everything and insert updated text value
                    e.RemoveNodes();
 
                    XElement t = new XElement(w + "t", customPropertyValue);
                    Text.PreserveSpace(t);
                    e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t));
                }
            }
        }

        /// <summary>
        /// Renumber all ins and del ids in this .docx file.
        /// </summary>
        internal static void RenumberIDs()
        {
            IEnumerable<XAttribute> trackerIDs =
                            (from d in mainDoc.Descendants()
                             where d.Name.LocalName == "ins" || d.Name.LocalName == "del"
                             select d.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")));

            for (int i = 0; i < trackerIDs.Count(); i++)
                trackerIDs.ElementAt(i).Value = i.ToString();
        }
    }
}