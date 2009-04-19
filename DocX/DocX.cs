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
    /// Represents a document.
    /// </summary>
    public class DocX
    {
        // A lookup for the runs in this paragraph
        internal static Dictionary<int, Paragraph> paragraphLookup = new Dictionary<int, Paragraph>();

        static internal string editSessionID;

        static internal XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        internal DocX()
        {
            editSessionID = Guid.NewGuid().ToString().Substring(0, 8);
        }

        #region Private static variable declarations
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
        private static List<Paragraph> paragraphs = new List<Paragraph>();
        // The collection of custom properties in this .docx
        private static List<CustomProperty> customProperties;
        private static List<Image> images;
        private static XNamespace customPropertiesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        private static XNamespace customVTypesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

        private Random random = new Random();
        #endregion
        
        /// <summary>
        /// Returns a list of Paragraphs in this document.
        /// </summary>
        /// <example>
        /// Write to Console the Text from each Paragraph in this document.
        /// <code>
        /// // Load a document
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        ///
        /// // Loop through each Paragraph in this document.
        /// foreach (Paragraph p in document.Paragraphs)
        /// {
        ///     // Write this Paragraphs Text to Console.
        ///     Console.WriteLine(p.Text);
        /// }
        ///
        /// // Always close your document when you are finished with it.
        /// document.Close(false);
        ///
        /// // Wait for the user to press a key before closing the console window.
        /// Console.ReadKey();
        /// </code>
        /// </example>
        /// <seealso cref="Paragraph.InsertText(string, bool)"/>
        /// <seealso cref="Paragraph.InsertText(int, string, bool)"/>
        /// <seealso cref="Paragraph.InsertText(string, bool, Formatting)"/>
        /// <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
        /// <seealso cref="Paragraph.RemoveText(int, bool)"/>
        /// <seealso cref="Paragraph.RemoveText(int, int, bool)"/>
        /// <seealso cref="Paragraph.ReplaceText(string, string, bool)"/>
        /// <seealso cref="Paragraph.ReplaceText(string, string, bool, RegexOptions)"/>
        /// <seealso cref="Paragraph.InsertPicture"/>
        public List<Paragraph> Paragraphs
        {
            get { return paragraphs; }      
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
        /// // Always close your document when you are finished with it.
        /// document.Close(false);
        /// </code>
        /// </example>
        /// <seealso cref="AddImage(string)"/>
        /// <seealso cref="AddImage(Stream)"/>
        /// <seealso cref="Paragraph.Pictures"/>
        /// <seealso cref="Paragraph.InsertPicture"/>
        public List<Image> Images
        {
            get { return images; }
        }

        /// <summary>
        /// Returns a list of custom properties in this document.
        /// </summary>
        /// <example>
        /// Get the name, type and value of each CustomProperty in this document.
        /// <code>
        /// // Load Example.docx
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        ///
        /// // Loop through each CustomProperty in this document
        /// foreach (CustomProperty cp in document.CustomProperties)
        /// {
        ///     string name = cp.Name;
        ///     CustomPropertyType type = cp.Type;
        ///     object value = cp.Value;
        /// }
        ///
        /// // Always close your document when you are finished with it.
        /// document.Close(false);
        /// </code>
        /// </example>
        /// <seealso cref="AddCustomProperty"/>
        public List<CustomProperty> CustomProperties
        {
            get { return customProperties; }
        }

        static internal void RebuildParagraphs()
        {
            paragraphLookup.Clear();
            paragraphs.Clear();

            // Get the runs in this paragraph
            IEnumerable<XElement> paras = mainDoc.Descendants(XName.Get("p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"));

            int startIndex = 0;

            // Loop through each run in this paragraph
            foreach (XElement par in paras)
            {
                Paragraph xp = new Paragraph(startIndex, par);
                
                // Add to paragraph list
                paragraphs.Add(xp);

                // Only add runs which contain text
                if (Paragraph.GetElementTextLength(par) > 0)
                {
                    paragraphLookup.Add(xp.endIndex, xp);
                    startIndex = xp.endIndex;
                }
            }
        }

        /// <summary>
        /// Insert a new Paragraph at the end of this document.
        /// </summary>
        /// <param name="text">The text of this Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <returns>A new Paragraph.</returns>
        /// <example>
        /// Inserting a new Paragraph at the end of a document with text formatting.
        /// <code>
        /// // Load a document.
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        /// 
        /// // Insert a new Paragraph at the end of this document.
        /// document.InsertParagraph("New text", false);
        ///
        /// // Always close your document when you are finished with it.
        /// document.Close(true);
        /// </code>
        /// </example>
        public Paragraph InsertParagraph(string text, bool trackChanges)
        {
            int index = 0;
            if (paragraphLookup.Keys.Count() > 0)
                index = paragraphLookup.Last().Key;

            return InsertParagraph(index, text, trackChanges, null);
        }
        
        /// <summary>
        /// Insert a new Paragraph at the end of a document with text formatting.
        /// </summary>
        /// <param name="text">The text of this Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <param name="formatting">The formatting for the text of this Paragraph.</param>
        /// <returns>A new Paragraph.</returns>
        /// <example>
        /// Inserting a new Paragraph at the end of a document with text formatting.
        /// <code>
        /// // Load a document.
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        ///
        /// // Create a Formatting object
        /// Formatting formatting = new Formatting();
        /// formatting.Bold = true;
        /// formatting.FontColor = Color.Red;
        /// formatting.Size = 30;
        /// 
        /// // Insert a new Paragraph at the end of this document with text formatting.
        /// document.InsertParagraph("New text", false, formatting);
        ///
        /// // Always close your document when you are finished with it.
        /// document.Close(true);
        /// </code>
        /// </example>
        public Paragraph InsertParagraph(string text, bool trackChanges, Formatting formatting)
        {
            int index = 0;
            if (paragraphLookup.Keys.Count() > 0)
                index = paragraphLookup.Last().Key;

            return InsertParagraph(index, text, trackChanges, formatting);
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
                string text = string.Empty;
                foreach (Paragraph p in paragraphs)
                {
                    text += p.Text + "\n";
                }
                return text;
            }
        
        }
        /// <summary>
        /// Insert a new Paragraph into this document at a specified index.
        /// </summary>
        /// <param name="index">The character index to insert this document at.</param>
        /// <param name="text">The text of this Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <returns>A new Paragraph.</returns>
        /// <example>
        /// Insert a new Paragraph into the middle of a document.
        /// <code>
        /// // Load a document.
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        ///
        /// // Middle character index of this document.
        /// int index = document.Text.Length / 2;
        ///            
        /// // Insert a new Paragraph at the middle of this document.
        /// document.InsertParagraph(index, "New text", false);
        ///
        /// // Always close your document when you are finished with it.
        /// document.Close(true);
        /// </code>
        /// </example>
        public Paragraph InsertParagraph(int index, string text, bool trackChanges)
        {
            return InsertParagraph(index, text, trackChanges, null);
        }

        /// <summary>
        /// Insert a new Paragraph into this document at a specified index with text formatting.
        /// </summary>
        /// <param name="index">The character index to insert this document at.</param>
        /// <param name="text">The text of this Paragraph.</param>
        /// <param name="trackChanges">Should this insertion be tracked as a change?</param>
        /// <param name="formatting">The formatting for the text of this Paragraph.</param>
        /// <returns>A new Paragraph.</returns>
        /// /// <example>
        /// Insert a new Paragraph into the middle of a document with text formatting.
        /// <code>
        /// // Load a document.
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        ///
        /// // Create a Formatting object
        /// Formatting formatting = new Formatting();
        /// formatting.Bold = true;
        /// formatting.FontColor = Color.Red;
        /// formatting.Size = 30;
        /// 
        /// //  Middle character index of this document.
        /// int index = document.Text.Length / 2;
        ///
        /// // Insert a new Paragraph in the middle of this document.
        /// document.InsertParagraph(index, "New text", false, formatting);
        /// 
        /// // Always close your document when you are finished with it.
        /// document.Close(true);
        /// </code>
        /// </example>
        public Paragraph InsertParagraph(int index, string text, bool trackChanges, Formatting formatting)
        {
            Paragraph newParagraph = new Paragraph(index, new XElement(w + "p"));
            newParagraph.InsertText(0, text, trackChanges, formatting);

            Paragraph firstPar = GetFirstParagraphEffectedByInsert(index);

            if (firstPar != null)
            {
                XElement[] splitParagraph = SplitParagraph(firstPar, index - firstPar.startIndex);

                firstPar.xml.ReplaceWith
                (
                    splitParagraph[0],
                    newParagraph.xml,
                    splitParagraph[1]
                );
            }

            else
                mainDoc.Descendants(XName.Get("body", DocX.w.NamespaceName)).First().Add(newParagraph.xml);

            return newParagraph;
        }

        static internal Paragraph GetFirstParagraphEffectedByInsert(int index)
        {
            // This document contains no Paragraphs and insertion is at index 0
            if (paragraphLookup.Keys.Count() == 0 && index == 0)
                return null;

            foreach (int paragraphEndIndex in paragraphLookup.Keys)
            {
                if (paragraphEndIndex >= index)
                    return paragraphLookup[paragraphEndIndex];
            }

            throw new ArgumentOutOfRangeException();
        }
  
        internal XElement[] SplitParagraph(Paragraph p, int index)
        {
            Run r = p.GetFirstRunEffectedByInsert(index);

            XElement[] split;
            XElement before, after;

            if (r.xml.Parent.Name.LocalName == "ins")
            {
                split = p.SplitEdit(r.xml.Parent, index, EditType.ins);
                before = new XElement(p.xml.Name, p.xml.Attributes(), r.xml.Parent.ElementsBeforeSelf(), split[0]);
                after = new XElement(p.xml.Name, p.xml.Attributes(), r.xml.Parent.ElementsAfterSelf(), split[1]);
            }

            else if (r.xml.Parent.Name.LocalName == "del")
            {
                split = p.SplitEdit(r.xml.Parent, index, EditType.del);

                before = new XElement(p.xml.Name, p.xml.Attributes(), r.xml.Parent.ElementsBeforeSelf(), split[0]);
                after = new XElement(p.xml.Name, p.xml.Attributes(), r.xml.Parent.ElementsAfterSelf(), split[1]);
            }

            else
            {
                split = Run.SplitRun(r, index);

                before = new XElement(p.xml.Name, p.xml.Attributes(), r.xml.ElementsBeforeSelf(), split[0]);
                after = new XElement(p.xml.Name, p.xml.Attributes(), r.xml.ElementsAfterSelf(), split[1]);
            }

            if (before.Elements().Count() == 0)
                before = null;

            if (after.Elements().Count() == 0)
                after = null;

            return new XElement[] { before, after };
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
        ///     DocX document = DocX.Load(fs);
        /// 
        ///     // Do something with the document here.
        /// 
        ///     // Always close your document when you are finished with it.
        ///     document.Close(true);
        /// }
        /// </code>
        /// </example>
        /// <example>
        /// Creating a document in a SharePoint site.
        /// <code>
        /// // Get the SharePoint site that you want to access.
        /// using(SPSite mySite = new SPSite("http://server/sites/site"))
        /// {
        ///     // Open a connection to the SharePoint site
        ///     using(SPWeb myWeb = mySite.OpenWeb())
        ///     {
        ///         // Create a MemoryStream ms.
        ///         using (MemoryStream ms = new MemoryStream())
        ///         {
        ///             // Create a document using ms.
        ///             DocX document = DocX.Create(ms);
        ///
        ///             // Do something with the document here.
        /// 
        ///             // Always close your document when you are finished with it.
        ///             document.Close(true);
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
        /// <seealso cref="DocX.Close"/>
        public static DocX Create(Stream stream)
        {
            // Create the docx package
            wdDoc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            PostCreation();
            return DocX.Load(stream);
        }

        /// <summary>
        /// Creates a document using a fully qualified or relative filename.
        /// </summary>
        /// <param name="filename">The fully qualified or relative filename.</param>
        /// <returns>Returns a DocX object which represents the document.</returns>
        /// <example>
        /// <code>
        /// // Create a document using a fully qualified filename
        /// DocX document = DocX.Create(@"C:\Example\Test.docx");
        /// 
        /// // Do something with the document here.
        /// 
        /// // Always close your document when you are finished with it.
        /// document.Close(true);
        /// </code>
        /// <code>
        /// // Create a document using a relative filename.
        /// DocX document = DocX.Create(@"..\Test.docx");
        ///
        /// // Do something with the document here.
        ///
        /// // Always close your document when you are finished with it.
        /// document.Close(true);
        /// </code>
        /// <seealso cref="DocX.Create(System.IO.Stream)"/>
        /// <seealso cref="DocX.Load(System.IO.Stream)"/>
        /// <seealso cref="DocX.Load(string)"/>
        /// <seealso cref="DocX.Close"/>
        /// </example>
        public static DocX Create(string filename)
        {
            // Create the docx package
            wdDoc = WordprocessingDocument.Create(filename, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            PostCreation();
            return DocX.Load(filename);
        }

        private static void PostCreation()
        {
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

            RebuildParagraphs();
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
                DocX.customProperties = (from cp in customPropDoc.Descendants(XName.Get("property", customPropertiesSchema.NamespaceName))
                                        select new CustomProperty(cp)).ToList();
            }

            #endregion

            DocX created = new DocX();
            created.Close(true);
        }

        private static DocX PostLoad()
        {
            #region MainDocumentPart
            // Get the main document part from the package
            mainDocumentPart = wdDoc.MainDocumentPart;

            // Load the document part into a XDocument object
            using (TextReader tr = new StreamReader(mainDocumentPart.GetStream(FileMode.Open, FileAccess.Read)))
            {
                mainDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);
            }

            RebuildParagraphs();

            DocX.images = (from i in mainDocumentPart.ImageParts
                          select new Image(i)).ToList();
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
                DocX.customProperties = (from cp in customPropDoc.Descendants(XName.Get("property", customPropertiesSchema.NamespaceName))
                                        select new CustomProperty(cp)).ToList();
            }

            #endregion

            return new DocX();
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
        /// using(FileStream fs = new FileStream(@"C:\Example\Test.docx", FileMode.Open))
        /// {
        ///     // Load the document using fs
        ///     DocX document = DocX.Load(fs);
        /// 
        ///     // Do something with the document here.
        /// 
        ///     // Always close your document when you are finished with it.
        ///     document.Close(true);
        /// }
        /// </code>
        /// </example>
        /// <example>
        /// Loading a document from a SharePoint site.
        /// <code>
        /// // Get the SharePoint site that you want to access.
        /// using(SPSite mySite = new SPSite("http://server/sites/site"))
        /// {
        ///     // Open a connection to the SharePoint site
        ///     using(SPWeb myWeb = mySite.OpenWeb())
        ///     {
        ///         // Grab a document stored on this site.
        ///         SPFile file = web.GetFile("Source_Folder_Name/Source_File");
        ///         
        ///         // DocX.Load requires a Stream, so open a Stream to this document.
        ///         Stream str = new MemoryStream(file.OpenBinary());
        ///         
        ///         // Load the file using the Stream str.
        ///         DocX document = DocX.Load(str);
        ///    
        ///         // Do something with the document here.
        /// 
        ///         // Always close your document when you are finished with it.
        ///         document.Close(true);
        ///     }
        /// }
        /// </code>
        /// </example>
        /// <seealso cref="DocX.Load(string)"/>
        /// <seealso cref="DocX.Create(System.IO.Stream)"/>
        /// <seealso cref="DocX.Create(string)"/>
        /// <seealso cref="DocX.Close"/>
        public static DocX Load(Stream stream)
        {
            // Open the docx package
            wdDoc = WordprocessingDocument.Open(stream, true);

            return PostLoad();
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
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        /// 
        /// // Do something with the document here
        /// 
        /// // Always close your document when you are finished with it.
        /// document.Close(true);
        /// </code>
        /// <code>
        /// // Load a document using its relative filename.
        /// DocX document = DocX.Load(@"..\..\Test.docx");
        /// 
        /// // Do something with the document here.
        /// 
        /// // Always close your document when you are finished with it.
        /// document.Close(true);
        /// </code>
        /// <seealso cref="DocX.Load(System.IO.Stream)"/>
        /// <seealso cref="DocX.Create(System.IO.Stream)"/>
        /// <seealso cref="DocX.Create(string)"/>
        /// <seealso cref="DocX.Close"/>
        /// </example>
        public static DocX Load(string filename)
        {
            if (!File.Exists(filename))
                throw new FileNotFoundException(string.Format("File could not be found {0}"));

            // Open the docx package
            wdDoc = WordprocessingDocument.Open(filename, true);

            return PostLoad();
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
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        ///
        /// // Add an Image from a file.
        /// document.AddImage(@"C:\Example\Image.png");
        ///
        /// // Close the document.
        /// document.Close(true);
        /// </code>
        /// </example>
        /// <seealso cref="AddImage(System.IO.Stream)"/>
        /// <seealso cref="Paragraph.InsertPicture"/>
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
        ///     DocX document = DocX.Load(@"C:\Example\Test.docx");
        ///
        ///     // Add an Image from a filestream fs.
        ///     document.AddImage(fs);
        ///
        ///     // Close the document.
        ///     document.Close(true);
        /// }
        /// </code>
        /// </example>
        /// <seealso cref="AddImage(string)"/>
        /// <seealso cref="Paragraph.InsertPicture"/>
        public Image AddImage(Stream stream)
        {
            ImagePart ip = mainDocumentPart.AddImagePart(ImagePartType.Jpeg);

            using (Stream s = ip.GetStream(FileMode.Create, FileAccess.Write))
            {
                byte[] bytes = new byte[stream.Length];
                stream.Read(bytes, 0, (int)stream.Length);
                s.Write(bytes, 0, (int)stream.Length);
            }

            return new Image(ip);
        }

        /// <summary>
        /// Close the document and optionally save pending changes.
        /// </summary>
        /// <param name="saveChanges">Should pending changes be saved?</param>
        /// <example>
        /// <code>
        ///  // Load a document using its fully qualified filename.
        ///  DocX document = DocX.Load(@"C:\Example\Test.docx");
        ///
        ///  // Insert a new Paragraph
        ///  document.InsertParagraph("Hello world!", false);
        ///     
        ///  // Should pending changes be saved?
        ///  bool saveChanges = true;
        ///     
        ///  // Close the document.
        ///  document.Close(saveChanges);
        /// </code>
        /// </example>
        /// <seealso cref="DocX.Create(System.IO.Stream)"/>
        /// <seealso cref="DocX.Create(string)"/>
        /// <seealso cref="DocX.Load(System.IO.Stream)"/>
        /// <seealso cref="DocX.Load(string)"/>
        public void Close(bool saveChanges)
        {
            if (saveChanges)
            {
                if (mainDocumentPart != null)
                {
                    // Save the main document
                    using (TextWriter tw = new StreamWriter(mainDocumentPart.GetStream(FileMode.Create, FileAccess.Write)))
                        mainDoc.Save(tw, SaveOptions.DisableFormatting);
                }

                if (customPropertiesPart != null)
                {
                    // Save the custom properties
                    using (TextWriter tw = new StreamWriter(customPropertiesPart.GetStream(FileMode.Create, FileAccess.Write)))
                        customPropDoc.Save(tw, SaveOptions.DisableFormatting);
                }
            }

            // Close and dispose of the file
            wdDoc.Close();
            wdDoc.Dispose();
        }

        /// <summary>
        /// Adds a new custom property to this document. If a custom property already exists with this name its value is updated. 
        /// </summary>
        /// <param name="name">The name of the custom property.</param>
        /// <param name="type">The type of the custom property.</param>
        /// <param name="value">The value of the new custom property.</param>
        /// <example>
        /// Add one of each custom property type to a document.
        /// <code>
        /// // Load a document.
        /// DocX document = DocX.Load(@"C:\Example\Test.docx");
        ///
        /// // Add a custom property 'Name'.
        /// document.AddCustomProperty("Name", CustomPropertyType.Text, "Helium");
        ///
        /// // Add a custom property 'Date discovered'.
        /// document.AddCustomProperty("Date discovered", CustomPropertyType.Date, new DateTime(1868, 08, 18));
        ///
        /// // Add a custom property 'Noble gas'.
        /// document.AddCustomProperty("Noble gas", CustomPropertyType.YesOrNo, true);
        ///
        /// // Add a custom property 'Atomic number'.
        /// document.AddCustomProperty("Atomic number", CustomPropertyType.NumberInteger, 2);
        ///
        /// // Add a custom property 'Boiling point'.
        /// document.AddCustomProperty("Boiling point", CustomPropertyType.NumberDecimal, -268.93);
        ///
        /// // Close the document.
        /// document.Close(true);
        /// </code>
        /// </example>
        /// <seealso cref="CustomProperties"/>
        public void AddCustomProperty(string name, CustomPropertyType type, object value)
        {
            string typeString = string.Empty;
            string valueString = string.Empty;

            switch (type)
            {
                case CustomPropertyType.Text:
                    {
                        if (!(value is string))
                            throw new Exception("Not a string");

                        typeString = "lpwstr";
                        valueString = value as string;
                        break;
                    }

                case CustomPropertyType.Date:
                    {
                        if (!(value is DateTime))
                            throw new Exception("Not a DateTime");

                        typeString = "filetime";
                        // Must be UTC time
                        valueString = (((DateTime)value).ToUniversalTime()).ToString();
                        break;
                    }

                case CustomPropertyType.YesOrNo:
                    {
                        if (!(value is bool))
                            throw new Exception("Not a Boolean");

                        typeString = "bool";
                        // Must be lower case either {true or false}
                        valueString = (((bool)value)).ToString().ToLower();
                        break;
                    }

                case CustomPropertyType.NumberInteger:
                    {
                        if (!(value is int))
                            throw new Exception("Not an int");

                        typeString = "i4";
                        valueString = value.ToString();
                        break;
                    }

                case CustomPropertyType.NumberDecimal:
                    {
                        if (!(value is double))
                            throw new Exception("Not a double");

                        typeString = "r8";
                        valueString = value.ToString();
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
                            new XAttribute("name", name),
                                new XElement(customVTypesSchema + typeString, value)
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
                                      where (d.Name.LocalName == "property") && (d.Attribute(XName.Get("name")).Value == name)
                                      select d).SingleOrDefault();

                if (customProperty != null)
                {
                    customProperty.Descendants(XName.Get(typeString, customVTypesSchema.NamespaceName)).SingleOrDefault().ReplaceWith(
                                new XElement(customVTypesSchema + typeString, value));
                }

                else
                {
                    XElement propertiesElement = customPropDoc.Element(XName.Get("Properties", customPropertiesSchema.NamespaceName));
                    propertiesElement.Add(new XElement(XName.Get("property", customPropertiesSchema.NamespaceName),
                            new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
                            new XAttribute("pid", pid + 1),
                            new XAttribute("name", name),
                                new XElement(customVTypesSchema + typeString, value)
                            ));
                }
            }

            UpdateCustomPropertyValue(name, value.ToString());
        }

        internal static void UpdateCustomPropertyValue(string customPropertyName, string customPropertyValue)
        {
            foreach (XElement e in mainDoc.Descendants(XName.Get("fldSimple", w.NamespaceName)))
            {
                if (e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Equals(string.Format(@" DOCPROPERTY  {0}  \* MERGEFORMAT ", customPropertyName), StringComparison.CurrentCultureIgnoreCase))
                {
                    XElement firstRun = e.Element(w + "r");

                    // Delete everything and insert updated text value
                    e.RemoveNodes();
 
                    XElement t = new XElement(w + "t", customPropertyValue);
                    Novacode.Text.PreserveSpace(t);
                    e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t));
                }
            }
        }

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