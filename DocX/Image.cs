using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO.Packaging;

namespace Novacode
{
    /// <summary>
    /// Represents an Image embedded in a document.
    /// </summary>
    public class Image
    {
        /// <summary>
        /// A unique id which identifies this Image.
        /// </summary>
        private string id;
        private DocX document;
        internal PackageRelationship pr;

        /// <summary>
        /// Returns the id of this Image.
        /// </summary>
        public string Id 
        { 
            get {return id;} 
        }

        internal Image(DocX document, PackageRelationship pr)
        {
            this.document = document;
            this.pr = pr;
            this.id = pr.Id;
        }

        /// <summary>
        /// Add an image to a document, create a custom view of that image (picture) and then insert it into a Paragraph using append.
        /// </summary>
        /// <returns></returns>
        /// <example>
        /// Add an image to a document, create a custom view of that image (picture) and then insert it into a Paragraph using append.
        /// <code>
        /// using (DocX document = DocX.Create("Test.docx"))
        /// {
        ///    // Add an image to the document. 
        ///    Image     i = document.AddImage(@"Image.jpg");
        ///    
        ///    // Create a picture i.e. (A custom view of an image)
        ///    Picture   p = i.CreatePicture();
        ///    p.FlipHorizontal = true;
        ///    p.Rotation = 10;
        ///
        ///    // Create a new Paragraph.
        ///    Paragraph par = document.InsertParagraph();
        ///    
        ///    // Append content to the Paragraph.
        ///    par.Append("Here is a cool picture")
        ///       .AppendPicture(p)
        ///       .Append(" don't you think so?");
        ///
        ///    // Save all changes made to this document.
        ///    document.Save();
        /// }
        /// </code>
        /// </example>
        public Picture CreatePicture()
        {
            return Paragraph.CreatePicture(document, id, string.Empty, string.Empty);
        }
    }
}
