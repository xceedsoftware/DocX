using System;
using System.IO.Packaging;
using System.IO;

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

        public Stream GetStream(FileMode mode, FileAccess access)
        {
            string temp = pr.SourceUri.OriginalString;
            string start = temp.Remove(temp.LastIndexOf('/'));
            string end = pr.TargetUri.OriginalString;
            string full = start + "/" + end;

            return(document.package.GetPart(new Uri(full, UriKind.Relative)).GetStream(mode, access));
        }

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
        public Picture CreatePicture(int height, int width) {
            Picture picture = Paragraph.CreatePicture(document, id, string.Empty, string.Empty);
            picture.Height = height;
            picture.Width = width;
            return picture;
        }

      ///<summary>
      /// Returns the name of the image file.
      ///</summary>
      public string FileName
      {
        get
        {
          return Path.GetFileName(this.pr.TargetUri.ToString());
        }
      }
    }
}
