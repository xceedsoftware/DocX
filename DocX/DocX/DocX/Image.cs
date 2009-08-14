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

        /// <summary>
        /// Returns the id of this Image.
        /// </summary>
        public string Id 
        { 
            get {return id;} 
        }

        internal Image(DocX document, PackageRelationship pr)
        {
            id = pr.Id;
        }
    }
}
