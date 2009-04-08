using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace Novacode
{
    public class Image
    {
        private string id;

        public string Id 
        { 
            get {return id;} 
        }

        internal Image(ImagePart ip)
        {
            id = DocX.mainDocumentPart.GetIdOfPart(ip);    
        }
    }
}
