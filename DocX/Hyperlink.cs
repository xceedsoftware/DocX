using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Novacode
{
    public class Hyperlink: DocXElement
    {
        private string text;
        public string Text { get{return text;} set{text = value;} }

        private Uri uri;
        public Uri Uri { get{return uri;} set{uri = value;} }

        internal Hyperlink(DocX document, XElement i): base(document, i)
        {

        }
    }
}
