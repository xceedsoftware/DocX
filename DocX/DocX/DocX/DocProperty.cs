using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace Novacode
{
    /// <summary>
    /// Represents a field of type document property. This field displays the value stored in a custom property.
    /// </summary>
    public class DocProperty
    {
        internal Regex extractName = new Regex(@"DOCPROPERTY  (?<name>.*)  ");
        internal XElement xml;
        private string name;

        /// <summary>
        /// The custom property to display.
        /// </summary>
        public string Name { get { return name; } }

        internal DocProperty(XElement xml)
        {
            this.xml = xml;
            
            string instr = xml.Attribute(XName.Get("instr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")).Value;
            this.name = extractName.Match(instr.Trim()).Groups["name"].Value;
        }
    }
}
