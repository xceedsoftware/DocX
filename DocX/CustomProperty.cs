using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Novacode
{
    /// <summary>
    /// Represents a document custom property.
    /// </summary>
    public class CustomProperty
    {
        // The underlying XElement which this CustomProperty wraps
        private XElement xml;
        // This customPropertys name
        private string name;
        // This customPropertys type
        private CustomPropertyType type;
        // This customPropertys value
        private object value;

        /// <summary>
        /// Custom Propertys name.
        /// </summary>
        public string Name { get { return name; } }

        /// <summary>
        /// Custom Propertys type.
        /// </summary>
        public CustomPropertyType Type { get { return type; } }

        /// <summary>
        /// Custom Propertys value.
        /// </summary>
        public object Value { get { return value; } }

        internal CustomProperty(XElement xml)
        {
            this.xml = xml;
            name = xml.Attribute(XName.Get("name")).Value;
            
            XElement p = xml.Elements().SingleOrDefault();

            switch (p.Name.LocalName)
            {
                case "i4":
                    type = CustomPropertyType.NumberInteger;
                    value = int.Parse(p.Value);
                    break;

                case "r8":
                    type = CustomPropertyType.NumberDecimal;
                    value = double.Parse(p.Value);
                    break;

                case "lpwstr":
                    type = CustomPropertyType.Text;
                    value = p.Value;
                    break;

                case "filetime":
                    type = CustomPropertyType.Date;
                    value = DateTime.Parse(p.Value);
                    break;
                
                case "bool":
                    type = CustomPropertyType.YesOrNo;
                    value = bool.Parse(p.Value);
                    break;

                default:
                    throw new Exception(string.Format("The custom property type {0} is not supported", p.Name.LocalName));
            }
        }
    }
}
