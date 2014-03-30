using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Novacode;


namespace Novacode
{
    public static class ExtensionsHeadings
    {
        public static Paragraph Heading(this Paragraph paragraph, HeadingType headingType)
        {
            string StyleName = headingType.EnumDescription();
            paragraph.StyleName = StyleName;
            return paragraph;
        }

        public static string EnumDescription(this Enum enumValue)
        {
            if (enumValue == null || enumValue.ToString() == "0")
            {
                return string.Empty;
            }
            FieldInfo enumInfo = enumValue.GetType().GetField(enumValue.ToString());
            DescriptionAttribute[] enumAttributes = (DescriptionAttribute[])enumInfo.GetCustomAttributes(typeof(DescriptionAttribute), false);
            if (enumAttributes.Length > 0)
            {
                return enumAttributes[0].Description;
            }
            else
            {
                return enumValue.ToString();
            }
        }
    }
    
}
