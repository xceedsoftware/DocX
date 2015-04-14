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

        /// <summary>
        /// From: http://stackoverflow.com/questions/4108828/generic-extension-method-to-see-if-an-enum-contains-a-flag
        /// Check to see if a flags enumeration has a specific flag set.
        /// </summary>
        /// <param name="variable">Flags enumeration to check</param>
        /// <param name="value">Flag to check for</param>
        /// <returns></returns>
        public static bool HasFlag(this Enum variable, Enum value)
        {
            if (variable == null)
                return false;

            if (value == null)
                throw new ArgumentNullException("value");

            // Not as good as the .NET 4 version of this function, but should be good enough
            if (!Enum.IsDefined(variable.GetType(), value))
            {
                throw new ArgumentException(string.Format(
                    "Enumeration type mismatch.  The flag is of type '{0}', was expecting '{1}'.",
                    value.GetType(), variable.GetType()));
            }

            ulong num = Convert.ToUInt64(value);
            return ((Convert.ToUInt64(variable) & num) == num);

        }
    }
    
}
