using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.IO;
using System.Xml.Linq;

namespace Novacode
{
    internal static class Extensions
    {
        internal static string ToHex(this Color source)
        {
            byte red = source.R;
            byte green = source.G;
            byte blue = source.B;

            string redHex = red.ToString("X");
            if (redHex.Length < 2)
                redHex = "0" + redHex;

            string blueHex = blue.ToString("X");
            if (blueHex.Length < 2)
                blueHex = "0" + blueHex;

            string greenHex = green.ToString("X");
            if (greenHex.Length < 2)
                greenHex = "0" + greenHex;

            return string.Format("{0}{1}{2}", redHex, greenHex, blueHex);
        }

        public static void Flatten(this XElement e, XName name, List<XElement> flat)
        {
            // Add this element (without its children) to the flat list.
            XElement clone = CloneElement(e);
            clone.Elements().Remove();

            // Filter elements using XName.
            if (clone.Name == name)
                flat.Add(clone);

            // Process the children.
            if (e.HasElements)
                foreach (XElement elem in e.Elements(name)) // Filter elements using XName
                    elem.Flatten(name, flat);
        }

        static XElement CloneElement(XElement element)
        {
            return new XElement(element.Name,
                element.Attributes(),
                element.Nodes().Select(n =>
                {
                    XElement e = n as XElement;
                    if (e != null)
                        return CloneElement(e);
                    return n;
                }
                )
            );
        }
    }
}
