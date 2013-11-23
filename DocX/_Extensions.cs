using System.Collections.Generic;
using System.Linq;
using System.Drawing;
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

        public static string GetAttribute(this XElement el, XName name, string defaultValue = "")
        {
            var attr = el.Attribute(name);
            if (attr != null)
                return attr.Value;
            return defaultValue;
        }

        /// <summary>
        /// Sets margin for all the pages in a Dox document in Inches. (Written by Shashwat Tripathi)
        /// </summary>
        /// <param name="document"></param>
        /// <param name="top">Margin from the Top. Leave -1 for no change</param>
        /// <param name="bottom">Margin from the Bottom. Leave -1 for no change</param>
        /// <param name="right">Margin from the Right. Leave -1 for no change</param>
        /// <param name="left">Margin from the Left. Leave -1 for no change</param>
        public static void SetMargin(this DocX document, float top, float bottom, float right, float left)
        {
            XNamespace ab = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var tempElement = document.PageLayout.Xml.Descendants(ab + "pgMar");
            var e = tempElement.GetEnumerator();

            foreach (var item in tempElement)
            {
                if (left != -1)
                    item.SetAttributeValue(ab + "left", (1440 * left) / 1);
                if (right != -1)
                    item.SetAttributeValue(ab + "right", (1440 * right) / 1);
                if (top != -1)
                    item.SetAttributeValue(ab + "top", (1440 * top) / 1);
                if (bottom != -1)
                    item.SetAttributeValue(ab + "bottom", (1440 * bottom) / 1);
            }
        }
    }
}
