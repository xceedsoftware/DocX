using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Novacode
{
    internal class Text
    {
        private int startIndex;
        private int endIndex;
        private string text;
        internal XElement xml;

        /// <summary>
        /// Gets the start index of this Text (text length before this text)
        /// </summary>
        public int StartIndex { get { return startIndex; } }

        /// <summary>
        /// Gets the end index of this Text (text length before this text + this texts length)
        /// </summary>
        public int EndIndex { get { return endIndex; } }

        /// <summary>
        /// The text value of this text element
        /// </summary>
        public string Value { get { return text; } }

        internal Text(int startIndex, XElement e)
        {
            this.startIndex = startIndex;
            this.xml = e;

            switch (e.Name.LocalName)
            {
                case "t":
                    {
                        goto case "delText";
                    }

                case "delText":
                    {
                        endIndex = startIndex + e.Value.Length;
                        text = e.Value;
                        break;
                    }

                case "br":
                        {
                            text = "\n";
                            endIndex = startIndex + 1;
                            break;
                        }

                case "tab":
                        {
                            text = "\t";
                            endIndex = startIndex + 1;
                            break;
                        }
                default:
                        {
                            break;
                        }
            }
        }

        internal static XElement[] SplitText(Text t, int index)
        {
            if (index < t.startIndex || index > t.EndIndex)
                throw new ArgumentOutOfRangeException("index");

            XElement splitLeft = null, splitRight = null;
            if (t.xml.Name.LocalName == "t" || t.xml.Name.LocalName == "delText")
            {
                // The origional text element, now containing only the text before the index point.
                splitLeft = new XElement(t.xml.Name, t.xml.Attributes(), t.xml.Value.Substring(0, index - t.startIndex));
                if (splitLeft.Value.Length == 0)
                    splitLeft = null;
                else
                    PreserveSpace(splitLeft);

                // The origional text element, now containing only the text after the index point.
                splitRight = new XElement(t.xml.Name, t.xml.Attributes(), t.xml.Value.Substring(index - t.startIndex, t.xml.Value.Length - (index - t.startIndex)));
                if (splitRight.Value.Length == 0)
                    splitRight = null;
                else
                    PreserveSpace(splitRight);
            }

            else
            {
                if (index == t.StartIndex)
                    splitLeft = t.xml;

                else
                    splitRight = t.xml;
            }

            return
            (
                new XElement[]
                {
                    splitLeft,
                    splitRight
                }
            );
        }

        /// <summary>
        /// If a text element or delText element, starts or ends with a space,
        /// it must have the attribute space, otherwise it must not have it.
        /// </summary>
        /// <param name="e">The (t or delText) element check</param>
        public static void PreserveSpace(XElement e)
        {
            // PreserveSpace should only be used on (t or delText) elements
            if (!e.Name.Equals(DocX.w + "t") && !e.Name.Equals(DocX.w + "delText"))
                throw new ArgumentException("SplitText can only split elements of type t or delText", "e");

            // Check if this w:t contains a space atribute
            XAttribute space = e.Attributes().Where(a => a.Name.Equals(XNamespace.Xml + "space")).SingleOrDefault();

            // This w:t's text begins or ends with whitespace
            if (e.Value.StartsWith(" ") || e.Value.EndsWith(" "))
            {
                // If this w:t contains no space attribute, add one.
                if (space == null)
                    e.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
            }

            // This w:t's text does not begin or end with a space
            else
            {
                // If this w:r contains a space attribute, remove it.
                if (space != null)
                    space.Remove();
            }
        }
    }
}
