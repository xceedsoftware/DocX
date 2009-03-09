using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Novacode
{
    public class Text
    {
        private int startIndex;
        private int endIndex;
        private string text;
        private XElement e;

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

        /// <summary>
        /// The underlying XElement of this run
        /// </summary>
        public XElement Xml { get { return e; } }

        /// <summary>
        /// A Text element
        /// </summary>
        /// <param name="startIndex">The index this text starts at</param>
        /// <param name="text">The index this text ends at</param>
        /// <param name="e">The underlying xml element that this text wraps</param>
        public Text(int startIndex, XElement e)
        {
            this.startIndex = startIndex;
            this.e = e;

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

        /// <summary>
        /// Splits a text element at a specified index
        /// </summary>
        /// <param name="e">The text element to split</param>
        /// <param name="index">The index to split at</param>
        /// <returns>A two element array which contains both sides of the split</returns>
        public static XElement[] SplitText(Text t, int index)
        {
            if (index < t.startIndex || index > t.EndIndex)
                throw new ArgumentOutOfRangeException("index");

            XElement splitLeft = null, splitRight = null;
            if (t.e.Name.LocalName == "t" || t.e.Name.LocalName == "delText")
            {
                // The origional text element, now containing only the text before the index point.
                splitLeft = new XElement(t.e.Name, t.e.Attributes(), t.e.Value.Substring(0, index - t.startIndex));
                if (splitLeft.Value.Length == 0)
                    splitLeft = null;
                else
                    PreserveSpace(splitLeft);

                // The origional text element, now containing only the text after the index point.
                splitRight = new XElement(t.e.Name, t.e.Attributes(), t.e.Value.Substring(index - t.startIndex, t.e.Value.Length - (index - t.startIndex)));
                if (splitRight.Value.Length == 0)
                    splitRight = null;
                else
                    PreserveSpace(splitRight);
            }

            else
            {
                if (index == t.StartIndex)
                    splitLeft = t.e;

                else
                    splitRight = t.e;
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
