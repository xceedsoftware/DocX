using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Novacode
{
    internal class Run
    {
        // A lookup for the text elements in this paragraph
        Dictionary<int, Text> textLookup = new Dictionary<int, Text>();

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
        internal string Value { set { value = text; } get { return text; } }

        internal Run(int startIndex, XElement xml)
        {
            this.startIndex = startIndex;
            this.xml = xml;

            // Get the text elements in this run
            IEnumerable<XElement> texts = xml.Descendants();

            int start = startIndex;

            // Loop through each text in this run
            foreach (XElement te in texts)
            {
                switch (te.Name.LocalName)
                {
                    case "tab":
                        {
                            textLookup.Add(start + 1, new Text(start, te));
                            text += "\t";
                            start++;
                            break;
                        }
                    case "br":
                        {
                            textLookup.Add(start + 1, new Text(start, te));
                            text += "\n";
                            start++;
                            break;
                        }
                    case "t": goto case "delText";
                    case "delText":
                        {
                            // Only add strings which are not empty
                            if (te.Value.Length > 0)
                            {
                                textLookup.Add(start + te.Value.Length, new Text(start, te));
                                text += te.Value;
                                start += te.Value.Length;
                            }
                            break;
                        }
                    default: break;
                }
            }

            endIndex = start;
        }

        static internal XElement[] SplitRun(Run r, int index)
        {
            Text t = r.GetFirstTextEffectedByEdit(index);
            XElement[] splitText = Text.SplitText(t, index);
            
            XElement splitLeft = new XElement(r.xml.Name, r.xml.Attributes(), r.xml.Element(XName.Get("rPr", DocX.w.NamespaceName)), t.xml.ElementsBeforeSelf().Where(n => n.Name.LocalName != "rPr"), splitText[0]);
            if(Paragraph.GetElementTextLength(splitLeft) == 0)
                splitLeft = null;

            XElement splitRight = new XElement(r.xml.Name, r.xml.Attributes(), r.xml.Element(XName.Get("rPr", DocX.w.NamespaceName)), splitText[1], t.xml.ElementsAfterSelf().Where(n => n.Name.LocalName != "rPr"));
            if(Paragraph.GetElementTextLength(splitRight) == 0)
                splitRight = null;

            return
            (
                new XElement[]
                {
                    splitLeft,
                    splitRight
                }
            );
        }

        internal Text GetFirstTextEffectedByEdit(int index)
        {
            foreach (int textEndIndex in textLookup.Keys)
            {
                if (textEndIndex > index)
                    return textLookup[textEndIndex];
            }

            if (textLookup.Last().Value.EndIndex == index)
                return textLookup.Last().Value;

            throw new ArgumentOutOfRangeException();
        }
    }
}
