using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Novacode
{
    public class Run
    {
        // A lookup for the text elements in this paragraph
        Dictionary<int, Text> textLookup = new Dictionary<int, Text>();

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
        private string Value { set { value = text; } get { return text; } }

        /// <summary>
        /// The underlying XElement of this run
        /// </summary>
        public XElement Xml { get { return e; } }

        /// <summary>
        /// A Run can contain text, delText, br, t and rPr elements
        /// </summary>
        /// <param name="startIndex">The start index of this run (text length before this run)</param>
        /// <param name="endIndex">The end index of this run (text length before this run, plus this runs text length)</param>
        /// <param name="runText">The text value of this run</param>
        /// <param name="e">The underlying XElement of this run</param>
        internal Run(int startIndex, XElement e)
        {
            this.startIndex = startIndex;
            this.e = e;

            // Get the text elements in this run
            IEnumerable<XElement> texts = e.Descendants();

            int start = startIndex;

            // Loop through each text in this run
            foreach (XElement text in texts)
            {
                switch (text.Name.LocalName)
                {
                    case "tab":
                        {
                            textLookup.Add(start + 1, new Text(start, text));
                            Value += "\t";
                            start++;
                            break;
                        }
                    case "br":
                        {
                            textLookup.Add(start + 1, new Text(start, text));
                            Value += "\n";
                            start++;
                            break;
                        }
                    case "t": goto case "delText";
                    case "delText":
                        {
                            // Only add strings which are not empty
                            if (text.Value.Length > 0)
                            {
                                textLookup.Add(start + text.Value.Length, new Text(start, text));
                                Value += text.Value;
                                start += text.Value.Length;
                            }
                            break;
                        }
                    default: break;
                }
            }

            endIndex = start;
        }

        /// <summary>
        /// Splits a run element at a specified index
        /// </summary>
        /// <param name="r">The run element to split</param>
        /// <param name="index">The index to split at</param>
        /// <returns>A two element array which contains both sides of the split</returns>
        static internal XElement[] SplitRun(Run r, int index)
        {
            Text t = r.GetFirstTextEffectedByEdit(index);
            XElement[] splitText = Text.SplitText(t, index);
            
            XElement splitLeft = new XElement(r.e.Name, r.e.Attributes(), r.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName)), t.Xml.ElementsBeforeSelf().Where(n => n.Name.LocalName != "rPr"), splitText[0]);
            if(Paragraph.GetElementTextLength(splitLeft) == 0)
                splitLeft = null;

            XElement splitRight = new XElement(r.e.Name, r.e.Attributes(), r.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName)), splitText[1], t.Xml.ElementsAfterSelf().Where(n => n.Name.LocalName != "rPr"));
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

        /// <summary>
        /// Return the first Run that will be effected by an edit at the index
        /// </summary>
        /// <param name="index">The index of this edit</param>
        /// <returns>The first Run that will be effected</returns>
        public Text GetFirstTextEffectedByEdit(int index)
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
