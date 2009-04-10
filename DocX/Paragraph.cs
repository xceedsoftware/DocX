using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Drawing;
using System.Security.Principal;
using System.Collections;

namespace Novacode
{
    /// <summary>
    /// Represents a .docx paragraph.
    /// </summary>
    public class Paragraph
    {
        // This paragraphs text alignment
        private Alignment alignment;

        // A lookup for the runs in this paragraph
        Dictionary<int, Run> runLookup = new Dictionary<int, Run>();

        // The underlying XElement which this Paragraph wraps
        private XElement p;

        // A collection of images in this paragraph
        private IEnumerable<Picture> pictures;
        public IEnumerable<Picture> Pictures { get { return pictures; } }
        /// <summary>
        /// Wraps a XElement as a Paragraph.
        /// </summary>
        /// <param name="p">The XElement to wrap.</param>
        internal Paragraph(XElement p)
        {
            this.p = p;

            BuildRunLookup(p);

            // Get all of the images in this document
            pictures = from i in p.Descendants(XName.Get("drawing", DocX.w.NamespaceName))
                     select new Picture(i);
        }

        /// <summary>
        /// Gets or set this paragraphs text alignment
        /// </summary>
        public Alignment Alignment 
        { 
            get { return alignment; }

            set 
            {
                alignment = value;

                XElement pPr = p.Element(XName.Get("pPr", DocX.w.NamespaceName));

                if (alignment != Novacode.Alignment.left)
                {
                    if (pPr == null)
                        p.Add(new XElement(XName.Get("pPr", DocX.w.NamespaceName)));
                    
                    pPr = p.Element(XName.Get("pPr", DocX.w.NamespaceName));

                    XElement jc = pPr.Element(XName.Get("jc", DocX.w.NamespaceName));

                    if(jc == null)
                        pPr.Add(new XElement(XName.Get("jc", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), alignment.ToString())));
                    else
                        jc.Attribute(XName.Get("val", DocX.w.NamespaceName)).Value = alignment.ToString();
                }

                else
                {
                    if (pPr != null)
                    {
                        XElement jc = pPr.Element(XName.Get("jc", DocX.w.NamespaceName));

                        if (jc != null)
                            jc.Remove();
                    }
                }
            } 
        }

        public void Delete(bool trackChanges)
        {
            if (trackChanges)
            {
                DateTime now = DateTime.Now.ToUniversalTime();

                List<XElement> elements = p.Elements().ToList();
                List<XElement> temp = new List<XElement>();
                for (int i = 0; i < elements.Count(); i++ )
                {
                    XElement e = elements[i];

                    if (e.Name.LocalName != "del")
                    {
                        temp.Add(e);
                        e.Remove();
                    }

                    else
                    {
                        if (temp.Count() > 0)
                        {
                            e.AddBeforeSelf(CreateEdit(EditType.del, now, temp.Elements()));
                            temp.Clear();
                        }
                    }
                }

                if (temp.Count() > 0)
                    p.Add(CreateEdit(EditType.del, now, temp));                   
            }

            else
            {
                // Remove this paragraph from the document
                p.Remove();
                p = null;

                runLookup.Clear();
                runLookup = null;        
            }

            DocX.RebuildParagraphs();
        }

        private void BuildRunLookup(XElement p)
        {
            // Get the runs in this paragraph
            IEnumerable<XElement> runs = p.Descendants(XName.Get("r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"));

            int startIndex = 0;

            // Loop through each run in this paragraph
            foreach (XElement run in runs)
            {
                // Only add runs which contain text
                if (GetElementTextLength(run) > 0)
                {
                    Run r = new Run(startIndex, run);
                    runLookup.Add(r.EndIndex, r);
                    startIndex = r.EndIndex;
                }
            }
        }

        /// <summary>
        /// Gets the value of this Novacode.DocX.Paragraph.
        /// </summary>
        public string Value
        {
            // Returns the underlying XElement's Value property.
            get
            {
                StringBuilder sb = new StringBuilder();

                // Loop through each run in this paragraph
                foreach (XElement r in p.Descendants(XName.Get("r", DocX.w.NamespaceName)))
                {
                    // Loop through each text item in this run
                    foreach (XElement descendant in r.Descendants())
                    {
                        switch (descendant.Name.LocalName)
                        {
                            case "tab":
                                sb.Append("\t");
                                break;
                            case "br":
                                sb.Append("\n");
                                break;
                            case "t":
                                goto case "delText";
                            case "delText":
                                sb.Append(descendant.Value);
                                break;
                            default: break;
                        }
                    }
                }

                return sb.ToString();
            }
        }

        public void InsertPicture(Picture picture, int index)
        {
            Run run = GetFirstRunEffectedByEdit(index);

            if (run == null)
                p.Add(picture.i);
            else
            {
                // Split this run at the point you want to insert
                XElement[] splitRun = Run.SplitRun(run, index);

                // Replace the origional run
                run.Xml.ReplaceWith
                (
                    splitRun[0],
                    picture.i,
                    splitRun[1]
                );
            }

            // Rebuild the run lookup for this paragraph
            runLookup.Clear();
            BuildRunLookup(p);
            DocX.RenumberIDs();
        }

        /// <summary>
        /// Creates an Edit either a ins or a del with the specified content and date
        /// </summary>
        /// <param name="t">The type of this edit (ins or del)</param>
        /// <param name="edit_time">The time stamp to use for this edit</param>
        /// <param name="content">The initial content of this edit</param>
        /// <returns></returns>
        private XElement CreateEdit(EditType t, DateTime edit_time, object content)
        {
            if (t == EditType.del)
            {
                foreach (object o in (IEnumerable<XElement>)content)
                {
                    if (o is XElement)
                    {
                       XElement e = (o as XElement);
                       IEnumerable<XElement> ts = e.DescendantsAndSelf(XName.Get("t", DocX.w.NamespaceName));
                       
                       for(int i = 0; i < ts.Count(); i ++)
                       {
                           XElement text = ts.ElementAt(i);
                           text.ReplaceWith(new XElement(DocX.w + "delText", text.Attributes(), text.Value));  
                       }
                    }
                }
            }

            return
            (
                new XElement(DocX.w + t.ToString(),
                    new XAttribute(DocX.w + "id", 0),
                    new XAttribute(DocX.w + "author", WindowsIdentity.GetCurrent().Name),
                    new XAttribute(DocX.w + "date", edit_time),
                content)
            );
        }

        /// <summary>
        /// Return the first Run that will be effected by an edit at the index
        /// </summary>
        /// <param name="index">The index of this edit</param>
        /// <returns>The first Run that will be effected</returns>
        public Run GetFirstRunEffectedByEdit(int index)
        {
            foreach (int runEndIndex in runLookup.Keys)
            {
                if (runEndIndex > index)
                    return runLookup[runEndIndex];
            }

            if (runLookup.Last().Value.EndIndex == index)
                return runLookup.Last().Value;

            throw new ArgumentOutOfRangeException();
        }

        /// <summary>
        /// Return the first Run that will be effected by an edit at the index
        /// </summary>
        /// <param name="index">The index of this edit</param>
        /// <returns>The first Run that will be effected</returns>
        public Run GetFirstRunEffectedByInsert(int index)
        {
            // This paragraph contains no Runs and insertion is at index 0
            if (runLookup.Keys.Count() == 0 && index == 0)
                return null;

            foreach (int runEndIndex in runLookup.Keys)
            {
                if (runEndIndex >= index)
                    return runLookup[runEndIndex];
            }

            throw new ArgumentOutOfRangeException();
        }

        /// <summary>
        /// If the value to be inserted contains tab elements or br elements, multiple runs will be inserted.
        /// </summary>
        /// <param name="text">The text to be inserted, this text can contain the special characters \t (tab) and \n (br)</param>
        /// <returns></returns>
        private List<XElement> formatInput(string text, XElement rPr)
        {
            List<XElement> newRuns = new List<XElement>();
            XElement tabRun = new XElement(DocX.w + "tab");

            string[] runTexts = text.Split('\t');
            XElement firstRun;
            if (runTexts[0] != String.Empty)
            {
                XElement firstText = new XElement(DocX.w + "t", runTexts[0]);

                Text.PreserveSpace(firstText);

                firstRun = new XElement(DocX.w + "r", rPr, firstText);

                newRuns.Add(firstRun);
            }

            if (runTexts.Length > 1)
            {
                for (int k = 1; k < runTexts.Length; k++)
                {
                    XElement newText = new XElement(DocX.w + "t", runTexts[k]);

                    XElement newRun;
                    if (runTexts[k] == String.Empty)
                        newRun = new XElement(DocX.w + "r", tabRun);

                    else
                    {
                        // Value begins or ends with a space
                        Text.PreserveSpace(newText);

                        newRun = new XElement(DocX.w + "r", rPr, tabRun, newText);
                    }

                    newRuns.Add(newRun);
                }
            }

            return newRuns;
        }

        /// <summary>
        /// Counts the text length of an element
        /// </summary>
        /// <param name="run">An element</param>
        /// <returns>The length of this elements text</returns>
        static internal int GetElementTextLength(XElement run)
        {
            int count = 0;
            
            if (run == null)
                return count;

            foreach (var d in run.Descendants())
            {
                switch (d.Name.LocalName)
                {
                    case "tab": goto case "br";
                    case "br": count++; break;
                    case "t": goto case "delText";
                    case "delText": count += d.Value.Length; break;
                    default: break;
                }
            }
            return count;
        }

        /// <summary>
        /// Splits an edit element at the specified run, at the specified index.
        /// </summary>
        /// <param name="edit">The edit element to split</param>
        /// <param name="run">The run element to split</param>
        /// <param name="index">The index to split at</param>
        /// <returns></returns>
        public XElement[] SplitEdit(XElement edit, int index, EditType type)
        {
            Run run;
            if(type == EditType.del)
                run = GetFirstRunEffectedByEdit(index);
            else
                run = GetFirstRunEffectedByInsert(index);

            XElement[] splitRun = Run.SplitRun(run, index);
            
            XElement splitLeft = new XElement(edit.Name, edit.Attributes(), run.Xml.ElementsBeforeSelf(), splitRun[0]);
            if (GetElementTextLength(splitLeft) == 0)
                splitLeft = null;

            XElement splitRight = new XElement(edit.Name, edit.Attributes(), splitRun[1], run.Xml.ElementsAfterSelf());
            if (GetElementTextLength(splitRight) == 0)
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

        public void Insert(int index, string value, bool trackChanges)
        {
            Insert(index, value, null, trackChanges);
        }

        /// <summary>
        /// Inserts a specified instance of System.String into a Novacode.DocX.Paragraph at a specified index position.
        /// </summary>
        /// <example>
        /// <code>
        ///  // Description: Simple string insertion
        ///  
        ///  // Load Example.docx
        ///  DocX dx = DocX.Load(@"C:\Example.docx");
        ///
        ///  // Iterate through the paragraphs
        ///  foreach (Paragraph p in dx.Paragraphs)
        ///  {
        ///     // Insert the string "Start: " at the begining of every paragraph and flag it as a change.
        ///     p.Insert(0, "Start: ", true);
        ///  }
        ///
        ///  // Save changes to Example.docx
        ///  dx.Save();
        /// </code>
        /// </example>
        /// <example>
        /// <code>
        ///  // Description: Inserting tabs using the \t switch
        ///  
        ///  // Load Example.docx
        ///  DocX dx = DocX.Load(@"C:\Example.docx");
        ///
        ///  // Iterate through the paragraphs
        ///  foreach (Paragraph p in dx.Paragraphs)
        ///  {
        ///     // Insert the string "\tStart:\t" at the begining of every paragraph and flag it as a change.
        ///     p.Insert(0, "\tStart:\t", true);
        ///  }
        ///
        ///  // Save changes to Example.docx
        ///  dx.Save();
        /// </code>
        /// </example>
        /// <seealso cref="Paragraph.Remove"/>
        /// <seealso cref="Paragraph.Replace"/>
        /// <param name="index">The index position of the insertion.</param>
        /// <param name="value">The System.String to insert.</param>
        /// <param name="trackChanges">Flag this insert as a change</param>
        public void Insert(int index, string value, Formatting formatting, bool trackChanges)
        {
            // Timestamp to mark the start of insert
            DateTime now = DateTime.Now;
            DateTime insert_datetime = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);

            // Get the first run effected by this Insert
            Run run = GetFirstRunEffectedByInsert(index);

            if (run == null)
            {
                object insert;
                if (formatting != null)
                    insert = formatInput(value, formatting.Xml);
                else
                    insert = formatInput(value, null);
                
                if (trackChanges)
                    insert = CreateEdit(EditType.ins, insert_datetime, insert);
                p.Add(insert);
            }

            else
            {
                object newRuns;
                if (formatting != null)
                    newRuns = formatInput(value, formatting.Xml);
                else
                    newRuns = formatInput(value, run.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName)));

                // The parent of this Run
                XElement parentElement = run.Xml.Parent;
                switch (parentElement.Name.LocalName)
                {
                    case "ins":
                        {
                            // The datetime that this ins was created
                            DateTime parent_ins_date = DateTime.Parse(parentElement.Attribute(XName.Get("date", DocX.w.NamespaceName)).Value);

                            /* 
                             * Special case: You want to track changes,
                             * and the first Run effected by this insert
                             * has a datetime stamp equal to now.
                            */
                            if (trackChanges && parent_ins_date.CompareTo(insert_datetime) == 0)
                            {
                                /*
                                 * Inserting into a non edit and this special case, is the same procedure.
                                */
                                goto default;
                            }

                            /*
                             * If not the special case above, 
                             * then inserting into an ins or a del, is the same procedure.
                            */
                            goto case "del";
                        }

                    case "del":
                        {
                            object insert = newRuns;
                            if (trackChanges)
                                insert = CreateEdit(EditType.ins, insert_datetime, newRuns);

                            // Split this Edit at the point you want to insert
                            XElement[] splitEdit = SplitEdit(parentElement, index, EditType.ins);

                            // Replace the origional run
                            parentElement.ReplaceWith
                            (
                                splitEdit[0],
                                insert,
                                splitEdit[1]
                            );

                            break;
                        }

                    default:
                        {
                            object insert = newRuns;
                            if (trackChanges && !parentElement.Name.LocalName.Equals("ins"))
                                insert = CreateEdit(EditType.ins, insert_datetime, newRuns);

                            // Split this run at the point you want to insert
                            XElement[] splitRun = Run.SplitRun(run, index);

                            // Replace the origional run
                            run.Xml.ReplaceWith
                            (
                                splitRun[0],
                                insert,
                                splitRun[1]
                            );

                            break;
                        }
                }
            }

            // Rebuild the run lookup for this paragraph
            runLookup.Clear();
            BuildRunLookup(p);
            DocX.RenumberIDs();
        }

        /// <summary>
        /// Removes characters from a Novacode.DocX.Paragraph.
        /// </summary>
        /// <example>
        /// <code>
        /// // Load Example.docx
        /// DocX dx = DocX.Load(@"C:\Example.docx");
        ///
        /// // Iterate through the paragraphs
        /// foreach (Paragraph p in dx.Paragraphs)
        /// {
        ///     // Remove the first two characters from every paragraph
        ///     p.Remove(0, 2);
        /// }
        ///
        /// // Save changes to Example.docx
        /// dx.Save();
        /// </code>
        /// </example>
        /// <seealso cref="Paragraph.Insert"/>
        /// <seealso cref="Paragraph.Replace"/>
        /// <param name="index">The position to begin deleting characters.</param>
        /// <param name="count">The number of characters to delete</param>
        /// <param name="trackChanges">Track changes</param>
        public void Remove(int index, int count, bool trackChanges)
        {
            // Timestamp to mark the start of insert
            DateTime now = DateTime.Now;
            DateTime remove_datetime = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);

            // The number of characters processed so far
            int processed = 0;

            do
            {
                // Get the first run effected by this Remove
                Run run = GetFirstRunEffectedByEdit(index + processed);

                // The parent of this Run
                XElement parentElement = run.Xml.Parent;
                switch (parentElement.Name.LocalName)
                {
                    case "ins":
                        {
                            XElement[] splitEditBefore = SplitEdit(parentElement, index + processed, EditType.del);
                            int min = Math.Min(count - processed, run.Xml.ElementsAfterSelf().Sum(e => GetElementTextLength(e)));
                            XElement[] splitEditAfter = SplitEdit(parentElement, index + processed + min, EditType.del);

                            XElement temp = SplitEdit(splitEditBefore[1], index + processed + min, EditType.del)[0];
                            object middle = CreateEdit(EditType.del, remove_datetime, temp.Elements());
                            processed += GetElementTextLength(middle as XElement);
                            
                            if (!trackChanges)
                                middle = null;
                                
                            parentElement.ReplaceWith
                            (
                                splitEditBefore[0],
                                middle,
                                splitEditAfter[1]
                            );

                            processed += GetElementTextLength(middle as XElement);
                            break;
                        }

                    case "del":
                        {
                            if (trackChanges)
                            {
                                // You cannot delete from a deletion, advance processed to the end of this del
                                processed += GetElementTextLength(parentElement);
                            }

                            else
                                goto case "ins";

                            break;
                        }

                    default:
                        {
                            XElement[] splitRunBefore = Run.SplitRun(run, index + processed);
                            int min = Math.Min(index + processed + (count - processed), run.EndIndex);
                            XElement[] splitRunAfter = Run.SplitRun(run, min);

                            object middle = CreateEdit(EditType.del, remove_datetime, new List<XElement>() { Run.SplitRun(new Run(run.StartIndex + GetElementTextLength(splitRunBefore[0]), splitRunBefore[1]), min)[0] });
                            processed += GetElementTextLength(middle as XElement);
                            
                            if (!trackChanges)
                                middle = null;

                            run.Xml.ReplaceWith
                            (
                                splitRunBefore[0],
                                middle,
                                splitRunAfter[1]
                            );

                            break;
                        }
                }

                // If after this remove the parent element is empty, remove it.
                if (GetElementTextLength(parentElement) == 0)
                    parentElement.Remove();
            }
            while (processed < count);

            // Rebuild the run lookup
            runLookup.Clear();
            BuildRunLookup(p);
            DocX.RenumberIDs();
        }

        /// <summary>
        /// Replaces all occurrences of a specified System.String in this instance, with another specified System.String.
        /// </summary>
        /// <example>
        /// <code>
        /// // Load Example.docx
        /// DocX dx = DocX.Load(@"C:\Example.docx");
        ///
        /// // Iterate through the paragraphs
        /// foreach (Paragraph p in dx.Paragraphs)
        /// {
        ///     // Replace all instances of the string "wrong" with the stirng "right"
        ///     p.Replace("wrong", "right");
        /// }
        ///
        /// // Save changes to Example.docx
        /// dx.Save();
        /// </code>
        /// </example>
        /// <seealso cref="Paragraph.Remove"/>
        /// <seealso cref="Paragraph.Insert"/>
        /// <param name="newValue">A System.String to replace all occurances of oldValue.</param>
        /// <param name="oldValue">A System.String to be replaced.</param>
        /// <param name="options">A bitwise OR combination of RegexOption enumeration options.</param>
        /// <param name="trackChanges">Track changes</param>
        public void Replace(string oldValue, string newValue, bool trackChanges, RegexOptions options)
        {
            MatchCollection mc = Regex.Matches(this.Value, Regex.Escape(oldValue), options);
            
            // Loop through the matches in reverse order
            foreach (Match m in mc.Cast<Match>().Reverse())
            {
                Insert(m.Index + oldValue.Length, newValue, trackChanges);
                Remove(m.Index, m.Length, trackChanges);
            }
        }

        public void Replace(string oldValue, string newValue, bool trackChanges)
        {
            Replace(oldValue, newValue, trackChanges, RegexOptions.None);
        }
    }
}
