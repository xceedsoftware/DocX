using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace Xceed.Document.NET.Src
{
    public class Footnote : DocumentElement
    {
        /*
         *
         *
        a complete end/footnote entry is a distinct structure that looks like this (with optional added []s)

        <w:footnote w:id="1">
            <w:p>
                <w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>
                <w:r>
                    <w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
                    <w:t>[</w:t>
                    <w:footnoteRef/>
                    <w:t>]</w:t>
                 </w:r>
                <w:r><w:t xml:space="preserve"> This is my footnote.</w:t></w:r>
            </w:p>
        </w:footnote>

        and a reference is a run (with added []s) that looks like this

        <w:r>
            <w:rPr>
                <w:rStyle w:val="FootnoteReference" />
            </w:rPr>
            <w:t>[</w:t>
            <w:footnoteReference w:id="2" />
            <w:t>]</w:t>
        </w:r>

         *
         *
         */
        public enum FragmentType
        {
            Text,
            Noteref,
            Hyperlink
        }

        internal class Fragment
        {
            public FragmentType Type { get; set; }
            public string Content { get; set; }
            public object DataObject { get; set; }
        }

        public static string BookmarkNamePattern = "_RefN{0}";
 
        public static string DefaultFootnoteStyle { get; set; }  = "FootnoteText";
        public static string DefaultFootnoteRefStyle { get; set; } = "FootnoteReference";

        public string NoteReferenceStyle { get; internal set; }
        public string NoteTextStyle { get; internal set; }
        public string NoteReferenceNode { get; internal set; }
        public string NoteRefNode { get; internal set; }
        public string NoteNode { get; internal set; }
        public bool IsApplied { get; internal set; }
        public int? Id => IsApplied ? id : (int?) null;
        public string BookmarkName { get; set; }
        public int BookmarkId { get; set; }
        public Footnote ReferenceNote { get; set; }
        internal List<Fragment> Fragments { get; set; }

        internal Document doc;
        internal string[] brackets;
        internal XElement noteElement;
        internal XElement noteRefElement;

        private int _id;

        internal int id
        {
            get => _id;
            set
            {
                if (_id == 0)
                    _id = value;
                else
                    throw new InvalidOperationException("footnote id is immutable once set");
            }
        }

        public Footnote(Document document, string noteText = null, string[] brackets = null) : base(document, null)
        {
            NoteReferenceStyle = DefaultFootnoteRefStyle;
            NoteTextStyle = DefaultFootnoteStyle;
            NoteReferenceNode = "footnoteReference";
            NoteRefNode = "footnoteRef";
            NoteNode = "footnote";

            Init(document, noteText, brackets);
        }

        internal void Init(Document document, string text, string[] pBrackets)
        {
            doc = document;
            if (!string.IsNullOrEmpty(text))
                (Fragments = new List<Fragment>()).Add(new Fragment() { Content = text, Type = FragmentType.Text });
            if (pBrackets == null) return;
            if (pBrackets.Length != 2)
                throw new ArgumentException("brackets parameter must be null or two elements");
            brackets = pBrackets;
        }

        public Footnote AppendText(string t)
        {
            (Fragments ?? (Fragments = new List<Fragment>()))
                .Add(new Fragment() { Type = FragmentType.Text, Content = t });
            return this;
        }

        public Footnote AppendNoteRef(Footnote other)
        {
            (Fragments ?? (Fragments = new List<Fragment>()))
                .Add(new Fragment() { Type = FragmentType.Noteref, DataObject = other});
            return this;
        }

        public Footnote AppendHyperlink(string t)
        {
            (Fragments ?? (Fragments = new List<Fragment>()))
                .Add(new Fragment() { Type = FragmentType.Hyperlink, Content = t });
            return this;
        }

        internal virtual void AssignNextId()
        {
            id = (doc.MaxFootnoteId() + 1);
        }

        internal virtual bool ApplyToDocument()
        {
            Xml = noteElement;
            return doc.AppendFootnote(noteElement);
        }

        public void Apply(Paragraph p, bool bookmarked = false)
        {
            if (IsApplied)
            {
                throw new InvalidOperationException("note has already been applied");
            }
            if (Fragments?[0] == null)
            {
                throw new InvalidOperationException("note has no content");
            }

            AssignNextId();

            // create the note element
            noteElement = new XElement(Document.w + NoteNode, new XAttribute(Document.w + "id", id));

            XElement np = new XElement(Document.w + "p");
            np.Add(new XElement(Document.w + "pPr", new XElement(Document.w + "pStyle", new XAttribute(Document.w + "val", NoteTextStyle))));
            noteElement.Add(np);

            XElement r = new XElement(Document.w + "r",
                new XElement(Document.w + "rPr", new XElement(Document.w + "rStyle", new XAttribute(Document.w + "val", NoteReferenceStyle))));
            if (brackets != null)
                r.Add(new XElement(Document.w + "t", brackets[0]));
            r.Add(new XElement(Document.w + NoteRefNode));
            if (brackets != null)
                r.Add(new XElement(Document.w + "t", brackets[1]));
            np.Add(r);

            // make sure there is separation between the fn# and the contents
            string space = (Fragments[0].Content ?? "").StartsWith(" ") ? "" : " ";
            foreach (Fragment fragment in Fragments)
            {
                switch (fragment.Type) 
                {
                    case FragmentType.Text:
                        r = new XElement(Document.w + "r", new XElement(Document.w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), $"{space}{fragment.Content}"));
                        np.Add(r);
                        space = "";
                        break;
                    case FragmentType.Noteref:
                        // insert a bookmark reference back to another foot/endnote
                        // this is limited to the note number, because any other text included 
                        // with it will be deleted if/when Word renumbers the notes, e.g. on <Ctrl>A, F9
                        if ((ReferenceNote ?? (ReferenceNote = fragment.DataObject as Footnote))?.BookmarkName == null)
                            break;
                        NoteRefField nrf = new NoteRefField(doc, null) { MarkName = ReferenceNote.BookmarkName, ReferenceText = $"{ReferenceNote.Id}", InsertHyperlink = true };
                        np.Add(nrf.Build().Xml);
                        break;
                    case FragmentType.Hyperlink:
                        try
                        {
                            Hyperlink h = BuildHyperlink(fragment);
                            np.Add(h?.Xml);
                        }
                        catch (UriFormatException ufx)
                        {
                            // giving up on it
                        }
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }

            // add the xNote to the document's collection
            if (!ApplyToDocument())
                return;
            IsApplied = true;

            ApplyInline(p, bookmarked);
        }

        internal virtual Hyperlink BuildHyperlink(Fragment fragment)
        {
            Hyperlink h =  doc.AddHyperlinkToFootnotes(fragment.Content, new Uri(fragment.Content));
            return h;
        }

        internal void ApplyInline(Paragraph p, bool bookmarked)
        {
            // create the reference element...
            noteRefElement = new XElement(Document.w + "r",
                new XElement(Document.w + "rPr",
                    new XElement(Document.w + "rStyle", new XAttribute(Document.w + "val", NoteReferenceStyle))));
            // ... optionally wrapped in brackets (choose when needed to distinguish footnotes from exponents etc.)
            if (brackets != null)
                noteRefElement.Add(new XElement(Document.w + "t", brackets[0]));
            // optionally wrapped in a bookmark marker
            if (bookmarked)
            {
                BookmarkId = Paragraph.NextBookmarkId;
                BookmarkName = string.Format(BookmarkNamePattern, BookmarkId);
                XElement wBookmarkStart = new XElement(
                    XName.Get("bookmarkStart", Document.w.NamespaceName),
                    new XAttribute(XName.Get("id", Document.w.NamespaceName), BookmarkId),
                    new XAttribute(XName.Get("name", Document.w.NamespaceName), BookmarkName));
                noteRefElement.Add(wBookmarkStart);
            }

            noteRefElement.Add(new XElement(Document.w + NoteReferenceNode, new XAttribute(Document.w + "id", id)));
            if (bookmarked)
            {
                XElement wBookmarkEnd = new XElement(
                    XName.Get("bookmarkEnd", Document.w.NamespaceName),
                    new XAttribute(XName.Get("id", Document.w.NamespaceName), BookmarkId),
                    new XAttribute(XName.Get("name", Document.w.NamespaceName), BookmarkName));
                noteRefElement.Add(wBookmarkEnd);
            }

            if (brackets != null)
                noteRefElement.Add(new XElement(Document.w + "t", brackets[1]));

            // append the reference run to the paragraph
            p.Xml.Add(noteRefElement);
        }
    }
}
