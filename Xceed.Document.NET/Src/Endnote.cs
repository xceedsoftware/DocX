using System;

namespace Xceed.Document.NET.Src
{
    public class Endnote : Footnote
    {

        public static string DefaultEndnoteStyle { get; set; } = "EndnoteText";
        public static string DefaultEndnoteRefStyle { get; set; } = "EndnoteReference";


        public Endnote(Document document, string noteText = null, string[] brackets = null) : base(document, null)
        {
            NoteReferenceStyle = DefaultEndnoteRefStyle;
            NoteTextStyle = DefaultEndnoteStyle;
            NoteReferenceNode = "endnoteReference";
            NoteRefNode = "endnoteRef";
            NoteNode = "endnote";

            Init(document, noteText, brackets);

        }
        internal override void AssignNextId()
        {
            id = (doc.MaxEndnoteId() + 1);
        }
        internal override bool ApplyToDocument()
        {
            return doc.AppendEndnote(noteElement);
        }
        internal override Hyperlink BuildHyperlink(Fragment fragment)
        {
            Hyperlink h = doc.AddHyperlinkToFootnotes(fragment.Content, new Uri(fragment.Content));
            return h;
        }
    }
}
