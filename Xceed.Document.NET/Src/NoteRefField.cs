using System.Text;
using System.Xml.Linq;

namespace Xceed.Document.NET.Src
{
    public class NoteRefField : AbstractField
    {
        public string MarkName { get; set; }
        public string ReferenceText { get; set; }
        public bool SameFormatting { get; set; }
        public bool InsertHyperlink { get; set; }
        public bool InsertRelativePosition { get; set; }
 

        #region Overrides of AbstractField

        public override AbstractField Build()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(" NOTEREF ").Append(MarkName).Append(' ');
            if (SameFormatting)
                sb.Append("\\f ");
            if (InsertRelativePosition)
                sb.Append("\\p ");
            if (InsertHyperlink)
                sb.Append("\\h ");

            Xml = Build(sb.ToString(), ReferenceText);
            return this;
        }

        #endregion

        public NoteRefField(Document document, XElement xml = null) : base(document, xml) { }
    }
}
