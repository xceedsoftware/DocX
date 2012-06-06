using System.Drawing;

namespace Novacode
{
    /// <summary>
    /// Represents a border of a table or table cell
    /// Added by lckuiper @ 20101117
    /// </summary>
    public class Border
    {
        public BorderStyle Tcbs { get; set; }
        public BorderSize Size { get; set; }
        public int Space { get; set; }
        public Color Color { get; set; }
        public Border()
        {
            this.Tcbs = BorderStyle.Tcbs_single;
            this.Size = BorderSize.one;
            this.Space = 0;
            this.Color = Color.Black;
        }

        public Border(BorderStyle tcbs, BorderSize size, int space, Color color)
        {
            this.Tcbs = tcbs;
            this.Size = size;
            this.Space = space;
            this.Color = color;
        }
    }
}