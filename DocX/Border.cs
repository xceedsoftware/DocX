using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace Novacode
{
    /// <summary>
    /// Represents a border of a table or table cell
    /// Added by lckuiper @ 20101117
    /// </summary>
    public class Border
    {
        BorderStyle tcbs;
        Color color;
        BorderSize size;
        int space;

        public BorderStyle Tcbs
        {
            get { return tcbs; }
            set { tcbs = value; }
        }

        public BorderSize Size
        {
            get { return size; }
            set { size = value; }
        }

        public int Space
        {
            get { return space; }
            set { space = value; }
        }

        public Color Color
        {
            get { return color; }
            set { color = value; }
        }

        public Border()
        {
            this.tcbs = BorderStyle.Tcbs_single;
            this.size = BorderSize.one;
            this.space = 0;
            this.color = Color.Black;
        }

        public Border(BorderStyle tcbs, BorderSize size, int space, Color color)
        {
            this.Tcbs = tcbs;
            this.size = size;
            this.space = space;
            this.Color = color;
        }
    }
}