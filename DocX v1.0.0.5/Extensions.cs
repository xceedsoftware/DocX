using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace Novacode
{
    internal static class Extensions
    {
        internal static string ToHex(this Color source)
        {
            byte red = source.R;
            byte green = source.G;
            byte blue = source.B;

            string redHex = red.ToString("X");
            if (redHex.Length < 2)
                redHex = "0" + redHex;

            string blueHex = blue.ToString("X");
            if (blueHex.Length < 2)
                blueHex = "0" + blueHex;

            string greenHex = green.ToString("X");
            if (greenHex.Length < 2)
                greenHex = "0" + greenHex;

            return string.Format("{0}{1}{2}", redHex, greenHex, blueHex);
        }
    }
}
