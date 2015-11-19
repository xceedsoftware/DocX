using System;
using System.Linq;
using System.Collections.Generic;

namespace Novacode
{
    public class Bookmark
    {
        public string Name { get; set; }
        public Paragraph Paragraph { get; set; }

        public void SetText(string newText)
        {
            Paragraph.ReplaceAtBookmark(newText, Name);
        }
    }
}
