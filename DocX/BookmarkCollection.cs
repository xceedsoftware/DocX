using System;
using System.Collections.Generic;
using System.Linq;

namespace Novacode
{
    public class BookmarkCollection : List<Bookmark>
    {
        public Bookmark this[string name]
        {
            get
            {
                return this.FirstOrDefault(bookmark => string.Equals(bookmark.Name, name, StringComparison.CurrentCultureIgnoreCase));
            }
        }
    }
}
