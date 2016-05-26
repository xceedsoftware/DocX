using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Novacode
{
    public class ContentCollection : List<Content>
    {
        public Content this[string name]
        {
            get
            {
                return this.FirstOrDefault(content => string.Equals(content.Name, name, StringComparison.CurrentCultureIgnoreCase));
            }
        }
    }
}
