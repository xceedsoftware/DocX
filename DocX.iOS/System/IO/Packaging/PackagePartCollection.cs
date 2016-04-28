using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace System.IO.Packaging
{
    public class PackagePartCollection : IEnumerable<PackagePart>, IEnumerable
    {
        internal List<PackagePart> Parts { get; private set; }

        internal PackagePartCollection()
        {
            Parts = new List<PackagePart>();
        }

        public IEnumerator<PackagePart> GetEnumerator()
        {
            return Parts.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return Parts.GetEnumerator();
        }
    }
}
