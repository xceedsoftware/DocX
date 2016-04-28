using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace System.IO.Packaging
{
    public class PackageRelationshipCollection : IEnumerable<PackageRelationship>, IEnumerable
    {
        internal List<PackageRelationship> Relationships { get; private set; }

        internal PackageRelationshipCollection()
        {
            Relationships = new List<PackageRelationship>();
        }

        public IEnumerator<PackageRelationship> GetEnumerator()
        {
            return Relationships.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
