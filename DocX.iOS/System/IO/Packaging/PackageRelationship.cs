using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace System.IO.Packaging
{
    public class PackageRelationship
    {
        internal PackageRelationship(string id, Package package, string relationshipType,
                                      Uri sourceUri, TargetMode targetMode, Uri targetUri)
        {
            Check.IdIsValid(id);
            Check.Package(package);
            Check.RelationshipTypeIsValid(relationshipType);
            Check.SourceUri(sourceUri);
            Check.TargetUri(targetUri);

            Id = id;
            Package = package;
            RelationshipType = relationshipType;
            SourceUri = sourceUri;
            TargetMode = targetMode;
            TargetUri = targetUri;
        }

        public string Id
        {
            get;
            private set;
        }
        public Package Package
        {
            get;
            private set;
        }
        public string RelationshipType
        {
            get;
            private set;
        }
        public Uri SourceUri
        {
            get;
            private set;
        }
        public TargetMode TargetMode
        {
            get;
            private set;
        }
        public Uri TargetUri
        {
            get;
            private set;
        }
    }
}
