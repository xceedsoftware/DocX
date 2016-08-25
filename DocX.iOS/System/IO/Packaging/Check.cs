using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using DocumentFormat.OpenXml.Packaging;

namespace System.IO.Packaging
{
    internal static class Check
    {
        static void NotNull(object o, string name)
        {
            if (o == null)
                throw new ArgumentNullException(name);
        }

        public static void ContentTypeIsValid(string contentType)
        {
            if (string.IsNullOrEmpty(contentType))
                return;

            // Must be in form of: type/subtype
            int index = contentType.IndexOf('/');
            bool result = (index > 0) && contentType.Length > (index + 1) && contentType.IndexOf('/', index + 1) == -1;

            if (!result)
                throw new ArgumentException("contentType", "contentType must be in the form of 'type/subtype'");
        }

        public static void Id(object id)
        {
            NotNull(id, "id");
        }

        public static void IdIsValid(string id)
        {
            if (id == null)
                return;

            // If the ID is a zero string, need to throw a ArgNullEx
            if (id.Length == 0)
                throw new ArgumentNullException("id", "Cannot be whitespace or empty");

            // FIXME: I need to XSD parse this to make sure it's valid
            // If it's not, throw an XmlException
        }

        private static bool EmptyOrBlank(string s)
        {
            return (s != null && (s == "" || s.Trim().Length == 0));
        }

        private static void PartUriDoesntEndWithSlash(Uri uri)
        {
            var s = !uri.IsAbsoluteUri ? uri.OriginalString
                : uri.GetComponents(UriComponents.Path, UriFormat.UriEscaped);

            // We allow '/' at uri's beggining.
            if ((s.Length > 1) && s.EndsWith("/"))
            {
                throw new ArgumentException("Part URI cannot end with a forward slash.");
            }
        }

        public static void Package(object package)
        {
            if (package == null)
                throw new ArgumentNullException(nameof(package));
        }


        public static void PackageUri(object packageUri)
        {
            NotNull(packageUri, "packageUri");
        }

        public static void PackageUriIsValid(Uri packageUri)
        {
            if (!packageUri.IsAbsoluteUri)
                throw new ArgumentException(nameof(packageUri), "Uri must be absolute");
        }

        public static void PackUriIsValid(Uri packUri)
        {
            if (!packUri.IsAbsoluteUri)
                throw new ArgumentException(nameof(packUri), "PackUris must be absolute");

            if (packUri.Scheme != PackUriHelper.UriSchemePack)
                throw new ArgumentException(nameof(packUri), "Uri scheme is not a valid PackUri scheme");
        }

        public static void PartUri(object partUri)
        {
            if (partUri == null)
                throw new ArgumentNullException(nameof(partUri));
        }

        public static void PartUriIsValid(Uri partUri)
        {
            if (!partUri.OriginalString.StartsWith("/"))
                throw new ArgumentException("PartUris must start with '/'");

            if (partUri.IsAbsoluteUri)
                throw new ArgumentException("PartUris cannot be absolute");
        }

        public static void RelationshipTypeIsValid(string relationshipType)
        {
            if (relationshipType == null)
                throw new ArgumentNullException(nameof(relationshipType));
            if (EmptyOrBlank(relationshipType))
                throw new ArgumentException("relationshipType", "Cannot be whitespace or empty");
        }

        public static void PartUri(Uri partUri)
        {
            if (partUri == null)
                throw new ArgumentNullException(nameof(partUri));
            if (partUri.IsAbsoluteUri)
                throw new ArgumentException(nameof(partUri), "Absolute URIs are not supported");
            if (string.IsNullOrEmpty(partUri.OriginalString))
                throw new ArgumentException(nameof(partUri), "Part uri cannot be an empty string");
        }

        public static void PackUri(Uri packUri)
        {
            NotNull(packUri, "packUri");
        }

        public static void SourcePartUri(Uri sourcePartUri)
        {
            NotNull(sourcePartUri, "sourcePartUri");
            PartUriDoesntEndWithSlash(sourcePartUri);
        }

        public static void TargetPartUri(Uri targetPartUri)
        {
            NotNull(targetPartUri, "targetPartUri");
            PartUriDoesntEndWithSlash(targetPartUri);
        }

        public static void SourceUri(Uri sourceUri)
        {
            if (sourceUri == null)
                throw new ArgumentNullException(nameof(sourceUri));
            //			if (sourceUri.IsAbsoluteUri)
            //				throw new ArgumentException ("sourceUri", "Absolute URIs are not supported");
            if (string.IsNullOrEmpty(sourceUri.OriginalString))
                throw new ArgumentException(nameof(sourceUri), "Part uri cannot be an empty string");
        }

        public static void TargetUri(Uri targetUri)
        {
            if (targetUri == null)
                throw new ArgumentNullException(nameof(targetUri));
            //			if (targetUri.IsAbsoluteUri)
            //				throw new ArgumentException ("targetUri", "Absolute URIs are not supported");
            if (string.IsNullOrEmpty(targetUri.OriginalString))
                throw new ArgumentException(nameof(targetUri), "Part uri cannot be an empty string");
        }
    }
}
