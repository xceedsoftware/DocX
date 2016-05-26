using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace System.IO.Packaging
{
    public static class PackUriHelper
    {
        public static readonly string UriSchemePack = "pack";
        static readonly Uri PackSchemeUri = new Uri("pack://", UriKind.Absolute);
        static readonly char[] _escapedChars = new char[] { '%', ',', '?', '@' };


        static PackUriHelper()
        {
            if (!UriParser.IsKnownScheme(UriSchemePack))
                UriParser.Register(new PackUriParser(), UriSchemePack, -1);
        }

        public static int ComparePackUri(Uri firstPackUri, Uri secondPackUri)
        {
            if (firstPackUri == null)
                return secondPackUri == null ? 0 : -1;
            if (secondPackUri == null)
                return 1;

            Check.PackUriIsValid(firstPackUri);
            Check.PackUriIsValid(secondPackUri);

            // FIXME: What exactly is compared. Lets assume originalstring
            return firstPackUri.OriginalString.CompareTo(secondPackUri.OriginalString);
        }

        public static int ComparePartUri(Uri firstPartUri, Uri secondPartUri)
        {
            if (firstPartUri == null)
                return secondPartUri == null ? 0 : -1;
            if (secondPartUri == null)
                return 1;

            Check.PartUriIsValid(firstPartUri);
            Check.PartUriIsValid(secondPartUri);

            return firstPartUri.OriginalString.CompareTo(secondPartUri.OriginalString);
        }

        public static Uri Create(Uri packageUri)
        {
            return Create(packageUri, null, null);
        }

        public static Uri Create(Uri packageUri, Uri partUri)
        {
            return Create(packageUri, partUri, null);
        }

        public static Uri Create(Uri packageUri, Uri partUri, string fragment)
        {
            Check.PackageUri(packageUri);
            Check.PackageUriIsValid(packageUri);

            if (partUri != null)
                Check.PartUriIsValid(partUri);

            if (fragment != null && (fragment.Length == 0 || fragment[0] != '#'))
                throw new ArgumentException("Fragment", "Fragment must not be empty and must start with '#'");

            // FIXME: Validate that partUri is a valid one? Must be relative, must start with '/'

            // First replace the slashes, then escape the special characters
            //string orig = packageUri.GetComponents(UriComponents.AbsoluteUri, UriFormat.UriEscaped);
            string orig = packageUri.OriginalString;

            foreach (var ch in _escapedChars)
            {
                orig = !orig.Contains(ch.ToString()) ? orig : orig.Replace(ch.ToString(), Uri.HexEscape(ch));
            }

            orig = orig.Replace('/', ',');

            if (partUri != null)
                orig += partUri.OriginalString;

            if ((fragment == null && partUri == null) && orig[orig.Length - 1] != '/')
                orig += '/';

            if (fragment != null)
                orig += fragment;

            return new Uri("pack://" + orig);
        }

        public static Uri CreatePartUri(Uri partUri)
        {
            Check.PartUri(partUri);

            if (partUri.OriginalString[0] != '/')
                partUri = new Uri("/" + partUri.ToString(), UriKind.Relative);
            return partUri;
        }

        public static Uri GetNormalizedPartUri(Uri partUri)
        {
            Check.PartUri(partUri);
            return new Uri(partUri.ToString().ToUpperInvariant(), UriKind.Relative);
        }

        public static Uri GetPackageUri(Uri packUri)
        {
            Check.PackUri(packUri);
            Check.PackUriIsValid(packUri);

            string s = packUri.Host.Replace(',', '/');
            return new Uri(Uri.UnescapeDataString(s), UriKind.RelativeOrAbsolute);
        }

        public static Uri GetPartUri(Uri packUri)
        {
            Check.PackUri(packUri);
            Check.PackUriIsValid(packUri);

            if (string.IsNullOrEmpty(packUri.AbsolutePath) || packUri.AbsolutePath == "/")
                return null;

            return new Uri(packUri.AbsolutePath, UriKind.Relative);
        }

        public static Uri GetRelationshipPartUri(Uri partUri)
        {
            Check.PartUri(partUri);
            Check.PartUriIsValid(partUri);

            int index = partUri.OriginalString.LastIndexOf("/");
            string s = partUri.OriginalString.Substring(0, index);
            s += "/_rels" + partUri.OriginalString.Substring(index) + ".rels";
            return new Uri(s, UriKind.Relative);
        }

        public static Uri GetRelativeUri(Uri sourcePartUri, Uri targetPartUri)
        {
            Check.SourcePartUri(sourcePartUri);
            Check.TargetPartUri(targetPartUri);

            Uri uri = new Uri("http://fake.com");
            Uri a = new Uri(uri, sourcePartUri.OriginalString);
            Uri b = new Uri(uri, targetPartUri.OriginalString);

            return a.MakeRelativeUri(b);
        }

        public static Uri GetSourcePartUriFromRelationshipPartUri(Uri relationshipPartUri)
        {
            //Check.RelationshipPartUri (relationshipPartUri);
            if (!IsRelationshipPartUri(relationshipPartUri))
                throw new Exception("is not a relationship part!?");
            return null;
        }

        public static bool IsRelationshipPartUri(Uri partUri)
        {
            Check.PartUri(partUri);
            return partUri.OriginalString.StartsWith("/_rels") && partUri.OriginalString.EndsWith(".rels");
        }

        public static Uri ResolvePartUri(Uri sourcePartUri, Uri targetUri)
        {
            Check.SourcePartUri(sourcePartUri);
            Check.TargetUri(targetUri);

            Check.PartUriIsValid(sourcePartUri);
         // commented out because on Android they are absolute file:///
			//   if (targetUri.IsAbsoluteUri)
         //       throw new ArgumentException("targetUri", "Absolute URIs are not supported");

            Uri uri = new Uri("http://fake.com");
            uri = new Uri(uri, sourcePartUri);
            uri = new Uri(uri, targetUri);

            // Trim out 'http://fake.com'
            return new Uri(uri.OriginalString.Substring(15), UriKind.Relative);
        }
    }
}
