using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
//using DocumentFormat.OpenXml.Packaging;

namespace System.IO.Packaging
{
    public abstract class Package : IDisposable
    {
        internal const string RelationshipContentType = "application/vnd.openxmlformats-package.relationships+xml";
        internal const string RelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        internal static readonly Uri RelationshipUri = new Uri("/_rels/.rels", UriKind.Relative);

        private PackageProperties packageProperties;
        private PackagePartCollection partsCollection;
        private Dictionary<string, PackageRelationship> relationships;
        private PackageRelationshipCollection relationshipsCollection = new PackageRelationshipCollection();
        private Uri Uri = new Uri("/", UriKind.Relative);

        private bool Disposed { get; set; }

        public FileAccess FileOpenAccess { get; private set; }

        public PackageProperties PackageProperties
        {
            get
            {
                // PackageProperties are loaded when the relationships are loaded.
                // Therefore ensure we've already loaded the relationships.
                int count = Relationships.Count;

                if (packageProperties == null)
                {
                    packageProperties = new PackagePropertiesPart();
                    packageProperties.Package = this;
                }
                return packageProperties;
            }
        }

        private PackagePartCollection PartsCollection
        {
            get
            {
                if (partsCollection == null)
                {
                    partsCollection = new PackagePartCollection();
                    partsCollection.Parts.AddRange(GetPartsCore());
                }
                return partsCollection;
            }
        }

        private int RelationshipId { get; set; }

        private Dictionary<string, PackageRelationship> Relationships
        {
            get
            {
                if (relationships == null)
                {
                    LoadRelationships();
                }
                return relationships;
            }
        }

        private bool Streaming { get; set; }


        protected Package(FileAccess openFileAccess)
            : this(openFileAccess, false)
        {

        }

        protected Package(FileAccess openFileAccess, bool streaming)
        {
            FileOpenAccess = openFileAccess;
            Streaming = streaming;
        }


        internal void CheckIsReadOnly()
        {
            if (FileOpenAccess == FileAccess.Read)
                throw new IOException("Operation not valid when package is read-only");
        }

        public void Close()
        {
            // FIXME: Ensure that Flush is actually called before dispose            
            ((IDisposable) this).Dispose();
        }

        public PackagePart CreatePart(Uri partUri, string contentType)
        {
            return CreatePart(partUri, contentType, CompressionOption.NotCompressed);
        }

        public PackagePart CreatePart(Uri partUri, string contentType, CompressionOption compressionOption)
        {
            CheckIsReadOnly();
            Check.PartUri(partUri);
            Check.ContentTypeIsValid(contentType);

            if (PartExists(partUri))
                throw new InvalidOperationException("This partUri is already contained in the package");

            PackagePart part = CreatePartCore(partUri, contentType, compressionOption);
            PartsCollection.Parts.Add(part);
            return part;
        }

        protected abstract PackagePart CreatePartCore(Uri partUri, string contentType,
            CompressionOption compressionOption);

        public PackageRelationship CreateRelationship(Uri targetUri, TargetMode targetMode, string relationshipType)
        {
            return CreateRelationship(targetUri, targetMode, relationshipType, null);
        }

        public PackageRelationship CreateRelationship(Uri targetUri, TargetMode targetMode, string relationshipType,
            string id)
        {
            return CreateRelationship(targetUri, targetMode, relationshipType, id, false);
        }

        internal PackageRelationship CreateRelationship(Uri targetUri, TargetMode targetMode, string relationshipType,
            string id, bool loading)
        {
            if (!loading)
                CheckIsReadOnly();

            Check.TargetUri(targetUri);
            if (targetUri.IsAbsoluteUri && targetMode == TargetMode.Internal)
                throw new ArgumentException("TargetUri cannot be absolute for an internal relationship");

            Check.RelationshipTypeIsValid(relationshipType);
            Check.IdIsValid(id);

            if (id == null)
                id = NextId();

            PackageRelationship r = new PackageRelationship(id, this, relationshipType, Uri, targetMode, targetUri);

            if (!PartExists(RelationshipUri))
                CreatePartCore(RelationshipUri, RelationshipContentType, CompressionOption.NotCompressed).IsRelationship
                    = true;

            Relationships.Add(r.Id, r);
            relationshipsCollection.Relationships.Add(r);

            if (!loading)
            {
                using (Stream s = GetPart(RelationshipUri).GetStream())
                    WriteRelationships(relationships, s);
            }

            return r;
        }


        public void DeletePart(Uri partUri)
        {
            CheckIsReadOnly();
            Check.PartUri(partUri);

            PackagePart part = GetPart(partUri);
            if (part != null)
            {
                if (part.Package == null)
                    throw new InvalidOperationException("This part has already been removed");

                // FIXME: MS.NET doesn't remove the relationship part
                // Instead it throws an exception if you try to use it
                if (PartExists(part.RelationshipsPartUri))
                    GetPart(part.RelationshipsPartUri).Package = null;

                part.Package = null;
                DeletePartCore(partUri);
                PartsCollection.Parts.RemoveAll(p => p.Uri == partUri);
            }
        }

        protected abstract void DeletePartCore(Uri partUri);

        public void DeleteRelationship(string id)
        {
            Check.Id(id);
            CheckIsReadOnly();

            Relationships.Remove(id);

            relationshipsCollection.Relationships.RemoveAll(r => r.Id == id);
            if (Relationships.Count > 0)
                using (Stream s = GetPart(RelationshipUri).GetStream())
                    WriteRelationships(relationships, s);
            else
                DeletePart(RelationshipUri);
        }

        void IDisposable.Dispose()
        {
            if (!Disposed)
            {
                Flush();
                Dispose(true);
                Disposed = true;
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            // Nothing here needs to be disposed of
        }

        private bool flushing = false;

        public void Flush()
        {
            if (FileOpenAccess == FileAccess.Read || flushing)
                return;

            flushing = true;

            // Ensure we've loaded the relationships, parts and properties
            int count = Relationships.Count;

            if (packageProperties != null)
                packageProperties.Flush();

            FlushCore();

            flushing = false;
        }

        protected abstract void FlushCore();

        public PackagePart GetPart(Uri partUri)
        {
            Check.PartUri(partUri);
            return GetPartCore(partUri);
        }

        protected abstract PackagePart GetPartCore(Uri partUri);

        public PackagePartCollection GetParts()
        {
            PartsCollection.Parts.Clear();
            PartsCollection.Parts.AddRange(GetPartsCore());
            return PartsCollection;
        }

        protected abstract PackagePart[] GetPartsCore();

        public PackageRelationship GetRelationship(string id)
        {
            return Relationships[id];
        }

        public PackageRelationshipCollection GetRelationships()
        {
            // Ensure the Relationships dict is instantiated first.
            ICollection<PackageRelationship> rels = Relationships.Values;
            relationshipsCollection.Relationships.Clear();
            relationshipsCollection.Relationships.AddRange(rels);
            return relationshipsCollection;
        }

        public PackageRelationshipCollection GetRelationshipsByType(string relationshipType)
        {
            PackageRelationshipCollection collection = new PackageRelationshipCollection();
            foreach (PackageRelationship r in Relationships.Values)
                if (r.RelationshipType == relationshipType)
                    collection.Relationships.Add(r);

            return collection;
        }

        private void LoadRelationships()
        {
            relationships = new Dictionary<string, PackageRelationship>();

            if (!PartExists(RelationshipUri))
                return;

            using (Stream stream = GetPart(RelationshipUri).GetStream())
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(stream);
                XmlNamespaceManager manager = new XmlNamespaceManager(doc.NameTable);
                manager.AddNamespace("rel", RelationshipNamespace);

                foreach (XmlNode node in doc.SelectNodes("/rel:Relationships/*", manager))
                {
                    TargetMode mode = TargetMode.Internal;
                    if (node.Attributes["TargetMode"] != null)
                        mode = (TargetMode) Enum.Parse(typeof (TargetMode), node.Attributes["TargetMode"].Value);

                    Uri uri;
                    try
                    {
                        uri = new Uri(node.Attributes["Target"].Value.ToString(), UriKind.Relative);
                    }
                    catch
                    {
                        uri = new Uri(node.Attributes["Target"].Value.ToString(), UriKind.Absolute);
                    }
                    CreateRelationship(uri,
                        mode,
                        node.Attributes["Type"].Value.ToString(),
                        node.Attributes["Id"].Value.ToString(),
                        true);
                }

                foreach (PackageRelationship r in relationships.Values)
                {
                    if (r.RelationshipType == PackageProperties.NSPackagePropertiesRelation)
                    {
                        PackagePart part = GetPart(PackUriHelper.ResolvePartUri(Uri, r.TargetUri));
                        packageProperties = new PackagePropertiesPart();
                        packageProperties.Package = this;
                        packageProperties.Part = part;
                        packageProperties.LoadFrom(part.GetStream());
                    }
                }
            }
        }

        private string NextId()
        {
            while (true)
            {
                string s = "Re" + RelationshipId.ToString();
                if (!Relationships.ContainsKey(s))
                    return s;

                RelationshipId++;
            }
        }

        public static Package Open(Stream stream)
        {
            return Open(stream, FileMode.Open);
        }

        public static Package Open(string path)
        {
            return Open(path, FileMode.OpenOrCreate);
        }

        public static Package Open(Stream stream, FileMode packageMode)
        {
            FileAccess access = packageMode == FileMode.Open ? FileAccess.Read : FileAccess.ReadWrite;
            return Open(stream, packageMode, access);
        }

        public static Package Open(string path, FileMode packageMode)
        {
            return Open(path, packageMode, FileAccess.ReadWrite);
        }

        public static Package Open(Stream stream, FileMode packageMode, FileAccess packageAccess)
        {
            return Open(stream, packageMode, packageAccess, false);
        }

        private static Package Open(Stream stream, FileMode packageMode, FileAccess packageAccess, bool ownsStream)
        {
            return OpenCore(stream, packageMode, packageAccess, ownsStream);
        }

        public static Package Open(string path, FileMode packageMode, FileAccess packageAccess)
        {
            return Open(path, packageMode, packageAccess, FileShare.None);
        }

        public static Package Open(string path, FileMode packageMode, FileAccess packageAccess, FileShare packageShare)
        {
            if (packageShare != FileShare.Read && packageShare != FileShare.None)
                throw new NotSupportedException("FileShare.Read and FileShare.None are the only supported options");

            FileInfo info = new FileInfo(path);

            // Bug - MS.NET appears to test for FileAccess.ReadWrite, not FileAccess.Write
            if (packageAccess != FileAccess.ReadWrite && !info.Exists)
                throw new ArgumentException("packageAccess", "Cannot create stream with FileAccess.Read");


            if (info.Exists && packageMode == FileMode.OpenOrCreate && info.Length == 0)
                throw new FileFormatException("Stream length cannot be zero with FileMode.Open");

            Stream s = File.Open(path, packageMode, packageAccess, packageShare);
            return Open(s, packageMode, packageAccess, true);
        }

        private static Package OpenCore(Stream stream, FileMode packageMode, FileAccess packageAccess, bool ownsStream)
        {
            if ((packageAccess & FileAccess.Read) == FileAccess.Read && !stream.CanRead)
                throw new IOException("Stream does not support reading");

            if ((packageAccess & FileAccess.Write) == FileAccess.Write && !stream.CanWrite)
                throw new IOException("Stream does not support reading");

            if (!stream.CanSeek)
                throw new ArgumentException("stream", "Stream must support seeking");

            if (packageMode == FileMode.Open && stream.Length == 0)
                throw new FileFormatException("Stream length cannot be zero with FileMode.Open");

            if (packageMode == FileMode.CreateNew && stream.Length > 0)
                throw new IOException("Cannot use CreateNew when stream contains data");

            if (packageMode == FileMode.Append || packageMode == FileMode.Truncate)
            {
                if (stream.CanWrite)
                    throw new NotSupportedException(string.Format("PackageMode.{0} is not supported", packageMode));
                else
                    throw new IOException(string.Format("PackageMode.{0} is not supported", packageMode));
            }

            return new ZipPackage(packageAccess, ownsStream, stream);
        }

        public virtual bool PartExists(Uri partUri)
        {
            return GetPart(partUri) != null;
        }

        public bool RelationshipExists(string id)
        {
            return Relationships.ContainsKey(id);
        }

        internal static void WriteRelationships(Dictionary<string, PackageRelationship> relationships, Stream stream)
        {
            XmlDocument doc = new XmlDocument();
            XmlNamespaceManager manager = new XmlNamespaceManager(doc.NameTable);
            manager.AddNamespace("rel", RelationshipNamespace);

            doc.AppendChild(doc.CreateNode(XmlNodeType.XmlDeclaration, "", ""));

            XmlNode root = doc.CreateNode(XmlNodeType.Element, "Relationships", RelationshipNamespace);
            doc.AppendChild(root);

            foreach (PackageRelationship relationship in relationships.Values)
            {
                XmlNode node = doc.CreateNode(XmlNodeType.Element, "Relationship", RelationshipNamespace);

                XmlAttribute idAtt = doc.CreateAttribute("Id");
                idAtt.Value = relationship.Id;
                node.Attributes.Append(idAtt);

                XmlAttribute targetAtt = doc.CreateAttribute("Target");
                targetAtt.Value = relationship.TargetUri.ToString();
                node.Attributes.Append(targetAtt);

                if (relationship.TargetMode != TargetMode.Internal)
                {
                    XmlAttribute modeAtt = doc.CreateAttribute("TargetMode");
                    modeAtt.Value = relationship.TargetMode.ToString();
                    node.Attributes.Append(modeAtt);
                }
                XmlAttribute typeAtt = doc.CreateAttribute("Type");
                typeAtt.Value = relationship.RelationshipType;
                node.Attributes.Append(typeAtt);

                root.AppendChild(node);
            }

            using (XmlTextWriter writer = new XmlTextWriter(stream, System.Text.Encoding.UTF8))
                doc.WriteTo(writer);
        }
    }

    internal class FileFormatException : Exception
    {
        public FileFormatException(string message): base(message)
        {
            
        }
    }
}
