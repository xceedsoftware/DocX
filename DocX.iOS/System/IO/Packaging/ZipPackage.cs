using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using DocX.iOS.Zip;

namespace System.IO.Packaging
{
	class UriComparer : IEqualityComparer<Uri>
	{
		public int GetHashCode(Uri uri)
		{
			return 1;
		}

		public bool Equals(Uri x, Uri y)
		{
			return x.OriginalString.Equals(y.OriginalString, StringComparison.OrdinalIgnoreCase);
		}
	}

	public sealed class ZipPackage : Package
	{
		public ZipStorer Archive { get; set; }
		private const string ContentNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
		private const string ContentUri = "[Content_Types].xml";

		bool OwnsStream
		{
			get;
			set;
		}

		Dictionary<Uri, ZipPackagePart> parts;
		internal Dictionary<Uri, MemoryStream> PartStreams = new Dictionary<Uri, MemoryStream>(new UriComparer());

		internal Stream PackageStream { get; set; }

		Dictionary<Uri, ZipPackagePart> Parts
		{
			get
			{
				if (parts == null)
					LoadParts();
				return parts;
			}
		}

		internal ZipPackage(FileAccess access, bool ownsStream, Stream stream)
			: base(access)
		{
			OwnsStream = ownsStream;
			PackageStream = stream;
		}

		internal ZipPackage(FileAccess access, bool ownsStream, Stream stream, bool streaming)
			: base(access, streaming)
		{
			OwnsStream = ownsStream;
			PackageStream = stream;
		}

		protected override void Dispose(bool disposing)
		{
			foreach (Stream s in PartStreams.Values)
				s.Close();

			if (Archive != null)	//	GZE		fixed bug where Archive == null
			{
				Archive.Close();
			}

			base.Dispose(disposing);

			if (OwnsStream)
				PackageStream.Close();
		}

		protected override void FlushCore()
		{
			// Ensure that all the data has been read out of the package
			// stream already. Otherwise we'll lose data when we recreate the zip

			foreach (ZipPackagePart part in Parts.Values)
			{
				part.GetStream();
			}
		    if (!PackageStream.CanSeek)
		        return;
			// Empty the package stream
			PackageStream.Position = 0;
			PackageStream.SetLength(0);

			// Recreate the zip file
			using (ZipStorer archive = ZipStorer.Create(PackageStream, "", false))
			{

				// Write all the part streams
				foreach (ZipPackagePart part in Parts.Values)
				{
					Stream partStream = part.GetStream();
					partStream.Seek(0, SeekOrigin.Begin);

					archive.AddStream(ZipStorer.Compression.Deflate, part.Uri.ToString().Substring(1), partStream,
						DateTime.UtcNow, "");
				}

				using (var ms = new MemoryStream())
				{
					WriteContentType(ms);
					ms.Seek(0, SeekOrigin.Begin);

					archive.AddStream(ZipStorer.Compression.Deflate, ContentUri, ms, DateTime.UtcNow, "");
				}
			}
		}


		protected override PackagePart CreatePartCore(Uri partUri, string contentType, CompressionOption compressionOption)
		{
			ZipPackagePart part = new ZipPackagePart(this, partUri, contentType, compressionOption);
			Parts.Add(part.Uri, part);
			return part;
		}

		protected override void DeletePartCore(Uri partUri)
		{
			Parts.Remove(partUri);
		}

		protected override PackagePart GetPartCore(Uri partUri)
		{
			ZipPackagePart part;
			Parts.TryGetValue(partUri, out part);
			return part;
		}

		protected override PackagePart[] GetPartsCore()
		{
			ZipPackagePart[] p = new ZipPackagePart[Parts.Count];
			Parts.Values.CopyTo(p, 0);
			return p;
		}

		void LoadParts()
		{
			parts = new Dictionary<Uri, ZipPackagePart>(new UriComparer());
			try
			{
				PackageStream.Seek(0, SeekOrigin.Begin);
				if (Archive == null)
				{
					Archive = ZipStorer.Open(PackageStream, FileAccess.Read, false);
				}
				List<ZipStorer.ZipFileEntry> dir = Archive.ReadCentralDir();

				// Load the content type map file
				XmlDocument doc = new XmlDocument();
				var content = dir.FirstOrDefault(x => x.FilenameInZip == ContentUri);
				using (var ms = new MemoryStream())
				{
					Archive.ExtractFile(content, ms);
					ms.Seek(0, SeekOrigin.Begin);
					doc.Load(ms);
				}

				XmlNamespaceManager manager = new XmlNamespaceManager(doc.NameTable);
				manager.AddNamespace("content", ContentNamespace);

				// The file names in the zip archive are not prepended with '/'
				foreach (var file in dir)
				{
					if (file.FilenameInZip.Equals(ContentUri, StringComparison.Ordinal))
						continue;

					XmlNode node;

					if (file.FilenameInZip == RelationshipUri.ToString().Substring(1))
					{
						CreatePartCore(RelationshipUri, RelationshipContentType, CompressionOption.Normal);
						continue;
					}

					string xPath = string.Format("/content:Types/content:Override[@PartName='/{0}']", file);
					node = doc.SelectSingleNode(xPath, manager);

					if (node == null)
					{
						string ext = Path.GetExtension(file.FilenameInZip);
						if (ext.StartsWith("."))
							ext = ext.Substring(1);
						xPath = string.Format("/content:Types/content:Default[@Extension='{0}']", ext);
						node = doc.SelectSingleNode(xPath, manager);
					}

					// What do i do if the node is null? This means some has tampered with the
					// package file manually
					if (node != null)
						CreatePartCore(new Uri("/" + file, UriKind.Relative), node.Attributes["ContentType"].Value,
							CompressionOption.Normal);
				}
			}
			catch
			{
				// The archive is invalid - therefore no parts
			}
		}

		void WriteContentType(Stream s)
		{
			XmlDocument doc = new XmlDocument();
			XmlNamespaceManager manager = new XmlNamespaceManager(doc.NameTable);
			Dictionary<string, string> mimes = new Dictionary<string, string>();

			manager.AddNamespace("content", ContentNamespace);

			doc.AppendChild(doc.CreateNode(XmlNodeType.XmlDeclaration, "", ""));

			XmlNode root = doc.CreateNode(XmlNodeType.Element, "Types", ContentNamespace);
			doc.AppendChild(root);
			foreach (ZipPackagePart part in Parts.Values)
			{
				XmlNode node = null;
				string existingMimeType;

				var extension = Path.GetExtension(part.Uri.OriginalString);
				if (extension.Length > 0)
					extension = extension.Substring(1);

				if (!mimes.TryGetValue(extension, out existingMimeType))
				{
					node = doc.CreateNode(XmlNodeType.Element, "Default", ContentNamespace);

					XmlAttribute ext = doc.CreateAttribute("Extension");
					ext.Value = extension;
					node.Attributes.Append(ext);
					mimes[extension] = part.ContentType;
				}
				else if (part.ContentType != existingMimeType)
				{
					node = doc.CreateNode(XmlNodeType.Element, "Override", ContentNamespace);

					XmlAttribute name = doc.CreateAttribute("PartName");
					name.Value = part.Uri.ToString();
					node.Attributes.Append(name);
				}

				if (node != null)
				{
					XmlAttribute contentType = doc.CreateAttribute("ContentType");
					contentType.Value = part.ContentType;
					node.Attributes.Prepend(contentType);

					root.AppendChild(node);
				}
			}

			XmlTextWriter writer = new XmlTextWriter(s, System.Text.Encoding.UTF8);
			doc.WriteTo(writer);
			writer.Flush();
		}
	}
}
