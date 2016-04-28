using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocX.iOS.Zip;

namespace System.IO.Packaging
{
    public sealed class ZipPackagePart : PackagePart
    {
        new ZipPackage Package
        {
            get { return (ZipPackage)base.Package; }
        }

        internal ZipPackagePart(Package package, Uri partUri)
            : base(package, partUri)
        {

        }

        internal ZipPackagePart(Package package, Uri partUri, string contentType)
            : base(package, partUri, contentType)
        {

        }

        internal ZipPackagePart(Package package, Uri partUri, string contentType, CompressionOption compressionOption)
            : base(package, partUri, contentType, compressionOption)
        {

        }

        protected override Stream GetStreamCore(FileMode mode, FileAccess access)
        {
            ZipPartStream zps;
            MemoryStream stream;
            if (Package.PartStreams.TryGetValue(Uri, out stream))
            {
                //zps = new ZipPartStream(Package, stream, access);
                if (mode == FileMode.Create)
                    stream.SetLength(0);
                return new ZipPartStream(Package, stream, access);
            }

            stream = new MemoryStream();
            try
            {
                if (Package.Archive == null)
                {
					Package.Archive = ZipStorer.Open(Package.PackageStream, access, false);
                }
                List<ZipStorer.ZipFileEntry> dir = Package.Archive.ReadCentralDir();
                foreach (ZipStorer.ZipFileEntry entry in dir)
                {
                    if (entry.FilenameInZip != Uri.ToString().Substring(1))
                        continue;

                    Package.Archive.ExtractFile(entry, stream);
                }
            }
            catch
            {
                // The zipfile is invalid, so just create the file
                // as if it didn't exist
                stream.SetLength(0);
            }

            Package.PartStreams.Add(Uri, stream);
            if (mode == FileMode.Create)
                stream.SetLength(0);
            return new ZipPartStream(Package, stream, access);
        }
    }
}
