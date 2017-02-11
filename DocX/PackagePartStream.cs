using System.IO;

namespace Novacode
{
    /// <summary>
    /// OpenXML Isolated Storage access is not thread safe.
    /// Use app domain wide lock for writing.
    /// </summary>
    public class PackagePartStream : Stream
    {
        private static readonly object lockObject = new object();

        private readonly Stream stream;

        public PackagePartStream(Stream stream)
        {
            this.stream = stream;
        }

        public override bool CanRead
        {
            get { return this.stream.CanRead; }
        }

        public override bool CanSeek
        {
            get { return this.stream.CanSeek; }
        }

        public override bool CanWrite
        {
            get { return this.stream.CanWrite; }
        }

        public override long Length
        {
            get { return this.stream.Length; }
        }

        public override long Position
        {
            get { return this.stream.Position; }
            set { this.stream.Position = value; }
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return this.stream.Seek(offset, origin);
        }

        public override void SetLength(long value)
        {
            this.stream.SetLength(value);
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return this.stream.Read(buffer, offset, count);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            lock (lockObject)
            {
                this.stream.Write(buffer, offset, count);
            }
        }

        public override void Flush()
        {
            lock (lockObject)
            {
                this.stream.Flush();
            }
        }

        public override void Close()
        {
            this.stream.Close();
        }

        protected override void Dispose(bool disposing)
        {
            this.stream.Dispose();
        }
    }
}
