using System.IO;
using System.Threading;

namespace Novacode
{
    /// <summary>
    /// See <a href="https://support.microsoft.com/en-gb/kb/951731" /> for explanation
    /// </summary>
    public class PackagePartStream : Stream
    {
        private static readonly Mutex Mutex = new Mutex(false);

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
            Mutex.WaitOne(Timeout.Infinite, false);
            this.stream.Write(buffer, offset, count);
            Mutex.ReleaseMutex();
        }

        public override void Flush()
        {
            Mutex.WaitOne(Timeout.Infinite, false);
            this.stream.Flush();
            Mutex.ReleaseMutex();
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
