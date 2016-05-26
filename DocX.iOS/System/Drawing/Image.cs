using System;
using System.IO;

using Foundation;
using UIKit;

namespace System.Drawing
{
	//	Image

	public class Image : MarshalByRefObject, IDisposable
	{
		//	FromStream

		public static Image FromStream(Stream stream)
		{
			return new Image (stream);
		}

		//	properties

		public bool Disposed { get; private set; }

		public int Height { get; private set; }

		public int Width { get; private set; } 

		//	constructor

		private Image (Stream stream)
		{
			using (var image = UIImage.LoadFromData (NSData.FromStream (stream)))
			{
				this.Width = (int)image.Size.Width;

				this.Height = (int)image.Size.Height;
			}
		}

		//	destructor

		~Image() 
		{
			Dispose(false);
		}

		//	Dispose

		public void Dispose()
		{
			this.Dispose(true);

			GC.SuppressFinalize(this);
		}

		//	Dispose

		protected virtual void Dispose(bool disposing)
		{
			if (this.Disposed)
			{
				return; 
			}

			try
			{
				try
				{
					if (disposing) 
					{
						//	dispose managed
					}
				}
				finally
				{
					//	dispose unmanaged
				}
			}
			finally
			{
				this.Disposed = true;
			}
		}
	}
}