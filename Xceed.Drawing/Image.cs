/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/



#if NET5
using SkiaSharp;
using System;

#else
using System;
using System.Drawing;
#endif

using System.IO;

namespace Xceed.Drawing
{
  public class Image : System.IDisposable
  {
    #region Private Members

#if NET5
    private SKImage m_image;
#else
    private System.Drawing.Image m_image;
#endif

    #endregion

    #region Constructors

    private Image(
#if NET5
            SKImage image
#else
            System.Drawing.Image image
#endif
      )
    {
      m_image = image;
    }

    #endregion

    #region Properties

    #region Height

    public int Height
    {
      get
      {
        return m_image.Height;
      }
    }

    #endregion

    #region HorizontalResolution

    public float HorizontalResolution
    {
      get
      {
#if NET5
        return -1f;
#else
        return m_image.HorizontalResolution;
#endif
      }
    }

    #endregion

    #region Value

#if NET5
    public SKImage Value
#else
    public System.Drawing.Image Value
#endif
    {
      get
      {
        return m_image;
      }
    }

    #endregion

    #region VerticalResolution

    public float VerticalResolution
    {
      get
      {
#if NET5
        return -1f;
#else
        return m_image.VerticalResolution;
#endif
      }
    }

    #endregion

    #region Width

    public int Width
    {
      get
      {
        return m_image.Width;
      }
    }

    #endregion

    #endregion

    #region Static Methods

    public static Image FromBitmap( Bitmap bitmap )
    {
#if NET5
      return new Image( SKImage.FromBitmap( bitmap.Value ) );
#else
      return new Image( bitmap.Value );
#endif
    }

    public static Image FromStream( Stream stream )
    {
#if NET5
      stream.Position = 0;
      return new Image( SKImage.FromEncodedData( stream ) );
      //var skBitmap = SKBitmap.Decode( new SKManagedStream( stream ) );
      //return new Image( SKImage.FromBitmap( skBitmap ) );
#else
      return new Image( System.Drawing.Image.FromStream( stream ) );
#endif
    }

    public static Image FromFile( string file )
    {
#if NET5
      return new Image( SKImage.FromEncodedData( file ) );
#else
      return new Image( System.Drawing.Image.FromFile( file ) );
#endif
    }

    public static string DetectFormat( Stream stream )
    {
      var image = Image.FromStream( stream ).Value;

#if NET5
        if( image.Encode( SKEncodedImageFormat.Jpeg, 100 ) != null )
          return "JPEG";
        else if( image.Encode( SKEncodedImageFormat.Png, 100 ) != null )
          return "PNG";
        else if( image.Encode( SKEncodedImageFormat.Gif, 100 ) != null )
          return "GIF";
        else if( image.Encode( SKEncodedImageFormat.Bmp, 100 ) != null )
          return "BMP";
        else
          throw new Exception( "Picture has an invalid format." );
#else
        if( System.Drawing.Imaging.ImageFormat.Jpeg.Equals( image.RawFormat ) )
          return "JPEG";
        else if( System.Drawing.Imaging.ImageFormat.Png.Equals( image.RawFormat ) )
          return "PNG";
        else if( System.Drawing.Imaging.ImageFormat.Gif.Equals( image.RawFormat ) )
          return "GIF";
        else if( System.Drawing.Imaging.ImageFormat.Tiff.Equals( image.RawFormat ) )
          return "TIFF";
         else if( System.Drawing.Imaging.ImageFormat.Bmp.Equals( image.RawFormat ) )
          return "BMP";
        else
          throw new Exception( "Picture has an invalid format." );
#endif
    }

    public static Image Clone( Image image )
    {
      if( image == null )
        return null;

#if NET5
      // Get the width and height of the original image
      int width = image.Width;
      int height = image.Height;

      // Create a new surface with the same dimensions
      using( var surface = SKSurface.Create( new SKImageInfo( width, height ) ) )
      {
        // Get the canvas from the new surface
        var canvas = surface.Canvas;

        // Clear the canvas (optional)
        canvas.Clear( SKColors.Transparent );

        // Draw the original image onto the canvas
        canvas.DrawImage( image.Value, 0, 0 );

        // Retrieve the drawn content as a new SKImage
        return new Image( surface.Snapshot() );
      }
      // or
      //var imageSize = image.Value.Info.Size;
      //var clonedBitmap = new SKBitmap( imageSize.Width, imageSize.Height, image.Value.ColorType, image.Value.AlphaType );
      //using( var canvas = new SKCanvas( clonedBitmap ) )
      //{
      //  canvas.Clear( SKColors.Transparent );
      //  canvas.DrawImage( image.Value, SKRect.Create( imageSize ) );
      //}
      //image = new Image( SKImage.FromBitmap( clonedBitmap ) );
#else
      return new Image( image.Value.Clone() as System.Drawing.Image );
#endif
    }

    public static MemoryStream CompressJPEG( Image image, int quality )
    {
      if( image == null )
        return null;

#if NET5
      var memoryStream = new MemoryStream();
      var data = image.Value.Encode( SKEncodedImageFormat.Jpeg, quality );
      data.SaveTo( memoryStream );

      return memoryStream;
#else
      System.Drawing.Imaging.ImageCodecInfo imageCodecInfo;
      System.Drawing.Imaging.Encoder encoder;
      System.Drawing.Imaging.EncoderParameter encoderParameter;
      System.Drawing.Imaging.EncoderParameters encoderParameters;

      imageCodecInfo = Image.GetEncoderInfo( "image/jpeg" );
      encoder = System.Drawing.Imaging.Encoder.Quality;
      encoderParameters = new System.Drawing.Imaging.EncoderParameters( 1 );
      encoderParameter = new System.Drawing.Imaging.EncoderParameter( encoder, (long)quality );
      encoderParameters.Param[ 0 ] = encoderParameter;

      var memoryStream = new System.IO.MemoryStream();
      image.Value.Save( memoryStream, imageCodecInfo, encoderParameters );
      memoryStream.Flush();

      return memoryStream;   
#endif
    }

#if !NET5
    private static System.Drawing.Imaging.ImageCodecInfo GetEncoderInfo( string mimeType )
    {
      int i;
      System.Drawing.Imaging.ImageCodecInfo[] encoders;
      encoders = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders();

      for( i = 0; i < encoders.Length; ++i )
      {
        if( encoders[ i ].MimeType == mimeType ) return encoders[ i ];
      }

      return null;
    }
#endif

#endregion

    #region Public Methods

    public Image Draw( uint rotation, RectangleF cropping, out float updatedWidthPercent, out float updatedHeightPercent )
    {
      updatedWidthPercent = 1f;
      updatedHeightPercent = 1f;

      var rotatedImage = new Bitmap( m_image.Width, m_image.Height );
#if NET5
      using( var canvas = new SKCanvas( rotatedImage.Value ) )
      {
        canvas.Clear( SKColors.White );

        canvas.Translate( rotatedImage.Value.Width / 2, rotatedImage.Value.Height / 2 ); // set the rotation point as the center into the matrix
        canvas.RotateDegrees( rotation ); // rotate
        canvas.Translate( -rotatedImage.Value.Width / 2, -rotatedImage.Value.Height / 2 ); // restore rotation point into the matrix

        var scaleX = (float)rotatedImage.Value.Width / m_image.Width;
        var scaleY = (float)rotatedImage.Value.Height / m_image.Height;

        if( cropping != RectangleF.Empty )
        {
          var cropLeft = (int)( cropping.X / 100f * m_image.Width );
          var cropTop = (int)( cropping.Y / 100f * m_image.Height );
          var cropWidth = (int)( m_image.Width - ( cropping.Width / 100f * m_image.Width ) - cropLeft );
          var cropHeight = (int)( m_image.Height - ( cropping.Height / 100f * m_image.Height ) - cropTop );

          var srcRect = SKRectI.Create( cropLeft, cropTop, cropWidth, cropHeight );
          var destRect = SKRect.Create( 0, 0, rotatedImage.Value.Width, rotatedImage.Value.Height );

          canvas.DrawImage( m_image, srcRect, destRect );
        }
        else
        {
          canvas.DrawImage( m_image, SKRect.Create( 0, 0, m_image.Width * scaleX, m_image.Height * scaleY ) ); // draw the image on the new bitmap
        }
      }
#else
      using( var g = System.Drawing.Graphics.FromImage( rotatedImage.Value ) )
      {
        g.Clear( Color.White.Value );  // Use the Background color of the paper to draw transparency.

        g.TranslateTransform( m_image.Width / 2, m_image.Height / 2 ); //set the rotation point as the center into the matrix
        g.RotateTransform( rotation ); //rotate
        g.TranslateTransform( g.VisibleClipBounds.X, g.VisibleClipBounds.Y ); //restore rotation point into the matrix

        updatedWidthPercent = g.VisibleClipBounds.Width / m_image.Width;
        updatedHeightPercent = g.VisibleClipBounds.Height / m_image.Height;

        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.High;
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

        if( cropping != RectangleF.Empty )
        {
          var cropLeft = System.Convert.ToInt32( cropping.X / 100f * m_image.Width );
          var cropTop = System.Convert.ToInt32( cropping.Y / 100f * m_image.Height );
          var cropWidth = System.Convert.ToInt32( m_image.Width - ( cropping.Width / 100f * m_image.Width ) - cropLeft );
          var cropHeight = System.Convert.ToInt32( m_image.Height - ( cropping.Height / 100f * m_image.Height ) - cropTop );
          g.DrawImage( m_image,
                       new Rectangle( 0, 0, m_image.Width, m_image.Height ),
                       new Rectangle( cropLeft, cropTop, cropWidth, cropHeight ),
                       System.Drawing.GraphicsUnit.Pixel );
        }
        else
        {
          g.DrawImage( m_image, 0f, 0f, m_image.Width * updatedWidthPercent, m_image.Height * updatedHeightPercent ); //draw the image on the new bitmap
        }
      }
#endif

      return Image.FromBitmap( rotatedImage );
    }

    public void Dispose()
    {
      m_image.Dispose();
    }

    #endregion
  }
}
