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
#endif

namespace Xceed.Drawing
{
  public class Bitmap
  {
    #region Private Members

#if NET5
    private SKBitmap m_bitmap;
#else
	  private System.Drawing.Bitmap m_bitmap;
#endif

    #endregion

    #region Constructors

    public Bitmap( int width, int height )
    {
#if NET5
      m_bitmap = new SKBitmap( width, height );
#else
      m_bitmap = new System.Drawing.Bitmap( width, height );
#endif
    }

    private Bitmap(
#if NET5
           SKBitmap bitmap
#else
           System.Drawing.Bitmap bitmap
#endif
     )
    {
      m_bitmap = bitmap;
    }

    #endregion

    #region Properties

    #region Value

#if NET5
    public SKBitmap Value
#else
    public System.Drawing.Bitmap Value
#endif
    {
      get
      {
        return m_bitmap;
      }
    }

    #endregion

    #endregion

    #region Static Methods

    public static Bitmap FromImage( Image image )
    {
#if NET5
      return new Bitmap( SKBitmap.FromImage( image.Value ) );
#else
      return new Bitmap( image.Value as System.Drawing.Bitmap );
#endif
    }

    #endregion

    #region Public Methods

    public void SetPixel( int x, int y, Color color )
    {
      m_bitmap.SetPixel( x, y, color.Value );
    }

    public void Dispose()
    {
      m_bitmap.Dispose();
    }

    #endregion
  }
}
