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
  public class Brush
  {
    #region Private Members

#if NET5
    private readonly SKPaint m_brush;
#else
    private readonly System.Drawing.Brush m_brush;
#endif

    #endregion

    #region Constructors

    public Brush( Color color )
    {
#if NET5
      m_brush = new SKPaint() { Color = color.Value };
#else
      m_brush = new System.Drawing.SolidBrush( color.Value );
#endif
    }

    #endregion

    #region Properties

    #region Color

    public Color Color
    {
      get
      {
#if NET5
        return new Color( m_brush.Color );
#else
        return new Color( (m_brush as System.Drawing.SolidBrush).Color );
#endif
      }
    }

    #endregion

    #endregion

    #region Methods

    public void Dispose()
    {
      m_brush.Dispose();
    }

    #endregion
  }
}
