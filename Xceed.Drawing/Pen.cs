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
  public enum DashStyle
  {
    Solid,
    Dash,
    Dot,
    DashDot,
    DashDotDot,
    Custom
  }

  public class Pen
  {
    #region Private Members

#if NET5
    private readonly SKPaint m_pen;
#else
    private readonly System.Drawing.Pen m_pen;
#endif

    #endregion

    #region Constructors
    public Pen( Color penColor, float penWidth )
    {
#if NET5
      m_pen = new SKPaint() { Color = penColor.Value, StrokeWidth = penWidth };
#else
      m_pen = new System.Drawing.Pen( penColor.Value, penWidth );
#endif
    }

    public Pen( Color penColor )
    {
#if NET5
      m_pen = new SKPaint() { Color = penColor.Value, StrokeWidth = 1 };
#else
      m_pen = new System.Drawing.Pen( penColor.Value, 1 );
#endif
    }

    #endregion

    #region Properties

    #region Color

    public Color Color
    {
      get
      {
        return new Color( m_pen.Color );
      }
      set
      {
        m_pen.Color = Color.Value;
      }
    }

    #endregion

    #region MiterLimit

    public float MiterLimit
    {
      get
      {
#if NET5
        return m_pen.StrokeMiter;
#else
        return m_pen.MiterLimit;
#endif
      }
    }

    #endregion

    #region DashPattern

    public float[] DashPattern
    {
      get;
      set;
    }

    #endregion

    #region DashOffset

    public float DashOffset
    {
      get;
    }

    #endregion

    #region DashStyle

    public DashStyle DashStyle
    {
      get;
      set;
    }

    #endregion

    #region Width

    public float Width
    {
      get
      {
#if NET5
        return m_pen.StrokeWidth;
#else
        return m_pen.Width;
#endif
      }
      set
      {
#if NET5
        m_pen.StrokeWidth = value;
#else
        m_pen.Width = value;
#endif
      }
    }

    #endregion

    #endregion

    #region Methods

    public int GetEndCap()
    {
#if NET5
      switch( m_pen.StrokeCap )
      {
        case SKStrokeCap.Butt:
          return 0;

        case SKStrokeCap.Round:
          return 1;

        case SKStrokeCap.Square:
          return 2;
      }
#else
      switch( m_pen.EndCap )
      {
        case System.Drawing.Drawing2D.LineCap.Flat:
          return 0;

        case System.Drawing.Drawing2D.LineCap.Round:
          return 1;

        case System.Drawing.Drawing2D.LineCap.Square:
          return 2;
      }
#endif

      return 0;
    }

    public int GetLineJoin()
    {
#if NET5
      switch( m_pen.StrokeJoin )
      {
        case SKStrokeJoin.Miter:
          return 0;

        case SKStrokeJoin.Round:
          return 1;

        case SKStrokeJoin.Bevel:
          return 2;
      }
#else
      switch( m_pen.LineJoin)
      {
        case System.Drawing.Drawing2D.LineJoin.Miter:
          return 0;

        case System.Drawing.Drawing2D.LineJoin.MiterClipped:
          return 0;

        case System.Drawing.Drawing2D.LineJoin.Round:
          return 1;

        case System.Drawing.Drawing2D.LineJoin.Bevel:
          return 2;
      }
#endif

      return 0;
    }

    #endregion
  }

}
