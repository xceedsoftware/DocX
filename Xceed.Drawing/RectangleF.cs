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
  public struct RectangleF
  {
    #region Private Members

#if NET5
    private SKRect m_rect;
#else
    private System.Drawing.RectangleF m_rect;
#endif

    #endregion

    #region Constructors

    public readonly static RectangleF Empty = new RectangleF();

    private RectangleF(
#if NET5
            SKRect rect
#else
            System.Drawing.RectangleF rect
#endif
      )
    {
      m_rect = rect;
    }

    public RectangleF( float x, float y, float width, float height )
    {
#if NET5
      m_rect = new SKRect() { Location = new SKPoint( x, y ), Size = new SKSize( width, height ) };
#else
      m_rect = new System.Drawing.RectangleF( x, y, width, height );
#endif
    }

    #endregion

    #region Properties

    #region Bottom

    public float Bottom
    {
      get
      {
        return m_rect.Bottom;
      }
    }

    #endregion

    #region Left

    public float Left
    {
      get
      {
        return m_rect.Left;
      }
    }

    #endregion

    #region Height

    public float Height
    {
      get
      {
        return m_rect.Height;
      }
      set
      {
#if NET5
        m_rect.Size = new SKSize( m_rect.Size.Height, value );
#else
        m_rect.Height = value;
#endif
      }
    }

    #endregion

    #region Right

    public float Right
    {
      get
      {
        return m_rect.Right;
      }
    }

    #endregion

    #region Top

    public float Top
    {
      get
      {
        return m_rect.Top;
      }
    }

    #endregion

    #region X

    public float X
    {
      get
      {
#if NET5
        return m_rect.Location.X;
#else
        return m_rect.X;
#endif
      }
      set
      {
#if NET5
        m_rect.Location = new SKPoint( value, m_rect.Location.Y );
#else
        m_rect.X = value;
#endif
      }
    }

    #endregion

    #region Y

    public float Y
    {
      get
      {
#if NET5
        return m_rect.Location.Y;
#else
        return m_rect.Y;
#endif
      }
      set
      {
#if NET5
        m_rect.Location = new SKPoint( m_rect.Location.X, value );
#else
        m_rect.Y = value;
#endif
      }
    }

    #endregion

    #region Value

#if NET5
    public SKRect Value
#else
    public System.Drawing.RectangleF Value
#endif
    {
      get
      {
        return m_rect;
      }
    }

    #endregion

    #region Width

    public float Width
    {
      get
      {
        return m_rect.Width;
      }
      set 
      {
#if NET5
        m_rect.Size = new SKSize( value, m_rect.Size.Width );
#else
        m_rect.Width = value;
#endif
      }
    }

    #endregion

    #endregion

    #region Public Methods

    public override bool Equals( object obj )
    {
      if( !( obj is RectangleF ) )
        return false;

      var other = (RectangleF)obj;

      return this.X == other.X
            && this.Y == other.Y
            && this.Width == other.Width
            && this.Height == other.Height
            && this.Left == other.Left
            && this.Right == other.Right
            && this.Top == other.Top
            && this.Bottom == other.Bottom
            && this.Value == other.Value;
    }

    public override int GetHashCode()
    {
      var hash = 17;
      hash = hash * 31 + this.X.GetHashCode();
      hash = hash * 31 + this.Y.GetHashCode();
      hash = hash * 31 + this.Width.GetHashCode();
      hash = hash * 31 + this.Height.GetHashCode();
      hash = hash * 31 + this.Left.GetHashCode();
      hash = hash * 31 + this.Right.GetHashCode();
      hash = hash * 31 + this.Top.GetHashCode();
      hash = hash * 31 + this.Bottom.GetHashCode();
      hash = hash * 31 + this.Value.GetHashCode();

      return hash;      
    }

    public static bool operator ==( RectangleF rect1, RectangleF rect2 )
    {
      return rect1.Equals( rect2 );
    }

    public static bool operator !=( RectangleF rect1, RectangleF rect2 )
    {
      return !rect1.Equals( rect2 );
    }

    public bool IntersectsWith( RectangleF rect )
    {
      return m_rect.IntersectsWith( rect.Value );
    }

    #endregion
  }
}
