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
  public struct PointF
  {
    #region Private Members

#if NET5
    private SKPoint m_point;
#else
    private System.Drawing.PointF m_point;
#endif

    #endregion

    #region Constructors

    public PointF( float x, float y )
    {
#if NET5
      m_point = new SKPoint( x, y );
#else
      m_point = new System.Drawing.PointF( x, y );
#endif
    }

    #endregion

    #region Properties

    #region X

    public float X
    {
      get
      {
        return m_point.X;
      }
      set
      {
        m_point.X = value;
      }
    }

    #endregion

    #region Y

    public float Y
    {
      get
      {
        return m_point.Y;
      }
      set
      {
        m_point.Y = value;
      }
    }

    #endregion

    #endregion

    #region Public Methods

    public override bool Equals( object obj )
    {
      if( !( obj is PointF ) )
        return false;

      var other = (PointF)obj;

      return this.X == other.X
           && this.Y == other.Y;
    }

    public override int GetHashCode()
    {
      var hash = 17;
      hash = hash * 31 + this.X.GetHashCode();
      hash = hash * 31 + this.Y.GetHashCode();

      return hash;
    }

    public static bool operator ==( PointF c1, PointF c2 )
    {
      return c1.Equals( c2 );
    }

    public static bool operator !=( PointF c1, PointF c2 )
    {
      return !c1.Equals( c2 );
    }

    #endregion
  }
}
