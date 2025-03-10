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

using System;

namespace Xceed.Drawing
{
    public struct Point
    {
    #region Private Members

#if NET5
    private SKPoint m_point;
#else
    private System.Drawing.Point m_point;
#endif

    #endregion

    #region Constructors

    public Point( int x, int y )
    {
#if NET5
      m_point = new SKPoint( x, y );
#else
      m_point = new System.Drawing.Point( x, y );
#endif
    }

    #endregion

    #region Properties

    #region X

    public int X
    {
      get 
      {
        return Convert.ToInt32( m_point.X );
      }
      set 
      {
        m_point.X = Convert.ToInt32( value );
      }
    }

    #endregion

    #region Y

    public int Y
    {
      get
      {
        return Convert.ToInt32( m_point.Y );
      }
      set
      {
        m_point.Y = Convert.ToInt32( value );
      }
    }

    #endregion

    #endregion

    #region Public Methods

    public override bool Equals( object obj )
    {
      if( !( obj is Point ) )
        return false;

      var other = (Point)obj;

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

    public static bool operator ==( Point c1, Point c2 )
    {
      return c1.Equals( c2 );
    }

    public static bool operator !=( Point c1, Point c2 )
    {
      return !c1.Equals( c2 );
    }

    #endregion
  }
}
