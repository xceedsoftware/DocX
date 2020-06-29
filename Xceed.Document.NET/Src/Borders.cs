/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2020 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


namespace Xceed.Document.NET
{
  public class Borders
  {
    #region Members

    private Border _left;
    private Border _top;
    private Border _right;
    private Border _bottom;

    #endregion

    #region Constructors

    public Borders()
    {
    }

    public Borders( Border border )
    {
      _left = border;
      _top = border;
      _right = border;
      _bottom = border;
    }

    public Borders( Border leftBorder, Border topBorder, Border rightBorder, Border bottomBorder )
    {
      _left = leftBorder;
      _top = topBorder;
      _right = rightBorder;
      _bottom = bottomBorder;
    }

    #endregion

    #region Public Properties

    public Border Left
    {
      get
      {
        return _left;
      }
      set
      {
        _left = value;
      }
    }

    public Border Top
    {
      get
      {
        return _top;
      }
      set
      {
        _top = value;
      }
    }

    public Border Right
    {
      get
      {
        return _right;
      }
      set
      {
        _right = value;
      }
    }

    public Border Bottom
    {
      get
      {
        return _bottom;
      }
      set
      {
        _bottom = value;
      }
    }

    #endregion
  }
}
