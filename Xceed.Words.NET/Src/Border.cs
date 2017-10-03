/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System.Drawing;

namespace Xceed.Words.NET
{
  /// <summary>
  /// Represents a border of a table or table cell
  /// </summary>
  public class Border
  {

    #region Public Properties

    public BorderStyle Tcbs { get; set; }
    public BorderSize Size { get; set; }
    public int Space { get; set; }
    public Color Color { get; set; }

    #endregion

    #region Constructors

    public Border()
    {
      this.Tcbs = BorderStyle.Tcbs_single;
      this.Size = BorderSize.one;
      this.Space = 0;
      this.Color = Color.Black;
    }

    public Border( BorderStyle tcbs, BorderSize size, int space, Color color )
    {
      this.Tcbs = tcbs;
      this.Size = size;
      this.Space = space;
      this.Color = color;
    }

    #endregion
  }
}
