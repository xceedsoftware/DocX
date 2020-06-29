/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2020 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System.Drawing;

namespace Xceed.Document.NET
{
  /// <summary>
  /// Represents a border of a table or table cell
  /// </summary>
  public class Border
  {

    #region Public Properties

    public BorderStyle Tcbs { get; set; }
    public BorderSize Size { get; set; }
    public float Space { get; set; }
    public Color Color { get; set; }

    #endregion

    #region Constructors

    public Border()
    {
      this.Tcbs = BorderStyle.Tcbs_single;
      this.Size = BorderSize.one;
      this.Space = 0f;
      this.Color = Color.Black;
    }

    public Border( BorderStyle tcbs, BorderSize size, float space, Color color )
    {
      this.Tcbs = tcbs;
      this.Size = size;
      this.Space = space;
      this.Color = color;
    }

    #endregion

    internal static string GetNumericSize( BorderSize borderSize )
    {
      var size = "2";
      switch( borderSize )
      {
        case BorderSize.two:
          size = "4";
        break;
        case BorderSize.three:
          size = "6";
        break;
        case BorderSize.four:
          size = "8";
        break;
        case BorderSize.five:
          size = "12";
        break;
        case BorderSize.six:
          size = "18";
        break;
        case BorderSize.seven:
          size = "24";
        break;
        case BorderSize.eight:
          size = "36";
        break;
        case BorderSize.nine:
          size = "48";
        break;
      case BorderSize.one:
        default:
          size = "2";
        break;
      }

      return size;
    }
  }
}
