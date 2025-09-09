/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System.ComponentModel;

namespace Xceed.Document.NET
{






  public enum ChartLegendPosition
  {
    [XmlName( "t" )]
    Top,
    [XmlName( "b" )]
    Bottom,
    [XmlName( "l" )]
    Left,
    [XmlName( "r" )]
    Right,
    [XmlName( "tr" )]
    TopRight
  }

  public enum DisplayBlanksAs
  {
    [XmlName( "gap" )]
    Gap,
    [XmlName( "span" )]
    Span,
    [XmlName( "zero" )]
    Zero
  }
}
