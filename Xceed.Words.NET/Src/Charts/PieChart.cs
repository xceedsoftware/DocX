/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System;
using System.Xml.Linq;

namespace Xceed.Words.NET
{
  /// <summary>
  /// This element contains the 2-D pie series for this chart.
  /// 21.2.2.141 pieChart (Pie Charts)
  /// </summary>
  public class PieChart : Chart
  {
    #region Overrides Properties

    public override Boolean IsAxisExist
    {
      get
      {
        return false;
      }
    }
    public override Int16 MaxSeriesCount
    {
      get
      {
        return 1;
      }
    }

    #endregion

    #region Overrides

    protected override XElement CreateChartXml()
    {
      return XElement.Parse(
          @"<c:pieChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                  </c:pieChart>" );
    }

    #endregion
  }
}
