/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System.Xml.Linq;

namespace Xceed.Words.NET
{
  /// <summary>
  /// This element contains the 2-D line chart series.
  /// 21.2.2.97 lineChart (Line Charts)
  /// </summary>
  public class LineChart : Chart
  {
    #region Public Properties

    /// <summary>
    /// Specifies the kind of grouping for a column, line, or area chart.
    /// </summary>
    public Grouping Grouping
    {
      get
      {
        return XElementHelpers.GetValueToEnum<Grouping>(
            ChartXml.Element( XName.Get( "grouping", DocX.c.NamespaceName ) ) );
      }
      set
      {
        XElementHelpers.SetValueFromEnum<Grouping>(
            ChartXml.Element( XName.Get( "grouping", DocX.c.NamespaceName ) ), value );
      }
    }

    #endregion

    #region Overrides

    protected override XElement CreateChartXml()
    {
      return XElement.Parse(
          @"<c:lineChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                    <c:grouping val=""standard""/>                    
                  </c:lineChart>" );
    }

    #endregion
  }

  /// <summary>
  /// Specifies the kind of grouping for a column, line, or area chart.
  /// 21.2.2.76 grouping (Grouping)
  /// </summary>
  public enum Grouping
  {
    [XmlName( "percentStacked" )]
    PercentStacked,
    [XmlName( "stacked" )]
    Stacked,
    [XmlName( "standard" )]
    Standard
  }
}
