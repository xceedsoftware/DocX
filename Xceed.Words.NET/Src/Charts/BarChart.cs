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
  /// This element contains the 2-D bar or column series on this chart.
  /// 21.2.2.16 barChart (Bar Charts)
  /// </summary>
  public class BarChart : Chart
  {
    #region Public Properties

    /// <summary>
    /// Specifies the possible directions for a bar chart.
    /// </summary>
    public BarDirection BarDirection
    {
      get
      {
        return XElementHelpers.GetValueToEnum<BarDirection>(
            ChartXml.Element( XName.Get( "barDir", DocX.c.NamespaceName ) ) );
      }
      set
      {
        XElementHelpers.SetValueFromEnum<BarDirection>(
            ChartXml.Element( XName.Get( "barDir", DocX.c.NamespaceName ) ), value );
      }
    }

    /// <summary>
    /// Specifies the possible groupings for a bar chart.
    /// </summary>
    public BarGrouping BarGrouping
    {
      get
      {
        return XElementHelpers.GetValueToEnum<BarGrouping>(
            ChartXml.Element( XName.Get( "grouping", DocX.c.NamespaceName ) ) );
      }
      set
      {
        XElementHelpers.SetValueFromEnum<BarGrouping>(
            ChartXml.Element( XName.Get( "grouping", DocX.c.NamespaceName ) ), value );
      }
    }

    /// <summary>
    /// Specifies that its contents contain a percentage between 0% and 500%.
    /// </summary>
    public Int32 GapWidth
    {
      get
      {
        return Convert.ToInt32(
            ChartXml.Element( XName.Get( "gapWidth", DocX.c.NamespaceName ) ).Attribute( XName.Get( "val" ) ).Value );
      }
      set
      {
        if( ( value < 1 ) || ( value > 500 ) )
          throw new ArgumentException( "GapWidth lay between 0% and 500%!" );
        ChartXml.Element( XName.Get( "gapWidth", DocX.c.NamespaceName ) ).Attribute( XName.Get( "val" ) ).Value = value.ToString();
      }
    }

    #endregion

    #region Overrides

    protected override XElement CreateChartXml()
    {
      return XElement.Parse(
          @"<c:barChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                    <c:barDir val=""col""/>
                    <c:grouping val=""clustered""/>                    
                    <c:gapWidth val=""150""/>
                  </c:barChart>" );
    }

    #endregion

  }

  /// <summary>
  /// Specifies the possible directions for a bar chart.
  /// 21.2.3.3 ST_BarDir (Bar Direction)
  /// </summary>
  public enum BarDirection
  {
    [XmlName( "col" )]
    Column,
    [XmlName( "bar" )]
    Bar
  }

  /// <summary>
  /// Specifies the possible groupings for a bar chart.
  /// 21.2.3.4 ST_BarGrouping (Bar Grouping)
  /// </summary>
  public enum BarGrouping
  {
    [XmlName( "clustered" )]
    Clustered,
    [XmlName( "percentStacked" )]
    PercentStacked,
    [XmlName( "stacked" )]
    Stacked,
    [XmlName( "standard" )]
    Standard
  }
}
