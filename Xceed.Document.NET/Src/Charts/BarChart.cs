/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System;
using System.Globalization;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace Xceed.Document.NET
{
  public class BarChart : Chart
  {
    #region Public Properties

    public BarDirection BarDirection
    {
      get
      {
        var chartXml = GetChartTypeXElement();

        return XElementHelpers.GetValueToEnum<BarDirection>(
            chartXml.Element( XName.Get( "barDir", Document.c.NamespaceName ) ) );
      }
      set
      {
        var chartXml = GetChartTypeXElement();

        XElementHelpers.SetValueFromEnum<BarDirection>(
            chartXml.Element( XName.Get( "barDir", Document.c.NamespaceName ) ), value );
      }
    }

    public BarGrouping BarGrouping
    {
      get
      {
        var chartXml = GetChartTypeXElement();

        return XElementHelpers.GetValueToEnum<BarGrouping>(
            chartXml.Element( XName.Get( "grouping", Document.c.NamespaceName ) ) );
      }
      set
      {
        var chartXml = GetChartTypeXElement();

        XElementHelpers.SetValueFromEnum<BarGrouping>(
            chartXml.Element( XName.Get( "grouping", Document.c.NamespaceName ) ), value );

        var overlapVal = ( ( value == BarGrouping.Stacked ) || ( value == BarGrouping.PercentStacked ) ) ? "100" : "0";
        var overlap = chartXml.Element( XName.Get( "overlap", Document.c.NamespaceName ) );
        if( overlap != null )
        {
          overlap.Attribute( XName.Get( "val" ) ).Value = overlapVal;
        }
      }
    }

    public Int32 GapWidth
    {
      get
      {
        var chartXml = GetChartTypeXElement();

        return Convert.ToInt32(
            chartXml.Element( XName.Get( "gapWidth", Document.c.NamespaceName ) ).Attribute( XName.Get( "val" ) ).Value );
      }
      set
      {
        var chartXml = GetChartTypeXElement();

        if( ( value < 1 ) || ( value > 500 ) )
          throw new ArgumentException( "GapWidth lay between 0% and 500%!" );
        chartXml.Element( XName.Get( "gapWidth", Document.c.NamespaceName ) ).Attribute( XName.Get( "val" ) ).Value = value.ToString( CultureInfo.InvariantCulture );
      }
    }

    #endregion

    #region Constructors
    [Obsolete("BarChart() is obsolete. Use Document.AddChart<BarChart>() instead.")]
    public BarChart()
    {
    }


    #endregion

    #region Overrides

    protected override XElement CreateExternalChartXml()
    {
      return XElement.Parse(
          @"<c:barChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                    <c:barDir val=""col""/>
                    <c:grouping val=""clustered""/>                    
                    <c:gapWidth val=""150""/>
                    <c:overlap val=""0""/>
                  </c:barChart>" );
    }

    protected override XElement GetChartTypeXElement()
    {
      if( this.ExternalXml == null )
        return null;

      return this.ExternalXml.Descendants().Where( chartElement => ( chartElement.Name.LocalName == "barChart" )
                                                                     || ( chartElement.Name.LocalName == "bar3DChart" ) ).SingleOrDefault();

    }

    #endregion
  }

  public enum BarDirection
  {
    [XmlName( "col" )]
    Column,
    [XmlName( "bar" )]
    Bar
  }

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
