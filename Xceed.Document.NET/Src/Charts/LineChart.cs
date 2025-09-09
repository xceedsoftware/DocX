/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System.IO.Packaging;
using System.Xml.Linq;
using System.Linq;
using System.Globalization;
using System.Collections.Generic;
using System;

namespace Xceed.Document.NET
{
  public class LineChart : Chart
  {
    #region Private Properties

    private List<BaseSeries> _series;

    #endregion // Private Properties

    #region Public Properties

    public override List<BaseSeries> Series
    {
      get
      {
        if( _series == null )
        {
          _series = new List<BaseSeries>();
          var chart = this.GetChartTypeXElement();
          var ser = chart.Elements( XName.Get( "ser", Document.c.NamespaceName ) );
          foreach( var element in ser )
          {
            var serie = new LineSeries( element );
            serie.PackagePart = this.PackagePart;
            _series.Add( serie );
          }
        }
        return _series;
      }
    }

    public Grouping Grouping
    {
      get
      {
        var chartXml = GetChartTypeXElement();

        return XElementHelpers.GetValueToEnum<Grouping>(
            chartXml.Element( XName.Get( "grouping", Document.c.NamespaceName ) ) );
      }
      set
      {
        var chartXml = GetChartTypeXElement();

        XElementHelpers.SetValueFromEnum<Grouping>(
            chartXml.Element( XName.Get( "grouping", Document.c.NamespaceName ) ), value );
      }
    }

    #endregion

    #region Protected Properties

    protected override Type AllowedSeriesType
    {
      get
      {
        return typeof( LineSeries );
      }
    }

    #endregion // Protected Properties

    #region Constructors
    [Obsolete( "LineChart() is obsolete. Use Document.AddChart<LineChart>() instead." )]
    public LineChart()
    {
    }

    #endregion

    #region Overrides

    public override void AddSeries( BaseSeries series )
    {
      // When the series has a Color set => LineChart series will color its line, not its content.
      var spPr = series.Xml.Element( XName.Get( "spPr", Document.c.NamespaceName ) );
      if( spPr != null )
      {
        if( spPr.Element( XName.Get( "ln", Document.a.NamespaceName ) ) == null )
        {
          var spPrContent = spPr.Elements().First(); // Only color tag is defined.

          var newSpPr = new XElement( XName.Get( "spPr", Document.c.NamespaceName ),
                              new XElement( XName.Get( "ln", Document.a.NamespaceName ), spPrContent ) );
          spPr.AddAfterSelf( newSpPr );
          spPr.Remove();
        }
      }

      series.PackagePart = this.PackagePart;
      base.AddSeries( series );
    }

    protected override XElement CreateExternalChartXml()
    {
      return XElement.Parse(
          @"<c:lineChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                    <c:grouping val=""standard""/>                    
                  </c:lineChart>" );
    }

    protected override XElement GetChartTypeXElement()
    {
      if( this.ExternalXml == null )
        return null;

      return this.ExternalXml.Descendants().Where( chartElement => ( chartElement.Name.LocalName == "lineChart" )
                                                                    || ( chartElement.Name.LocalName == "line3DChart" ) ).SingleOrDefault();

    }
    #endregion


  }
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
