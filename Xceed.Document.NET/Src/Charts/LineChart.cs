/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2020 Xceed Software Inc.
 
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
        var chartXml = GetChartTypeXElement();

        return XElementHelpers.GetValueToEnum<Grouping>(
            chartXml.Element(XName.Get("grouping", Document.c.NamespaceName)));
      }
      set
      {
        var chartXml = GetChartTypeXElement();

        XElementHelpers.SetValueFromEnum<Grouping>(
            chartXml.Element(XName.Get("grouping", Document.c.NamespaceName)), value);
      }
    }

    #endregion

    #region Constructors
    [Obsolete("LineChart() is obsolete. Use Document.AddChart<LineChart>() instead.")]
    public LineChart()
    {
    }

    #endregion

    #region Overrides

    public override void AddSeries(Series series)
    {
      // When the series has a Color set => LineChart series will color its line, not its content.
      var spPr = series.Xml.Element(XName.Get("spPr", Document.c.NamespaceName));
      if (spPr != null)
      {
        if (spPr.Element(XName.Get("ln", Document.a.NamespaceName)) == null)
        {
          var spPrContent = spPr.Elements().First(); // Only color tag is defined.

          var newSpPr = new XElement(XName.Get("spPr", Document.c.NamespaceName),
                              new XElement(XName.Get("ln", Document.a.NamespaceName), spPrContent));
          spPr.AddAfterSelf(newSpPr);
          spPr.Remove();
        }
      }

      series.PackagePart = this.PackagePart;
      base.AddSeries(series);
    }

    protected override XElement CreateExternalChartXml()
    {
      return XElement.Parse(
          @"<c:lineChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                    <c:grouping val=""standard""/>                    
                  </c:lineChart>");
    }

    protected override XElement GetChartTypeXElement()
    {
      if (this.ExternalXml == null)
        return null;

      return this.ExternalXml.Descendants().Where(chartElement => (chartElement.Name.LocalName == "lineChart")
                                                                    || (chartElement.Name.LocalName == "line3DChart")).SingleOrDefault();

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
