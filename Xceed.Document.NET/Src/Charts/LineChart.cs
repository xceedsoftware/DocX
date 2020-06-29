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
        return XElementHelpers.GetValueToEnum<Grouping>(
            ChartXml.Element( XName.Get( "grouping", Document.c.NamespaceName ) ) );
      }
      set
      {
        XElementHelpers.SetValueFromEnum<Grouping>(
            ChartXml.Element( XName.Get( "grouping", Document.c.NamespaceName ) ), value );
      }
    }

    #endregion

    #region Constructors

    public LineChart()
    {
    }


    #endregion

    #region Overrides

    public override void AddSeries( Series series )
    {
      // When the series has a Color set => LineChart series will color its line, not its content.
      var spPr = series.Xml.Element( XName.Get( "spPr", Document.c.NamespaceName ) );
      if( spPr != null )
      {
        var spPrContent = spPr.Elements().First();
        var newSpPr = new XElement( XName.Get( "spPr", Document.c.NamespaceName ), new XElement( XName.Get( "ln", Document.a.NamespaceName ), spPrContent ) );
        spPr.AddAfterSelf( newSpPr );
        spPr.Remove();
      }

      base.AddSeries( series );
    }

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
