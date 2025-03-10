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
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.IO.Packaging;
using System.IO;
using Xceed.Drawing;

namespace Xceed.Document.NET
{
  public abstract class Chart
  {
    #region Private Members

    private Paragraph _parentParagraph;
    private PackageRelationship _packageRelationship;
    private List<Series> _series;

    private XDocument _chartDocument;

    #endregion

    #region Public Properties


    public XElement ExternalXml
    {
      get; private set;
    }

    public XElement Xml
    {
      get; private set;
    }

    public List<Series> Series
    {
      get
      {
        if(_series== null)
        {
          _series = new List<Series>();
          var chart = GetChartTypeXElement();
          var ser = chart.Elements(XName.Get("ser", Document.c.NamespaceName));
          foreach (var element in ser)
          {
            var serie = new Series(element);
            serie.PackagePart = this.PackagePart;
            _series.Add(serie);
          }
        }
        return _series;
      }
    }

    public virtual Int16 MaxSeriesCount
    {
      get
      {
        return Int16.MaxValue;
      }
    }

    public ChartLegend Legend
    {
      get; internal set;
    }

    public CategoryAxis CategoryAxis
    {
      get
      {
        var catAxXML = this.ExternalXml.Descendants( XName.Get( "catAx", Document.c.NamespaceName ) ).SingleOrDefault();

        return ( catAxXML != null ) ? new CategoryAxis( catAxXML ) : null;
      }
    }

    public ValueAxis ValueAxis
    {
      get
      {
        var valAxXML = this.ExternalXml.Descendants( XName.Get( "valAx", Document.c.NamespaceName ) ).SingleOrDefault();

        return ( valAxXML != null ) ? new ValueAxis( valAxXML ) : null;
      }
    }

    public virtual Boolean IsAxisExist
    {
      get
      {
        return true;
      }
    }

    public Boolean View3D
    {
      get
      {
        var chartXml = GetChartTypeXElement();
        return chartXml != null && chartXml.Name.LocalName.Contains( "3D" );
      }
      set
      {
        var chartXml = GetChartTypeXElement();
        if( chartXml != null )
        {
          if( value )
          {
            if( !View3D )
            {
              String currentName = chartXml.Name.LocalName;
              chartXml.Name = XName.Get( currentName.Replace( "Chart", "3DChart" ), Document.c.NamespaceName );
            }
          }
          else
          {
            if( View3D )
            {
              String currentName = chartXml.Name.LocalName;
              chartXml.Name = XName.Get( currentName.Replace( "3DChart", "Chart" ), Document.c.NamespaceName );
            }
          }
        }
      }
    }

    public DisplayBlanksAs DisplayBlanksAs
    {
      get
      {
        var chart = this.ExternalXml.Element( XName.Get( "chart" ) );
        return XElementHelpers.GetValueToEnum<DisplayBlanksAs>(
            chart.Element( XName.Get( "dispBlanksAs", Document.c.NamespaceName ) ) );
      }
      set
      {
        var chart = this.ExternalXml.Element( XName.Get( "chart" ) );

        XElementHelpers.SetValueFromEnum<DisplayBlanksAs>(
            chart.Element( XName.Get( "dispBlanksAs", Document.c.NamespaceName ) ), value );
      }
    }

    #endregion

    #region Internal Properties
    internal PackagePart PackagePart
    {
      get; set;
    }

    internal PackageRelationship RelationPackage
    {
      get; set;
    }

    #endregion



































    #region Constructors

    public Chart()
    {

      // Create global xml
      this.ExternalXml = XElement.Parse
          ( @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                   <c:chartSpace xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">  
                       <c:roundedCorners val=""0""/>
                       <c:chart>
                           <c:autoTitleDeleted val=""0""/>
                           <c:plotVisOnly val=""1""/>
                           <c:dispBlanksAs val=""gap""/>
                           <c:showDLblsOverMax val=""0""/>
                       </c:chart>
                   </c:chartSpace>" );


      // Create internal chart Xml
      this.Xml = XElement.Parse
        ( @"<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
            <w:drawing xmlns=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
              <wp:inline xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"">
                <wp:extent cx=""1270000"" cy=""1270000""/>
                <wp:effectExtent l=""0"" t=""0"" r=""19050"" b=""19050""/>
                <wp:docPr id=""1"" name=""chart""/>
                <a:graphic xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
                    <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                      <c:chart p6:id=""rIdX"" xmlns:p6=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""/>
                    </a:graphicData>
                </a:graphic>
            </wp:inline>
          </w:drawing>
        </w:r>" );

      // Create a real chart xml in an inheritor
      var chartXml = this.CreateExternalChartXml();

      // Create result plotarea element
      var plotAreaXml = new XElement( XName.Get( "plotArea", Document.c.NamespaceName ),
                                      new XElement( XName.Get( "layout", Document.c.NamespaceName ) ),
                                      chartXml );

      // Set labels 
      var dLblsXml = XElement.Parse(
          @"<c:dLbls xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                    <c:showLegendKey val=""0""/>
                    <c:showVal val=""0""/>
                    <c:showCatName val=""0""/>
                    <c:showSerName val=""0""/>
                    <c:showPercent val=""0""/>
                    <c:showBubbleSize val=""0""/>
                    <c:showLeaderLines val=""1""/>
                </c:dLbls>" );
      chartXml.Add( dLblsXml );

      // if axes exists, create their
      if( this.IsAxisExist )
      {
        var categoryAxis = new CategoryAxis( "148921728" );
        var valueAxis = new ValueAxis( "154227840" );

        var axIDcatXml = XElement.Parse( String.Format( @"<c:axId val=""{0}"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""/>", categoryAxis.Id ) );
        var axIDvalXml = XElement.Parse( String.Format( @"<c:axId val=""{0}"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""/>", valueAxis.Id ) );

        var gapWidth = chartXml.Element( XName.Get( "gapWidth", Document.c.NamespaceName ) );
        if( gapWidth != null )
        {
          gapWidth.AddAfterSelf( axIDvalXml );
          gapWidth.AddAfterSelf( axIDcatXml );
        }
        else
        {
          chartXml.Add( axIDcatXml );
          chartXml.Add( axIDvalXml );
        }

        plotAreaXml.Add( categoryAxis.Xml );
        plotAreaXml.Add( valueAxis.Xml );
      }

      var chartRootXml = this.ExternalXml.Element( XName.Get( "chart", Document.c.NamespaceName ) );
      chartRootXml.Element( XName.Get( "autoTitleDeleted", Document.c.NamespaceName ) ).AddAfterSelf( plotAreaXml );


      }
    internal Chart( Paragraph parentParagraph, PackageRelationship packageRelationship, PackagePart packagePart, XDocument chartDocument )
      : this()
    {
      _parentParagraph = parentParagraph;
      _packageRelationship = packageRelationship;
      this.PackagePart = packagePart;
      _chartDocument = chartDocument;


    }


    #endregion

    #region Public Methods

    public virtual void AddSeries( Series series )
    {
      Series.Add(series);
      series.PackagePart = this.PackagePart;

      var chart = GetChartTypeXElement();
      if( chart != null )
      {
        var seriesCount = chart.Elements( XName.Get( "ser", Document.c.NamespaceName ) ).Count();
        if( seriesCount >= this.MaxSeriesCount )
          throw new InvalidOperationException( "Maximum series for this chart is" + this.MaxSeriesCount.ToString() + "and have exceeded!" );

        //To work in Words, all series need an Index and Order.
        var value = seriesCount + 1;
        var content = new XAttribute( XName.Get( "val" ), value.ToString() );
        series.Xml.AddFirst( new XElement( XName.Get( "order", Document.c.NamespaceName ), content ) );
        series.Xml.AddFirst( new XElement( XName.Get( "idx", Document.c.NamespaceName ), content ) );
        chart.Add( series.Xml );
      }
    }

    public void AddLegend()
    {
      AddLegend( ChartLegendPosition.Right, false );
    }

    public void AddLegend( ChartLegendPosition position, Boolean overlay )
    {
      if( this.Legend != null )
      {
        this.RemoveLegend();
      }
      this.Legend = new ChartLegend( position, overlay );
      var chart = this.ExternalXml.Element( XName.Get( "chart", Document.c.NamespaceName ) );
      if( chart != null )
      {
        chart.Element( XName.Get( "plotArea", Document.c.NamespaceName ) ).AddAfterSelf( Legend.Xml );
      }
    }

    public void RemoveLegend()
    {
      if( this.Legend != null )
      {
        this.Legend.Xml.Remove();
        this.Legend = null;
      }
    }

    public void Remove()
    {
      if( _packageRelationship.Package != null )
      {
        var uriString = _packageRelationship.TargetUri.OriginalString;
        if( !uriString.StartsWith( "/" ) )
        {
          uriString = "/" + uriString;
        }
        if( !uriString.StartsWith( "/word/" ) )
        {
          uriString = "/word" + uriString;
        }

        var uri = new Uri( uriString, UriKind.Relative );

        _packageRelationship.Package.DeletePart( uri );
      }

      if( _parentParagraph.Document.PackagePart != null )
      {
        _parentParagraph.Document.PackagePart.DeleteRelationship( _packageRelationship.Id );
      }

      // Remove the Xml from document.
      var parentParagraphChart = _parentParagraph.Xml.Descendants( XName.Get( "chart", Document.c.NamespaceName ) )
                                                   .FirstOrDefault( c => c.GetAttribute( XName.Get( "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships" ) ) == _packageRelationship.Id );
      if( parentParagraphChart != null )
      {
        var parentDrawing = parentParagraphChart.Ancestors( XName.Get( "drawing", Document.w.NamespaceName ) ).FirstOrDefault();
        if( parentDrawing != null )
        {
          parentDrawing.Remove();
        }
      }
    }



    #endregion

    #region Protected Methods

    protected abstract XElement CreateExternalChartXml();

    protected abstract XElement GetChartTypeXElement();

    #endregion

    #region Internal Method

    static internal IEnumerable<XElement> GetChartsXml( XElement xml )
    {
      if( xml == null )
        return null;

      return xml.Elements().Where( chartElement => ( chartElement.Name.LocalName == "barChart" )
                                                        || ( chartElement.Name.LocalName == "bar3DChart" )
                                                        || ( chartElement.Name.LocalName == "lineChart" )
                                                        || ( chartElement.Name.LocalName == "line3DChart" )
                                                        || ( chartElement.Name.LocalName == "pieChart" )
                                                        || ( chartElement.Name.LocalName == "pie3DChart" ) );

    }

    internal void SetXml( XElement externalChartXml, XElement internalChartXml )
    {
      this.ExternalXml = externalChartXml;
      this.Xml = internalChartXml;
    }

    internal void SetInternalChartSettings( string relId, float chartWidth, float chartHeight )
    {
      if( string.IsNullOrEmpty( relId ) )
        throw new ArgumentNullException( "relId" );

      var width = chartWidth * Picture.EmusInPixel;
      var height = chartHeight * Picture.EmusInPixel;

      var extent = this.Xml.Descendants( XName.Get( "extent", Document.wp.NamespaceName ) ).FirstOrDefault();
      if( extent != null )
      {
        extent.Attribute( "cx" ).Value = width.ToString();
        extent.Attribute( "cy" ).Value = height.ToString();
      }

      var chart = this.Xml.Descendants( XName.Get( "chart", Document.c.NamespaceName ) ).SingleOrDefault();
      if( chart != null )
      {
        var idAttribute = chart.Attribute( XName.Get( "id", Document.r.NamespaceName ) );
        if( idAttribute != null )
        {
          idAttribute.Value = relId;
        }
      }
    }


#endregion
  }


  public class ChartLegend
  {
    #region Public Properties

    public Boolean Overlay
    {
      get
      {
        return Xml.Element( XName.Get( "overlay", Document.c.NamespaceName ) ).Attribute( "val" ).Value == "1";
      }
      set
      {
        Xml.Element( XName.Get( "overlay", Document.c.NamespaceName ) ).Attribute( "val" ).Value = GetOverlayValue( value );
      }
    }

    public ChartLegendPosition Position
    {
      get
      {
        return XElementHelpers.GetValueToEnum<ChartLegendPosition>(
            Xml.Element( XName.Get( "legendPos", Document.c.NamespaceName ) ) );
      }
      set
      {
        XElementHelpers.SetValueFromEnum<ChartLegendPosition>(
            Xml.Element( XName.Get( "legendPos", Document.c.NamespaceName ) ), value );
      }
    }

    #endregion

    #region Internal Properties

    internal XElement Xml
    {
      get; private set;
    }

    #endregion

    #region Constructors

    internal ChartLegend( ChartLegendPosition position, Boolean overlay )
    {
      Xml = new XElement(
          XName.Get( "legend", Document.c.NamespaceName ),
          new XElement( XName.Get( "legendPos", Document.c.NamespaceName ), new XAttribute( "val", XElementHelpers.GetXmlNameFromEnum<ChartLegendPosition>( position ) ) ),
          new XElement( XName.Get( "overlay", Document.c.NamespaceName ), new XAttribute( "val", GetOverlayValue( overlay ) ) )
          );
    }


    #endregion

    #region Internal Methods


    #endregion

    #region Private Methods

    private String GetOverlayValue( Boolean overlay )
    {
      if( overlay )
        return "1";
      else
        return "0";
    }

    #endregion
  }

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
