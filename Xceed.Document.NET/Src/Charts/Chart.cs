/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2020 Xceed Software Inc.
 
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
using System.Collections;
using System.Drawing;
using System.Globalization;
using System.IO.Packaging;
using System.IO;

namespace Xceed.Document.NET
{
  /// <summary>
  /// Represents every Chart in this document.
  /// </summary>
  public abstract class Chart
  {
    #region Private Members


#endregion

    #region Public Properties

    /// <summary>
    /// The xml representation of this chart
    /// </summary>
    public XDocument Xml
    {
      get; private set;
    }

    /// <summary>
    /// Chart's series
    /// </summary>
    public List<Series> Series
    {
      get
      {
        var series = new List<Series>();
        var ser = XName.Get( "ser", Document.c.NamespaceName );
        int index = 1;
        foreach( var element in ChartXml.Elements( ser ) )
        {
          element.Add( new XElement( XName.Get("idx", Document.c.NamespaceName ) ), index.ToString() );
          series.Add( new Series( element ) );
          ++index;
        }
        return series;
      }
    }

    /// <summary>
    /// Return maximum count of series
    /// </summary>
    public virtual Int16 MaxSeriesCount
    {
      get
      {
        return Int16.MaxValue;
      }
    }

    /// <summary>
    /// Chart's legend.
    /// If legend doesn't exist property is null.
    /// </summary>
    public ChartLegend Legend
    {
      get; internal set;
    }

    /// <summary>
    /// Represents the category axis
    /// </summary>
    public CategoryAxis CategoryAxis
    {
      get; private set;
    }

    /// <summary>
    /// Represents the values axis
    /// </summary>
    public ValueAxis ValueAxis
    {
      get; private set;
    }

    /// <summary>
    /// Represents existing the axis
    /// </summary>
    public virtual Boolean IsAxisExist
    {
      get
      {
        return true;
      }
    }

    /// <summary>
    /// Get or set 3D view for this chart
    /// </summary>
    public Boolean View3D
    {
      get
      {
        return ChartXml.Name.LocalName.Contains( "3D" );
      }
      set
      {
        if( value )
        {
          if( !View3D )
          {
            String currentName = ChartXml.Name.LocalName;
            ChartXml.Name = XName.Get( currentName.Replace( "Chart", "3DChart" ), Document.c.NamespaceName );
          }
        }
        else
        {
          if( View3D )
          {
            String currentName = ChartXml.Name.LocalName;
            ChartXml.Name = XName.Get( currentName.Replace( "3DChart", "Chart" ), Document.c.NamespaceName );
          }
        }
      }
    }

    /// <summary>
    /// Specifies how blank cells shall be plotted on a chart
    /// </summary>
    public DisplayBlanksAs DisplayBlanksAs
    {
      get
      {
        return XElementHelpers.GetValueToEnum<DisplayBlanksAs>(
            ChartRootXml.Element( XName.Get( "dispBlanksAs", Document.c.NamespaceName ) ) );
      }
      set
      {
        XElementHelpers.SetValueFromEnum<DisplayBlanksAs>(
            ChartRootXml.Element( XName.Get( "dispBlanksAs", Document.c.NamespaceName ) ), value );
      }
    }

    #endregion

    #region Protected Properties

    protected internal XElement ChartXml
    {
      get; private set;
    }
    protected internal XElement ChartRootXml
    {
      get; private set;
    }

    #endregion

    #region Constructors

    /// <summary>
    /// Create an Chart for this document
    /// </summary>        
    public Chart()
    {

      // Create global xml
      this.Xml = XDocument.Parse
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

      // Create a real chart xml in an inheritor
      this.ChartXml = this.CreateChartXml();

      // Create result plotarea element
      var plotAreaXml = new XElement( XName.Get( "plotArea", Document.c.NamespaceName ),
                                      new XElement( XName.Get( "layout", Document.c.NamespaceName ) ),
                                      this.ChartXml );

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
      this.ChartXml.Add( dLblsXml );

      // if axes exists, create their
      if( this.IsAxisExist )
      {
        this.CategoryAxis = new CategoryAxis( "148921728" );
        this.ValueAxis = new ValueAxis( "154227840" );

        var axIDcatXml = XElement.Parse( String.Format( @"<c:axId val=""{0}"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""/>", this.CategoryAxis.Id ) );
        var axIDvalXml = XElement.Parse( String.Format( @"<c:axId val=""{0}"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""/>", this.ValueAxis.Id ) );

        var gapWidth = this.ChartXml.Element( XName.Get( "gapWidth", Document.c.NamespaceName ) );
        if( gapWidth != null )
        {
          gapWidth.AddAfterSelf( axIDvalXml );
          gapWidth.AddAfterSelf( axIDcatXml );
        }
        else
        {
          this.ChartXml.Add( axIDcatXml );
          this.ChartXml.Add( axIDvalXml );
        }

        plotAreaXml.Add( this.CategoryAxis.Xml );
        plotAreaXml.Add( this.ValueAxis.Xml );
      }

      this.ChartRootXml = this.Xml.Root.Element( XName.Get( "chart", Document.c.NamespaceName ) );
      this.ChartRootXml.Element( XName.Get( "autoTitleDeleted", Document.c.NamespaceName ) ).AddAfterSelf( plotAreaXml );
    }


#endregion

    #region Public Methods

    /// <summary>
    /// Add a new series to this chart
    /// </summary>
    public virtual void AddSeries( Series series )
    {
      var seriesCount = this.ChartXml.Elements( XName.Get( "ser", Document.c.NamespaceName ) ).Count();
      if( seriesCount >= this.MaxSeriesCount )
        throw new InvalidOperationException( "Maximum series for this chart is" + this.MaxSeriesCount.ToString() + "and have exceeded!" );

      //To work in Words, all series need an Index and Order.
      var value = seriesCount + 1;
      var content = new XAttribute( XName.Get( "val" ), value.ToString() );
      series.Xml.AddFirst( new XElement( XName.Get( "order", Document.c.NamespaceName ), content ) );
      series.Xml.AddFirst( new XElement( XName.Get( "idx", Document.c.NamespaceName ), content ) );
      this.ChartXml.Add( series.Xml );
    }

    /// <summary>
    /// Add standart legend to the chart.
    /// </summary>
    public void AddLegend()
    {
      AddLegend( ChartLegendPosition.Right, false );
    }

    /// <summary>
    /// Add a legend with parameters to the chart.
    /// </summary>
    public void AddLegend( ChartLegendPosition position, Boolean overlay )
    {
      if( this.Legend != null )
      {
        this.RemoveLegend();
      }
      this.Legend = new ChartLegend( position, overlay );
      this.ChartRootXml.Element( XName.Get( "plotArea", Document.c.NamespaceName ) ).AddAfterSelf( Legend.Xml );
    }

    /// <summary>
    /// Remove the legend from the chart.
    /// </summary>
    public void RemoveLegend()
    {
      if( this.Legend != null )
      {
        this.Legend.Xml.Remove();
        this.Legend = null;
      }
    }


#endregion

    #region Protected Methods

    /// <summary>
    /// An abstract method which creates the current chart xml
    /// </summary>
    protected abstract XElement CreateChartXml();

    #endregion

    #region Internal Method





#endregion
  }

  /// <summary>
  /// Represents a chart series
  /// </summary>
  public class Series
  {
    #region Private Members

    private XElement _strCache;
    private XElement _numCache;

    #endregion

    #region Public Properties

    public Color Color
    {
      get
      {
        var spPr = this.Xml.Element( XName.Get( "spPr", Document.c.NamespaceName ) );
        if( spPr == null )
          return Color.Transparent;

        var srgbClr = spPr.Descendants( XName.Get( "srgbClr", Document.a.NamespaceName ) ).FirstOrDefault();
        if( srgbClr != null )
        {
          var val = srgbClr.Attribute( XName.Get( "val" ) );
          if( val != null )
          {
            var rgb = Color.FromArgb( Int32.Parse( val.Value, NumberStyles.HexNumber ) );
            return Color.FromArgb( 255, rgb );
          }
        }

        return Color.Transparent;
      }
      set
      {
        var colorElement = this.Xml.Element( XName.Get( "spPr", Document.c.NamespaceName ) );
        if( colorElement != null )
        {
          colorElement.Remove();
        }

        var colorData = new XElement( XName.Get( "solidFill", Document.a.NamespaceName ),
                                   new XElement( XName.Get( "srgbClr", Document.a.NamespaceName ), new XAttribute( XName.Get( "val" ), value.ToHex() ) ) );

        // When the chart containing this series is a lineChart, the line will be colored, else the shape will be colored.
        colorElement = ( ( this.Xml.Parent != null ) && ( this.Xml.Parent.Name != null ) && (this.Xml.Parent.Name.LocalName == "lineChart" ) )
                       ? new XElement( XName.Get( "spPr", Document.c.NamespaceName ),
                                  new XElement( XName.Get( "ln", Document.a.NamespaceName ), colorData ) )
                       : new XElement( XName.Get( "spPr", Document.c.NamespaceName ), colorData );
        this.Xml.Element( XName.Get( "tx", Document.c.NamespaceName ) ).AddAfterSelf( colorElement );
      }
    }














#endregion

    #region Internal Properties

    /// <summary>
    /// Series xml element
    /// </summary>
    internal XElement Xml
    {
      get; private set;
    }

    #endregion

    #region Constructors

    internal Series( XElement xml )
    {
      this.Xml = xml;

      var cat = xml.Element( XName.Get( "cat", Document.c.NamespaceName ) );
      if( cat != null )
      {
        _strCache = cat.Descendants( XName.Get( "strCache", Document.c.NamespaceName ) ).FirstOrDefault();
        if( _strCache == null )
        {
          _strCache = cat.Descendants( XName.Get( "strLit", Document.c.NamespaceName ) ).FirstOrDefault();
        }
      }

      var val = xml.Element( XName.Get( "val", Document.c.NamespaceName ) );
      if( val != null )
      {      
        _numCache = val.Descendants( XName.Get( "numCache", Document.c.NamespaceName ) ).FirstOrDefault();
        if( _numCache == null )
        {
          _numCache = val.Descendants( XName.Get( "numLit", Document.c.NamespaceName ) ).FirstOrDefault();
        }
      }
    }

    public Series( String name )
    {
      _strCache = new XElement( XName.Get( "strCache", Document.c.NamespaceName ) );
      _numCache = new XElement( XName.Get( "numCache", Document.c.NamespaceName ) );

      this.Xml = new XElement( XName.Get( "ser", Document.c.NamespaceName ),
                               new XElement( XName.Get( "tx", Document.c.NamespaceName ),
                                             new XElement( XName.Get( "strRef", Document.c.NamespaceName ), 
                                                           new XElement( XName.Get( "f", Document.c.NamespaceName ), "" ),
                                                           new XElement( XName.Get( "strCache", Document.c.NamespaceName ),
                                                                         new XElement( XName.Get( "pt", Document.c.NamespaceName ), 
                                                                                       new XAttribute( XName.Get( "idx" ), "0" ), 
                                                                                       new XElement( XName.Get( "v", Document.c.NamespaceName ), name ) ) ) ) ),
                               new XElement( XName.Get( "invertIfNegative", Document.c.NamespaceName ), "0" ),
                               new XElement( XName.Get( "cat", Document.c.NamespaceName ), 
                                             new XElement( XName.Get( "strRef", Document.c.NamespaceName ),
                                                           new XElement( XName.Get( "f", Document.c.NamespaceName ), "" ),
                                                           _strCache ) ),
                               new XElement( XName.Get( "val", Document.c.NamespaceName ), 
                                             new XElement( XName.Get( "numRef", Document.c.NamespaceName ),
                                                           new XElement( XName.Get( "f", Document.c.NamespaceName ), "" ),
                                                           _numCache ) )
          );
    }

    #endregion

    #region Public Methods

    public void Bind( ICollection list, String categoryPropertyName, String valuePropertyName )
    {
      var ptCount = new XElement( XName.Get( "ptCount", Document.c.NamespaceName ), new XAttribute( XName.Get( "val" ), list.Count ) );
      var formatCode = new XElement( XName.Get( "formatCode", Document.c.NamespaceName ), "General" );

      _strCache.RemoveAll();
      _numCache.RemoveAll();

      _strCache.Add( ptCount );
      _numCache.Add( formatCode );
      _numCache.Add( ptCount );      

      Int32 index = 0;
      XElement pt;
      foreach( var item in list )
      {
        pt = new XElement( XName.Get( "pt", Document.c.NamespaceName ), new XAttribute( XName.Get( "idx" ), index ),
                           new XElement( XName.Get( "v", Document.c.NamespaceName ), item.GetType().GetProperty( categoryPropertyName ).GetValue( item, null ) ) );
        _strCache.Add( pt );
        pt = new XElement( XName.Get( "pt", Document.c.NamespaceName ), new XAttribute( XName.Get( "idx" ), index ),
                           new XElement( XName.Get( "v", Document.c.NamespaceName ), item.GetType().GetProperty( valuePropertyName ).GetValue( item, null ) ) );
        _numCache.Add( pt );
        index++;
      }
    }

    public void Bind( IList categories, IList values )
    {
      if( categories.Count != values.Count )
        throw new ArgumentException( "Categories count must equal to Values count" );

      var ptCount = new XElement( XName.Get( "ptCount", Document.c.NamespaceName ), new XAttribute( XName.Get( "val" ), categories.Count ) );
      var formatCode = new XElement( XName.Get( "formatCode", Document.c.NamespaceName ), "General" );

      _strCache.RemoveAll();
      _numCache.RemoveAll();

      _strCache.Add( ptCount );
      _numCache.Add( formatCode );
      _numCache.Add( ptCount );

      XElement pt;
      for( int index = 0; index < categories.Count; index++ )
      {
        pt = new XElement( XName.Get( "pt", Document.c.NamespaceName ), new XAttribute( XName.Get( "idx" ), index ),
                           new XElement( XName.Get( "v", Document.c.NamespaceName ), categories[ index ].ToString() ) );
        _strCache.Add( pt );
        pt = new XElement( XName.Get( "pt", Document.c.NamespaceName ), new XAttribute( XName.Get( "idx" ), index ),
                           new XElement( XName.Get( "v", Document.c.NamespaceName ), values[ index ].ToString() ) );
        _numCache.Add( pt );
      }
    }

    #endregion
  }

  /// <summary>
  /// Represents a chart legend
  /// More: http://msdn.microsoft.com/ru-ru/library/cc845123.aspx
  /// </summary>
  public class ChartLegend
  {
    #region Public Properties

    /// <summary>
    /// Specifies that other chart elements shall be allowed to overlap this chart element
    /// </summary>
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

    /// <summary>
    /// Specifies the possible positions for a legend
    /// </summary>
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

    /// <summary>
    /// Legend xml element
    /// </summary>
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

    /// <summary>
    /// ECMA-376, page 3840
    /// 21.2.2.132 overlay (Overlay)
    /// </summary>
    private String GetOverlayValue( Boolean overlay )
    {
      if( overlay )
        return "1";
      else
        return "0";
    }

    #endregion
  }

  /// <summary>
  /// Specifies the possible positions for a legend.
  /// 21.2.3.24 ST_LegendPos (Legend Position)
  /// </summary>
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

  /// <summary>
  /// Specifies the possible ways to display blanks.
  /// 21.2.3.10 ST_DispBlanksAs (Display Blanks As)
  /// </summary>
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
