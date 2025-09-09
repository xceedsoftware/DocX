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
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace Xceed.Document.NET
{
  public class PieChart : Chart
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
            var serie = new PieSeries( element );
            serie.PackagePart = this.PackagePart;
            _series.Add( serie );
          }
        }
        return _series;
      }
    }

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

    #endregion // Public Properties

    #region Protected Properties

    protected override Type AllowedSeriesType
    {
      get
      {
        return typeof( PieSeries );
      }
    }

    #endregion // Protected Properties

    #region Constructors
    [Obsolete( "PieChart() is obsolete. Use Document.AddChart<PieChart>() instead." )]
    public PieChart()
    {
    }


    #endregion

    #region Overrides

    protected override XElement CreateExternalChartXml()
    {
      return XElement.Parse(
          @"<c:pieChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                  </c:pieChart>" );
    }

    protected override XElement GetChartTypeXElement()
    {
      if( this.ExternalXml == null )
        return null;

      return this.ExternalXml.Descendants().Where( chartElement => ( chartElement.Name.LocalName == "pieChart" )
                                                                     || ( chartElement.Name.LocalName == "pie3DChart" ) ).SingleOrDefault();

    }

    #endregion
  }
}
