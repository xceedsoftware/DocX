/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2022 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace Xceed.Document.NET
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

    #region Constructors
    [Obsolete("PieChart() is obsolete. Use Document.AddChart<PieChart>() instead.")]
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
