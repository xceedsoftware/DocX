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
using System.Xml.Linq;

namespace Xceed.Document.NET
{
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
}
