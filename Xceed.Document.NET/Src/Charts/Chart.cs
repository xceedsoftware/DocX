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
using System.Linq;
using System.Xml.Linq;
using System.IO.Packaging;

namespace Xceed.Document.NET
{
  public abstract class Chart : BaseChart
  {
    #region Public Properties

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

    #endregion

    #region Constructors

    public Chart()
    : base()
    {
    }

    internal Chart( Paragraph parentParagraph, PackageRelationship packageRelationship, PackagePart packagePart, XDocument chartDocument )
        : base( parentParagraph, packageRelationship, packagePart, chartDocument )
    {
    }

    #endregion
  }

}
