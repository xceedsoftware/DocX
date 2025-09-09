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
using System.IO.Packaging;
using System.Xml.Linq;

namespace Xceed.Document.NET
{
  public abstract class DataPointBase
  {
    #region Internal Properties

    internal XElement Xml
    {
      get; set;
    }

    internal PackagePart PackagePart
    {
      get; set;
    }

    internal virtual int Idx
    {
      get
      {
        return Convert.ToInt32( this.Xml.Element( XName.Get( "idx", Document.c.NamespaceName ) ).Attribute( XName.Get( "val" ) ).Value );
      }
    }

    #endregion

    #region Construtors

    internal DataPointBase( int idx )
    {
      this.Xml = new XElement( XName.Get( "dPt", Document.c.NamespaceName ),
                    new XElement( XName.Get( "idx", Document.c.NamespaceName ),
                      new XAttribute( XName.Get( "val" ), idx.ToString() ) ) );
    }

    #endregion
  }
}
