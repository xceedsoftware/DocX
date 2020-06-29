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
using System.Xml.Linq;

namespace Xceed.Document.NET
{
  public class PageLayout : DocumentElement
  {
    #region Constructors

    internal PageLayout( Document document, XElement xml ) : base( document, xml )
    {

    }

    #endregion

    #region Public Properties

    public Orientation Orientation
    {
      get
      {
        /*
         * Get the pgSz (page size) element for this Section,
         * null will be return if no such element exists.
         */
        XElement pgSz = Xml.Element( XName.Get( "pgSz", Document.w.NamespaceName ) );

        if( pgSz == null )
          return Orientation.Portrait;

        // Get the attribute of the pgSz element.
        XAttribute val = pgSz.Attribute( XName.Get( "orient", Document.w.NamespaceName ) );

        // If val is null, this cell contains no information.
        if( val == null )
          return Orientation.Portrait;

        if( val.Value.Equals( "Landscape", StringComparison.CurrentCultureIgnoreCase ) )
          return Orientation.Landscape;
        else
          return Orientation.Portrait;
      }

      set
      {
        // Check if already correct value.
        if( Orientation == value )
          return;

        /*
         * Get the pgSz (page size) element for this Section,
         * null will be return if no such element exists.
         */
        XElement pgSz = Xml.Element( XName.Get( "pgSz", Document.w.NamespaceName ) );

        if( pgSz == null )
        {
          Xml.SetElementValue( XName.Get( "pgSz", Document.w.NamespaceName ), string.Empty );
          pgSz = Xml.Element( XName.Get( "pgSz", Document.w.NamespaceName ) );
        }

        pgSz.SetAttributeValue( XName.Get( "orient", Document.w.NamespaceName ), value.ToString().ToLower() );

        if( value == Xceed.Document.NET.Orientation.Landscape )
        {
          pgSz.SetAttributeValue( XName.Get( "w", Document.w.NamespaceName ), "16838" );
          pgSz.SetAttributeValue( XName.Get( "h", Document.w.NamespaceName ), "11906" );
        }

        else if( value == Xceed.Document.NET.Orientation.Portrait )
        {
          pgSz.SetAttributeValue( XName.Get( "w", Document.w.NamespaceName ), "11906" );
          pgSz.SetAttributeValue( XName.Get( "h", Document.w.NamespaceName ), "16838" );
        }
      }
    }

    #endregion
  }
}
