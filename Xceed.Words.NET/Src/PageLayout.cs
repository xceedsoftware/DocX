/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System;
using System.Xml.Linq;

namespace Xceed.Words.NET
{
  public class PageLayout : DocXElement
  {
    #region Constructors

    internal PageLayout( DocX document, XElement xml ) : base( document, xml )
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
        XElement pgSz = Xml.Element( XName.Get( "pgSz", DocX.w.NamespaceName ) );

        if( pgSz == null )
          return Orientation.Portrait;

        // Get the attribute of the pgSz element.
        XAttribute val = pgSz.Attribute( XName.Get( "orient", DocX.w.NamespaceName ) );

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
        XElement pgSz = Xml.Element( XName.Get( "pgSz", DocX.w.NamespaceName ) );

        if( pgSz == null )
        {
          Xml.SetElementValue( XName.Get( "pgSz", DocX.w.NamespaceName ), string.Empty );
          pgSz = Xml.Element( XName.Get( "pgSz", DocX.w.NamespaceName ) );
        }

        pgSz.SetAttributeValue( XName.Get( "orient", DocX.w.NamespaceName ), value.ToString().ToLower() );

        if( value == Xceed.Words.NET.Orientation.Landscape )
        {
          pgSz.SetAttributeValue( XName.Get( "w", DocX.w.NamespaceName ), "16838" );
          pgSz.SetAttributeValue( XName.Get( "h", DocX.w.NamespaceName ), "11906" );
        }

        else if( value == Xceed.Words.NET.Orientation.Portrait )
        {
          pgSz.SetAttributeValue( XName.Get( "w", DocX.w.NamespaceName ), "11906" );
          pgSz.SetAttributeValue( XName.Get( "h", DocX.w.NamespaceName ), "16838" );
        }
      }
    }

    #endregion
  }
}
