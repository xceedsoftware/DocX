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
using System.Collections;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace Xceed.Document.NET
{
  public abstract class Series : CommonSeries
  {

    #region Private Members

    private XElement _strCache;
    private XElement _numCache;
    private PackagePart _packagePart;
    #endregion

    #region Public Properties
























    #endregion

    #region Internal Properties

    internal override PackagePart PackagePart
    {
      get
      {
        return _packagePart;
      }
      set
      {
        _packagePart = value;
      }
    }
    #endregion

    #region Constructors

    [Obsolete( "Series() is obsolete. Use new LineSeries, new BarSeries, new PieSeries...  instead." )]
    public Series( String name ) : base( name )
    {
      _strCache = new XElement( XName.Get( "strCache", Document.c.NamespaceName ) );
      _numCache = new XElement( XName.Get( "numCache", Document.c.NamespaceName ) );

      var serXml = new XElement( XName.Get( "ser", Document.c.NamespaceName ),
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

      this.SetXml( serXml );
    }

    internal Series( XElement xml ) : base( xml )
    {
      this.SetXml( xml );

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

    #region Internal Method







    #endregion





  }
}
