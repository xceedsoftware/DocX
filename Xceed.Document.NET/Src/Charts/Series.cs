/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2023 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace Xceed.Document.NET
{
  public class Series
  {
    /// <summary>
    /// Represents a chart series
    /// </summary>
    /// 

    #region Private Members

    private XElement _strCache;
    private XElement _numCache;
    private PackagePart _packagePart;
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
        var spPrElement = this.Xml.Element( XName.Get( "spPr", Document.c.NamespaceName ) );
        string widthValue = string.Empty;

        if( spPrElement != null )
        {
          var ln = spPrElement.Element( XName.Get( "ln", Document.a.NamespaceName ) );
          if( ln != null )
          {
            var val = ln.Attribute( XName.Get( "w" ) );
            if( val != null )
            {
              widthValue = val.Value;
            }
          }
          spPrElement.Remove();
        }

        var colorData = new XElement( XName.Get( "solidFill", Document.a.NamespaceName ),
                                   new XElement( XName.Get( "srgbClr", Document.a.NamespaceName ), new XAttribute( XName.Get( "val" ), value.ToHex() ) ) );

        // When the chart containing this series is a lineChart, the line will be colored, else the shape will be colored.
        if( string.IsNullOrEmpty( widthValue ) )
        {
          spPrElement = ( ( this.Xml.Parent != null ) && ( this.Xml.Parent.Name != null ) && ( this.Xml.Parent.Name.LocalName == "lineChart" ) )
               ? new XElement( XName.Get( "spPr", Document.c.NamespaceName ),
                          new XElement( XName.Get( "ln", Document.a.NamespaceName ), colorData ) )
               : new XElement( XName.Get( "spPr", Document.c.NamespaceName ), colorData );
        }
        else
        {
          spPrElement = new XElement( XName.Get( "spPr", Document.c.NamespaceName ),
                          new XElement( XName.Get( "ln", Document.a.NamespaceName ),
                                       new XAttribute( XName.Get( "w" ), widthValue ), colorData ) );
        }

        this.Xml.Element( XName.Get( "tx", Document.c.NamespaceName ) ).AddAfterSelf( spPrElement );
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

    internal PackagePart PackagePart
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
