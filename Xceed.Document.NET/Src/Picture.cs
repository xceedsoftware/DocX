/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.IO.Packaging;
using System.Diagnostics;
using System;
using System.Globalization;
using System.IO;
using Xceed.Drawing;

namespace Xceed.Document.NET
{
  public class Picture : DocumentElement
  {
    #region Private Members

    private string _id;
    private string _name;
    private string _descr;
    private long _cx, _cy;
    private uint _rotation;
    private bool _hFlip, _vFlip;
    private object _pictureShape;
    private XElement _xfrm;
    private XElement _prstGeom;

    // Calculating Height & Width in Inches
    // https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
    // http://lcorneliussen.de/raw/dashboards/ooxml/
    private const int InchToEmuFactor = 914400;
    private const double EmuToInchFactor = 1d / InchToEmuFactor;

    #endregion

    #region Internal Members

    internal const int EmusInPixel = ( Picture.InchToEmuFactor / 72 ); // 12700, Result of : 914400 EMUs per inch / 72 pixels per inch.

    internal Image _img;

    #endregion

    #region Public Properties

    public string Id
    {
      get
      {
        return _id;
      }
    }

    public bool FlipHorizontal
    {
      get
      {
        return _hFlip;
      }

      set
      {
        _hFlip = value;

        var flipH = _xfrm.Attribute( XName.Get( "flipH" ) );
        if( flipH == null )
        {
          _xfrm.Add( new XAttribute( XName.Get( "flipH" ), "0" ) );
        }

        _xfrm.Attribute( XName.Get( "flipH" ) ).Value = _hFlip ? "1" : "0";
      }
    }

    public bool FlipVertical
    {
      get
      {
        return _vFlip;
      }

      set
      {
        _vFlip = value;

        var flipV = _xfrm.Attribute( XName.Get( "flipV" ) );
        if( flipV == null )
        {
          _xfrm.Add( new XAttribute( XName.Get( "flipV" ), "0" ) );
        }

        _xfrm.Attribute( XName.Get( "flipV" ) ).Value = _vFlip ? "1" : "0";
      }
    }

    public uint Rotation
    {
      get
      {
        return _rotation / 60000;
      }

      set
      {
        _rotation = ( value % 360 ) * 60000;
        var xfrm =
            ( from d in Xml.Descendants()
              where d.Name.LocalName.Equals( "xfrm" )
              select d ).Single();

        var rot = xfrm.Attribute( XName.Get( "rot" ) );
        if( rot == null )
        {
          xfrm.Add( new XAttribute( XName.Get( "rot" ), 0 ) );
        }

        xfrm.Attribute( XName.Get( "rot" ) ).Value = _rotation.ToString();
      }
    }














    public string Name
    {
      get
      {
        return _name;
      }

      set
      {
        _name = value;

        foreach( XAttribute a in Xml.Descendants().Attributes( XName.Get( "name" ) ) )
        {
          a.Value = _name;
        }
      }
    }

    public string Description
    {
      get
      {
        return _descr;
      }

      set
      {
        _descr = value;

        foreach( XAttribute a in Xml.Descendants().Attributes( XName.Get( "descr" ) ) )
        {
          a.Value = _descr;
        }
      }
    }

    public string FileName
    {
      get
      {
        return _img.FileName;
      }
    }

    public float Width
    {
      get
      {
        return _cx / EmusInPixel;
      }

      set
      {
        _cx = Convert.ToInt64(value * EmusInPixel);

        foreach( XAttribute a in Xml.Descendants().Attributes( XName.Get( "cx" ) ) )
          a.Value = _cx.ToString();
      }
    }

    public float WidthInches
    {
      get
      {
        return this.Width / 72f;
      }

      set
      {
        Width = ( value * 72f );
      }
    }

    public float Height
    {
      get
      {
        return _cy / EmusInPixel;
      }

      set
      {
        _cy = Convert.ToInt64(value * EmusInPixel);

        foreach( XAttribute a in Xml.Descendants().Attributes( XName.Get( "cy" ) ) )
          a.Value = _cy.ToString();
      }
    }

    public float HeightInches
    {
      get
      {
        return Height / 72f;
      }

      set
      {
        Height = ( value * 72f );
      }
    }

    public Stream Stream
    {
      get
      {
        return _img.GetStream( FileMode.Open, FileAccess.Read );
      }
    }

    #endregion

    #region Constructors

    internal Picture( Document document, XElement i, Image image ) : base( document, i )
    {
      _img = image;

      _id = ( image != null ) 
           ? image.Id
           : (
                from e in Xml.Descendants()
                where e.Name.LocalName.Equals( "imagedata" )
                select e.Attribute( XName.Get( "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships" ) ).Value
             ).FirstOrDefault();

      var pic = this.Xml.Descendants( XName.Get( "pic", Document.pic.NamespaceName ) ).FirstOrDefault( p =>
      {
        var id =
        (
          from e in p.Descendants()
          where e.Name.LocalName.Equals( "blip" )
          select e.Attribute( XName.Get( "embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships" ) ).Value
        );

        return (id.FirstOrDefault() == _id);
      } );

      if( pic == null )
      {
        pic = this.Xml;
      }

      var nameToFind =
      (
          from e in pic.Descendants()
          let a = e.Attribute( XName.Get( "name" ) )
          where ( a != null )
          select a.Value
      ).FirstOrDefault();

      _name = ( nameToFind != null )
             ? nameToFind
             : (
                  from e in pic.Descendants()
                  let a = e.Attribute( XName.Get( "title" ) )
                  where ( a != null )
                  select a.Value
               ).FirstOrDefault();

      _descr =
      (
          from e in pic.Descendants()
          let a = e.Attribute( XName.Get( "descr" ) )
          where ( a != null )
          select a.Value
      ).FirstOrDefault();

      _cx =
      (
          from e in pic.Descendants()
          let a = e.Attribute( XName.Get( "cx" ) )
          where ( a != null )
          select long.Parse( a.Value )
      ).FirstOrDefault();

      if( _cx == 0 )
      {
        var style = 
        (
            from e in pic.Descendants()
            let a = e.Attribute( XName.Get( "style" ) )
            where ( a != null )
            select a
        ).FirstOrDefault();

        if( style != null )
        {
          var widthString = style.Value.Substring( style.Value.IndexOf( "width:" ) + 6 );
          var widthIndex = widthString.IndexOf( "pt" );
          Debug.Assert( widthIndex >= 0, "widthString has a wrong format." );
          if( widthIndex >= 0 )
          {
            var widthValueString = widthString.Substring( 0, widthIndex );
            var widthDouble = double.Parse( widthValueString, CultureInfo.InvariantCulture ) * EmusInPixel;
            _cx = System.Convert.ToInt64( widthDouble );
          }
        }
      }

      _cy =
      (
          from e in pic.Descendants()
          let a = e.Attribute( XName.Get( "cy" ) )
          where ( a != null )
          select long.Parse( a.Value )
      ).FirstOrDefault();

      if( _cy == 0 )
      {
        var style =
        (
            from e in pic.Descendants()
            let a = e.Attribute( XName.Get( "style" ) )
            where ( a != null )
            select a
        ).FirstOrDefault();

        if( style != null )
        {
          var heightString = style.Value.Substring( style.Value.IndexOf( "height:" ) + 7 );
          var heightIndex = heightString.IndexOf( "pt" );
          Debug.Assert( heightIndex >= 0, "heightString has a wrong format." );
          if( heightIndex >= 0 )
          {
            var heightValueString = heightString.Substring( 0, heightIndex );
            var heightDouble = double.Parse( heightValueString, CultureInfo.InvariantCulture ) * EmusInPixel;
            _cy = System.Convert.ToInt64( heightDouble );
          }
        }
      }

      _xfrm =
      (
          from d in pic.Descendants()
          where d.Name.LocalName.Equals( "xfrm" )
          select d
      ).FirstOrDefault();

      _prstGeom =
      (
          from d in pic.Descendants()
          where d.Name.LocalName.Equals( "prstGeom" )
          select d
      ).FirstOrDefault();

      if( _xfrm != null )
      {
        _rotation = _xfrm.Attribute( XName.Get( "rot" ) ) == null ? 0 : uint.Parse( _xfrm.Attribute( XName.Get( "rot" ) ).Value );
      }







    }

    #endregion

    #region Public Methods

    public void Remove()
    {
      if( this.Xml.Parent != null )
      {
        this.Xml.Remove();
        _img.Remove();
      }
    }

    public void SetPictureShape( BasicShapes shape )
    {
      SetPictureShape( ( object )shape );
    }

    public void SetPictureShape( RectangleShapes shape )
    {
      SetPictureShape( ( object )shape );
    }

    public void SetPictureShape( BlockArrowShapes shape )
    {
      SetPictureShape( ( object )shape );
    }

    public void SetPictureShape( EquationShapes shape )
    {
      SetPictureShape( ( object )shape );
    }

    public void SetPictureShape( FlowchartShapes shape )
    {
      SetPictureShape( ( object )shape );
    }

    public void SetPictureShape( StarAndBannerShapes shape )
    {
      SetPictureShape( ( object )shape );
    }

    public void SetPictureShape( CalloutShapes shape )
    {
      SetPictureShape( ( object )shape );
    }

    public Paragraph InsertCaptionAfterSelf( string caption )
    {
      var parentXElement = this.GetParentP();

      if( parentXElement == null )
        throw new Exception( "Cannot find the parent XElement" );

      var p = new Paragraph(this.Document, parentXElement, 0);

      return p.InsertCaptionAfterSelf( caption );
    }

    #endregion

    #region Internal Methods

    internal bool IsJpegImage()
    {
      return ( !string.IsNullOrEmpty( this.FileName ) && 
        ( this.FileName.EndsWith( "jpg" ) || this.FileName.EndsWith( "jpeg" ) ) );
    }

    internal XElement GetParentP()
    {
      return this.Xml.Ancestors( XName.Get( "p", Document.w.NamespaceName ) ).FirstOrDefault();
    }

    #endregion

    #region Private Methods

    private void SetPictureShape( object shape )
    {
      _pictureShape = shape;

      XAttribute prst = _prstGeom.Attribute( XName.Get( "prst" ) );
      if( prst == null )
      {
        _prstGeom.Add( new XAttribute( XName.Get( "prst" ), "rectangle" ) );
      }

      _prstGeom.Attribute( XName.Get( "prst" ) ).Value = shape.ToString();
    }

    #endregion
































  }















}
