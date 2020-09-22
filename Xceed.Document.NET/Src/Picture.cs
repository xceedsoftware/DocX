/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2020 Xceed Software Inc.
 
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
using System.Drawing;
using System.Globalization;
using System.IO;

namespace Xceed.Document.NET
{
  /// <summary>
  /// Represents a Picture in this document, a Picture is a customized view of an Image.
  /// </summary>
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

    /// <summary>
    /// A unique id that identifies an Image embedded in this document.
    /// </summary>
    public string Id
    {
      get
      {
        return _id;
      }
    }

    /// <summary>
    /// Flip this Picture Horizontally.
    /// </summary>
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

    /// <summary>
    /// Flip this Picture Vertically.
    /// </summary>
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

    /// <summary>
    /// The rotation in degrees of this image, actual value = value % 360
    /// </summary>
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














    /// <summary>
    /// Gets or sets the name of this Image.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the description for this Image.
    /// </summary>
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

    ///<summary>
    /// Returns the name of the image file for the picture.
    ///</summary>
    public string FileName
    {
      get
      {
        return _img.FileName;
      }
    }

    /// <summary>
    /// Gets or sets the Width of this Image.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the Width of this Image (in Inches)
    /// </summary>
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

    /// <summary>
    /// Gets or sets the height of this Image.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the Height of this Image (in Inches)
    /// </summary>
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

    /// <summary>
    /// Wraps an XElement as an Image
    /// </summary>
    /// <param name="document"></param>
    /// <param name="i">The XElement i to wrap</param>
    /// <param name="image"></param>
    internal Picture( Document document, XElement i, Image image ) : base( document, i )
    {
      _img = image;

      var imageId =
      (
          from e in Xml.Descendants()
          where e.Name.LocalName.Equals( "blip" )
          select e.Attribute( XName.Get( "embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships" ) ).Value
      ).SingleOrDefault();

      _id = ( imageId != null ) 
           ? imageId
           : (
                from e in Xml.Descendants()
                where e.Name.LocalName.Equals( "imagedata" )
                select e.Attribute( XName.Get( "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships" ) ).Value
             ).FirstOrDefault();

      var nameToFind =
      (
          from e in Xml.Descendants()
          let a = e.Attribute( XName.Get( "name" ) )
          where ( a != null )
          select a.Value
      ).FirstOrDefault();

      _name = ( nameToFind != null )
             ? nameToFind
             : (
                  from e in Xml.Descendants()
                  let a = e.Attribute( XName.Get( "title" ) )
                  where ( a != null )
                  select a.Value
               ).FirstOrDefault();

      _descr =
      (
          from e in Xml.Descendants()
          let a = e.Attribute( XName.Get( "descr" ) )
          where ( a != null )
          select a.Value
      ).FirstOrDefault();

      _cx =
      (
          from e in Xml.Descendants()
          let a = e.Attribute( XName.Get( "cx" ) )
          where ( a != null )
          select long.Parse( a.Value )
      ).FirstOrDefault();

      if( _cx == 0 )
      {
        var style = 
        (
            from e in Xml.Descendants()
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
            _cx = long.Parse( widthValueString, CultureInfo.InvariantCulture ) * EmusInPixel;
          }
        }
      }

      _cy =
      (
          from e in Xml.Descendants()
          let a = e.Attribute( XName.Get( "cy" ) )
          where ( a != null )
          select long.Parse( a.Value )
      ).FirstOrDefault();

      if( _cy == 0 )
      {
        var style =
        (
            from e in Xml.Descendants()
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
            _cy = long.Parse( heightValueString, CultureInfo.InvariantCulture ) * EmusInPixel;
          }
        }
      }

      _xfrm =
      (
          from d in Xml.Descendants()
          where d.Name.LocalName.Equals( "xfrm" )
          select d
      ).FirstOrDefault();

      _prstGeom =
      (
          from d in Xml.Descendants()
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

    /// <summary>
    /// Remove this Picture from this document.
    /// </summary>
    public void Remove()
    {
      Xml.Remove();
    }

    /// <summary>
    /// Set the shape of this Picture to one in the BasicShapes enumeration.
    /// </summary>
    /// <param name="shape">A shape from the BasicShapes enumeration.</param>
    public void SetPictureShape( BasicShapes shape )
    {
      SetPictureShape( ( object )shape );
    }

    /// <summary>
    /// Set the shape of this Picture to one in the RectangleShapes enumeration.
    /// </summary>
    /// <param name="shape">A shape from the RectangleShapes enumeration.</param>
    public void SetPictureShape( RectangleShapes shape )
    {
      SetPictureShape( ( object )shape );
    }

    /// <summary>
    /// Set the shape of this Picture to one in the BlockArrowShapes enumeration.
    /// </summary>
    /// <param name="shape">A shape from the BlockArrowShapes enumeration.</param>
    public void SetPictureShape( BlockArrowShapes shape )
    {
      SetPictureShape( ( object )shape );
    }

    /// <summary>
    /// Set the shape of this Picture to one in the EquationShapes enumeration.
    /// </summary>
    /// <param name="shape">A shape from the EquationShapes enumeration.</param>
    public void SetPictureShape( EquationShapes shape )
    {
      SetPictureShape( ( object )shape );
    }

    /// <summary>
    /// Set the shape of this Picture to one in the FlowchartShapes enumeration.
    /// </summary>
    /// <param name="shape">A shape from the FlowchartShapes enumeration.</param>
    public void SetPictureShape( FlowchartShapes shape )
    {
      SetPictureShape( ( object )shape );
    }

    /// <summary>
    /// Set the shape of this Picture to one in the StarAndBannerShapes enumeration.
    /// </summary>
    /// <param name="shape">A shape from the StarAndBannerShapes enumeration.</param>
    public void SetPictureShape( StarAndBannerShapes shape )
    {
      SetPictureShape( ( object )shape );
    }

    /// <summary>
    /// Set the shape of this Picture to one in the CalloutShapes enumeration.
    /// </summary>
    /// <param name="shape">A shape from the CalloutShapes enumeration.</param>
    public void SetPictureShape( CalloutShapes shape )
    {
      SetPictureShape( ( object )shape );
    }

    //public void Delete()
    //{
    //    // Remove xml
    //    i.Remove();

    //    // Rebuild the image collection for this paragraph
    //    // Requires that every Image have a link to its paragraph

    //}

    #endregion

    #region Internal Methods

    internal bool IsJpegImage()
    {
      return ( !string.IsNullOrEmpty( this.FileName ) && 
        ( this.FileName.EndsWith( "jpg" ) || this.FileName.EndsWith( "jpeg" ) ) );
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
