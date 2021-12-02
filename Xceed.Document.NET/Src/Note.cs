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
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace Xceed.Document.NET
{
  public abstract class Note : Container, IParagraphContainer
  {
    #region Constants

    internal static int DefaultCustomCharCode = 61440;  //"F000"

    #endregion

    #region Private Members

    private int _id;
    private Symbol _customMark;

    #endregion

    #region Public Properties

    #region CustomMark

    public Symbol CustomMark
    {
      get
      {
        if( _customMark != null )
          return _customMark;

        var sym = this.Xml.Descendants( XName.Get( "sym", Document.w.NamespaceName ) ).FirstOrDefault();
        if( sym != null )
        {
          _customMark = new Symbol();

          var font = sym.GetAttribute( XName.Get( "font", Document.w.NamespaceName ) );
          if( !string.IsNullOrEmpty( font ) )
          {
            _customMark.Font = new Font( font );
          }

          var code = sym.GetAttribute( XName.Get( "char", Document.w.NamespaceName ) );
          if( !string.IsNullOrEmpty( code ) )
          {
            // Do code - "F000".
            _customMark.Code = int.Parse( code, NumberStyles.HexNumber ) - DefaultCustomCharCode;
          }

          _customMark.PropertyChanged += this.CustomMark_PropertyChanged;

          return _customMark;
        }

        return null;
      }

      set
      {
        if( _customMark != null )
        {
          _customMark.PropertyChanged -= this.CustomMark_PropertyChanged;
        }

        _customMark = value;

        if( _customMark != null )
        {
          _customMark.PropertyChanged += this.CustomMark_PropertyChanged;
        }

        this.UpdateCustomMarkXml();
      }
    }

    #endregion  //CustomMark

    #region Paragraphs

    public override ReadOnlyCollection<Paragraph> Paragraphs
    {
      get
      {
        var paragraphs = base.Paragraphs;
        foreach( var paragraph in paragraphs )
        {
          paragraph.PackagePart = this.PackagePart;
        }
        return paragraphs;
      }
    }

    #endregion  // Paragraphs

    #endregion

    #region Internal Properties

    #region Id

    internal int Id
    {
      get
      {
        return _id;
      }
    }

    #endregion  //Id

    #endregion

    #region Constructor

    internal Note( Document document, PackagePart part, XElement xml ) : base( document, xml )
    {
      this.PackagePart = part;

      var id = this.Xml.Attribute( XName.Get( "id", Document.w.NamespaceName ) );
      _id = ( id != null ) ? Int32.Parse( id.Value ) : 0;
    }

    #endregion 

    #region Internal Methods

    internal abstract string GetNoteRefType();

    internal abstract XElement CreateReferenceRunCore( bool customMarkFollows, XElement symbol, Formatting noteNumberFormatting );

    internal XElement CreateReferenceRun( Formatting noteNumberFormatting )
    {
      var customMarkFollows = false;
      XElement symbol = null;

      if( this.CustomMark != null )
      {
        customMarkFollows = true;

        var font = this.CustomMark.Font.Name;
        var code = this.CustomMark.HexCode;
        symbol = XElement.Parse( string.Format( @"<w:sym xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" w:font=""{0}"" w:char=""{1}""/>", font, code ) );
      }

      // Create new reference for paragraph.
      var referenceRun = this.CreateReferenceRunCore( customMarkFollows, symbol, noteNumberFormatting );

      // When number formatting is used, update the number at the beginning of the note.
      if( noteNumberFormatting != null )
      {
        var rPr = this.Xml.Descendants( XName.Get( "rPr", Document.w.NamespaceName ) ).FirstOrDefault();
        if( rPr != null )
        {
          foreach( var numberFormattingElement in noteNumberFormatting.Xml.Elements() )
          {
            var runElement = rPr.Element( numberFormattingElement.Name );
            // run doesn't contains the property, add it.
            if( runElement == null )
            {
              rPr.Add( numberFormattingElement );
            }
            else
            {
              runElement.Remove();
              rPr.Add( numberFormattingElement );
            }
          }
        }
      }

      return referenceRun;
    }

    #endregion

    #region Private Methods

    private void UpdateCustomMarkXml()
    {
      if( _customMark != null )
      {
        var sym = new XElement( XName.Get( "sym", Document.w.NamespaceName ) );
        sym.SetAttributeValue( XName.Get( "char", Document.w.NamespaceName ), _customMark.HexCode );
        sym.SetAttributeValue( XName.Get( "font", Document.w.NamespaceName ), _customMark.Font.Name );

        var noteRef = this.Xml.Descendants( XName.Get( this.GetNoteRefType(), Document.w.NamespaceName ) ).FirstOrDefault();
        if( noteRef != null )
        {
          noteRef.ReplaceWith( sym );
        }
        else
        {
          var currentSym = this.Xml.Descendants( XName.Get( "sym", Document.w.NamespaceName ) ).FirstOrDefault();
          if( currentSym != null )
          {
            currentSym.ReplaceWith( sym );
          }
        }
      }
      else
      {
        var currentSym = this.Xml.Descendants( XName.Get( "sym", Document.w.NamespaceName ) ).FirstOrDefault();
        if( currentSym != null )
        {
          currentSym.ReplaceWith( new XElement( XName.Get( this.GetNoteRefType(), Document.w.NamespaceName ) ) );
        }
      }
    }

    #endregion

    #region Event Handlers

    private void CustomMark_PropertyChanged( object sender, PropertyChangedEventArgs e )
    {
      this.UpdateCustomMarkXml();
    }

    #endregion
  }


  public class Symbol : INotifyPropertyChanged
  {
    #region Private Members

    private int _code;
    private Font _font;

    #endregion

    #region Public Properties

    public int Code
    {
      get
      {
        return _code;
      }
      set
      {
        _code = value;
        OnPropertyChanged( "Code" );
      }
    }

    public Font Font
    {
      get
      {
        return _font;
      }
      set
      {
        _font = value;
        OnPropertyChanged( "Font" );
      }
    }

    #endregion

    #region Internal Properties

    internal string HexCode
    {
      get
      {
        // Do "F000" + this.code.
        return ( Footnote.DefaultCustomCharCode + this.Code ).ToString( "X" );
      }
    }

    #endregion

    #region INotifyPropertyChanged

    public event PropertyChangedEventHandler PropertyChanged;
    protected void OnPropertyChanged( string propertyName )
    {
      if( PropertyChanged != null )
      {
        PropertyChanged( this, new PropertyChangedEventArgs( propertyName ) );
      }
    }

    #endregion
  }
}
