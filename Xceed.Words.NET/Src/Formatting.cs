/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System;
using System.Linq;
using System.Xml.Linq;
using System.Drawing;
using System.Globalization;

namespace Xceed.Words.NET
{
  /// <summary>
  /// A text formatting.
  /// </summary>
  public class Formatting : IComparable
  {
    #region Private Members

    private XElement _rPr;
    private bool? _hidden;
    private bool? _bold;
    private bool? _italic;
    private StrikeThrough? _strikethrough;
    private Script? _script;
    private Highlight? _highlight;
    private Color? _shading;
    private double? _size;
    private Color? _fontColor;
    private Color? _underlineColor;
    private UnderlineStyle? _underlineStyle;
    private Misc? _misc;
    private CapsStyle? _capsStyle;
    private Font _fontFamily;
    private int? _percentageScale;
    private int? _kerning;
    private int? _position;
    private double? _spacing;
    private string _styleName;

    private CultureInfo _language;

    #endregion

    #region Constructors

    /// <summary>
    /// A text formatting.
    /// </summary>
    public Formatting()
    {
      _capsStyle = Xceed.Words.NET.CapsStyle.none;
      _strikethrough = Xceed.Words.NET.StrikeThrough.none;
      _script = Xceed.Words.NET.Script.none;
      _highlight = Xceed.Words.NET.Highlight.none;
      _underlineStyle = Xceed.Words.NET.UnderlineStyle.none;
      _misc = Xceed.Words.NET.Misc.none;

      // Use current culture by default
      _language = CultureInfo.CurrentCulture;

      _rPr = new XElement( XName.Get( "rPr", DocX.w.NamespaceName ) );
    }

    #endregion

    #region Public Properties

    /// <summary>
    /// Text language
    /// </summary>
    public CultureInfo Language
    {
      get
      {
        return _language;
      }

      set
      {
        _language = value;
      }
    }

    /// <summary>
    /// This formatting will apply Bold.
    /// </summary>
    public bool? Bold
    {
      get
      {
        return _bold;
      }
      set
      {
        _bold = value;
      }
    }

    /// <summary>
    /// This formatting will apply Italic.
    /// </summary>
    public bool? Italic
    {
      get
      {
        return _italic;
      }
      set
      {
        _italic = value;
      }
    }

    /// <summary>
    /// This formatting will apply StrickThrough.
    /// </summary>
    public StrikeThrough? StrikeThrough
    {
      get
      {
        return _strikethrough;
      }
      set
      {
        _strikethrough = value;
      }
    }

    /// <summary>
    /// The script that this formatting should be, normal, superscript or subscript.
    /// </summary>
    public Script? Script
    {
      get
      {
        return _script;
      }
      set
      {
        _script = value;
      }
    }

    /// <summary>
    /// The Size of this text, must be between 0 and 1638.
    /// </summary>
    public double? Size
    {
      get
      {
        return _size;
      }

      set
      {
        double? temp = value * 2;

        if( temp - ( int )temp == 0 )
        {
          if( value > 0 && value < 1639 )
          {
            _size = value;
          }
          else
            throw new ArgumentException( "Size", "Value must be in the range 0 - 1638" );
        }
        else
          throw new ArgumentException( "Size", "Value must be either a whole or half number, examples: 32, 32.5" );
      }
    }

    /// <summary>
    /// Percentage scale must be one of the following values 200, 150, 100, 90, 80, 66, 50 or 33.
    /// </summary>
    public int? PercentageScale
    {
      get
      {
        return _percentageScale;
      }

      set
      {
        if( ( new int?[] { 200, 150, 100, 90, 80, 66, 50, 33 } ).Contains( value ) )
        {
          _percentageScale = value;
        }
        else
          throw new ArgumentOutOfRangeException( "PercentageScale", "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33" );
      }
    }

    /// <summary>
    /// The Kerning to apply to this text must be one of the following values 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72.
    /// </summary>
    public int? Kerning
    {
      get
      {
        return _kerning;
      }

      set
      {
        if( new int?[] { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 }.Contains( value ) )
          _kerning = value;
        else
          throw new ArgumentOutOfRangeException( "Kerning", "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72" );
      }
    }

    /// <summary>
    /// Text position must be in the range (-1585 - 1585).
    /// </summary>
    public int? Position
    {
      get
      {
        return _position;
      }

      set
      {
        if( value > -1585 && value < 1585 )
          _position = value;
        else
          throw new ArgumentOutOfRangeException( "Position", "Value must be in the range -1585 - 1585" );
      }
    }

    /// <summary>
    /// Text spacing must be in the range (-1585 - 1585).
    /// </summary>
    public double? Spacing
    {
      get
      {
        return _spacing;
      }

      set
      {
        double? temp = value * 20;

        if( temp - ( int )temp == 0 )
        {
          if( value > -1585 && value < 1585 )
            _spacing = value;
          else
            throw new ArgumentException( "Spacing", "Value must be in the range: -1584 - 1584" );
        }

        else
          throw new ArgumentException( "Spacing", "Value must be either a whole or acurate to one decimal, examples: 32, 32.1, 32.2, 32.9" );
      }
    }

    /// <summary>
    /// The colour of the text.
    /// </summary>
    public Color? FontColor
    {
      get
      {
        return _fontColor;
      }
      set
      {
        _fontColor = value;
      }
    }

    /// <summary>
    /// Highlight colour.
    /// </summary>
    public Highlight? Highlight
    {
      get
      {
        return _highlight;
      }
      set
      {
        _highlight = value;
      }
    }

      /// <summary>
    /// Shading color.
    /// </summary>
    public Color? Shading
    {
      get
      {
        return _shading;
      }
      set
      {
        _shading = value;
      }
    }

    public string StyleName
    {
      get
      {
        return _styleName;
      }
      set
      {
        _styleName = value;
      }
    }

    /// <summary>
    /// The Underline style that this formatting applies.
    /// </summary>
    public UnderlineStyle? UnderlineStyle
    {
      get
      {
        return _underlineStyle;
      }
      set
      {
        _underlineStyle = value;
      }
    }

    /// <summary>
    /// The underline colour.
    /// </summary>
    public Color? UnderlineColor
    {
      get
      {
        return _underlineColor;
      }
      set
      {
        _underlineColor = value;
      }
    }

    /// <summary>
    /// Misc settings.
    /// </summary>
    public Misc? Misc
    {
      get
      {
        return _misc;
      }
      set
      {
        _misc = value;
      }
    }

    /// <summary>
    /// Is this text hidden or visible.
    /// </summary>
    public bool? Hidden
    {
      get
      {
        return _hidden;
      }
      set
      {
        _hidden = value;
      }
    }

    /// <summary>
    /// Capitalization style.
    /// </summary>
    public CapsStyle? CapsStyle
    {
      get
      {
        return _capsStyle;
      }
      set
      {
        _capsStyle = value;
      }
    }

    /// <summary>
    /// The font Family of this formatting.
    /// </summary>
    /// <!-- 
    /// Bug found and fixed by krugs525 on August 12 2009.
    /// Use TFS compare to see exact code change.
    /// -->
    public Font FontFamily
    {
      get
      {
        return _fontFamily;
      }
      set
      {
        _fontFamily = value;
      }
    }

    #endregion

    #region Internal Properties

    internal XElement Xml
    {
      get
      {
        _rPr = new XElement( XName.Get( "rPr", DocX.w.NamespaceName ) );

        if( _language != null )
        {
          _rPr.Add( new XElement( XName.Get( "lang", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), _language.Name ) ) );
        }

        if( _spacing.HasValue )
        {
          _rPr.Add( new XElement( XName.Get( "spacing", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), _spacing.Value * 20 ) ) );
        }

        if( !string.IsNullOrEmpty( _styleName ) )
        {
          _rPr.Add( new XElement( XName.Get( "rStyle", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), _styleName ) ) );
        }

        if( _position.HasValue )
        {
          _rPr.Add( new XElement( XName.Get( "position", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), _position.Value * 2 ) ) );
        }

        if( _kerning.HasValue )
        {
          _rPr.Add( new XElement( XName.Get( "kern", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), _kerning.Value * 2 ) ) );
        }

        if( _percentageScale.HasValue )
        {
          _rPr.Add( new XElement( XName.Get( "w", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), _percentageScale ) ) );
        }

        if( _fontFamily != null )
        {
          _rPr.Add
          (
              new XElement( XName.Get( "rFonts", DocX.w.NamespaceName ), new XAttribute( XName.Get( "ascii", DocX.w.NamespaceName ), _fontFamily.Name ),
                                                                         new XAttribute( XName.Get( "hAnsi", DocX.w.NamespaceName ), _fontFamily.Name ),
                                                                         new XAttribute( XName.Get( "cs", DocX.w.NamespaceName ), _fontFamily.Name ), 
                                                                         new XAttribute( XName.Get( "eastAsia", DocX.w.NamespaceName ), _fontFamily.Name ) )
          );
        }

        if( _hidden.HasValue && _hidden.Value )
        {
          _rPr.Add( new XElement( XName.Get( "vanish", DocX.w.NamespaceName ) ) );
        }

        if( _bold.HasValue && _bold.Value )
        {
          _rPr.Add( new XElement( XName.Get( "b", DocX.w.NamespaceName ) ) );
        }

        if( _italic.HasValue && _italic.Value )
        {
          _rPr.Add( new XElement( XName.Get( "i", DocX.w.NamespaceName ) ) );
        }

        if( _underlineStyle.HasValue )
        {
          switch( _underlineStyle )
          {
            case Xceed.Words.NET.UnderlineStyle.none:
              break;
            case Xceed.Words.NET.UnderlineStyle.singleLine:
              _rPr.Add( new XElement( XName.Get( "u", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), "single" ) ) );
              break;
            case Xceed.Words.NET.UnderlineStyle.doubleLine:
              _rPr.Add( new XElement( XName.Get( "u", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), "double" ) ) );
              break;
            default:
              _rPr.Add( new XElement( XName.Get( "u", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), _underlineStyle.ToString() ) ) );
              break;
          }
        }        

        if( _underlineColor.HasValue )
        {
          // If an underlineColor has been set but no underlineStyle has been set
          if( _underlineStyle == Xceed.Words.NET.UnderlineStyle.none )
          {
            // Set the underlineStyle to the default
            _underlineStyle = Xceed.Words.NET.UnderlineStyle.singleLine;
            _rPr.Add( new XElement( XName.Get( "u", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), "single" ) ) );
          }

          _rPr.Element( XName.Get( "u", DocX.w.NamespaceName ) ).Add( new XAttribute( XName.Get( "color", DocX.w.NamespaceName ), _underlineColor.Value.ToHex() ) );
        }

        if( _strikethrough.HasValue )
        {
          switch( _strikethrough )
          {
            case Xceed.Words.NET.StrikeThrough.none:
              break;
            case Xceed.Words.NET.StrikeThrough.strike:
              _rPr.Add( new XElement( XName.Get( "strike", DocX.w.NamespaceName ) ) );
              break;
            case Xceed.Words.NET.StrikeThrough.doubleStrike:
              _rPr.Add( new XElement( XName.Get( "dstrike", DocX.w.NamespaceName ) ) );
              break;
            default:
              break;
          }
        }

        if( _script.HasValue )
        {
          switch( _script )
          {
            case Xceed.Words.NET.Script.none:
              break;
            default:
              _rPr.Add( new XElement( XName.Get( "vertAlign", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), _script.ToString() ) ) );
              break;
          }
        }        

        if( _size.HasValue )
        {
          _rPr.Add( new XElement( XName.Get( "sz", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), ( _size * 2 ).ToString() ) ) );
          _rPr.Add( new XElement( XName.Get( "szCs", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), ( _size * 2 ).ToString() ) ) );
        }

        if( _fontColor.HasValue )
        {
          _rPr.Add( new XElement( XName.Get( "color", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), _fontColor.Value.ToHex() ) ) );
        }

        if( _highlight.HasValue )
        {
          switch( _highlight )
          {
            case Xceed.Words.NET.Highlight.none:
              break;
            default:
              _rPr.Add( new XElement( XName.Get( "highlight", DocX.w.NamespaceName ), new XAttribute( XName.Get( "val", DocX.w.NamespaceName ), _highlight.ToString() ) ) );
              break;
          }
        }

        if( _shading.HasValue )
        {
          _rPr.Add( new XElement( XName.Get( "shd", DocX.w.NamespaceName ), new XAttribute( XName.Get( "fill", DocX.w.NamespaceName ), _shading.Value.ToHex() ) ) );
        }

        if( _capsStyle.HasValue )
        {
          switch( _capsStyle )
          {
            case Xceed.Words.NET.CapsStyle.none:
              break;
            default:
              _rPr.Add( new XElement( XName.Get( _capsStyle.ToString(), DocX.w.NamespaceName ) ) );
              break;
          }
        }

        if( _misc.HasValue )
        {
          switch( _misc )
          {
            case Xceed.Words.NET.Misc.none:
              break;
            case Xceed.Words.NET.Misc.outlineShadow:
              _rPr.Add( new XElement( XName.Get( "outline", DocX.w.NamespaceName ) ) );
              _rPr.Add( new XElement( XName.Get( "shadow", DocX.w.NamespaceName ) ) );
              break;
            case Xceed.Words.NET.Misc.engrave:
              _rPr.Add( new XElement( XName.Get( "imprint", DocX.w.NamespaceName ) ) );
              break;
            default:
              _rPr.Add( new XElement( XName.Get( _misc.ToString(), DocX.w.NamespaceName ) ) );
              break;
          }
        }

        return _rPr;
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Returns a cloned instance of Formatting.
    /// </summary>
    /// <returns></returns>
    public Formatting Clone()
    {
      var clone = new Formatting();
      clone.Bold = _bold;
      clone.CapsStyle = _capsStyle;
      clone.FontColor = _fontColor;
      clone.FontFamily = _fontFamily;
      clone.Hidden = _hidden;
      clone.Highlight = _highlight;
      clone.Shading = _shading;
      clone.Italic = _italic;
      if( _kerning.HasValue )
      {
        clone.Kerning = _kerning;
      }
      clone.Language = _language;
      clone.Misc = _misc;
      if( _percentageScale.HasValue )
      {
        clone.PercentageScale = _percentageScale;
      }
      if( _position.HasValue )
      {
        clone.Position = _position;
      }
      clone.Script = _script;
      if( _size.HasValue )
      {
        clone.Size = _size;
      }
      if( _spacing.HasValue )
      {
        clone.Spacing = _spacing;
      }
      if( !string.IsNullOrEmpty( _styleName ) )
      {
        clone.StyleName = _styleName;
      }
      clone.StrikeThrough = _strikethrough;
      clone.UnderlineColor = _underlineColor;
      clone.UnderlineStyle = _underlineStyle;

      return clone;
    }


    public static Formatting Parse( XElement rPr, Formatting formatting = null )
    {
      if( formatting == null )
      {
        formatting = new Formatting();
      }

      if( rPr == null )
        return formatting;

      // Build up the Formatting object.
      foreach( XElement option in rPr.Elements() )
      {
        switch( option.Name.LocalName )
        {
          case "lang":
            formatting.Language = new CultureInfo(option.GetAttribute(XName.Get("val", DocX.w.NamespaceName), null) ?? option.GetAttribute(XName.Get("eastAsia", DocX.w.NamespaceName), null) ?? option.GetAttribute(XName.Get("bidi", DocX.w.NamespaceName)));
            break;
          case "spacing":
            formatting.Spacing = Double.Parse(option.GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 20.0;
            break;
          case "position":
            formatting.Position = Int32.Parse(option.GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 2;
            break;
          case "kern":
            formatting.Kerning = Int32.Parse(option.GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 2;
            break;
          case "w":
            formatting.PercentageScale = Int32.Parse(option.GetAttribute(XName.Get("val", DocX.w.NamespaceName)));
            break;
          case "sz":
            formatting.Size = Int32.Parse(option.GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 2;
            break;
          case "rFonts":
            var fontName = option.GetAttribute( XName.Get( "cs", DocX.w.NamespaceName ), null )
                                               ?? option.GetAttribute( XName.Get( "ascii", DocX.w.NamespaceName ), null )
                                               ?? option.GetAttribute( XName.Get( "hAnsi", DocX.w.NamespaceName ), null )
                                               ?? option.GetAttribute( XName.Get( "hint", DocX.w.NamespaceName ), null )
                                               ?? option.GetAttribute( XName.Get( "eastAsia", DocX.w.NamespaceName ), null );
            formatting.FontFamily = new Font( fontName ?? "Calibri" );
            break;
          case "color":
            try
            {
              var color = option.GetAttribute(XName.Get("val", DocX.w.NamespaceName));
              formatting.FontColor = ( color == "auto") ? Color.Black : ColorTranslator.FromHtml(string.Format("#{0}", color));
            }
            catch (Exception)
            {
              // ignore
            }
            break;
          case "vanish":
            formatting._hidden = true;
            break;
          case "b":
            formatting.Bold = option.GetAttribute( XName.Get( "val", DocX.w.NamespaceName ) ) != "0";
            break;
          case "i":
            formatting.Italic = true;
            break;
          case "highlight":
            switch( option.GetAttribute( XName.Get( "val", DocX.w.NamespaceName ) ) )
            {
              case "yellow":
                formatting.Highlight = NET.Highlight.yellow;
                break;
              case "green":
                formatting.Highlight = NET.Highlight.green;
                break;
              case "cyan":
                formatting.Highlight = NET.Highlight.cyan;
                break;
              case "magenta":
                formatting.Highlight = NET.Highlight.magenta;
                break;
              case "blue":
                formatting.Highlight = NET.Highlight.blue;
                break;
              case "red":
                formatting.Highlight = NET.Highlight.red;
                break;
              case "darkBlue":
                formatting.Highlight = NET.Highlight.darkBlue;
                break;
              case "darkCyan":
                formatting.Highlight = NET.Highlight.darkCyan;
                break;
              case "darkGreen":
                formatting.Highlight = NET.Highlight.darkGreen;
                break;
              case "darkMagenta":
                formatting.Highlight = NET.Highlight.darkMagenta;
                break;
              case "darkRed":
                formatting.Highlight = NET.Highlight.darkRed;
                break;
              case "darkYellow":
                formatting.Highlight = NET.Highlight.darkYellow;
                break;
              case "darkGray":
                formatting.Highlight = NET.Highlight.darkGray;
                break;
              case "lightGray":
                formatting.Highlight = NET.Highlight.lightGray;
                break;
              case "black":
                formatting.Highlight = NET.Highlight.black;
                break;
            }
            break;
          case "strike":
            formatting.StrikeThrough = NET.StrikeThrough.strike;
            break;
          case "dstrike":
            formatting.StrikeThrough = NET.StrikeThrough.doubleStrike;
            break;
          case "u":
            formatting.UnderlineStyle = HelperFunctions.GetUnderlineStyle(option.GetAttribute(XName.Get("val", DocX.w.NamespaceName)));
            try
            {
              var color = option.GetAttribute( XName.Get( "color", DocX.w.NamespaceName ) );
              if( !string.IsNullOrEmpty( color ) )
              {
                formatting.UnderlineColor = System.Drawing.ColorTranslator.FromHtml( string.Format( "#{0}", color ) );
              }
            }
            catch( Exception )
            {
              // ignore
            }
            break;
          case "vertAlign": //script
            var script = option.GetAttribute(XName.Get("val", DocX.w.NamespaceName), null);
            formatting.Script = (Script)Enum.Parse(typeof(Script), script);
            break;
          case "caps":
            formatting.CapsStyle = NET.CapsStyle.caps;
            break;
          case "smallCaps":
            formatting.CapsStyle = NET.CapsStyle.smallCaps;
            break;
          case "shd":
            var fill = option.GetAttribute( XName.Get( "fill", DocX.w.NamespaceName ) );
            if( !string.IsNullOrEmpty( fill ) )
            {
              formatting.Shading = System.Drawing.ColorTranslator.FromHtml( string.Format( "#{0}", fill ) );
            }
            break;
          case "rStyle":
            var style = option.GetAttribute( XName.Get( "val", DocX.w.NamespaceName ), null );
            formatting.StyleName = style;
            break;
          default:
            break;
        }
      }

      return formatting;
    }

    public int CompareTo( object obj )
    {
      Formatting other = ( Formatting )obj;

      if( other._hidden != _hidden )
        return -1;

      if( other._bold != _bold )
        return -1;

      if( other._italic != _italic )
        return -1;

      if( other._strikethrough != _strikethrough )
        return -1;

      if( other._script != _script )
        return -1;

      if( other._highlight != _highlight )
        return -1;

      if( other._shading != _shading )
        return -1;

      if( other._size != _size )
        return -1;

      if( other._fontColor != _fontColor )
        return -1;

      if( other._underlineColor != _underlineColor )
        return -1;

      if( other._underlineStyle != _underlineStyle )
        return -1;

      if( other._misc != _misc )
        return -1;

      if( other._capsStyle != _capsStyle )
        return -1;

      if( other._fontFamily != _fontFamily )
        return -1;

      if( other._percentageScale != _percentageScale )
        return -1;

      if( other._kerning != _kerning )
        return -1;

      if( other._position != _position )
        return -1;

      if( other._spacing != _spacing )
        return -1;

      if( other._styleName != _styleName )
        return -1;

      if( !other._language.Equals(_language) )
        return -1;

      return 0;
    }

    #endregion
  }
}
