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
using System.Linq;
using System.Xml.Linq;
using System.Drawing;
using System.Globalization;

namespace Xceed.Document.NET
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
    private Border _border;
    private double? _size;
    private Color? _fontColor;
    private Color? _underlineColor;
    private UnderlineStyle? _underlineStyle;
    private Misc? _misc;
    private CapsStyle? _capsStyle;
    private Font _fontFamily;
    private float? _percentageScale;
    private float? _kerning;
    private float? _position;
    private double? _spacing;
    private string _styleId;

    private CultureInfo _language;

    #endregion

    #region Constructors

    /// <summary>
    /// A text formatting.
    /// </summary>
    public Formatting()
    {
      // Use current culture by default
      _language = CultureInfo.CurrentCulture;

      _rPr = new XElement( XName.Get( "rPr", Document.w.NamespaceName ) );
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

        if( temp - (int)temp == 0 )
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
    /// Percentage scale must be between 1 and 600.
    /// </summary>
    public float? PercentageScale
    {
      get
      {
        return _percentageScale;
      }

      set
      {
        if( value == null )
        {
          _percentageScale = null;
        }
        else
        {
          if( ( value >= 1f ) && ( value <= 600f ) )
          {
            _percentageScale = value;
          }
          else
            throw new ArgumentException( "PercentageScale", "Value must be in the range 1 - 600" );
        }
      }
    }

    /// <summary>
    /// The Kerning to apply to this text.
    /// </summary>
    public float? Kerning
    {
      get
      {
        return _kerning;
      }

      set
      {
        _kerning = value;
      }
    }

    /// <summary>
    /// Text position must be in the range (-1585 - 1585).
    /// </summary>
    public float? Position
    {
      get
      {
        return _position;
      }

      set
      {
        if( value > -1585f && value < 1585f )
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

        if( temp - (int)temp == 0 )
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

    public Border Border
    {
      get
      {
        return _border;
      }
      set
      {
        _border = value;
      }
    }

    [Obsolete( "This property is obsolete and should no longer be used. Use StyleId instead." )]
    public string StyleName
    {
      get
      {
        return _styleId;
      }
      set
      {
        _styleId = value;
      }
    }

    public string StyleId
    {
      get
      {
        return _styleId;
      }
      set
      {
        _styleId = value;
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
        _rPr = new XElement( XName.Get( "rPr", Document.w.NamespaceName ) );

        if( _language != null )
        {
          _rPr.Add( new XElement( XName.Get( "lang", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), _language.Name ) ) );
        }

        if( _spacing.HasValue )
        {
          _rPr.Add( new XElement( XName.Get( "spacing", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), _spacing.Value * 20 ) ) );
        }

        if( !string.IsNullOrEmpty( _styleId ) )
        {
          _rPr.Add( new XElement( XName.Get( "rStyle", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), _styleId ) ) );
        }

        if( _position.HasValue )
        {
          _rPr.Add( new XElement( XName.Get( "position", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), _position.Value * 2f ) ) );
        }

        if( _kerning.HasValue )
        {
          _rPr.Add( new XElement( XName.Get( "kern", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), _kerning.Value * 2f ) ) );
        }

        if( _percentageScale.HasValue )
        {
          _rPr.Add( new XElement( XName.Get( "w", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), _percentageScale ) ) );
        }

        if( _fontFamily != null )
        {
          _rPr.Add
          (
              new XElement( XName.Get( "rFonts", Document.w.NamespaceName ), new XAttribute( XName.Get( "ascii", Document.w.NamespaceName ), _fontFamily.Name ),
                                                                         new XAttribute( XName.Get( "hAnsi", Document.w.NamespaceName ), _fontFamily.Name ),
                                                                         new XAttribute( XName.Get( "cs", Document.w.NamespaceName ), _fontFamily.Name ),
                                                                         new XAttribute( XName.Get( "eastAsia", Document.w.NamespaceName ), _fontFamily.Name ) )
          );
        }

        if( _hidden.HasValue && _hidden.Value )
        {
          _rPr.Add( new XElement( XName.Get( "vanish", Document.w.NamespaceName ) ) );
        }

        if( _bold.HasValue && _bold.Value )
        {
          _rPr.Add( new XElement( XName.Get( "b", Document.w.NamespaceName ) ) );
        }

        if( _italic.HasValue && _italic.Value )
        {
          _rPr.Add( new XElement( XName.Get( "i", Document.w.NamespaceName ) ) );
        }

        if( _underlineStyle.HasValue )
        {
          switch( _underlineStyle )
          {
            case Xceed.Document.NET.UnderlineStyle.none:
              break;
            case Xceed.Document.NET.UnderlineStyle.singleLine:
              _rPr.Add( new XElement( XName.Get( "u", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), "single" ) ) );
              break;
            case Xceed.Document.NET.UnderlineStyle.doubleLine:
              _rPr.Add( new XElement( XName.Get( "u", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), "double" ) ) );
              break;
            default:
              _rPr.Add( new XElement( XName.Get( "u", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), _underlineStyle.ToString() ) ) );
              break;
          }
        }

        if( _underlineColor.HasValue )
        {
          // If an underlineColor has been set but no underlineStyle has been set
          if( !_underlineStyle.HasValue || ( _underlineStyle == Xceed.Document.NET.UnderlineStyle.none ) )
          {
            // Set the underlineStyle to the default
            _underlineStyle = Xceed.Document.NET.UnderlineStyle.singleLine;
            _rPr.Add( new XElement( XName.Get( "u", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), "single" ) ) );
          }

          _rPr.Element( XName.Get( "u", Document.w.NamespaceName ) ).Add( new XAttribute( XName.Get( "color", Document.w.NamespaceName ), _underlineColor.Value.ToHex() ) );
        }

        if( _strikethrough.HasValue )
        {
          switch( _strikethrough )
          {
            case Xceed.Document.NET.StrikeThrough.none:
              break;
            case Xceed.Document.NET.StrikeThrough.strike:
              _rPr.Add( new XElement( XName.Get( "strike", Document.w.NamespaceName ) ) );
              break;
            case Xceed.Document.NET.StrikeThrough.doubleStrike:
              _rPr.Add( new XElement( XName.Get( "dstrike", Document.w.NamespaceName ) ) );
              break;
            default:
              break;
          }
        }

        if( _script.HasValue )
        {
          switch( _script )
          {
            case Xceed.Document.NET.Script.none:
              break;
            default:
              _rPr.Add( new XElement( XName.Get( "vertAlign", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), _script.ToString() ) ) );
              break;
          }
        }

        if( _size.HasValue )
        {
          _rPr.Add( new XElement( XName.Get( "sz", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), ( _size * 2 ).ToString() ) ) );
          _rPr.Add( new XElement( XName.Get( "szCs", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), ( _size * 2 ).ToString() ) ) );
        }

        if( _fontColor.HasValue )
        {
          _rPr.Add( new XElement( XName.Get( "color", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), _fontColor.Value.ToHex() ) ) );
        }

        if( _highlight.HasValue )
        {
          switch( _highlight )
          {
            case Xceed.Document.NET.Highlight.none:
              break;
            default:
              _rPr.Add( new XElement( XName.Get( "highlight", Document.w.NamespaceName ), new XAttribute( XName.Get( "val", Document.w.NamespaceName ), _highlight.ToString() ) ) );
              break;
          }
        }

        if( _shading.HasValue )
        {
          _rPr.Add( new XElement( XName.Get( "shd", Document.w.NamespaceName ), new XAttribute( XName.Get( "fill", Document.w.NamespaceName ), _shading.Value.ToHex() ) ) );
        }

        if( _border != null )
        {
          _rPr.Add( new XElement( XName.Get( "bdr", Document.w.NamespaceName ),
                    new object[] { new XAttribute( XName.Get( "color", Document.w.NamespaceName ), _border.Color ),
                                   new XAttribute( XName.Get( "space", Document.w.NamespaceName ), _border.Space ),
                                   new XAttribute( XName.Get( "sz", Document.w.NamespaceName ), _border.Size ),
                                   new XAttribute( XName.Get( "val", Document.w.NamespaceName ), _border.Tcbs )
                                 } ) );
        }

        if( _capsStyle.HasValue )
        {
          switch( _capsStyle )
          {
            case Xceed.Document.NET.CapsStyle.none:
              break;
            default:
              _rPr.Add( new XElement( XName.Get( _capsStyle.ToString(), Document.w.NamespaceName ) ) );
              break;
          }
        }

        if( _misc.HasValue )
        {
          switch( _misc )
          {
            case Xceed.Document.NET.Misc.none:
              break;
            case Xceed.Document.NET.Misc.outlineShadow:
              _rPr.Add( new XElement( XName.Get( "outline", Document.w.NamespaceName ) ) );
              _rPr.Add( new XElement( XName.Get( "shadow", Document.w.NamespaceName ) ) );
              break;
            case Xceed.Document.NET.Misc.engrave:
              _rPr.Add( new XElement( XName.Get( "imprint", Document.w.NamespaceName ) ) );
              break;
            default:
              _rPr.Add( new XElement( XName.Get( _misc.ToString(), Document.w.NamespaceName ) ) );
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
      clone.Border = _border;
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
      if( !string.IsNullOrEmpty( _styleId ) )
      {
        clone.StyleId = _styleId;
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
            var cultureString = option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ), null ) ?? option.GetAttribute( XName.Get( "eastAsia", Document.w.NamespaceName ), null ) ?? option.GetAttribute( XName.Get( "bidi", Document.w.NamespaceName ) );
            try
            {
              formatting.Language = new CultureInfo( cultureString );
            }
            catch( Exception )
            {
              formatting.Language = CultureInfo.CurrentCulture;
            }
            break;
          case "spacing":
            formatting.Spacing = Double.Parse( option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) ) / 20.0;
            break;
          case "position":
            formatting.Position = Int32.Parse( option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) ) / 2f;
            break;
          case "kern":
            formatting.Kerning = Int32.Parse( option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) ) / 2f;
            break;
          case "w":
            formatting.PercentageScale = Int32.Parse( option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) );
            break;
          case "sz":
            formatting.Size = Double.Parse( option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) ) / 2;
            break;
          case "rFonts":
            var fontName = option.GetAttribute( XName.Get( "ascii", Document.w.NamespaceName ), null )
                            ?? option.GetAttribute( XName.Get( "hAnsi", Document.w.NamespaceName ), null )
                            ?? option.GetAttribute( XName.Get( "cs", Document.w.NamespaceName ), null )
                            ?? option.GetAttribute( XName.Get( "hint", Document.w.NamespaceName ), null )
                            ?? option.GetAttribute( XName.Get( "eastAsia", Document.w.NamespaceName ), null );

            formatting.FontFamily = ( fontName != null )
                                    ? new Font( fontName )
                                    : ( formatting.FontFamily == null ) ?
                                      new Font( "Calibri" ) : formatting.FontFamily;
            break;
          case "color":
            try
            {
              var color = option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) );
              formatting.FontColor = ( color == "auto" ) ? Color.Black : HelperFunctions.GetColorFromHtml( color );
            }
            catch( Exception )
            {
              // ignore
            }
            break;
          case "vanish":
            formatting._hidden = option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) != "0";
            break;
          case "b":
            formatting.Bold = option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) != "0";
            break;
          case "i":
            formatting.Italic = option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) != "0";
            break;
          case "highlight":
            switch( option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) )
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
              default:
                formatting.Highlight = NET.Highlight.none;
                break;
            }
            break;
          case "strike":
            formatting.StrikeThrough = ( option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) == "0" ) ? NET.StrikeThrough.none : NET.StrikeThrough.strike;
            break;
          case "dstrike":
            formatting.StrikeThrough = ( option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) == "0" ) ? NET.StrikeThrough.none : NET.StrikeThrough.doubleStrike;
            break;
          case "u":
            formatting.UnderlineStyle = HelperFunctions.GetUnderlineStyle( option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) );
            try
            {
              var color = option.GetAttribute( XName.Get( "color", Document.w.NamespaceName ) );
              if( !string.IsNullOrEmpty( color ) )
              {
                formatting.UnderlineColor = HelperFunctions.GetColorFromHtml( color );
              }
              else
              {
                var fontColor = rPr.Element( XName.Get( "color", Document.w.NamespaceName ) );
                if( fontColor != null )
                {
                  var val = fontColor.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) );
                  formatting.UnderlineColor = ( val == "auto" ) || ( val == "" ) ? Color.Black : HelperFunctions.GetColorFromHtml( val );
                }
              }
            }
            catch( Exception )
            {
              // ignore
            }
            break;
          case "vertAlign": //script
            var script = option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ), null );
            Script enumScript;
            formatting.Script = Enum.TryParse( script, out enumScript ) ? enumScript : NET.Script.none;
            break;
          case "caps":
            formatting.CapsStyle = ( option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) == "0" ) ? NET.CapsStyle.none : NET.CapsStyle.caps;
            break;
          case "smallCaps":
            formatting.CapsStyle = ( option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ) ) == "0" ) ? NET.CapsStyle.none : NET.CapsStyle.smallCaps;
            break;
          case "shd":
            var fill = option.GetAttribute( XName.Get( "fill", Document.w.NamespaceName ) );
            if( !string.IsNullOrEmpty( fill ) )
            {
              formatting.Shading = HelperFunctions.GetColorFromHtml( fill );
            }
            break;
          case "bdr":
            formatting.Border = HelperFunctions.GetBorderFromXml( option );
            break;
          case "rStyle":
            var style = option.GetAttribute( XName.Get( "val", Document.w.NamespaceName ), null );
            formatting.StyleId = style;
            break;
          default:
            break;
        }
      }

      return formatting;
    }

    public int CompareTo( object obj )
    {
      Formatting other = (Formatting)obj;

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

      if( other._border != _border )
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

      if( other._styleId != _styleId )
        return -1;

      if( !other._language.Equals( _language ) )
        return -1;

      return 0;
    }

    #endregion
  }
}
