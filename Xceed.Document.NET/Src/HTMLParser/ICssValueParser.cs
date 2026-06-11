/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2026 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


// #OPEN_SOURCE_EXCLUDE_FILE - This is not part of the open source version
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Xceed.Document.NET
{
  #region Interfaces

  public interface ICssValueParser
  {
    bool TryParse( string token, out string outValue );
  }

  #endregion

  #region Shared Regex Patterns

  internal static class CssRegexPatterns
  {
    // Standard length/percentage pattern used by multiple parsers
    internal static readonly Regex LengthPercentage = new Regex( @"^[+-]?(\d+|\d*\.\d+)(px|em|rem|pt|pc|in|cm|mm|%|vw|vh|vmin|vmax)?$", RegexOptions.Compiled | RegexOptions.IgnoreCase );

    // Extended length/percentage with additional units (ex, ch)
    internal static readonly Regex ExtendedLengthPercentage = new Regex( @"^[+-]?(\d+|\d*\.\d+)(px|em|rem|pt|pc|in|cm|mm|ex|ch|vw|vh|vmin|vmax|%)?$", RegexOptions.Compiled | RegexOptions.IgnoreCase );

    // Simple length without percentage
    internal static readonly Regex LengthOnly = new Regex( @"^[+-]?(\d+|\d*\.\d+)(px|em|rem|pt|pc|in|cm|mm)?$", RegexOptions.Compiled | RegexOptions.IgnoreCase );

    // Pure number (no units)
    internal static readonly Regex Number = new Regex( @"^[+-]?(\d+|\d*\.\d+)$", RegexOptions.Compiled );

    // Number or percentage (no length units)
    internal static readonly Regex NumberOrPercent = new Regex( @"^(\d+|\d*\.\d+)(%)?$", RegexOptions.Compiled );

    // Hex color pattern
    internal static readonly Regex HexColor = new Regex( @"^#([0-9a-fA-F]{3,4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$", RegexOptions.Compiled );

    // Color function pattern (rgb, rgba, hsl, hsla)
    internal static readonly Regex ColorFunction = new Regex( @"^(rgb|rgba|hsl|hsla)\(", RegexOptions.IgnoreCase | RegexOptions.Compiled );

    // URL pattern
    internal static readonly Regex Url = new Regex( @"^url\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled );

    // Gradient pattern
    internal static readonly Regex Gradient = new Regex( @"^(linear-gradient|radial-gradient|conic-gradient|repeating-linear-gradient|repeating-radial-gradient|repeating-conic-gradient|image|image-set|cross-fade|element)\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled );

    // URL or gradient combined (for border-image-source)
    internal static readonly Regex UrlOrGradient = new Regex( @"^(url|linear-gradient|radial-gradient|conic-gradient|repeating-linear-gradient|repeating-radial-gradient)\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled );

    // Oblique angle pattern
    internal static readonly Regex ObliqueAngle = new Regex( @"^oblique(\s+-?\d+(\.\d+)?(deg|rad|grad|turn))?$", RegexOptions.IgnoreCase | RegexOptions.Compiled );
  }

  #endregion

  #region Base Parser Classes

  internal class KeywordParser : ICssValueParser
  {
    private readonly HashSet<string> _keywords;

    internal KeywordParser( IEnumerable<string> keywords )
    {
      _keywords = new HashSet<string>( keywords, StringComparer.OrdinalIgnoreCase );
    }

    internal KeywordParser( params string[] keywords )
        : this( (IEnumerable<string>)keywords )
    {
    }

    public bool TryParse( string token, out string outValue )
    {
      if( _keywords.Contains( token ) )
      {
        outValue = token.ToLowerInvariant();
        return true;
      }
      outValue = null;
      return false;
    }
  }

  internal abstract class KeywordOrRegexParser : ICssValueParser
  {
    private readonly HashSet<string> _keywords;
    private readonly Regex _regex;
    private readonly bool _lowercaseKeywords;
    private readonly bool _lowercaseRegexMatch;

    protected KeywordOrRegexParser( IEnumerable<string> keywords = null, Regex regex = null, bool lowercaseKeywords = true, bool lowercaseRegexMatch = false )
    {
      _keywords = keywords != null ? new HashSet<string>( keywords, StringComparer.OrdinalIgnoreCase ) : null;
      _regex = regex;
      _lowercaseKeywords = lowercaseKeywords;
      _lowercaseRegexMatch = lowercaseRegexMatch;
    }

    public virtual bool TryParse( string token, out string outValue )
    {
      if( string.IsNullOrWhiteSpace( token ) )
      {
        outValue = null;
        return false;
      }

      // Check keywords first
      if( _keywords != null && _keywords.Contains( token ) )
      {
        outValue = _lowercaseKeywords ? token.ToLowerInvariant() : token;
        return true;
      }

      // Check regex pattern
      if( _regex != null && _regex.IsMatch( token ) )
      {
        outValue = _lowercaseRegexMatch ? token.ToLowerInvariant() : token;
        return true;
      }

      outValue = null;
      return false;
    }
  }

  internal class MultiParser : ICssValueParser
  {
    private readonly ICssValueParser[] _parsers;

    internal MultiParser( params ICssValueParser[] parsers )
    {
      _parsers = parsers ?? new ICssValueParser[ 0 ];
    }

    public bool TryParse( string token, out string outValue )
    {
      foreach( var parser in _parsers )
      {
        if( parser.TryParse( token, out outValue ) )
          return true;
      }
      outValue = null;
      return false;
    }
  }

  internal class AnyTokenParser : ICssValueParser
  {
    public bool TryParse( string token, out string outValue )
    {
      if( string.IsNullOrEmpty( token ) )
      {
        outValue = null;
        return false;
      }
      outValue = token;
      return true;
    }
  }

  #endregion

  #region Specific Value Parsers

  internal class CssColorParser : ICssValueParser
  {
    private static readonly HashSet<string> ColorKeywords = new HashSet<string>( StringComparer.OrdinalIgnoreCase )
    {
      "currentcolor", "transparent", "black", "silver", "gray", "white", "maroon", "red", "purple", "fuchsia", "green", "lime", "olive", "yellow", "navy", "blue",
      "teal", "aqua", "orange", "aliceblue", "antiquewhite", "aquamarine", "azure", "beige", "bisque", "blanchedalmond", "blueviolet", "brown", "burlywood", "cadetblue",
      "chartreuse", "chocolate", "coral", "cornflowerblue", "cornsilk", "crimson", "cyan", "darkblue", "darkcyan", "darkgoldenrod", "darkgray", "darkgreen", "darkgrey",
      "darkkhaki", "darkmagenta", "darkolivegreen", "darkorange", "darkorchid", "darkred", "darksalmon", "darkseagreen", "darkslateblue", "darkslategray", "darkslategrey",
      "darkturquoise", "darkviolet", "deeppink", "deepskyblue", "dimgray", "dimgrey", "dodgerblue", "firebrick", "floralwhite", "forestgreen", "gainsboro", "ghostwhite",
      "gold", "goldenrod", "greenyellow", "grey", "honeydew", "hotpink", "indianred", "indigo", "ivory", "khaki", "lavender", "lavenderblush", "lawngreen", "lemonchiffon",
      "lightblue", "lightcoral", "lightcyan", "lightgoldenrodyellow", "lightgray", "lightgreen", "lightgrey", "lightpink", "lightsalmon", "lightseagreen", "lightskyblue",
      "lightslategray", "lightslategrey", "lightsteelblue", "lightyellow", "limegreen", "linen", "magenta", "mediumaquamarine", "mediumblue", "mediumorchid", "mediumpurple",
      "mediumseagreen", "mediumslateblue", "mediumspringgreen", "mediumturquoise", "mediumvioletred", "midnightblue", "mintcream", "mistyrose", "moccasin", "navajowhite",
      "oldlace", "olivedrab", "orangered", "orchid", "palegoldenrod", "palegreen", "paleturquoise", "palevioletred", "papayawhip", "peachpuff", "peru", "pink", "plum",
      "powderblue", "rosybrown", "royalblue", "saddlebrown", "salmon", "sandybrown", "seagreen", "seashell", "sienna", "skyblue", "slateblue", "slategray", "slategrey", "snow",
      "springgreen", "steelblue", "tan", "thistle", "tomato", "turquoise", "violet", "wheat", "whitesmoke", "yellowgreen"
    };

    public bool TryParse( string token, out string outValue )
    {
      if( string.IsNullOrWhiteSpace( token ) )
      {
        outValue = null;
        return false;
      }

      if( ColorKeywords.Contains( token ) )
      {
        outValue = token.ToLowerInvariant();
        return true;
      }

      if( CssRegexPatterns.HexColor.IsMatch( token ) || CssRegexPatterns.ColorFunction.IsMatch( token ) )
      {
        outValue = token;
        return true;
      }

      outValue = null;
      return false;
    }
  }

  internal class ImageParser : ICssValueParser
  {
    public bool TryParse( string token, out string outValue )
    {
      if( string.IsNullOrWhiteSpace( token ) )
      {
        outValue = null;
        return false;
      }

      if( string.Equals( token, "none", StringComparison.OrdinalIgnoreCase ) )
      {
        outValue = "none";
        return true;
      }

      if( CssRegexPatterns.Url.IsMatch( token ) || CssRegexPatterns.Gradient.IsMatch( token ) )
      {
        outValue = token;
        return true;
      }

      outValue = null;
      return false;
    }
  }

  internal class LengthOrKeywordParser : KeywordOrRegexParser
  {
    private static readonly string[] DefaultKeywords = { "auto", "from-font", "thin", "medium", "thick" };

    internal LengthOrKeywordParser() : base( DefaultKeywords, CssRegexPatterns.LengthPercentage )
    {
    }

    internal LengthOrKeywordParser( IEnumerable<string> additionalKeywords ) : base( CombineKeywords( additionalKeywords ), CssRegexPatterns.LengthPercentage )
    {
    }

    private static IEnumerable<string> CombineKeywords( IEnumerable<string> additional )
    {
      var combined = new List<string>( DefaultKeywords );
      if( additional != null )
        combined.AddRange( additional );
      return combined;
    }
  }

  internal class LengthParser : KeywordOrRegexParser
  {
    internal LengthParser() : base( null, CssRegexPatterns.LengthOnly )
    {
    }
  }

  internal class BackgroundPositionParser : KeywordOrRegexParser
  {
    private static readonly string[] PositionKeywords = { "left", "center", "right", "top", "bottom", "x-start", "x-end", "y-start", "y-end", "block-start", "block-end", "inline-start", "inline-end", "start", "end" };

    internal BackgroundPositionParser() : base( PositionKeywords, CssRegexPatterns.LengthPercentage )
    {
    }
  }

  internal class BackgroundSizeParser : KeywordOrRegexParser
  {
    private static readonly string[] SizeKeywords = { "auto", "cover", "contain" };

    internal BackgroundSizeParser() : base( SizeKeywords, CssRegexPatterns.LengthPercentage )
    {
    }
  }

  internal class BackgroundRepeatParser : KeywordParser
  {
    internal BackgroundRepeatParser() : base( "repeat", "repeat-x", "repeat-y", "no-repeat", "space", "round" )
    {
    }
  }

  internal class BackgroundAttachmentParser : KeywordParser
  {
    internal BackgroundAttachmentParser() : base( "scroll", "fixed", "local" )
    {
    }
  }

  internal class VisualBoxParser : KeywordParser
  {
    internal VisualBoxParser() : base( "content-box", "padding-box", "border-box", "text", "border-area" )
    {
    }
  }

  internal class FontSizeParser : ICssValueParser
  {
    private static readonly HashSet<string> SizeKeywords = new HashSet<string>( StringComparer.OrdinalIgnoreCase )
    {
      "xx-small", "x-small", "small", "medium", "large", "x-large", "xx-large", "xxx-large", "larger", "smaller", "math"
    };

    public bool TryParse( string token, out string outValue )
    {
      if( string.IsNullOrWhiteSpace( token ) )
      {
        outValue = null;
        return false;
      }

      if( SizeKeywords.Contains( token ) )
      {
        outValue = token.ToLowerInvariant();
        return true;
      }

      if( CssRegexPatterns.ExtendedLengthPercentage.IsMatch( token ) )
      {
        outValue = token;
        return true;
      }

      outValue = null;
      return false;
    }
  }

  internal class LineHeightParser : ICssValueParser
  {
    public bool TryParse( string token, out string outValue )
    {
      if( string.IsNullOrWhiteSpace( token ) )
      {
        outValue = null;
        return false;
      }

      if( string.Equals( token, "normal", StringComparison.OrdinalIgnoreCase ) )
      {
        outValue = "normal";
        return true;
      }

      if( CssRegexPatterns.Number.IsMatch( token ) || CssRegexPatterns.ExtendedLengthPercentage.IsMatch( token ) )
      {
        outValue = token;
        return true;
      }

      outValue = null;
      return false;
    }
  }

  internal class FontFamilyParser : ICssValueParser
  {
    public bool TryParse( string token, out string outValue )
    {
      if( string.IsNullOrWhiteSpace( token ) )
      {
        outValue = null;
        return false;
      }
      outValue = token;
      return true;
    }
  }

  internal class ObliqueAngleParser : KeywordOrRegexParser
  {
    internal ObliqueAngleParser() : base( null, CssRegexPatterns.ObliqueAngle, lowercaseRegexMatch: true )
    {
    }
  }

  internal class NumericRangeParser : ICssValueParser
  {
    private readonly int _min;
    private readonly int _max;

    internal NumericRangeParser( int min, int max )
    {
      _min = min;
      _max = max;
    }

    public bool TryParse( string token, out string outValue )
    {
      int value;
      if( int.TryParse( token, out value ) && value >= _min && value <= _max )
      {
        outValue = token;
        return true;
      }
      outValue = null;
      return false;
    }
  }

  internal class BorderRadiusValueParser : ICssValueParser
  {
    public bool TryParse( string token, out string outValue )
    {
      if( string.IsNullOrWhiteSpace( token ) )
      {
        outValue = null;
        return false;
      }

      // Handle "/" for elliptical radii
      if( token == "/" )
      {
        outValue = "/";
        return true;
      }

      if( CssRegexPatterns.LengthPercentage.IsMatch( token ) )
      {
        outValue = token;
        return true;
      }

      outValue = null;
      return false;
    }
  }

  internal class BorderImageSourceParser : ICssValueParser
  {
    public bool TryParse( string token, out string outValue )
    {
      if( string.IsNullOrWhiteSpace( token ) )
      {
        outValue = null;
        return false;
      }

      if( string.Equals( token, "none", StringComparison.OrdinalIgnoreCase ) )
      {
        outValue = "none";
        return true;
      }

      if( CssRegexPatterns.UrlOrGradient.IsMatch( token ) )
      {
        outValue = token;
        return true;
      }

      outValue = null;
      return false;
    }
  }

  internal class BorderImageSliceParser : ICssValueParser
  {
    public bool TryParse( string token, out string outValue )
    {
      if( string.IsNullOrWhiteSpace( token ) )
      {
        outValue = null;
        return false;
      }

      if( string.Equals( token, "fill", StringComparison.OrdinalIgnoreCase ) )
      {
        outValue = "fill";
        return true;
      }

      if( CssRegexPatterns.NumberOrPercent.IsMatch( token ) )
      {
        outValue = token;
        return true;
      }

      outValue = null;
      return false;
    }
  }

  internal class BorderImageRepeatParser : KeywordParser
  {
    internal BorderImageRepeatParser() : base( "stretch", "repeat", "round", "space" )
    {
    }
  }

  #endregion
}
