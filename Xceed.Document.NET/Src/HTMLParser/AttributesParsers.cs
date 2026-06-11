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
using System.Text;

namespace Xceed.Document.NET
{
  #region Shorthand Definition Model

  internal class ShorthandComponent
  {
    internal string PropertyName { get; set; }
    internal List<ICssValueParser> Parsers { get; set; }
    internal bool AllowMultiple { get; set; }
    internal bool ClaimToken { get; set; }

    internal ShorthandComponent()
    {
      Parsers = new List<ICssValueParser>();
      AllowMultiple = false;
      ClaimToken = true;
    }
  }

  internal class CssShorthandDefinition
  {
    internal string Name { get; set; }
    internal Dictionary<string, string> DefaultValues { get; set; }
    internal List<ShorthandComponent> Components { get; set; }

    internal CssShorthandDefinition()
    {
      DefaultValues = new Dictionary<string, string>( StringComparer.OrdinalIgnoreCase );
      Components = new List<ShorthandComponent>();
    }
  }

  #endregion

  #region Parser Engine

  internal static class CssShorthandParser
  {
    internal static Dictionary<string, string> Parse( string shorthandValue, CssShorthandDefinition definition )
    {
      var result = new Dictionary<string, string>( StringComparer.OrdinalIgnoreCase );

      // Initialize with defaults
      foreach( var kv in definition.DefaultValues )
      {
        result[ kv.Key ] = kv.Value;
      }

      if( !string.IsNullOrWhiteSpace( shorthandValue ) )
      {
        var tokens = Tokenize( shorthandValue );
        var collected = new Dictionary<ShorthandComponent, List<string>>();

        foreach( var comp in definition.Components )
        {
          collected[ comp ] = new List<string>();
        }

        foreach( var token in tokens )
        {
          bool tokenClaimed = false;

          foreach( var comp in definition.Components )
          {
            foreach( var parser in comp.Parsers )
            {
              string parsed;
              if( parser.TryParse( token, out parsed ) )
              {
                collected[ comp ].Add( parsed );
                tokenClaimed = comp.ClaimToken;
                break; // Move to next token
              }
            }

            if( tokenClaimed )
            {
              break;
            }
          }
        }

        // Materialize collected values
        foreach( var comp in definition.Components )
        {
          var list = collected[ comp ];
          if( list.Count != 0 )
          {
            result[ comp.PropertyName ] = comp.AllowMultiple ? string.Join( " ", list ) : list[ list.Count - 1 ];
          }
        }

        return result;
      }

      return result;
    }

    internal static List<string> Tokenize( string input )
    {
      var tokens = new List<string>();
      if( string.IsNullOrEmpty( input ) )
        return tokens;

      int i = 0;
      int n = input.Length;

      while( i < n )
      {
        // Skip whitespace
        while( i < n && char.IsWhiteSpace( input[ i ] ) )
          i++;
        if( i >= n )
          break;

        char c = input[ i ];

        // Handle quoted strings
        if( c == '"' || c == '\'' )
        {
          int start = i++;
          while( i < n && input[ i ] != c )
            i++;
          if( i < n )
            i++;
          tokens.Add( input.Substring( start, i - start ) );
        }
        // Handle comma separator
        else if( c == ',' )
        {
          tokens.Add( "," );
          i++;
        }
        // Handle regular tokens and function calls
        else
        {
          int start = i;
          while( i < n && !char.IsWhiteSpace( input[ i ] ) && input[ i ] != ',' )
          {
            if( input[ i ] == '(' )
            {
              // Capture entire function including nested parentheses
              int parenDepth = 1;
              i++;
              while( i < n && parenDepth > 0 )
              {
                if( input[ i ] == '(' )
                  parenDepth++;
                else if( input[ i ] == ')' )
                  parenDepth--;
                i++;
              }
              break;
            }
            i++;
          }
          tokens.Add( input.Substring( start, i - start ) );
        }
      }

      return tokens;
    }

    internal static Dictionary<string, string> ParseFont( string shorthandValue )
    {
      var result = new Dictionary<string, string>( StringComparer.OrdinalIgnoreCase )
      {
        [ "font-style" ] = "normal",
        [ "font-variant" ] = "normal",
        [ "font-weight" ] = "normal",
        [ "font-stretch" ] = "normal",
        [ "font-size" ] = "medium",
        [ "line-height" ] = "normal",
        [ "font-family" ] = ""
      };

      if( !string.IsNullOrWhiteSpace( shorthandValue ) )
      {
        var tokens = Tokenize( shorthandValue );
        if( tokens.Count == 0 )
        {
          return result;
        }

        // Check for system font
        var systemFontParser = new KeywordParser( "caption", "icon", "menu", "message-box", "small-caption", "status-bar" );

        if( tokens.Count == 1 )
        {
          string parsed;
          if( systemFontParser.TryParse( tokens[ 0 ], out parsed ) )
          {
            result[ "system-font" ] = parsed;
            return result;
          }
        }

        var styleParser = new KeywordParser( "normal", "italic", "oblique" );
        var variantParser = new KeywordParser( "normal", "small-caps" );
        var weightParser = new MultiParser( new KeywordParser( "normal", "bold", "bolder", "lighter" ), new NumericRangeParser( 1, 1000 ) );
        var stretchParser = new KeywordParser( "normal", "ultra-condensed", "extra-condensed", "condensed", "semi-condensed", "semi-expanded", "expanded", "extra-expanded", "ultra-expanded" );
        var sizeParser = new FontSizeParser();
        var lineHeightParser = new LineHeightParser();

        int tokenIndex = 0;
        var setFlags = new HashSet<string>();

        // Parse optional style, variant, weight, stretch
        while( tokenIndex < tokens.Count )
        {
          string token = tokens[ tokenIndex ];
          string parsed;
          bool matched = false;

          if( !setFlags.Contains( "style" ) && styleParser.TryParse( token, out parsed ) )
          {
            result[ "font-style" ] = parsed;
            setFlags.Add( "style" );
            matched = true;
          }
          else if( !setFlags.Contains( "variant" ) && variantParser.TryParse( token, out parsed ) )
          {
            result[ "font-variant" ] = parsed;
            setFlags.Add( "variant" );
            matched = true;
          }
          else if( !setFlags.Contains( "weight" ) && weightParser.TryParse( token, out parsed ) )
          {
            result[ "font-weight" ] = parsed;
            setFlags.Add( "weight" );
            matched = true;
          }
          else if( !setFlags.Contains( "stretch" ) && stretchParser.TryParse( token, out parsed ) )
          {
            result[ "font-stretch" ] = parsed;
            setFlags.Add( "stretch" );
            matched = true;
          }

          if( !matched )
            break;
          tokenIndex++;
        }

        // Parse required font-size (with optional /line-height)
        if( tokenIndex >= tokens.Count )
          return result;

        string sizeToken = tokens[ tokenIndex ];

        if( sizeToken.Contains( "/" ) )
        {
          var parts = sizeToken.Split( new[] { '/' }, 2 );
          string size, lineHeight;
          if( parts.Length == 2 && sizeParser.TryParse( parts[ 0 ], out size )
              && lineHeightParser.TryParse( parts[ 1 ], out lineHeight ) )
          {
            result[ "font-size" ] = size;
            result[ "line-height" ] = lineHeight;
            tokenIndex++;
          }
        }
        else
        {
          string size;
          if( sizeParser.TryParse( sizeToken, out size ) )
          {
            result[ "font-size" ] = size;
            tokenIndex++;

            // Check for /line-height
            if( tokenIndex < tokens.Count
                && tokens[ tokenIndex ].StartsWith( "/" ) )
            {
              string lhToken = tokens[ tokenIndex ].Substring( 1 );
              string lineHeight;
              if( lineHeightParser.TryParse( lhToken, out lineHeight ) )
              {
                result[ "line-height" ] = lineHeight;
                tokenIndex++;
              }
            }
          }
        }

        // Parse required font-family (remaining tokens)
        if( tokenIndex < tokens.Count )
        {
          var familyTokens = new List<string>();
          for( int i = tokenIndex; i < tokens.Count; i++ )
            familyTokens.Add( tokens[ i ] );
          result[ "font-family" ] = string.Join( " ", familyTokens );
        }

        return result;
      }

      return result;
    }

    internal static Dictionary<string, string> ParseBackground( string shorthandValue )
    {
      var result = new Dictionary<string, string>( StringComparer.OrdinalIgnoreCase )
      {
        [ "background-image" ] = "none",
        [ "background-position" ] = "0% 0%",
        [ "background-size" ] = "auto",
        [ "background-repeat" ] = "repeat",
        [ "background-origin" ] = "padding-box",
        [ "background-clip" ] = "border-box",
        [ "background-attachment" ] = "scroll",
        [ "background-color" ] = "transparent"
      };

      if( string.IsNullOrWhiteSpace( shorthandValue ) )
        return result;

      var layers = SplitByCommasOutsideParens( shorthandValue );
      if( layers.Count == 0 )
        return result;

      var allImages = new List<string>();
      var allPositions = new List<string>();
      var allSizes = new List<string>();
      var allRepeats = new List<string>();
      var allAttachments = new List<string>();
      var allOrigins = new List<string>();
      var allClips = new List<string>();

      for( int layerIdx = 0; layerIdx < layers.Count; layerIdx++ )
      {
        bool isLastLayer = ( layerIdx == layers.Count - 1 );
        var layerResult = ParseBackgroundLayer( layers[ layerIdx ], isLastLayer );

        if( layerResult.ContainsKey( "background-image" ) )
          allImages.Add( layerResult[ "background-image" ] );
        if( layerResult.ContainsKey( "background-position" ) )
          allPositions.Add( layerResult[ "background-position" ] );
        if( layerResult.ContainsKey( "background-size" ) )
          allSizes.Add( layerResult[ "background-size" ] );
        if( layerResult.ContainsKey( "background-repeat" ) )
          allRepeats.Add( layerResult[ "background-repeat" ] );
        if( layerResult.ContainsKey( "background-attachment" ) )
          allAttachments.Add( layerResult[ "background-attachment" ] );
        if( layerResult.ContainsKey( "background-origin" ) )
          allOrigins.Add( layerResult[ "background-origin" ] );
        if( layerResult.ContainsKey( "background-clip" ) )
          allClips.Add( layerResult[ "background-clip" ] );

        if( isLastLayer && layerResult.ContainsKey( "background-color" ) )
          result[ "background-color" ] = layerResult[ "background-color" ];
      }

      if( allImages.Count > 0 )
        result[ "background-image" ] = string.Join( ", ", allImages );
      if( allPositions.Count > 0 )
        result[ "background-position" ] = string.Join( ", ", allPositions );
      if( allSizes.Count > 0 )
        result[ "background-size" ] = string.Join( ", ", allSizes );
      if( allRepeats.Count > 0 )
        result[ "background-repeat" ] = string.Join( ", ", allRepeats );
      if( allAttachments.Count > 0 )
        result[ "background-attachment" ] = string.Join( ", ", allAttachments );
      if( allOrigins.Count > 0 )
        result[ "background-origin" ] = string.Join( ", ", allOrigins );
      if( allClips.Count > 0 )
        result[ "background-clip" ] = string.Join( ", ", allClips );

      return result;
    }

    private static Dictionary<string, string> ParseBackgroundLayer( string layer, bool isLastLayer )
    {
      var result = new Dictionary<string, string>( StringComparer.OrdinalIgnoreCase );

      if( !string.IsNullOrWhiteSpace( layer ) )
      {
        var tokens = Tokenize( layer );
        if( tokens.Count == 0 )
          return result;

        var imageParser = new ImageParser();
        var positionParser = new BackgroundPositionParser();
        var sizeParser = new BackgroundSizeParser();
        var repeatParser = new BackgroundRepeatParser();
        var attachmentParser = new BackgroundAttachmentParser();
        var boxParser = new VisualBoxParser();
        var colorParser = new CssColorParser();

        var positionTokens = new List<string>();
        var sizeTokens = new List<string>();
        var repeatTokens = new List<string>();
        string firstVisualBox = null;
        string secondVisualBox = null;
        int visualBoxCount = 0;

        for( int i = 0; i < tokens.Count; i++ )
        {
          string token = tokens[ i ];
          string parsed;

          if( imageParser.TryParse( token, out parsed ) )
          {
            result[ "background-image" ] = parsed;
          }
          else if( repeatParser.TryParse( token, out parsed ) )
          {
            repeatTokens.Add( parsed );
            // Check for second repeat value
            if( i + 1 < tokens.Count
              && !token.Equals( "repeat-x", StringComparison.OrdinalIgnoreCase )
              && !token.Equals( "repeat-y", StringComparison.OrdinalIgnoreCase )
              && repeatParser.TryParse( tokens[ i + 1 ], out parsed ) )
            {
              repeatTokens.Add( parsed );
              i++;
            }
          }
          else if( attachmentParser.TryParse( token, out parsed ) )
          {
            result[ "background-attachment" ] = parsed;
          }
          else if( boxParser.TryParse( token, out parsed ) )
          {
            visualBoxCount++;
            if( visualBoxCount == 1 )
            {
              firstVisualBox = parsed;
            }
            else if( visualBoxCount == 2 )
            {
              secondVisualBox = parsed;
            }
          }
          else if( isLastLayer && colorParser.TryParse( token, out parsed ) )
          {
            result[ "background-color" ] = parsed;
          }
          else if( positionParser.TryParse( token, out parsed ) )
          {
            positionTokens.Add( parsed );

            // Check for second position value
            if( i + 1 < tokens.Count
                && positionParser.TryParse( tokens[ i + 1 ], out parsed ) )
            {
              positionTokens.Add( parsed );
              i++;
            }

            // Check for / followed by size
            if( i + 1 < tokens.Count && tokens[ i + 1 ] == "/" )
            {
              i++; // skip /
              if( i + 1 < tokens.Count
                  && sizeParser.TryParse( tokens[ i + 1 ], out parsed ) )
              {
                sizeTokens.Add( parsed );
                i++;
                if( i + 1 < tokens.Count
                    && sizeParser.TryParse( tokens[ i + 1 ], out parsed ) )
                {
                  sizeTokens.Add( parsed );
                  i++;
                }
              }
            }
          }
        }

        if( positionTokens.Count > 0 )
          result[ "background-position" ] = string.Join( " ", positionTokens );
        if( sizeTokens.Count > 0 )
          result[ "background-size" ] = string.Join( " ", sizeTokens );
        if( repeatTokens.Count > 0 )
          result[ "background-repeat" ] = string.Join( " ", repeatTokens );

        // Handle visual box values
        if( visualBoxCount == 1 )
        {
          result[ "background-origin" ] = firstVisualBox;
          result[ "background-clip" ] = firstVisualBox;
        }
        else if( visualBoxCount == 2 )
        {
          result[ "background-origin" ] = firstVisualBox;
          result[ "background-clip" ] = secondVisualBox;
        }

        return result;
      }

      return result;
    }

    private static List<string> SplitByCommasOutsideParens( string input )
    {
      var layers = new List<string>();
      if( string.IsNullOrEmpty( input ) )
        return layers;

      var currentLayer = new StringBuilder();
      int parenDepth = 0;

      for( int i = 0; i < input.Length; i++ )
      {
        char c = input[ i ];

        if( c == '(' )
          parenDepth++;
        else if( c == ')' )
          parenDepth--;
        else if( c == ',' && parenDepth == 0 )
        {
          string layer = currentLayer.ToString().Trim();
          if( !string.IsNullOrEmpty( layer ) )
          {
            layers.Add( layer );
          }

          currentLayer.Clear();
          continue;
        }

        currentLayer.Append( c );
      }

      string finalLayer = currentLayer.ToString().Trim();
      if( !string.IsNullOrEmpty( finalLayer ) )
      {
        layers.Add( finalLayer );
      }

      return layers;
    }
  }

  #endregion

  #region Utilities &Examples

  #region CSS Helpers

  internal static class CssHelpers
  {
    internal static void ExpandFourSides( Dictionary<string, string> result, string collectKey,
                     string topKey, string rightKey, string bottomKey,
                     string leftKey )
    {
      string raw;
      if( !result.TryGetValue( collectKey, out raw ) )
        return;

      var tokens = raw.Split( new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries );
      string top, right, bottom, left;

      switch( tokens.Length )
      {
        case 1:
          top = right = bottom = left = tokens[ 0 ];
          break;
        case 2:
          top = bottom = tokens[ 0 ];
          right = left = tokens[ 1 ];
          break;
        case 3:
          top = tokens[ 0 ];
          right = left = tokens[ 1 ];
          bottom = tokens[ 2 ];
          break;
        default:
          top = tokens[ 0 ];
          right = tokens[ 1 ];
          bottom = tokens[ 2 ];
          left = tokens.Length > 3 ? tokens[ 3 ] : tokens[ 1 ];
          break;
      }

      result[ topKey ] = top;
      result[ rightKey ] = right;
      result[ bottomKey ] = bottom;
      result[ leftKey ] = left;

      result.Remove( collectKey );
    }

    internal static void ExpandTwoSides( Dictionary<string, string> result, string collectKey, string startProp, string endProp )
    {
      string raw;
      if( !result.TryGetValue( collectKey, out raw ) )
        return;

      var tokens = raw.Split( new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries );
      result[ startProp ] = tokens[ 0 ];
      result[ endProp ] = tokens.Length > 1 ? tokens[ 1 ] : tokens[ 0 ];
      result.Remove( collectKey );
    }

    // Convenience methods for specific properties
    internal static void ExpandBorderWidth( Dictionary<string, string> result )
    {
      CssHelpers.ExpandFourSides( result, "border-width-collect", "border-top-width", "border-right-width", "border-bottom-width", "border-left-width" );
    }

    internal static void ExpandBorderStyle( Dictionary<string, string> result )
    {
      CssHelpers.ExpandFourSides( result, "border-style-collect", "border-top-style", "border-right-style", "border-bottom-style", "border-left-style" );
    }

    internal static void ExpandBorderColor( Dictionary<string, string> result )
    {
      CssHelpers.ExpandFourSides( result, "border-color-collect", "border-top-color", "border-right-color", "border-bottom-color", "border-left-color" );
    }

    internal static void ExpandBorderRadius( Dictionary<string, string> result )
    {
      CssHelpers.ExpandFourSides( result, "border-radius-collect", "border-top-left-radius", "border-top-right-radius", "border-bottom-right-radius", "border-bottom-left-radius" );
    }

    internal static void ExpandMargin( Dictionary<string, string> result )
    {
      CssHelpers.ExpandFourSides( result, "margin-collect", "margin-top", "margin-right", "margin-bottom", "margin-left" );
    }

    internal static void ExpandPadding( Dictionary<string, string> result )
    {
      CssHelpers.ExpandFourSides( result, "padding-collect", "padding-top", "padding-right", "padding-bottom", "padding-left" );
    }

    internal static void ExpandBorderBlockWidth( Dictionary<string, string> result )
    {
      CssHelpers.ExpandTwoSides( result, "border-block-width-collect", "border-block-start-width", "border-block-end-width" );
    }

    internal static void ExpandBorderBlockStyle( Dictionary<string, string> result )
    {
      CssHelpers.ExpandTwoSides( result, "border-block-style-collect", "border-block-start-style", "border-block-end-style" );
    }

    internal static void ExpandBorderBlockColor( Dictionary<string, string> result )
    {
      CssHelpers.ExpandTwoSides( result, "border-block-color-collect", "border-block-start-color", "border-block-end-color" );
    }

    internal static void ExpandBorderInlineWidth( Dictionary<string, string> result )
    {
      CssHelpers.ExpandTwoSides( result, "border-inline-width-collect", "border-inline-start-width", "border-inline-end-width" );
    }

    internal static void ExpandBorderInlineStyle( Dictionary<string, string> result )
    {
      CssHelpers.ExpandTwoSides( result, "border-inline-style-collect", "border-inline-start-style", "border-inline-end-style" );
    }

    internal static void ExpandBorderInlineColor( Dictionary<string, string> result )
    {
      CssHelpers.ExpandTwoSides( result, "border-inline-color-collect", "border-inline-start-color", "border-inline-end-color" );
    }

    internal static void ExpandBorderSpacing( Dictionary<string, string> result )
    {
      CssHelpers.ExpandTwoSides( result, "border-spacing-collect", "border-spacing-horizontal", "border-spacing-vertical" );
    }

    internal static bool IsNumericValue( string token )
    {
      if( !string.IsNullOrWhiteSpace( token ) )
      {
        double tmp;

        return double.TryParse( token, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out tmp );
      }

      return false;
    }
  }

  #endregion
  // Example usage
  internal class Example
  {
    internal static void Run()
    {
      var td = "underline dotted #00f 2px";
      var parsed = CssShorthandParser.Parse( td, CssDefinitions.TextDecoration );
      Console.WriteLine( "text-decoration parsed:" );
      foreach( var kv in parsed )
        Console.WriteLine( "  {0}: {1}", kv.Key, kv.Value );
      Console.WriteLine();

      var border = "2px solid red";
      var bParsed = CssShorthandParser.Parse( border, CssDefinitions.Border );
      Console.WriteLine( "border parsed:" );
      foreach( var kv in bParsed )
        Console.WriteLine( "  {0}: {1}", kv.Key, kv.Value );
      Console.WriteLine();

      var br = "1em 2em";
      var brParsed = CssShorthandParser.Parse( br, CssDefinitions.BorderRadius );
      // the parser will put collected tokens into border-radius-collect key
      CssHelpers.ExpandBorderRadius( brParsed );
      Console.WriteLine( "border-radius parsed:" );
      foreach( var kv in brParsed )
        Console.WriteLine( "  {0}: {1}", kv.Key, kv.Value );
      Console.WriteLine();

      var marginValue = "10px 20px";
      var marginParsed = CssShorthandParser.Parse( marginValue, CssDefinitions.Margin );
      CssHelpers.ExpandMargin( marginParsed );
      Console.WriteLine( "margin parsed:" );
      foreach( var kv in marginParsed )
        Console.WriteLine( "  {0}: {1}", kv.Key, kv.Value );
      Console.WriteLine();

      var paddingValue = "5px 10px 15px";
      var paddingParsed = CssShorthandParser.Parse( paddingValue, CssDefinitions.Padding );
      CssHelpers.ExpandPadding( paddingParsed );
      Console.WriteLine( "padding parsed:" );
      foreach( var kv in paddingParsed )
        Console.WriteLine( "  {0}: {1}", kv.Key, kv.Value );
      Console.WriteLine();

      var listStyleValue = "square inside url('bullet.png')";
      var listStyleParsed = CssShorthandParser.Parse( listStyleValue, CssDefinitions.ListStyle );
      Console.WriteLine( "list-style parsed:" );
      foreach( var kv in listStyleParsed )
        Console.WriteLine( "  {0}: {1}", kv.Key, kv.Value );
      Console.WriteLine();

      var outlineValue = "thick double #32a1ce";
      var outlineParser = CssShorthandParser.Parse( outlineValue, CssDefinitions.Outline );
      Console.WriteLine( "outline parsed:" );
      foreach( var kv in outlineParser )
        Console.WriteLine( "  {0}: {1}", kv.Key, kv.Value );
      Console.WriteLine();

      var outlineValue2 = "8px ridge rgb(170 50 220 / 0.6)";
      var outlineParser2 = CssShorthandParser.Parse( outlineValue2, CssDefinitions.Outline );
      Console.WriteLine( "outline parsed:" );
      foreach( var kv in outlineParser2 )
        Console.WriteLine( "  {0}: {1}", kv.Key, kv.Value );
      Console.WriteLine();

      parsed = CssShorthandParser.Parse( "2px solid red", CssDefinitions.Border );
      // border-width: 2px, border-style: solid, border-color: red
      Console.WriteLine( "border parsed:" );
      foreach( var kv in parsed )
        Console.WriteLine( "  {0}: {1}", kv.Key, kv.Value );
      Console.WriteLine();

      // border-width (1-4 values)
      parsed = CssShorthandParser.Parse( "1px 2px 3px 4px", CssDefinitions.BorderWidth );
      CssHelpers.ExpandBorderWidth( parsed );
      // border-top-width: 1px, border-right-width: 2px, etc.
      Console.WriteLine( "border-width parsed:" );
      foreach( var kv in parsed )
        Console.WriteLine( "  {0}: {1}", kv.Key, kv.Value );
      Console.WriteLine();

      // border-style (2 values)
      parsed = CssShorthandParser.Parse( "solid dashed", CssDefinitions.BorderStyle );
      CssHelpers.ExpandBorderStyle( parsed );
      // border-top-style: solid, border-right-style: dashed, etc.
      Console.WriteLine( "border-style parsed:" );
      foreach( var kv in parsed )
        Console.WriteLine( "  {0}: {1}", kv.Key, kv.Value );
      Console.WriteLine();

      // border-radius
      parsed = CssShorthandParser.Parse( "10px 20px", CssDefinitions.BorderRadius );
      CssHelpers.ExpandBorderRadius( parsed );
      // border-top-left-radius: 10px, border-top-right-radius: 20px, etc.
      Console.WriteLine( "border-radius parsed:" );
      foreach( var kv in parsed )
        Console.WriteLine( "  {0}: {1}", kv.Key, kv.Value );
      Console.WriteLine();

      // border-block
      parsed = CssShorthandParser.Parse( "1px solid blue",
                                         CssDefinitions.BorderBlock );
      Console.WriteLine( "border-block parsed:" );
      foreach( var kv in parsed )
        Console.WriteLine( "  {0}: {1}", kv.Key, kv.Value );
      Console.WriteLine();

      // border-block-width
      parsed = CssShorthandParser.Parse( "1px 2px", CssDefinitions.BorderBlockWidth );
      CssHelpers.ExpandBorderBlockWidth( parsed );
      // border-block-start-width: 1px, border-block-end-width: 2px
      Console.WriteLine( "border-block-width parsed:" );
      foreach( var kv in parsed )
        Console.WriteLine( "  {0}: {1}", kv.Key, kv.Value );
      Console.WriteLine();

      // border-image
      parsed = CssShorthandParser.Parse( "url('border.png') 30 round", CssDefinitions.BorderImage );
      Console.WriteLine( "border-image parsed:" );
      foreach( var kv in parsed )
        Console.WriteLine( "  {0}: {1}", kv.Key, kv.Value );
      Console.WriteLine();

      // border-spacing
      parsed = CssShorthandParser.Parse( "5px 10px", CssDefinitions.BorderSpacing );
      CssHelpers.ExpandBorderSpacing( parsed );
      // border-spacing-horizontal: 5px, border-spacing-vertical: 10px
      Console.WriteLine( "border-spacing parsed:" );
      foreach( var kv in parsed )
        Console.WriteLine( "  {0}: {1}", kv.Key, kv.Value );
      Console.WriteLine();
    }

    internal static void RunExamples()
    {
      var examples = new[]
      {
        "1.2em \"Fira Sans\", sans-serif",
        "italic 1.2rem \"Fira Sans\", serif",
        "italic small-caps bold 16px/2 cursive",
        "small-caps bold 24px/1 sans-serif",
        "caption",
        "12px/14px sans-serif",
        "80% sans-serif",
        "bold italic large serif",
        "status-bar",
        "ultra-condensed small-caps 1.2em \"Fira Sans\", sans-serif"
     };

      foreach( var example in examples )
      {
        Console.WriteLine( $"\nParsing: {example}" );
        var parsed = CssShorthandParser.ParseFont( example );
        foreach( var kv in parsed )
        {
          if( !string.IsNullOrEmpty( kv.Value ) )
          {
            Console.WriteLine( $"  {kv.Key}: {kv.Value}" );
          }
        }
      }
    }

    internal static void RunBackgroundExamples()
    {
      var examples = new[]
      {
        "green",
        "content-box radial-gradient(crimson, skyblue)",
        "no-repeat url(\"/images/lizard.png\")",
        "left 5% / 15% 60% repeat-x url(\"/images/star.png\")",
        "center / contain no-repeat url(\"/images/firefox-logo.svg\"), "
            + "#eeeeee 35% url(\"/images/lizard.png\")",
        "url(\"test.jpg\") repeat-y",
        "border-box red",
        "no-repeat center/80% url(\"../img/image.png\")",
        "fixed top left / cover no-repeat url(\"bg.jpg\")",
        "padding-box content-box repeat url(\"pattern.png\")"
      };

      foreach( var example in examples )
      {
        Console.WriteLine( $"\nParsing background: {example}" );
        var parsed = CssShorthandParser.ParseBackground( example );
        foreach( var kv in parsed )
        {
          Console.WriteLine( $"  {kv.Key}: {kv.Value}" );
        }
      }
    }
  }

  #endregion
}
