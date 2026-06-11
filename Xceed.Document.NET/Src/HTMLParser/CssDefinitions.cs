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
using System.Collections.Generic;

namespace Xceed.Document.NET
{
  internal static class CssDefinitions
  {
    // Shared parser instances
    private static readonly CssColorParser ColorParser = new CssColorParser();
    private static readonly LengthOrKeywordParser BorderWidthParser = new LengthOrKeywordParser();
    private static readonly KeywordParser BorderStyleParser = new KeywordParser( "none", "hidden", "dotted", "dashed", "solid", "double", "groove", "ridge", "inset", "outset" );

    #region Factory Methods

    private static CssShorthandDefinition CreateBorderDefinition( string name, string widthProp, string styleProp, string colorProp )
    {
      var def = new CssShorthandDefinition { Name = name };
      def.DefaultValues[ widthProp ] = "medium";
      def.DefaultValues[ styleProp ] = "none";
      def.DefaultValues[ colorProp ] = "currentcolor";

      def.Components.Add( new ShorthandComponent { PropertyName = widthProp, Parsers = new List<ICssValueParser> { BorderWidthParser }, AllowMultiple = false } );

      def.Components.Add( new ShorthandComponent { PropertyName = styleProp, Parsers = new List<ICssValueParser> { BorderStyleParser }, AllowMultiple = false } );

      def.Components.Add( new ShorthandComponent { PropertyName = colorProp, Parsers = new List<ICssValueParser> { ColorParser }, AllowMultiple = false } );

      return def;
    }

    private static CssShorthandDefinition CreateFourSideDefinition( string name, string topProp, string rightProp, string bottomProp, string leftProp, string defaultValue, ICssValueParser parser )
    {
      var def = new CssShorthandDefinition { Name = name };
      def.DefaultValues[ topProp ] = defaultValue;
      def.DefaultValues[ rightProp ] = defaultValue;
      def.DefaultValues[ bottomProp ] = defaultValue;
      def.DefaultValues[ leftProp ] = defaultValue;

      def.Components.Add( new ShorthandComponent { PropertyName = name + "-collect", Parsers = new List<ICssValueParser> { parser }, AllowMultiple = true, ClaimToken = true } );

      return def;
    }

    private static CssShorthandDefinition CreateTwoSideDefinition( string name, string startProp, string endProp, string defaultValue, ICssValueParser parser )
    {
      var def = new CssShorthandDefinition { Name = name };
      def.DefaultValues[ startProp ] = defaultValue;
      def.DefaultValues[ endProp ] = defaultValue;

      def.Components.Add( new ShorthandComponent { PropertyName = name + "-collect", Parsers = new List<ICssValueParser> { parser }, AllowMultiple = true, ClaimToken = true } );

      return def;
    }

    #endregion

    #region Border Definitions

    internal static CssShorthandDefinition Border
    {
      get
      {
        return CssDefinitions.CreateBorderDefinition(
          "border",
          "border-width",
          "border-style",
          "border-color" );
      }
    }

    internal static CssShorthandDefinition BorderTop
    {
      get
      {
        return CssDefinitions.CreateBorderDefinition(
          "border-top",
          "border-top-width",
          "border-top-style",
          "border-top-color" );
      }
    }

    internal static CssShorthandDefinition BorderRight
    {
      get
      {
        return CssDefinitions.CreateBorderDefinition(
          "border-right",
          "border-right-width",
          "border-right-style",
          "border-right-color" );
      }
    }

    internal static CssShorthandDefinition BorderBottom
    {
      get
      {
        return CssDefinitions.CreateBorderDefinition(
          "border-bottom",
          "border-bottom-width",
          "border-bottom-style",
          "border-bottom-color" );
      }
    }

    internal static CssShorthandDefinition BorderLeft
    {
      get
      {
        return CssDefinitions.CreateBorderDefinition(
          "border-left",
          "border-left-width",
          "border-left-style",
          "border-left-color" );
      }
    }

    internal static CssShorthandDefinition BorderBlock
    {
      get
      {
        return CssDefinitions.CreateBorderDefinition(
          "border-block",
          "border-block-width",
          "border-block-style",
          "border-block-color" );
      }
    }

    internal static CssShorthandDefinition BorderBlockStart
    {
      get
      {
        return CssDefinitions.CreateBorderDefinition(
          "border-block-start",
          "border-block-start-width",
          "border-block-start-style",
          "border-block-start-color" );
      }
    }

    internal static CssShorthandDefinition BorderBlockEnd
    {
      get
      {
        return CssDefinitions.CreateBorderDefinition(
          "border-block-end",
          "border-block-end-width",
          "border-block-end-style",
          "border-block-end-color" );
      }
    }

    internal static CssShorthandDefinition BorderInline
    {
      get
      {
        return CssDefinitions.CreateBorderDefinition(
          "border-inline",
          "border-inline-width",
          "border-inline-style",
          "border-inline-color" );
      }
    }

    internal static CssShorthandDefinition BorderInlineStart
    {
      get
      {
        return CssDefinitions.CreateBorderDefinition(
          "border-inline-start",
          "border-inline-start-width",
          "border-inline-start-style",
          "border-inline-start-color" );
      }
    }

    internal static CssShorthandDefinition BorderInlineEnd
    {
      get
      {
        return CssDefinitions.CreateBorderDefinition(
          "border-inline-end",
          "border-inline-end-width",
          "border-inline-end-style",
          "border-inline-end-color" );
      }
    }

    #endregion

    #region Border Width / Style / Color Definitions

    internal static CssShorthandDefinition BorderWidth
    {
      get
      {
        return CssDefinitions.CreateFourSideDefinition(
          "border-width",
          "border-top-width",
          "border-right-width",
          "border-bottom-width",
          "border-left-width",
          "medium",
          BorderWidthParser );
      }
    }

    internal static CssShorthandDefinition BorderStyle
    {
      get
      {
        return CssDefinitions.CreateFourSideDefinition(
          "border-style",
          "border-top-style",
          "border-right-style",
          "border-bottom-style",
          "border-left-style",
          "none",
          BorderStyleParser );
      }
    }

    internal static CssShorthandDefinition BorderColor
    {
      get
      {
        return CssDefinitions.CreateFourSideDefinition(
          "border-color",
          "border-top-color",
          "border-right-color",
          "border-bottom-color",
          "border-left-color",
          "currentcolor",
          ColorParser );
      }
    }

    internal static CssShorthandDefinition BorderBlockWidth
    {
      get
      {
        return CssDefinitions.CreateTwoSideDefinition(
          "border-block-width",
          "border-block-start-width",
          "border-block-end-width",
          "medium",
          BorderWidthParser );
      }
    }

    internal static CssShorthandDefinition BorderBlockStyle
    {
      get
      {
        return CssDefinitions.CreateTwoSideDefinition(
          "border-block-style",
          "border-block-start-style",
          "border-block-end-style",
          "none",
          BorderStyleParser );
      }
    }

    internal static CssShorthandDefinition BorderBlockColor
    {
      get
      {
        return CssDefinitions.CreateTwoSideDefinition(
          "border-block-color",
          "border-block-start-color",
          "border-block-end-color",
          "currentcolor",
          ColorParser );
      }
    }

    internal static CssShorthandDefinition BorderInlineWidth
    {
      get
      {
        return CssDefinitions.CreateTwoSideDefinition(
          "border-inline-width",
          "border-inline-start-width",
          "border-inline-end-width",
          "medium",
          BorderWidthParser );
      }
    }

    internal static CssShorthandDefinition BorderInlineStyle
    {
      get
      {
        return CssDefinitions.CreateTwoSideDefinition(
          "border-inline-style",
          "border-inline-start-style",
          "border-inline-end-style",
          "none",
          BorderStyleParser );
      }
    }

    internal static CssShorthandDefinition BorderInlineColor
    {
      get
      {
        return CssDefinitions.CreateTwoSideDefinition(
          "border-inline-color",
          "border-inline-start-color",
          "border-inline-end-color",
          "currentcolor",
          ColorParser );
      }
    }

    #endregion


    #region Other Definitions

    internal static CssShorthandDefinition BorderRadius
    {
      get
      {
        var def = new CssShorthandDefinition { Name = "border-radius" };
        def.DefaultValues[ "border-top-left-radius" ] = "0";
        def.DefaultValues[ "border-top-right-radius" ] = "0";
        def.DefaultValues[ "border-bottom-right-radius" ] = "0";
        def.DefaultValues[ "border-bottom-left-radius" ] = "0";

        def.Components.Add( new ShorthandComponent { PropertyName = "border-radius-collect", Parsers = new List<ICssValueParser> { new BorderRadiusValueParser() }, AllowMultiple = true, ClaimToken = true } );

        return def;
      }
    }

    internal static CssShorthandDefinition BorderImage
    {
      get
      {
        var def = new CssShorthandDefinition { Name = "border-image" };
        def.DefaultValues[ "border-image-source" ] = "none";
        def.DefaultValues[ "border-image-slice" ] = "100%";
        def.DefaultValues[ "border-image-width" ] = "1";
        def.DefaultValues[ "border-image-outset" ] = "0";
        def.DefaultValues[ "border-image-repeat" ] = "stretch";

        def.Components.Add( new ShorthandComponent { PropertyName = "border-image-source", Parsers = new List<ICssValueParser> { new BorderImageSourceParser() }, AllowMultiple = false, ClaimToken = true } );

        def.Components.Add( new ShorthandComponent { PropertyName = "border-image-slice", Parsers = new List<ICssValueParser> { new BorderImageSliceParser() }, AllowMultiple = true, ClaimToken = true } );

        def.Components.Add( new ShorthandComponent { PropertyName = "border-image-repeat", Parsers = new List<ICssValueParser> { new BorderImageRepeatParser() }, AllowMultiple = true, ClaimToken = true } );

        return def;
      }
    }

    internal static CssShorthandDefinition BorderSpacing
    {
      get
      {
        return CreateTwoSideDefinition( "border-spacing", "border-spacing-horizontal", "border-spacing-vertical", "0", new LengthParser() );
      }
    }

    internal static CssShorthandDefinition Margin
    {
      get
      {
        var def = new CssShorthandDefinition { Name = "margin" };
        def.DefaultValues[ "margin-top" ] = "0";
        def.DefaultValues[ "margin-right" ] = "0";
        def.DefaultValues[ "margin-bottom" ] = "0";
        def.DefaultValues[ "margin-left" ] = "0";

        def.Components.Add( new ShorthandComponent { PropertyName = "margin-collect", Parsers = new List<ICssValueParser> { new LengthOrKeywordParser() }, AllowMultiple = true, ClaimToken = true } );

        return def;
      }
    }

    internal static CssShorthandDefinition Padding
    {
      get
      {
        var def = new CssShorthandDefinition { Name = "padding" };
        def.DefaultValues[ "padding-top" ] = "0";
        def.DefaultValues[ "padding-right" ] = "0";
        def.DefaultValues[ "padding-bottom" ] = "0";
        def.DefaultValues[ "padding-left" ] = "0";

        def.Components.Add( new ShorthandComponent { PropertyName = "padding-collect", Parsers = new List<ICssValueParser> { new LengthOrKeywordParser() }, AllowMultiple = true, ClaimToken = true } );

        return def;
      }
    }

    internal static CssShorthandDefinition TextDecoration
    {
      get
      {
        var def = new CssShorthandDefinition { Name = "text-decoration" };
        def.DefaultValues[ "text-decoration-line" ] = "none";
        def.DefaultValues[ "text-decoration-color" ] = "currentcolor";
        def.DefaultValues[ "text-decoration-style" ] = "solid";
        def.DefaultValues[ "text-decoration-thickness" ] = "auto";

        var lineParser = new KeywordParser( "underline", "overline", "line-through", "blink", "spelling-error", "grammar-error", "none" );
        var lineStyleParser = new KeywordParser( "solid", "double", "dotted", "dashed", "wavy" );
        var thicknessParser = new LengthOrKeywordParser();

        def.Components.Add( new ShorthandComponent { PropertyName = "text-decoration-line", Parsers = new List<ICssValueParser> { lineParser }, AllowMultiple = true, ClaimToken = true } );

        def.Components.Add( new ShorthandComponent { PropertyName = "text-decoration-color", Parsers = new List<ICssValueParser> { ColorParser }, AllowMultiple = false, ClaimToken = true } );

        def.Components.Add( new ShorthandComponent { PropertyName = "text-decoration-style", Parsers = new List<ICssValueParser> { lineStyleParser }, AllowMultiple = false, ClaimToken = true } );

        def.Components.Add( new ShorthandComponent { PropertyName = "text-decoration-thickness", Parsers = new List<ICssValueParser> { thicknessParser }, AllowMultiple = false, ClaimToken = true } );

        return def;
      }
    }

    internal static CssShorthandDefinition ListStyle
    {
      get
      {
        var def = new CssShorthandDefinition { Name = "list-style" };
        def.DefaultValues[ "list-style-type" ] = "disc";
        def.DefaultValues[ "list-style-position" ] = "outside";
        def.DefaultValues[ "list-style-image" ] = "none";

        var listTypeParser = new KeywordParser( "none", "disc", "circle", "square", "decimal", "decimal-leading-zero", "lower-roman", "upper-roman", "lower-greek", "lower-latin", "upper-latin", "armenian", "georgian", "lower-alpha", "upper-alpha" );
        var listPositionParser = new KeywordParser( "inside", "outside" );
        var listImageParser = new ImageParser();

        def.Components.Add( new ShorthandComponent { PropertyName = "list-style-type", Parsers = new List<ICssValueParser> { listTypeParser }, AllowMultiple = false, ClaimToken = true } );

        def.Components.Add( new ShorthandComponent { PropertyName = "list-style-position", Parsers = new List<ICssValueParser> { listPositionParser }, AllowMultiple = false, ClaimToken = true } );

        def.Components.Add( new ShorthandComponent { PropertyName = "list-style-image", Parsers = new List<ICssValueParser> { listImageParser }, AllowMultiple = false, ClaimToken = true } );

        return def;
      }
    }

    internal static CssShorthandDefinition Outline
    {
      get
      {
        var def = new CssShorthandDefinition { Name = "outline" };
        def.DefaultValues[ "outline-color" ] = "invert";
        def.DefaultValues[ "outline-style" ] = "none";
        def.DefaultValues[ "outline-width" ] = "medium";

        def.Components.Add( new ShorthandComponent { PropertyName = "outline-color", Parsers = new List<ICssValueParser> { ColorParser }, AllowMultiple = false } );

        def.Components.Add( new ShorthandComponent { PropertyName = "outline-style", Parsers = new List<ICssValueParser> { BorderStyleParser }, AllowMultiple = false } );

        def.Components.Add( new ShorthandComponent { PropertyName = "outline-width", Parsers = new List<ICssValueParser> { new LengthOrKeywordParser() }, AllowMultiple = false } );

        return def;
      }
    }

    internal static CssShorthandDefinition Font
    {
      get
      {
        var def = new CssShorthandDefinition { Name = "font" };
        def.DefaultValues[ "font-style" ] = "normal";
        def.DefaultValues[ "font-variant" ] = "normal";
        def.DefaultValues[ "font-weight" ] = "normal";
        def.DefaultValues[ "font-stretch" ] = "normal";
        def.DefaultValues[ "font-size" ] = "medium";
        def.DefaultValues[ "line-height" ] = "normal";
        def.DefaultValues[ "font-family" ] = "";

        // Note: Font shorthand is complex and uses ParseFont() method instead of generic parsing
        return def;
      }
    }

    #endregion
  }
}
