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
using System.ComponentModel;

namespace Xceed.Document.NET
{

  public enum ListItemType
  {
    Numbered,
    Bulleted
  }

  public enum SectionBreakType
  {
    defaultNextPage,
    evenPage,
    oddPage,
    continuous
  }

  public enum ContainerType
  {
    None,
    TOC,
    Section,
    Cell,
    Table,
    Header,
    Footer,
    Paragraph,
    Body,
    Shape
  }

  public enum ShadingType
  {
    Text,
    Paragraph
  }

  public enum PageNumberFormat
  {
    normal,
    roman
  }

  public enum ChapterSeperator
  {
    colon,
    emDash,
    enDash,
    hyphen,
    period
  }
  public enum BorderSize
  {
    one,
    two,
    three,
    four,
    five,
    six,
    seven,
    eight,
    nine
  }

  public enum EditRestrictions
  {
    none,
    readOnly,
    forms,
    comments,
    trackedChanges
  }

  public enum BorderStyle
  {
    Tcbs_none = 0,
    Tcbs_single,
    Tcbs_thick,
    Tcbs_double,
    Tcbs_dotted,
    Tcbs_dashed,
    Tcbs_dotDash,
    Tcbs_dotDotDash,
    Tcbs_triple,
    Tcbs_thinThickSmallGap,
    Tcbs_thickThinSmallGap,
    Tcbs_thinThickThinSmallGap,
    Tcbs_thinThickMediumGap,
    Tcbs_thickThinMediumGap,
    Tcbs_thinThickThinMediumGap,
    Tcbs_thinThickLargeGap,
    Tcbs_thickThinLargeGap,
    Tcbs_thinThickThinLargeGap,
    Tcbs_wave,
    Tcbs_doubleWave,
    Tcbs_dashSmallGap,
    Tcbs_dashDotStroked,
    Tcbs_threeDEmboss,
    Tcbs_threeDEngrave,
    Tcbs_outset,
    Tcbs_inset,
    Tcbs_nil
  }

  public enum TableCellBorderType
  {
    Top,
    Bottom,
    Left,
    Right,
    InsideH,
    InsideV,
    TopLeftToBottomRight,
    TopRightToBottomLeft
  }

  public enum TableBorderType
  {
    Top,
    Bottom,
    Left,
    Right,
    InsideH,
    InsideV
  }

  public enum VerticalAlignment
  {
    Top,
    Center,
    Bottom
  };

  public enum Orientation
  {
    Portrait,
    Landscape
  };

  public enum MatchFormattingOptions
  {
    ExactMatch,
    SubsetMatch
  };

  public enum Script
  {
    superscript,
    subscript,
    none
  }

  public enum Highlight
  {
    yellow,
    green,
    cyan,
    magenta,
    blue,
    red,
    darkBlue,
    darkCyan,
    darkGreen,
    darkMagenta,
    darkRed,
    darkYellow,
    darkGray,
    lightGray,
    black,
    none
  };

  public enum UnderlineStyle
  {
    none = 0,
    singleLine = 1,
    words = 2,
    doubleLine = 3,
    dotted = 4,
    thick = 6,
    dash = 7,
    dotDash = 9,
    dotDotDash = 10,
    wave = 11,
    dottedHeavy = 20,
    dashedHeavy = 23,
    dashDotHeavy = 25,
    dashDotDotHeavy = 26,
    dashLongHeavy = 27,
    dashLong = 39,
    wavyDouble = 43,
    wavyHeavy = 55
  };

  public enum StrikeThrough
  {
    none,
    strike,
    doubleStrike
  };

  public enum Misc
  {
    none,
    shadow,
    outline,
    outlineShadow,
    emboss,
    engrave
  };

  public enum CapsStyle
  {
    none,
    caps,
    smallCaps
  };

  public enum TableDesign
  {
    Custom,
    TableNormal,
    TableGrid,
    LightShading,
    LightShadingAccent1,
    LightShadingAccent2,
    LightShadingAccent3,
    LightShadingAccent4,
    LightShadingAccent5,
    LightShadingAccent6,
    LightList,
    LightListAccent1,
    LightListAccent2,
    LightListAccent3,
    LightListAccent4,
    LightListAccent5,
    LightListAccent6,
    LightGrid,
    LightGridAccent1,
    LightGridAccent2,
    LightGridAccent3,
    LightGridAccent4,
    LightGridAccent5,
    LightGridAccent6,
    MediumShading1,
    MediumShading1Accent1,
    MediumShading1Accent2,
    MediumShading1Accent3,
    MediumShading1Accent4,
    MediumShading1Accent5,
    MediumShading1Accent6,
    MediumShading2,
    MediumShading2Accent1,
    MediumShading2Accent2,
    MediumShading2Accent3,
    MediumShading2Accent4,
    MediumShading2Accent5,
    MediumShading2Accent6,
    MediumList1,
    MediumList1Accent1,
    MediumList1Accent2,
    MediumList1Accent3,
    MediumList1Accent4,
    MediumList1Accent5,
    MediumList1Accent6,
    MediumList2,
    MediumList2Accent1,
    MediumList2Accent2,
    MediumList2Accent3,
    MediumList2Accent4,
    MediumList2Accent5,
    MediumList2Accent6,
    MediumGrid1,
    MediumGrid1Accent1,
    MediumGrid1Accent2,
    MediumGrid1Accent3,
    MediumGrid1Accent4,
    MediumGrid1Accent5,
    MediumGrid1Accent6,
    MediumGrid2,
    MediumGrid2Accent1,
    MediumGrid2Accent2,
    MediumGrid2Accent3,
    MediumGrid2Accent4,
    MediumGrid2Accent5,
    MediumGrid2Accent6,
    MediumGrid3,
    MediumGrid3Accent1,
    MediumGrid3Accent2,
    MediumGrid3Accent3,
    MediumGrid3Accent4,
    MediumGrid3Accent5,
    MediumGrid3Accent6,
    DarkList,
    DarkListAccent1,
    DarkListAccent2,
    DarkListAccent3,
    DarkListAccent4,
    DarkListAccent5,
    DarkListAccent6,
    ColorfulShading,
    ColorfulShadingAccent1,
    ColorfulShadingAccent2,
    ColorfulShadingAccent3,
    ColorfulShadingAccent4,
    ColorfulShadingAccent5,
    ColorfulShadingAccent6,
    ColorfulList,
    ColorfulListAccent1,
    ColorfulListAccent2,
    ColorfulListAccent3,
    ColorfulListAccent4,
    ColorfulListAccent5,
    ColorfulListAccent6,
    ColorfulGrid,
    ColorfulGridAccent1,
    ColorfulGridAccent2,
    ColorfulGridAccent3,
    ColorfulGridAccent4,
    ColorfulGridAccent5,
    ColorfulGridAccent6,
    None
  };

  public enum AutoFit
  {
    Contents,
    Window,
    ColumnWidth,
    Fixed
  };

  public enum RectangleShapes
  {
    rect,
    roundRect,
    snip1Rect,
    snip2SameRect,
    snip2DiagRect,
    snipRoundRect,
    round1Rect,
    round2SameRect,
    round2DiagRect
  };

  public enum BasicShapes
  {
    ellipse,
    triangle,
    rtTriangle,
    parallelogram,
    trapezoid,
    diamond,
    pentagon,
    hexagon,
    heptagon,
    octagon,
    decagon,
    dodecagon,
    pie,
    chord,
    teardrop,
    frame,
    halfFrame,
    corner,
    diagStripe,
    plus,
    plaque,
    can,
    cube,
    bevel,
    donut,
    noSmoking,
    blockArc,
    foldedCorner,
    smileyFace,
    heart,
    lightningBolt,
    sun,
    moon,
    cloud,
    arc,
    backetPair,
    bracePair,
    leftBracket,
    rightBracket,
    leftBrace,
    rightBrace
  };

  public enum BlockArrowShapes
  {
    rightArrow,
    leftArrow,
    upArrow,
    downArrow,
    leftRightArrow,
    upDownArrow,
    quadArrow,
    leftRightUpArrow,
    bentArrow,
    uturnArrow,
    leftUpArrow,
    bentUpArrow,
    curvedRightArrow,
    curvedLeftArrow,
    curvedUpArrow,
    curvedDownArrow,
    stripedRightArrow,
    notchedRightArrow,
    homePlate,
    chevron,
    rightArrowCallout,
    downArrowCallout,
    leftArrowCallout,
    upArrowCallout,
    leftRightArrowCallout,
    quadArrowCallout,
    circularArrow
  };

  public enum EquationShapes
  {
    mathPlus,
    mathMinus,
    mathMultiply,
    mathDivide,
    mathEqual,
    mathNotEqual
  };

  public enum FlowchartShapes
  {
    flowChartProcess,
    flowChartAlternateProcess,
    flowChartDecision,
    flowChartInputOutput,
    flowChartPredefinedProcess,
    flowChartInternalStorage,
    flowChartDocument,
    flowChartMultidocument,
    flowChartTerminator,
    flowChartPreparation,
    flowChartManualInput,
    flowChartManualOperation,
    flowChartConnector,
    flowChartOffpageConnector,
    flowChartPunchedCard,
    flowChartPunchedTape,
    flowChartSummingJunction,
    flowChartOr,
    flowChartCollate,
    flowChartSort,
    flowChartExtract,
    flowChartMerge,
    flowChartOnlineStorage,
    flowChartDelay,
    flowChartMagneticTape,
    flowChartMagneticDisk,
    flowChartMagneticDrum,
    flowChartDisplay
  };

  public enum StarAndBannerShapes
  {
    irregularSeal1,
    irregularSeal2,
    star4,
    star5,
    star6,
    star7,
    star8,
    star10,
    star12,
    star16,
    star24,
    star32,
    ribbon,
    ribbon2,
    ellipseRibbon,
    ellipseRibbon2,
    verticalScroll,
    horizontalScroll,
    wave,
    doubleWave
  };

  public enum CalloutShapes
  {
    wedgeRectCallout,
    wedgeRoundRectCallout,
    wedgeEllipseCallout,
    cloudCallout,
    borderCallout1,
    borderCallout2,
    borderCallout3,
    accentCallout1,
    accentCallout2,
    accentCallout3,
    callout1,
    callout2,
    callout3,
    accentBorderCallout1,
    accentBorderCallout2,
    accentBorderCallout3
  };

  public enum Alignment
  {
    left,

    center,

    right,

    both
  };

  public enum Direction
  {
    LeftToRight,
    RightToLeft
  };

  internal enum EditType
  {
    ins,
    del
  }

  internal enum CustomPropertyType
  {
    Text,
    Date,
    NumberInteger,
    NumberDecimal,
    YesOrNo
  }

  public enum RunTextType
  {
    Text,
    DelText,
  }

  public enum NoteType
  {
    Footnote,
    Endnote
  }

  public enum LineSpacingType
  {
    Line,
    Before,
    After
  }

  public enum LineSpacingTypeAuto
  {
    AutoBefore,
    AutoAfter,
    Auto,
    None
  }

  public enum DocumentTypes
  {
    Document,
    Template,
    Pdf
  }

  public enum HeadingType
  {
    [Description( "Heading1" )]
    Heading1,

    [Description( "Heading2" )]
    Heading2,

    [Description( "Heading3" )]
    Heading3,

    [Description( "Heading4" )]
    Heading4,

    [Description( "Heading5" )]
    Heading5,

    [Description( "Heading6" )]
    Heading6,

    [Description( "Heading7" )]
    Heading7,

    [Description( "Heading8" )]
    Heading8,

    [Description( "Heading9" )]
    Heading9

    // The following headings appear in the same list in Word, but they do not work in the same way (they are character based headings, not paragraph based headings)
    // NoSpacing
    // Title, Subtitle
    // Quote, IntenseQuote
    // Emphasis, IntenseEmphasis
    // Strong
    // ListParagraph
    // SubtleReference, IntenseReference
    // BookTitle
  }

  public enum TextDirection
  {
    btLr,
    right
  }

  [Flags]
  public enum TableOfContentsSwitches
  {
    None = 0 << 0,

    [Description( "\\a" )]
    A = 1 << 0,

    [Description( "\\b" )]
    B = 1 << 1,

    [Description( "\\c" )]
    C = 1 << 2,

    [Description( "\\d" )]
    D = 1 << 3,

    [Description( "\\f" )]
    F = 1 << 4,

    [Description( "\\h" )]
    H = 1 << 5,

    [Description( "\\l" )]
    L = 1 << 6,

    [Description( "\\n" )]
    N = 1 << 7,

    [Description( "\\o" )]
    O = 1 << 8,

    [Description( "\\p" )]
    P = 1 << 9,

    [Description( "\\s" )]
    S = 1 << 10,

    [Description( "\\t" )]
    T = 1 << 11,

    [Description( "\\u" )]
    U = 1 << 12,

    [Description( "\\w" )]
    W = 1 << 13,

    [Description( "\\x" )]
    X = 1 << 14,

    [Description( "\\z" )]
    Z = 1 << 15
  }

  public enum TableCellMarginType
  {
    left,
    right,
    bottom,
    top
  }

  public enum PatternStyle
  {
    Clear,
    Solid,
    Percent5,
    Percent10,
    Percent12,
    Percent15,
    Percent20,
    Percent25,
    Percent30,
    Percent35,
    Percent37,
    Percent40,
    Percent45,
    Percent50,
    Percent55,
    Percent60,
    Percent62,
    Percent65,
    Percent70,
    Percent75,
    Percent80,
    Percent85,
    Percent87,
    Percent90,
    Percent95,
    DkHorizonal,
    DkVertical,
    DkDwnDiagonal,
    DkUpDiagonal,
    DkGrid,
    DkTrellis,
    LtHorizonal,
    LtVertical,
    LtDwnDiagonal,
    LtUpDiagonal,
    LtGrid,
    LtTrellis,
  }

  public enum HorizontalBorderPosition
  {
    top,
    bottom
  }

  public enum TabStopPositionLeader
  {
    none,
    dot,
    underscore,
    hyphen
  }

  public enum NoteNumberFormat
  {
    number,
    lowerLetter,
    upperLetter,
    lowerRoman,
    upperRoman,
    chicago
  }

  public enum NumberingFormat
  {
    aiueo,
    aiueoFullWidth,
    arabicAlpha,
    bullet,
    cardinalText,
    chicago,
    chineseCounting,
    chineseCountingThousand,
    chineseLegalSimplified,
    chosung,
    decimalNormal,
    decimalEnclosedCircle,
    decimalEnclosedCircleChinese,
    decimalEnclosedFullstop,
    decimalEnclosedParen,
    decimalFullWidth,
    decimalFullWidth2,
    decimalHalfWidth,
    decimalZero,
    ganada,
    hebrew1,
    hebrew2,
    hex,
    hindiConsonants,
    hindiCounting,
    hindiNumbers,
    hindiVowels,
    ideographDigital,
    ideographEnclosedCircle,
    ideographLegalTraditional,
    ideographTraditional,
    ideographZodiac,
    ideographZodiacTraditional,
    iroha,
    irohaFullWidth,
    japaneseCounting,
    japaneseDigitalTenThousand,
    japaneseLegal,
    koreanCounting,
    koreanDigital,
    koreanDigital2,
    koreanLegal,
    lowerLetter,
    lowerRoman,
    none,
    numberInDash,
    ordinal,
    ordinalText,
    russianLower,
    russianUpper,
    taiwaneseCounting,
    taiwaneseCountingThousand,
    taiwaneseDigital,
    thaiCounting,
    thaiLetters,
    thaiNumbers,
    upperLetter,
    upperRoman,
    vietnameseCounting
  }

  public enum ReplaceTextContainer
  {
    Header = 1,
    Footer = 2,
    Body = 4,
    All = Header | Footer | Body
  }





















  public enum RemoveParagraphFlags
  {
    None = 0,
    Tables = 1,
    Pictures = 2,
    Charts = 4,
    All = RemoveParagraphFlags.Tables | RemoveParagraphFlags.Pictures | RemoveParagraphFlags.Charts
  }
}
