using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Novacode
{
    public enum MatchFormattingOptions { ExactMatch, SubsetMatch};
    public enum Script { superscript, subscript, none }
    public enum Highlight { yellow, green, cyan, magenta, blue, red, darkBlue, darkCyan, darkGreen, darkMagenta, darkRed, darkYellow, darkGray, lightGray, black, none };
    public enum UnderlineStyle { none, singleLine, doubleLine, thick, dotted, dottedHeavy, dash, dashedHeavy, dashLong, dashLongHeavy, dotDash, dashDotHeavy, dotDotDash, dashDotDotHeavy, wave, wavyHeavy, wavyDouble, words };
    public enum StrickThrough { none, strike, doubleStrike };
    public enum Misc { none, shadow, outline, outlineShadow, emboss, engrave };
    public enum CapsStyle { none, caps, smallCaps };

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

    /// <summary>
    /// Text alignment of a paragraph
    /// </summary>
    public enum Alignment
    {
        /// <summary>
        /// Align text to the left.
        /// </summary>
        left,

        /// <summary>
        /// Center text.
        /// </summary>
        center,

        /// <summary>
        /// Align text to the right.
        /// </summary>
        right,

        /// <summary>
        /// Align text to both the left and right margins, adding extra space between words as necessary.
        /// </summary>
        both
    };
    
    /// <summary>
    /// Paragraph edit types
    /// </summary>
    internal enum EditType
    {
        /// <summary>
        /// A ins is a tracked insertion
        /// </summary>
        ins,
        /// <summary>
        /// A del is  tracked deletion
        /// </summary>
        del
    }

    /// <summary>
    /// Custom property types.
    /// </summary>
    internal enum CustomPropertyType
    {
        /// <summary>
        /// System.String
        /// </summary>
        Text,
        /// <summary>
        /// System.DateTime
        /// </summary>
        Date,
        /// <summary>
        /// System.Int32
        /// </summary>
        NumberInteger,
        /// <summary>
        /// System.Double
        /// </summary>
        NumberDecimal,
        /// <summary>
        /// System.Boolean
        /// </summary>
        YesOrNo
    }

    /// <summary>
    /// Text types in a Run
    /// </summary>
    public enum RunTextType
    {
        /// <summary>
        /// System.String
        /// </summary>
        Text,
        /// <summary>
        /// System.String
        /// </summary>
        DelText,
    }
}
