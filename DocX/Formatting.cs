using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Drawing;

namespace Novacode
{
    public enum Script { superscript, subscript, none }
    public enum Highlight { yellow, green, cyan, magenta, blue, red, darkBlue, darkCyan, darkGreen, darkMagenta, darkRed, darkYellow, darkGray, lightGray, black, none};
    public enum UnderlineStyle { none, singleLine, doubleLine, thick, dotted, dottedHeavy, dash, dashedHeavy, dashLong, dashLongHeavy, dotDash, dashDotHeavy, dotDotDash, dashDotDotHeavy, wave, wavyHeavy, wavyDouble, words};
    public enum StrickThrough { none, strike, doubleStrike };
    public enum Misc { none, shadow, outline, outlineShadow, emboss, engrave};
    public enum CapsStyle { none, caps, smallCaps };

    public class Formatting
    {
        private XElement rPr;
        private bool hidden;
        private bool bold;
        private bool italic;
        private StrickThrough strikethrough;
        private Script script;
        private Highlight highlight;
        private double? size;
        private Color? fontColor;
        private Color? underlineColor;
        private UnderlineStyle underlineStyle;
        private Misc misc;
        private CapsStyle capsStyle;
        private FontFamily fontFamily;
        private int? percentageScale;
        private int? kerning;
        private int? position;
        private double? spacing;

        public Formatting()
        {
            capsStyle = CapsStyle.none;
            strikethrough = StrickThrough.none;
            script = Script.none;
            highlight = Highlight.none;
            underlineStyle = UnderlineStyle.none;
            misc = Misc.none;

            rPr = new XElement(XName.Get("rPr", DocX.w.NamespaceName));
        }

        public XElement Xml
        {
            get
            {
                if(spacing.HasValue)
                    rPr.Add(new XElement(XName.Get("spacing", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), spacing.Value * 20)));

                if(position.HasValue)
                    rPr.Add(new XElement(XName.Get("position", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), position.Value * 2)));                   

                if (kerning.HasValue)
                    rPr.Add(new XElement(XName.Get("kern", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), kerning.Value * 2)));                   

                if (percentageScale.HasValue)
                    rPr.Add(new XElement(XName.Get("w", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), percentageScale)));

                if(fontFamily != null)
                    rPr.Add(new XElement(XName.Get("rFonts", DocX.w.NamespaceName), new XAttribute(XName.Get("ascii", DocX.w.NamespaceName), fontFamily.Name)));

                if(hidden)
                    rPr.Add(new XElement(XName.Get("vanish", DocX.w.NamespaceName)));

                if (bold)
                    rPr.Add(new XElement(XName.Get("b", DocX.w.NamespaceName)));

                if (italic)
                    rPr.Add(new XElement(XName.Get("i", DocX.w.NamespaceName)));

                switch (underlineStyle)
                {
                    case UnderlineStyle.none:
                        break;
                    case UnderlineStyle.singleLine:
                        rPr.Add(new XElement(XName.Get("u", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "single")));
                        break;
                    case UnderlineStyle.doubleLine:
                        rPr.Add(new XElement(XName.Get("u", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "double")));
                        break;
                    default:
                        rPr.Add(new XElement(XName.Get("u", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), underlineStyle.ToString())));
                        break;
                }

                if(underlineColor.HasValue)
                {
                    // If an underlineColor has been set but no underlineStyle has been set
                    if (underlineStyle == UnderlineStyle.none)
                    {
                        // Set the underlineStyle to the default
                        underlineStyle = UnderlineStyle.singleLine;
                        rPr.Add(new XElement(XName.Get("u", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "single")));
                    }

                    rPr.Element(XName.Get("u", DocX.w.NamespaceName)).Add(new XAttribute(XName.Get("color", DocX.w.NamespaceName), underlineColor.Value.ToHex()));
                }

                switch (strikethrough)
                {
                    case StrickThrough.none:
                        break;
                    case StrickThrough.strike:
                        rPr.Add(new XElement(XName.Get("strike", DocX.w.NamespaceName)));
                        break;
                    case StrickThrough.doubleStrike:
                        rPr.Add(new XElement(XName.Get("dstrike", DocX.w.NamespaceName)));
                        break;
                    default:
                        break;
                }
                  
                switch (script)
                {
                    case Script.none:
                        break;
                    default:
                        rPr.Add(new XElement(XName.Get("vertAlign", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), script.ToString())));
                        break;
                }

                if (size.HasValue)
                {
                    rPr.Add(new XElement(XName.Get("sz", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), (size * 2).ToString())));
                    rPr.Add(new XElement(XName.Get("szCs", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), (size * 2).ToString())));
                }

                if(fontColor.HasValue)
                    rPr.Add(new XElement(XName.Get("color", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), fontColor.Value.ToHex())));

                switch (highlight)
                {
                    case Highlight.none:
                        break;
                    default:
                        rPr.Add(new XElement(XName.Get("highlight", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), highlight.ToString())));
                        break;
                }

                switch (capsStyle)
                {
                    case CapsStyle.none:
                        break;
                    default:
                        rPr.Add(new XElement(XName.Get(capsStyle.ToString(), DocX.w.NamespaceName)));
                        break;
                }

                switch (misc)
                {
                    case Misc.none:
                        break;
                    case Misc.outlineShadow:
                        rPr.Add(new XElement(XName.Get("outline", DocX.w.NamespaceName)));
                        rPr.Add(new XElement(XName.Get("shadow", DocX.w.NamespaceName)));
                        break;
                    case Misc.engrave:
                        rPr.Add(new XElement(XName.Get("imprint", DocX.w.NamespaceName)));
                        break;
                    default:
                        rPr.Add(new XElement(XName.Get(misc.ToString(), DocX.w.NamespaceName)));
                        break;
                }

                return rPr;
            }
        }

        public bool Bold { get { return bold; } set { bold = value;} }
        public bool Italic { get { return Italic; } set { italic = value; } }
        public StrickThrough StrikeThrough { get { return strikethrough; } set { strikethrough = value; } }
        public Script Script { get { return script; } set { script = value; } }
        
        public double? Size 
        { 
            get { return size; } 
            
            set 
            { 
                double? temp = value * 2;

                if (temp - (int)temp == 0)
                {
                    if(value > 0 && value < 1639)
                        size = value;
                    else
                        throw new ArgumentException("Size", "Value must be in the range 0 - 1638");
                }

                else
                    throw new ArgumentException("Size", "Value must be either a whole or half number, examples: 32, 32.5");
            } 
        }

        public int? PercentageScale
        { 
            get { return percentageScale; } 
            
            set 
            {
                if ((new int?[] { 200, 150, 100, 90, 80, 66, 50, 33 }).Contains(value))
                    percentageScale = value; 
                else
                    throw new ArgumentOutOfRangeException("PercentageScale", "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33");
            } 
        }

        public int? Kerning 
        { 
            get { return kerning; } 
            
            set 
            { 
                if(new int?[] {8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72}.Contains(value))
                    kerning = value; 
                else
                    throw new ArgumentOutOfRangeException("Kerning", "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72");
            } 
        }

        public int? Position
        {
            get { return position; }

            set
            {
                if (value > -1585 && value < 1585)
                    position = value;
                else
                    throw new ArgumentOutOfRangeException("Position", "Value must be in the range -1585 - 1585");
            }
        }

        public double? Spacing
        {
            get { return spacing; }

            set
            {
                double? temp = value * 20;

                if (temp - (int)temp == 0)
                {
                    if (value > -1585 && value < 1585)
                        spacing = value;
                    else
                        throw new ArgumentException("Spacing", "Value must be in the range: -1584 - 1584");
                }

                else
                    throw new ArgumentException("Spacing", "Value must be either a whole or acurate to one decimal, examples: 32, 32.1, 32.2, 32.9");
            } 
        }

        public Color? FontColor { get { return fontColor; } set { fontColor = value; } }
        public Highlight Highlight { get { return highlight; } set { highlight = value; } }
        public UnderlineStyle UnderlineStyle { get { return underlineStyle; } set { underlineStyle = value; } }
        public Color? UnderlineColor { get { return underlineColor; } set { underlineColor = value; } }
        public Misc Misc { get { return misc; } set { misc = value; } }
        public bool Hidden { get { return hidden; } set { hidden = value; } }
        public CapsStyle CapsStyle { get { return capsStyle; } set { capsStyle = value; } }
        public FontFamily FontFamily { get { return FontFamily; } set { fontFamily = value; } }

    }
}
