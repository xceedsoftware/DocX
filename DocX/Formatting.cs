using System;
using System.Linq;
using System.Xml.Linq;
using System.Drawing;
using System.Globalization;
namespace Novacode
{
    /// <summary>
    /// A text formatting.
    /// </summary>
    public class Formatting : IComparable
    {
        private XElement rPr;
        private bool hidden;
        private bool bold;
        private bool italic;
        private StrikeThrough strikethrough;
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

        private CultureInfo language;

        /// <summary>
        /// A text formatting.
        /// </summary>
        public Formatting()
        {
            capsStyle = CapsStyle.none;
            strikethrough = StrikeThrough.none;
            script = Script.none;
            highlight = Highlight.none;
            underlineStyle = UnderlineStyle.none;
            misc = Misc.none;

            // Use current culture by default
            language = CultureInfo.CurrentCulture;

            rPr = new XElement(XName.Get("rPr", DocX.w.NamespaceName));
        }

        /// <summary>
        /// Text language
        /// </summary>
        public CultureInfo Language 
        { 
            get 
            { 
                return language; 
            } 
            
            set 
            { 
                language = value; 
            } 
        }

        public static Formatting Parse(XElement rPr)
        {
            Formatting formatting = new Formatting();

            // Build up the Formatting object.
            foreach (XElement option in rPr.Elements())
            {
                switch (option.Name.LocalName)
                {
                    case "lang": formatting.Language = new CultureInfo(option.GetAttribute(XName.Get("val", DocX.w.NamespaceName), null) ?? option.GetAttribute(XName.Get("eastAsia", DocX.w.NamespaceName), null) ?? option.GetAttribute(XName.Get("bidi", DocX.w.NamespaceName))); break;
                    case "spacing": formatting.Spacing = Double.Parse(option.GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 20.0; break;
                    case "position": formatting.Position = Int32.Parse(option.GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 2; break;
                    case "kern": formatting.Position = Int32.Parse(option.GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 2; break;
                    case "w": formatting.PercentageScale = Int32.Parse(option.GetAttribute(XName.Get("val", DocX.w.NamespaceName))); break;
                    case "rFonts": formatting.FontFamily = new FontFamily(option.GetAttribute(XName.Get("cs", DocX.w.NamespaceName))); break;
                    case "vanish": formatting.hidden = true; break;
                    case "b": formatting.Bold = true; break;
                    case "i": formatting.Italic = true; break;
                    case "u": formatting.UnderlineStyle = HelperFunctions.GetUnderlineStyle(option.GetAttribute(XName.Get("val", DocX.w.NamespaceName)));
                              break;
                    default: break;
                }
            }

            return formatting;
        }

        internal XElement Xml
        {
            get
            {
                rPr = new XElement(XName.Get("rPr", DocX.w.NamespaceName));

                if (language != null)
                    rPr.Add(new XElement(XName.Get("lang", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), language.Name)));
                
                if(spacing.HasValue)
                    rPr.Add(new XElement(XName.Get("spacing", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), spacing.Value * 20)));

                if(position.HasValue)
                    rPr.Add(new XElement(XName.Get("position", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), position.Value * 2)));                   

                if (kerning.HasValue)
                    rPr.Add(new XElement(XName.Get("kern", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), kerning.Value * 2)));                   

                if (percentageScale.HasValue)
                    rPr.Add(new XElement(XName.Get("w", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), percentageScale)));

                if (fontFamily != null)
                {
                    rPr.Add
                    (
                        new XElement
                        (
                            XName.Get("rFonts", DocX.w.NamespaceName), 
                            new XAttribute(XName.Get("ascii", DocX.w.NamespaceName), fontFamily.Name),
                            new XAttribute(XName.Get("hAnsi", DocX.w.NamespaceName), fontFamily.Name), // Added by Maurits Elbers to support non-standard characters. See http://docx.codeplex.com/Thread/View.aspx?ThreadId=70097&ANCHOR#Post453865
                            new XAttribute(XName.Get("cs", DocX.w.NamespaceName), fontFamily.Name)    // Added by Maurits Elbers to support non-standard characters. See http://docx.codeplex.com/Thread/View.aspx?ThreadId=70097&ANCHOR#Post453865
                        )
                    );
                }

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
                    case StrikeThrough.none:
                        break;
                    case StrikeThrough.strike:
                        rPr.Add(new XElement(XName.Get("strike", DocX.w.NamespaceName)));
                        break;
                    case StrikeThrough.doubleStrike:
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

        /// <summary>
        /// This formatting will apply Bold.
        /// </summary>
        public bool Bold { get { return bold; } set { bold = value;} }

        /// <summary>
        /// This formatting will apply Italic.
        /// </summary>
        public bool Italic { get { return italic; } set { italic = value; } }

        /// <summary>
        /// This formatting will apply StrickThrough.
        /// </summary>
        public StrikeThrough StrikeThrough { get { return strikethrough; } set { strikethrough = value; } }

        /// <summary>
        /// The script that this formatting should be, normal, superscript or subscript.
        /// </summary>
        public Script Script { get { return script; } set { script = value; } }
        
        /// <summary>
        /// The Size of this text, must be between 0 and 1638.
        /// </summary>
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

        /// <summary>
        /// Percentage scale must be one of the following values 200, 150, 100, 90, 80, 66, 50 or 33.
        /// </summary>
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

        /// <summary>
        /// The Kerning to apply to this text must be one of the following values 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72.
        /// </summary>
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

        /// <summary>
        /// Text position must be in the range (-1585 - 1585).
        /// </summary>
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

        /// <summary>
        /// Text spacing must be in the range (-1585 - 1585).
        /// </summary>
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

        /// <summary>
        /// The colour of the text.
        /// </summary>
        public Color? FontColor { get { return fontColor; } set { fontColor = value; } }

        /// <summary>
        /// Highlight colour.
        /// </summary>
        public Highlight Highlight { get { return highlight; } set { highlight = value; } }
       
        /// <summary>
        /// The Underline style that this formatting applies.
        /// </summary>
        public UnderlineStyle UnderlineStyle { get { return underlineStyle; } set { underlineStyle = value; } }
        
        /// <summary>
        /// The underline colour.
        /// </summary>
        public Color? UnderlineColor { get { return underlineColor; } set { underlineColor = value; } }
        
        /// <summary>
        /// Misc settings.
        /// </summary>
        public Misc Misc { get { return misc; } set { misc = value; } }
        
        /// <summary>
        /// Is this text hidden or visible.
        /// </summary>
        public bool Hidden { get { return hidden; } set { hidden = value; } }
        
        /// <summary>
        /// Capitalization style.
        /// </summary>
        public CapsStyle CapsStyle { get { return capsStyle; } set { capsStyle = value; } }
        
        /// <summary>
        /// The font familt of this formatting.
        /// </summary>
        /// <!-- 
        /// Bug found and fixed by krugs525 on August 12 2009.
        /// Use TFS compare to see exact code change.
        /// -->
        public FontFamily FontFamily { get { return fontFamily; } set { fontFamily = value; } }

        public int CompareTo(object obj)
        {
            Formatting other = (Formatting)obj;

            if(other.hidden != this.hidden)
                return -1;

            if(other.bold != this.bold)
                return -1;

            if(other.italic != this.italic)
                return -1;

            if(other.strikethrough != this.strikethrough)
                return -1;

            if(other.script != this.script)
                return -1;

            if(other.highlight != this.highlight)
                return -1;

            if(other.size != this.size)
                return -1;

            if(other.fontColor != this.fontColor)
                return -1;

            if(other.underlineColor != this.underlineColor)
                return -1;

            if(other.underlineStyle != this.underlineStyle)
                return -1;

            if(other.misc != this.misc)
                return -1;

            if(other.capsStyle != this.capsStyle)
                return -1;

            if(other.fontFamily != this.fontFamily)
                return -1;

            if(other.percentageScale != this.percentageScale)
                return -1;

            if(other.kerning != this.kerning)
                return -1;

            if(other.position != this.position)
                return -1;

            if(other.spacing != this.spacing)
                return -1;

            if (!other.language.Equals(this.language))
                return -1;

            return 0;
        }
    }
}
