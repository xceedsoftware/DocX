/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


#if NET5
using SkiaSharp;
using System;
using System.IO;
using System.Linq;
#else
using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
#endif

namespace Xceed.Drawing
{
  public struct Color
  {
    #region Private Members

#if NET5
    private readonly SKColor m_color;
#else
    private readonly System.Drawing.Color m_color;
#endif

    #endregion

    #region Constructors

    public Color(
#if NET5
            SKColor color
#else
            System.Drawing.Color color
#endif
      )
    {
      m_color = color;
    }

    private Color( string value )
    {
#if NET5
      var skColor = typeof( SKColors ).GetField( value ); 
      if( skColor == null )
        throw new InvalidDataException( "Unknown color name." );

      m_color = (SKColor)skColor.GetValue( null );      
#else
      m_color = System.Drawing.Color.FromName( value );
#endif
    }

    #endregion

    #region Static Colors 

    public static readonly Color AliceBlue = new Color( "AliceBlue" );

    public static readonly Color AntiqueWhite = new Color( "AntiqueWhite" );

    public static readonly Color Aqua = new Color( "Aqua" );

    public static readonly Color Aquamarine = new Color( "Aquamarine" );

    public static readonly Color Azure = new Color( "Azure" );

    public static readonly Color Beige = new Color( "Beige" );

    public static readonly Color Bisque = new Color( "Bisque" );

    public static readonly Color Black = new Color( "Black" );

    public static readonly Color BlanchedAlmond = new Color( "BlanchedAlmond" );

    public static readonly Color Blue = new Color( "Blue" );

    public static readonly Color BlueViolet = new Color( "BlueViolet" );

    public static readonly Color Brown = new Color( "Brown" );

    public static readonly Color BurlyWood = new Color( "BurlyWood" );

    public static readonly Color CadetBlue = new Color( "CadetBlue" );

    public static readonly Color Chartreuse = new Color( "Chartreuse" );

    public static readonly Color Chocolate = new Color( "Chocolate" );

    public static readonly Color Coral = new Color( "Coral" );

    public static readonly Color CornflowerBlue = new Color( "CornflowerBlue" );

    public static readonly Color Cornsilk = new Color( "Cornsilk" );

    public static readonly Color Crimson = new Color( "Crimson" );

    public static readonly Color Cyan = new Color( "Cyan" );

    public static readonly Color DarkBlue = new Color( "DarkBlue" );

    public static readonly Color DarkCyan = new Color( "DarkCyan" );

    public static readonly Color DarkGoldenrod = new Color( "DarkGoldenrod" );

    public static readonly Color DarkGray = new Color( "DarkGray" );

    public static readonly Color DarkGreen = new Color( "DarkGreen" );

    public static readonly Color DarkKhaki = new Color( "DarkKhaki" );

    public static readonly Color DarkMagenta = new Color( "DarkMagenta" );

    public static readonly Color DarkOliveGreen = new Color( "DarkOliveGreen" );

    public static readonly Color DarkOrange = new Color( "DarkOrange" );

    public static readonly Color DarkOrchid = new Color( "DarkOrchid" );

    public static readonly Color DarkRed = new Color( "DarkRed" );

    public static readonly Color DarkSalmon = new Color( "DarkSalmon" );

    public static readonly Color DarkSeaGreen = new Color( "DarkSeaGreen" );

    public static readonly Color DarkSlateBlue = new Color( "DarkSlateBlue" );

    public static readonly Color DarkSlateGray = new Color( "DarkSlateGray" );

    public static readonly Color DarkTurquoise = new Color( "DarkTurquoise" );

    public static readonly Color DarkViolet = new Color( "DarkViolet" );

    public static readonly Color DeepPink = new Color( "DeepPink" );

    public static readonly Color DeepSkyBlue = new Color( "DeepSkyBlue" );

    public static readonly Color DimGray = new Color( "DimGray" );

    public static readonly Color DodgerBlue = new Color( "DodgerBlue" );

    public static readonly Color Firebrick = new Color( "Firebrick" );

    public static readonly Color FloralWhite = new Color( "FloralWhite" );

    public static readonly Color ForestGreen = new Color( "ForestGreen" );

    public static readonly Color Fuchsia = new Color( "Fuchsia" );

    public static readonly Color Gainsboro = new Color( "Gainsboro" );

    public static readonly Color GhostWhite = new Color( "GhostWhite" );

    public static readonly Color Gold = new Color( "Gold" );

    public static readonly Color Goldenrod = new Color( "Goldenrod" );

    public static readonly Color Gray = new Color( "Gray" );

    public static readonly Color Green = new Color( "Green" );

    public static readonly Color GreenYellow = new Color( "GreenYellow" );

    public static readonly Color Honeydew = new Color( "Honeydew" );

    public static readonly Color HotPink = new Color( "HotPink" );

    public static readonly Color IndianRed = new Color( "IndianRed" );

    public static readonly Color Indigo = new Color( "Indigo" );

    public static readonly Color Ivory = new Color( "Ivory" );

    public static readonly Color Khaki = new Color( "Khaki" );

    public static readonly Color Lavender = new Color( "Lavender" );

    public static readonly Color LavenderBlush = new Color( "LavenderBlush" );

    public static readonly Color LawnGreen = new Color( "LawnGreen" );

    public static readonly Color LemonChiffon = new Color( "LemonChiffon" );

    public static readonly Color LightBlue = new Color( "LightBlue" );

    public static readonly Color LightCoral = new Color( "LightCoral" );

    public static readonly Color LightCyan = new Color( "LightCyan" );

    public static readonly Color LightGoldenrodYellow = new Color( "LightGoldenrodYellow" );

    public static readonly Color LightGray = new Color( "LightGray" );

    public static readonly Color LightGreen = new Color( "LightGreen" );

    public static readonly Color LightPink = new Color( "LightPink" );

    public static readonly Color LightSalmon = new Color( "LightSalmon" );

    public static readonly Color LightSeaGreen = new Color( "LightSeaGreen" );

    public static readonly Color LightSkyBlue = new Color( "LightSkyBlue" );

    public static readonly Color LightSlateGray = new Color( "LightSlateGray" );

    public static readonly Color LightSteelBlue = new Color( "LightSteelBlue" );

    public static readonly Color LightYellow = new Color( "LightYellow" );

    public static readonly Color Lime = new Color( "Lime" );

    public static readonly Color LimeGreen = new Color( "LimeGreen" );

    public static readonly Color Linen = new Color( "Linen" );

    public static readonly Color Magenta = new Color( "Magenta" );

    public static readonly Color Maroon = new Color( "Maroon" );

    public static readonly Color MediumAquamarine = new Color( "MediumAquamarine" );

    public static readonly Color MediumBlue = new Color( "MediumBlue" );

    public static readonly Color MediumOrchid = new Color( "MediumOrchid" );

    public static readonly Color MediumPurple = new Color( "MediumPurple" );

    public static readonly Color MediumSeaGreen = new Color( "MediumSeaGreen" );

    public static readonly Color MediumSlateBlue = new Color( "MediumSlateBlue" );

    public static readonly Color MediumSpringGreen = new Color( "MediumSpringGreen" );

    public static readonly Color MediumTurquoise = new Color( "MediumTurquoise" );

    public static readonly Color MediumVioletRed = new Color( "MediumVioletRed" );

    public static readonly Color MidnightBlue = new Color( "MidnightBlue" );

    public static readonly Color MintCream = new Color( "MintCream" );

    public static readonly Color MistyRose = new Color( "MistyRose" );

    public static readonly Color Moccasin = new Color( "Moccasin" );

    public static readonly Color NavajoWhite = new Color( "NavajoWhite" );

    public static readonly Color Navy = new Color( "Navy" );

    public static readonly Color OldLace = new Color( "OldLace" );

    public static readonly Color Olive = new Color( "Olive" );

    public static readonly Color OliveDrab = new Color( "OliveDrab" );

    public static readonly Color Orange = new Color( "Orange" );

    public static readonly Color OrangeRed = new Color( "OrangeRed" );

    public static readonly Color Orchid = new Color( "Orchid" );

    public static readonly Color PaleGoldenrod = new Color( "PaleGoldenrod" );

    public static readonly Color PaleGreen = new Color( "PaleGreen" );

    public static readonly Color PaleTurquoise = new Color( "PaleTurquoise" );

    public static readonly Color PaleVioletRed = new Color( "PaleVioletRed" );

    public static readonly Color PapayaWhip = new Color( "PapayaWhip" );

    public static readonly Color PeachPuff = new Color( "PeachPuff" );

    public static readonly Color Peru = new Color( "Peru" );

    public static readonly Color Pink = new Color( "Pink" );

    public static readonly Color Plum = new Color( "Plum" );

    public static readonly Color PowderBlue = new Color( "PowderBlue" );

    public static readonly Color Purple = new Color( "Purple" );

    public static readonly Color Red = new Color( "Red" );

    public static readonly Color RosyBrown = new Color( "RosyBrown" );

    public static readonly Color RoyalBlue = new Color( "RoyalBlue" );

    public static readonly Color SaddleBrown = new Color( "SaddleBrown" );

    public static readonly Color Salmon = new Color( "Salmon" );

    public static readonly Color SandyBrown = new Color( "SandyBrown" );

    public static readonly Color SeaGreen = new Color( "SeaGreen" );

    public static readonly Color SeaShell = new Color( "SeaShell" );

    public static readonly Color Sienna = new Color( "Sienna" );

    public static readonly Color Silver = new Color( "Silver" );

    public static readonly Color SkyBlue = new Color( "SkyBlue" );

    public static readonly Color SlateBlue = new Color( "SlateBlue" );

    public static readonly Color SlateGray = new Color( "SlateGray" );

    public static readonly Color Snow = new Color( "Snow" );

    public static readonly Color SpringGreen = new Color( "SpringGreen" );

    public static readonly Color SteelBlue = new Color( "SteelBlue" );

    public static readonly Color Tan = new Color( "Tan" );

    public static readonly Color Teal = new Color( "Teal" );

    public static readonly Color Thistle = new Color( "Thistle" );

    public static readonly Color Tomato = new Color( "Tomato" );

    public static readonly Color Transparent = new Color( "Transparent" );

    public static readonly Color Turquoise = new Color( "Turquoise" );

    public static readonly Color Violet = new Color( "Violet" );

    public static readonly Color Wheat = new Color( "Wheat" );

    public static readonly Color White = new Color( "White" );

    public static readonly Color WhiteSmoke = new Color( "WhiteSmoke" );

    public static readonly Color Yellow = new Color( "Yellow" );

    public static readonly Color YellowGreen = new Color( "YellowGreen" );

    public static readonly Color Empty =
#if NET5
      Color.Parse( 0, Color.Black );
#else
      new Color( System.Drawing.Color.Empty );
#endif

    public static readonly Color ScrollBar =
#if NET5
      Color.Parse( 200, 200, 200 );
#else
      new Color( SystemColors.ScrollBar );
#endif

    public static readonly Color Desktop =
#if NET5
     new Color( SKColors.Black );
#else
    new Color( SystemColors.Desktop );
#endif

    public static readonly Color ActiveCaption =
#if NET5
      Color.Parse( 153, 180, 209 );
#else
      new Color( SystemColors.ActiveCaption );
#endif

    public static readonly Color InactiveCaption =
#if NET5
      Color.Parse( 191, 205, 219 );
#else
     new Color( SystemColors.InactiveCaption );
#endif

    public static readonly Color Menu =
#if NET5
      Color.Parse( 240, 240, 240 );
#else
     new Color( SystemColors.Menu );
#endif

    public static readonly Color Window =
#if NET5
      Color.White;
#else
    new Color( SystemColors.Window );
#endif

    public static readonly Color WindowFrame =
#if NET5
      Color.Parse( 100, 100, 100 );
#else
    new Color( SystemColors.WindowFrame );
#endif

    public static readonly Color MenuText =
#if NET5
      Color.Black;
#else
    new Color( SystemColors.MenuText );
#endif

    public static readonly Color WindowText =
#if NET5
      Color.Black;
#else
   new Color( SystemColors.WindowText );
#endif

    public static readonly Color ActiveCaptionText =
#if NET5
      Color.Black;
#else
    new Color( SystemColors.ActiveCaptionText );
#endif

    public static readonly Color ActiveBorder =
#if NET5
      Color.Parse( 180, 180, 180 );
#else
    new Color( SystemColors.ActiveBorder );
#endif

    public static readonly Color InactiveBorder =
#if NET5
      Color.Parse( 210, 210, 210 );
#else
      new Color( SystemColors.InactiveBorder );
#endif

    public static readonly Color AppWorkspace =
#if NET5
      Color.Parse( 171, 171, 171 );
#else
     new Color( SystemColors.AppWorkspace );
#endif

    public static readonly Color Highlight =
#if NET5
      Color.Parse( 0, 120, 215 );
#else
    new Color( SystemColors.Highlight );
#endif

    public static readonly Color HighlightText =
#if NET5
      Color.White;
#else
    new Color( SystemColors.HighlightText );
#endif

    public static readonly Color ButtonFace =
#if NET5
      Color.Parse( 240, 240, 240 );
#else
    new Color( SystemColors.ButtonFace );
#endif

    public static readonly Color ButtonShadow =
#if NET5
      Color.Parse( 160, 160, 160 );
#else
    new Color( SystemColors.ButtonShadow );
#endif

    public static readonly Color GrayText =
#if NET5
      Color.Parse( 109, 109, 109 );
#else
      new Color( SystemColors.GrayText );
#endif

    public static readonly Color ControlText =
#if NET5
      Color.Black;
#else
    new Color( SystemColors.ControlText );
#endif

    public static readonly Color InactiveCaptionText =
#if NET5
      Color.Black;
#else
    new Color( SystemColors.InactiveCaptionText );
#endif

    public static readonly Color ButtonHighlight =
#if NET5
      Color.White;
#else
    new Color( SystemColors.ButtonHighlight );
#endif

    public static readonly Color ControlLight =
#if NET5
      Color.Parse( 227, 227, 227 );
#else
    new Color( SystemColors.ControlLight );
#endif

    public static readonly Color InfoText =
#if NET5
      Color.Black;
#else
      new Color( SystemColors.InfoText );
#endif

    public static readonly Color Info =
#if NET5
      Color.Parse( 255, 255, 225 );
#else
      new Color( SystemColors.Info );
#endif

    public static readonly Color HotTrack =
#if NET5
      Color.Parse( 0, 102, 204 );
#else
      new Color( SystemColors.HotTrack );
#endif

    public static readonly Color GradientActiveCaption =
#if NET5
      Color.Parse( 185, 209, 234 );
#else
     new Color( SystemColors.GradientActiveCaption );
#endif

    public static readonly Color GradientInactiveCaption =
#if NET5
      Color.Parse( 215, 228, 242 );
#else
      new Color( SystemColors.GradientInactiveCaption );
#endif

    public static readonly Color MenuHighlight =
#if NET5
      Color.Parse( 51, 153, 255 );
#else
      new Color( SystemColors.MenuHighlight );
#endif

    public static readonly Color MenuBar =
#if NET5
      Color.Parse( 240, 240, 240 );
#else
     new Color( SystemColors.MenuBar );
#endif

    #endregion

    #region Properties

    #region A

    public byte A
    {
      get
      {
#if NET5
        return m_color.Alpha;
#else
        return m_color.A;
#endif
      }
    }

    #endregion

    #region B
    public byte B
    {
      get
      {
#if NET5
        return m_color.Blue;
#else
        return m_color.B;
#endif
      }
    }

    #endregion

    #region G

    public byte G
    {
      get
      {
#if NET5
        return m_color.Green;
#else
        return m_color.G;
#endif
      }
    }

    #endregion

    #region IsIsEmpty

    public bool IsEmpty
    {
      get
      {
        return this == Color.Empty;
      }
    }

    #endregion

    #region Name

    public string Name
    {
      get
      {
#if NET5
        return m_color.ToString();
#else
        return m_color.Name;
#endif
      }
    }

    #endregion

    #region R

    public byte R
    {
      get
      {
#if NET5
        return m_color.Red;
#else
        return m_color.R;
#endif
      }
    }

    #endregion

    #region Value

#if NET5
    public SKColor Value
#else
    public System.Drawing.Color Value
#endif
    {
      get
      {
        return m_color;
      }
    }

    #endregion

    #endregion

    #region Static Methods

    public static bool IsKnownColor( string value )
    {
#if NET5
      return Enum.TryParse( value, true, out SKColors knownColor );
#else
      return Enum.TryParse( value, true, out KnownColor knownColor );
#endif
    }

    public static Color Parse( string stringColor )
    {
#if NET5
      return new Color( SKColor.Parse( stringColor ) );
#else
      var rgb = System.Drawing.Color.FromArgb( Int32.Parse( stringColor, NumberStyles.HexNumber ) );
      return new Color( System.Drawing.Color.FromArgb( 255, rgb ) );
#endif
    }

    public static Color Parse( int r, int g, int b )
    {
#if NET5
      return new Color( new SKColor( Convert.ToByte( r ), Convert.ToByte( g ), Convert.ToByte( b ) ) );
#else
      return new Color( System.Drawing.Color.FromArgb( r, g, b ) );
#endif
    }

    public static Color Parse( int a, int r, int g, int b )
    {
#if NET5
      return new Color( new SKColor( Convert.ToByte( a ), Convert.ToByte( r ), Convert.ToByte( g ), Convert.ToByte( b ) ) );
#else
      return new Color( System.Drawing.Color.FromArgb( a, r, g, b ) );
#endif
    }

    public static Color Parse( int a, Color color )
    {
#if NET5
      return new Color( new SKColor( color.R, color.G, color.B, Convert.ToByte( a ) ) );
#else
      return new Color( System.Drawing.Color.FromArgb( a, color.Value ) );
#endif
    }

    public static Color Parse( int argb )
    {
#if NET5
      return new Color( new SKColor( (uint)argb ) );
#else
      return new Color( System.Drawing.Color.FromArgb( argb ) );
#endif
    }

    public static string GetColorName( int argbColor )
    {
#if NET5
      var baseColor = Color.Parse( argbColor ).Value;

      // Get all predefined colors from SKColors
    var knownColors = typeof(SKColors).GetProperties()
        .Select(p => (SKColor)p.GetValue(null))
        .ToList();

    // Find the color that matches the ARGB value
    var matchingColor = knownColors.Where(color => color == baseColor )
                                   .FirstOrDefault();

    // If no match is found, return null or a default value
    if (matchingColor == SKColors.Empty)
        return "Unknown"; // or handle as needed

    // Return the name of the color by finding the property name
    var colorName = typeof(SKColors).GetProperties()
        .FirstOrDefault(p => (SKColor)p.GetValue(null) == matchingColor)?
        .Name;

    return colorName ?? "Unknown";
#else
      var knownColors = (KnownColor[])Enum.GetValues( typeof( KnownColor ) );
      var knownColor = knownColors.Where( col => System.Drawing.Color.FromKnownColor( col ).ToArgb() == argbColor ).FirstOrDefault();
      return System.Drawing.Color.FromKnownColor( knownColor ).Name;
#endif
    }

    public static Color FromColorName( string colorName )
    {
#if NET5
      // Use reflection to find the property in SKColors that matches the color name
      var skColor = typeof( SKColors ).GetField( colorName,
          System.Reflection.BindingFlags.Public |
          System.Reflection.BindingFlags.Static |
          System.Reflection.BindingFlags.IgnoreCase );

      // If the property exists, return the corresponding SKColor
      if( skColor != null )
        return new Color( (SKColor)skColor.GetValue( null ) );

      return Color.Black;
#else
      var knownColors = (KnownColor[])Enum.GetValues( typeof( KnownColor ) );
      var knownColor = knownColors.Where( col => System.Drawing.Color.FromKnownColor( col ).Name == colorName ).SingleOrDefault();
      return new Color( System.Drawing.Color.FromKnownColor( knownColor ) );
#endif
    }

    public int ToArgb()
    {
#if NET5
      return ( m_color.Alpha << 24 ) | ( m_color.Red << 16 ) | ( m_color.Green << 8 ) | m_color.Blue;
#else
      return m_color.ToArgb();
#endif
    }

    public float GetBrightness()
    {
#if NET5
      m_color.ToHsv( out float hue, out float saturation, out float brightness );

      return brightness;
#else
      return m_color.GetBrightness();
#endif
    }

    public float GetHue()
    {
#if NET5
      return m_color.Hue;
#else
      return m_color.GetHue();
#endif
    }

    public float GetSaturation()
    {
#if NET5
      m_color.ToHsv( out float hue, out float saturation, out float brightness );

      return saturation;
#else
      return m_color.GetSaturation();
#endif
    }

    #endregion

    #region Public Methods

    public override bool Equals( object obj )
    {
      if( !( obj is Color ) )
        return false;

      var other = (Color)obj;

      return this.Value == other.Value
           && this.A == other.A
           && this.R == other.R
           && this.G == other.G
           && this.B == other.B
           && this.Name == other.Name;
    }

    public override int GetHashCode()
    {
      var hash = 17;
      hash = hash * 31 + this.R.GetHashCode();
      hash = hash * 31 + this.G.GetHashCode();
      hash = hash * 31 + this.B.GetHashCode();
      hash = hash * 31 + this.A.GetHashCode();
      hash = hash * 31 + this.Name.GetHashCode();
      hash = hash * 31 + this.Value.GetHashCode();

      return hash;
    }

    public static bool operator ==( Color c1, Color c2 )
    {
      return c1.Equals( c2 );
    }

    public static bool operator !=( Color c1, Color c2 )
    {
      return !c1.Equals( c2 );
    }

    #endregion
  }
}
