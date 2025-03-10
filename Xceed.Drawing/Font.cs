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
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
#else
using System;
using System.Drawing;
using System.Drawing.Text;
using System.Runtime.InteropServices;
#endif

namespace Xceed.Drawing
{
  [Flags]
  public enum FontStyle
  {
    Regular = 0,
    Bold = 1,
    Italic = 2,
    Underline = 4,
    Strikeout = 8
  }

  public class Font
  {
#if !NET5
		[DllImport("user32.dll")]
		private static extern IntPtr GetDC(IntPtr hWnd);

    [DllImport("gdi32.dll")]
		private static extern IntPtr SelectObject(IntPtr hDC, IntPtr hGDIObj);

    [DllImport("gdi32.dll")]
		private static extern IntPtr GetFontData(IntPtr hDC, int table, int offset, [In, Out] byte[] buffer, int length);

    [DllImport("gdi32.dll")]
		private static extern IntPtr DeleteObject(IntPtr hGDIObj);

    [DllImport("user32.dll")]
		private static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDC);
#endif
    #region Private Members

#if NET5
    private readonly SKFont m_font;
    private bool m_isUnderline;
    private bool m_isStrikeout;
#else
    private readonly System.Drawing.Font m_font;
    private static System.Drawing.Graphics Graphics;
#endif

    #endregion

    #region Constructors

    public Font( string fontFamily, float fontSize, FontStyle fontStyle )
    {
#if NET5
      m_isUnderline = ( (fontStyle & FontStyle.Underline) == FontStyle.Underline );
      m_isStrikeout = ( ( fontStyle & FontStyle.Strikeout ) == FontStyle.Strikeout );
      var typefaceWeight = ( ( fontStyle & FontStyle.Bold ) == FontStyle.Bold ) ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal;
      var typefaceSlant = ( ( fontStyle & FontStyle.Italic ) == FontStyle.Italic ) ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright;
      var typeface = Font.GetSKTypeface( fontFamily, typefaceWeight, SKFontStyleWidth.Normal, typefaceSlant );
      m_font = new SKFont( typeface, fontSize );
#else
      var systemFontStyle = System.Drawing.FontStyle.Regular;
      if( ( fontStyle & FontStyle.Bold ) == FontStyle.Bold )
      {
        systemFontStyle |= System.Drawing.FontStyle.Bold;
      }
      if( ( fontStyle & FontStyle.Italic ) == FontStyle.Italic )
      {
        systemFontStyle |= System.Drawing.FontStyle.Italic;
      }
      if( ( fontStyle & FontStyle.Underline ) == FontStyle.Underline )
      {
        systemFontStyle |= System.Drawing.FontStyle.Underline;
      }
      if( ( fontStyle & FontStyle.Strikeout ) == FontStyle.Strikeout )
      {
        systemFontStyle |= System.Drawing.FontStyle.Strikeout;
      }

      m_font = new System.Drawing.Font( fontFamily, System.Convert.ToSingle( fontSize ), systemFontStyle );
#endif
    }

    public Font( string fontFamily, float fontSize )
    {
#if NET5
      var typeface = Font.GetSKTypeFace( fontFamily );
      m_font = new SKFont( typeface, fontSize );
#else
      m_font = new System.Drawing.Font( fontFamily, System.Convert.ToSingle( fontSize ) );
#endif
    }

    public Font( string fontFamily, string fontPath, float fontSize, FontStyle fontStyle )
    {
#if NET5
      var typeface = SKTypeface.FromFile( fontPath );
      m_font = new SKFont( typeface, fontSize );
#else
      var systemFontStyle = System.Drawing.FontStyle.Regular;
      if( ( fontStyle & FontStyle.Bold ) == FontStyle.Bold )
      {
        systemFontStyle |= System.Drawing.FontStyle.Bold;
      }
      if( ( fontStyle & FontStyle.Italic ) == FontStyle.Italic )
      {
        systemFontStyle |= System.Drawing.FontStyle.Italic;
      }
      if( ( fontStyle & FontStyle.Underline ) == FontStyle.Underline )
      {
        systemFontStyle |= System.Drawing.FontStyle.Underline;
      }
      if( ( fontStyle & FontStyle.Strikeout ) == FontStyle.Strikeout )
      {
        systemFontStyle |= System.Drawing.FontStyle.Strikeout;
      }

      var fontCollection = new PrivateFontCollection();
      fontCollection.AddFontFile( fontPath );
      if( fontCollection.Families.Length < 0 )
      {
        throw new InvalidOperationException( "No font family found when loading font." );
      }

      m_font = new System.Drawing.Font( fontCollection.Families[ 0 ], System.Convert.ToSingle( fontSize ), systemFontStyle, GraphicsUnit.Pixel );
#endif
    }


    #endregion

    #region Properties

    #region Name

    public string Name
    {
      get
      {
#if NET5
        return m_font.Typeface.FamilyName;
#else
        return m_font.Name;
#endif
      }
    }

    #endregion

    #region Size

    public float Size
    {
      get 
      {
        return m_font.Size;
      }      
    }

    #endregion

    #region Style

    public FontStyle Style
    {
      get
      {
        var result = FontStyle.Regular;

#if NET5        
        if( m_font.Typeface.FontWeight == (int)SKFontStyleWeight.Bold )
        {
          result |= FontStyle.Bold;
        }
        if( m_font.Typeface.FontSlant == SKFontStyleSlant.Italic )
        {
          result |= FontStyle.Italic;
        }
        if( m_isUnderline )
        {
          result |= FontStyle.Underline;
        }
        if ( m_isStrikeout )
        {
          result |= FontStyle.Strikeout;
        }       
#else
        var fontStyle = m_font.Style;
        if( (fontStyle & System.Drawing.FontStyle.Bold) == System.Drawing.FontStyle.Bold )
        {
          result |= FontStyle.Bold;
        }
        if( ( fontStyle & System.Drawing.FontStyle.Italic ) == System.Drawing.FontStyle.Italic )
        {
          result |= FontStyle.Italic;
        }
        if( ( fontStyle & System.Drawing.FontStyle.Underline ) == System.Drawing.FontStyle.Underline )
        {
          result |= FontStyle.Underline;
        }
        if( ( fontStyle & System.Drawing.FontStyle.Strikeout ) == System.Drawing.FontStyle.Strikeout )
        {
          result |= FontStyle.Strikeout;
        }
#endif

        return result;
      }
    }

    #endregion

    #region Value

#if NET5
    public SKFont Value
#else
    public System.Drawing.Font Value
#endif
    {
      get
      {
        return m_font;
      }
    }

    #endregion

    #endregion

    #region Static Methods

    public static float GetTextWidth( string text, string fontName, float fontSize )
    {
#if NET5
      var paint = Font.GetSKPaintObject( fontName, fontSize );

      return paint.MeasureText( text );
#else
      if( Font.Graphics == null )
      {
        Font.SetGraphicObject();
      }

      var defaultFont = new System.Drawing.Font( fontName, fontSize, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point );

      var size = Font.Graphics.MeasureString( text, defaultFont, int.MaxValue, StringFormat.GenericTypographic );
      return size.Width;
#endif
    }

    public static float GetTextHeight( string text, string fontName, float fontSize, int columnWidthInPixels )
    {
#if NET5
      var paint = Font.GetSKPaintObject( fontName, fontSize );

      // Split the text into words
      var words = text.Split( ' ' );

      var lines = new List<string>();
      string currentLine = "";

      // Break text into lines that fit within the max width
      foreach( string word in words )
      {
        string testLine = string.IsNullOrEmpty( currentLine ) ? word : currentLine + " " + word;
        float lineWidth = paint.MeasureText( testLine );

        if( lineWidth <= columnWidthInPixels )
        {
          currentLine = testLine;
        }
        else
        {
          lines.Add( currentLine );
          currentLine = word;
        }
      }

      // Add the last line
      if( !string.IsNullOrEmpty( currentLine ) )
      {
        lines.Add( currentLine );
      }

      // Retrieve font metrics from the paint object
      var fontMetrics = paint.FontMetrics;

      // Calculate the height of a single line
      var lineHeight = fontMetrics.Descent - fontMetrics.Ascent;

      // Calculate the total height for all lines
      var totalHeight = lineHeight * lines.Count;

      // Add leading (spacing between lines), if applicable
      totalHeight += fontMetrics.Leading * ( lines.Count - 1 );

      return totalHeight;
#else
      if( Font.Graphics == null )
      {
        Font.SetGraphicObject();
      }

      var defaultFont = new System.Drawing.Font( fontName, fontSize, System.Drawing.FontStyle.Regular, GraphicsUnit.Point );

      var format = new StringFormat()
      {
        FormatFlags = StringFormatFlags.LineLimit, // Ensure text is measured within the limit
      };

      var size = Font.Graphics.MeasureString( text, defaultFont, columnWidthInPixels, format );
      return size.Height;
#endif
    }

    public static void GetFontImage( Font font, out byte[] image, out int imageSize, out bool corrupted )
    {
      image = null;
      imageSize = 0;
      corrupted = false;

#if NET5
      var fontStream = font.Value.Typeface.OpenStream();
      imageSize = fontStream.Length;
      if( imageSize > 0 )
      {
        image = new byte[ fontStream.Length ];
        fontStream.Read( image, imageSize );
      }
      else
      {
        image = null;
        imageSize = 0;
        corrupted = true;
      }
#else
			IntPtr hWnd = new IntPtr(0);
			IntPtr hDC = GetDC(hWnd);
			IntPtr hFont = font.Value.ToHfont();
			IntPtr oldObj = SelectObject(hDC, hFont);

			imageSize = (int)GetFontData(hDC, 0, 0, null, 0);

			if (imageSize > 0)
			{
				image = new byte[imageSize];
				GetFontData(hDC, 0, 0, image, imageSize);
			}
			else
			{
				image = null;
				imageSize = 0;
				corrupted = true;
			}

			SelectObject(hDC, oldObj);
			DeleteObject(hFont);
			ReleaseDC(hWnd, hDC);
#endif
    }

    public static string FromFile( string file )
    {
#if NET5
      return SKTypeface.FromFile( file ).FamilyName;
#else
      var fontCollection = new PrivateFontCollection();
      fontCollection.AddFontFile( file );
      if( fontCollection.Families.Length < 0 )
        throw new InvalidOperationException( "No font familiy found when loading font." );

      return fontCollection.Families[ 0 ].Name;
#endif
    }

#endregion

    #region Public Methods

    public float GetHeight( out float ascent, out float descent )
    {
#if NET5
      var paint = new SKPaint();
      paint.Typeface = Font.GetSKTypeface( m_font.Typeface.FamilyName, (SKFontStyleWeight)m_font.Typeface.FontWeight, (SKFontStyleWidth)m_font.Typeface.FontWidth, m_font.Typeface.FontSlant );
      paint.TextSize = m_font.Size;
      var metrics = paint.FontMetrics;

      ascent = Math.Abs( metrics.Ascent );
      descent = Math.Abs( metrics.Descent );
      var fontHeight = ascent + descent;

      return fontHeight;
#else
      int cellAscent = m_font.FontFamily.GetCellAscent( m_font.Style );
      int cellDescent = m_font.FontFamily.GetCellDescent( m_font.Style );
      int cellHeight = cellAscent + cellDescent;
      int emHeight = m_font.FontFamily.GetEmHeight( m_font.Style );
      int lineSpacing = m_font.FontFamily.GetLineSpacing( m_font.Style );

      ascent = ( m_font.Size * cellAscent / emHeight );
      descent = ( m_font.Size * cellDescent / emHeight );
      return ( m_font.Size * lineSpacing / emHeight );
#endif
    }

    public void Dispose()
    {
      m_font.Dispose();
    }

    #endregion

    #region Private Methods

#if NET5
    private static SKPaint GetSKPaintObject( string fontName, float fontSize )
    {
      return new SKPaint
      {
        TextSize = fontSize,
        Typeface = Font.GetSKTypeFace( fontName ),
        IsAntialias = true,
        TextScaleX = 1.33f,  
      };
    }

    private static SKTypeface GetSKTypeFace( string fontFamily )
    {
      var typeface = SKTypeface.FromFamilyName( fontFamily );
      typeface = Font.ValidateSKTypeface( typeface, fontFamily );     

      return typeface;
    }

    private static SKTypeface GetSKTypeface( string fontFamily, SKFontStyleWeight typefaceWeight, SKFontStyleWidth fontStyleWidth, SKFontStyleSlant fontStyleSlant )
    {
      var typeface = SKTypeface.FromFamilyName( fontFamily, typefaceWeight, fontStyleWidth, fontStyleSlant );
      typeface = Font.ValidateSKTypeface( typeface, fontFamily );

      return typeface;
    }

    private static SKTypeface ValidateSKTypeface( SKTypeface typeface, string fontFamily )
    {
      // Make sure the non-Windows environment will use a known font for PDF Conversion.
      if( (typeface == null) || (typeface.FamilyName != fontFamily) )
      {
        var fontWeight = (typeface != null) ? typeface.FontWeight : 400;
        var fontWidth = ( typeface != null ) ? typeface.FontWidth : 5;
        var fontSlant = ( typeface != null ) ? typeface.FontSlant : SKFontStyleSlant.Upright;

        if( RuntimeInformation.IsOSPlatform( OSPlatform.Windows ) )
        {
          return SKTypeface.FromFamilyName( "Arial", fontWeight, fontWidth, fontSlant );
        }
        else if( RuntimeInformation.IsOSPlatform( OSPlatform.OSX ) )
        {
          return SKTypeface.FromFamilyName( "Arial", fontWeight, fontWidth, fontSlant );
        }
        else if( RuntimeInformation.IsOSPlatform( OSPlatform.Linux ) )
        {
          var newTypeface = SKTypeface.FromFamilyName( "Ubuntu Sans", fontWeight, fontWidth, fontSlant );
          if( newTypeface == null )
          {
            newTypeface = SKTypeface.FromFamilyName( "DejaVu Sans", fontWeight, fontWidth, fontSlant );
          }
          if( newTypeface == null )
          {
            newTypeface = SKTypeface.FromFamilyName( "Cantarell", fontWeight, fontWidth, fontSlant );
          }
          if( newTypeface == null )
          {
            newTypeface = SKTypeface.FromFamilyName( "Noto Sans", fontWeight, fontWidth, fontSlant );
          }
          if( newTypeface == null )
            throw new InvalidDataException( "Unknown system font under Linux." );

          return newTypeface;
        }
        else if( RuntimeInformation.IsOSPlatform( OSPlatform.Create( "ANDROID" ) ) )
        {
          var newTypeface = SKTypeface.FromFamilyName( "sans-serif", fontWeight, fontWidth, fontSlant );
          if( newTypeface == null )
            throw new InvalidDataException("Unknown system font under Android.");

          return newTypeface;
        }
        else if( RuntimeInformation.IsOSPlatform( OSPlatform.Create( "IOS" ) ) )
        {
          return SKTypeface.FromFamilyName( "Arial", fontWeight, fontWidth, fontSlant );
        }
        else
          throw new Exception( "Unknown OS. Can't set a default font to use. Please specify the font to use." );
      }

      return typeface;
    }
#else
    private static void SetGraphicObject()
    {
      var fakeImage = new System.Drawing.Bitmap( 1, 1 );
      Font.Graphics = Graphics.FromImage( fakeImage );
      Font.Graphics.TextRenderingHint = TextRenderingHint.AntiAlias;
      Font.Graphics.PageUnit = GraphicsUnit.Pixel;
    }
#endif

    #endregion
  }
}
