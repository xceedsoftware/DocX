/***************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2017 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  *************************************************************************************/
using System;
using System.Drawing;
using System.IO;

namespace Xceed.Words.NET.Examples
{
  public class LineSample
  {
    #region Private Members

    private const string LineSampleOutputDirectory = Program.SampleDirectory + @"Line\Output\";

    #endregion

    #region Constructors

    static LineSample()
    {
      if( !Directory.Exists( LineSample.LineSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( LineSample.LineSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Create a document and add different lines under paragraphs.
    /// </summary>
    public static void InsertHorizontalLine()
    {
      Console.WriteLine( "\tInsertHorizontalLine()" );

      using( var document = DocX.Create( LineSample.LineSampleOutputDirectory + @"InsertHorizontalLine.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Adding top or bottom Horizontal lines" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a paragraph with a single line.
        var p = document.InsertParagraph();
        p.Append( "This is a paragraph with a single bottom line." ).Font( new Font( "Arial" ) ).FontSize( 15 );
        p.InsertHorizontalLine( HorizontalBorderPosition.bottom, "single", 6, 1, "auto" );
        p.SpacingAfter( 20 );

        // Add a paragraph with a double green line.
        var p2 = document.InsertParagraph();
        p2.Append( "This is a paragraph with a double bottom colored line." ).Font( new Font( "Arial" ) ).FontSize( 15 );
        p2.InsertHorizontalLine( HorizontalBorderPosition.bottom, "double", 6, 1, "green" );
        p2.SpacingAfter( 20 );

        // Add a paragraph with a triple red line.
        var p3 = document.InsertParagraph();
        p3.Append( "This is a paragraph with a triple bottom colored line." ).Font( new Font( "Arial" ) ).FontSize( 15 );
        p3.InsertHorizontalLine( HorizontalBorderPosition.bottom, "triple", 6, 1, "red" );
        p3.SpacingAfter( 20 );

        // Add a paragraph with a single spaced line.
        var p4 = document.InsertParagraph();
        p4.Append( "This is a paragraph with a spaced bottom line." ).Font( new Font( "Arial" ) ).FontSize( 15 );
        p4.InsertHorizontalLine( HorizontalBorderPosition.bottom, "single", 6, 12, "auto" );
        p4.SpacingAfter( 20 );

        // Add a paragraph with a single large line.
        var p5 = document.InsertParagraph();
        p5.Append( "This is a paragraph with a large bottom line." ).Font( new Font( "Arial" ) ).FontSize( 15 );
        p5.InsertHorizontalLine( HorizontalBorderPosition.bottom, "single", 25, 1, "auto" );
        p5.SpacingAfter( 60 );

        // Add a paragraph with a single blue top line.
        var p6 = document.InsertParagraph();
        p6.Append( "This is a paragraph with a blue top line." ).Font( new Font( "Arial" ) ).FontSize( 15 );
        p6.InsertHorizontalLine( HorizontalBorderPosition.top, "single", 6, 1, "blue" );
        p5.SpacingAfter( 20 );

        document.Save();
        Console.WriteLine( "\tCreated: InsertHorizontalLine.docx\n" );
      }
    }

    #endregion
  }
}
