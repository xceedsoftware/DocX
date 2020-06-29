/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Line Sample Application
Copyright (c) 2009-2020 - Xceed Software Inc.

This application demonstrates how to insert horizontal lines when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/
using System;
using System.Drawing;
using System.IO;
using Xceed.Document.NET;

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
        p.Append( "This is a paragraph with a single bottom line." ).Font( new Xceed.Document.NET.Font( "Arial" ) ).FontSize( 15 );
        p.InsertHorizontalLine( HorizontalBorderPosition.bottom, BorderStyle.Tcbs_single );
        p.SpacingAfter( 20 );

        // Add a paragraph with a double green line.
        var p2 = document.InsertParagraph();
        p2.Append( "This is a paragraph with a double bottom colored line." ).Font( new Xceed.Document.NET.Font( "Arial" ) ).FontSize( 15 );
        p2.InsertHorizontalLine( HorizontalBorderPosition.bottom, BorderStyle.Tcbs_double, 6, 1, Color.Green );
        p2.SpacingAfter( 20 );

        // Add a paragraph with a triple red line.
        var p3 = document.InsertParagraph();
        p3.Append( "This is a paragraph with a triple bottom colored line." ).Font( new Xceed.Document.NET.Font( "Arial" ) ).FontSize( 15 );
        p3.InsertHorizontalLine( HorizontalBorderPosition.bottom, BorderStyle.Tcbs_triple, 6, 1, Color.Red );
        p3.SpacingAfter( 20 );

        // Add a paragraph with a single spaced line.
        var p4 = document.InsertParagraph();
        p4.Append( "This is a paragraph with a spaced bottom line." ).Font( new Xceed.Document.NET.Font( "Arial" ) ).FontSize( 15 );
        p4.InsertHorizontalLine( HorizontalBorderPosition.bottom, BorderStyle.Tcbs_single, 6, 12 );
        p4.SpacingAfter( 20 );

        // Add a paragraph with a single large line.
        var p5 = document.InsertParagraph();
        p5.Append( "This is a paragraph with a large bottom line." ).Font( new Xceed.Document.NET.Font( "Arial" ) ).FontSize( 15 );
        p5.InsertHorizontalLine( HorizontalBorderPosition.bottom, BorderStyle.Tcbs_single, 25 );
        p5.SpacingAfter( 60 );

        // Add a paragraph with a wave blue top line.
        var p6 = document.InsertParagraph();
        p6.Append( "This is a paragraph with a wave blue top line." ).Font( new Xceed.Document.NET.Font( "Arial" ) ).FontSize( 15 );
        p6.InsertHorizontalLine( HorizontalBorderPosition.top, BorderStyle.Tcbs_wave, 6, 1, Color.FromArgb( 0, 0, 255 ) );
        p5.SpacingAfter( 20 );

        document.Save();
        Console.WriteLine( "\tCreated: InsertHorizontalLine.docx\n" );
      }
    }

    #endregion
  }
}
