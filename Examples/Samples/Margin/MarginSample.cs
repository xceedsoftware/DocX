/***************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2017 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  *************************************************************************************/

using System;
using System.IO;

namespace Xceed.Words.NET.Examples
{
  public class MarginSample
  {
    #region Private Members

    private const string MarginSampleOutputDirectory = Program.SampleDirectory + @"Margin\Output\";

    #endregion

    #region Constructors

    static MarginSample()
    {
      if( !Directory.Exists( MarginSample.MarginSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( MarginSample.MarginSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Modify the direction of text in a paragraph or document.
    /// </summary>
    public static void SetDirection()
    {
      Console.WriteLine( "\tSetDirection()" );

      // Create a document.
      using( DocX document = DocX.Create( MarginSample.MarginSampleOutputDirectory + @"SetDirection.docx" ) )
      {
        // Add a title.
        document.InsertParagraph( "Modify direction of paragraphs" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add first paragraph.
        var p = document.InsertParagraph("This is the first paragraph.");
        p.SpacingAfter( 30 );

        // Add second paragraph.
        var p2 = document.InsertParagraph( "This is the second paragraph." );
        p2.SpacingAfter( 30 );
        // Make this Paragraph flow right to left. Default is left to right.
        p2.Direction = Direction.RightToLeft;

        // Add third paragraph.
        var p3 = document.InsertParagraph( "This is the third paragraph." );
        p3.SpacingAfter( 30 );

        // Add fourth paragraph.
        var p4 = document.InsertParagraph( "This is the fourth paragraph." );
        p4.SpacingAfter( 30 );

        // To modify the direction of each paragraph in a document, just set the direction on the document.
        document.SetDirection( Direction.RightToLeft );

        document.Save();
        Console.WriteLine( "\tCreated: SetDirection.docx\n" );
      }
    }

    /// <summary>
    /// Add indentations on paragraphs.
    /// </summary>
    public static void Indentation()
    {
      Console.WriteLine( "\tIndentation()" );

      // Create a document.
      using( DocX document = DocX.Create( MarginSample.MarginSampleOutputDirectory + @"Indentation.docx" ) )
      {
        // Add a title.
        document.InsertParagraph( "Paragraph indentation" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Set a smaller page width.
        document.PageWidth = 250f;

        // Add the first paragraph.
        var p = document.InsertParagraph( "This is the first paragraph. It doesn't contain any indentation." );
        p.SpacingAfter( 30 );

        // Add the second paragraph.
        var p2 = document.InsertParagraph( "This is the second paragraph. It contains an indentation on the first line." );        
        // Indent only the first line of the Paragraph.
        p2.IndentationFirstLine = 1.0f;
        p2.SpacingAfter( 30 );

        // Add the third paragraph.
        var p3 = document.InsertParagraph( "This is the third paragraph. It contains an indentation on all the lines except the first one." );
        // Indent all the lines of the Paragraph, except the first.
        p3.IndentationHanging = 1.0f;     

        document.Save();
        Console.WriteLine( "\tCreated: Indentation.docx\n" );
      }
    }

    /// <summary>
    /// Add margins for a document.
    /// </summary>
    public static void Margins()
    {
      Console.WriteLine( "\tMargins()" );

      // Create a document.
      using( DocX document = DocX.Create( MarginSample.MarginSampleOutputDirectory + @"Margins.docx" ) )
      {
        // Add a title.
        document.InsertParagraph( "Document margins" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Set the page width to be smaller.
        document.PageWidth = 350f;

        // Set the document margins.
        document.MarginLeft = 85f;
        document.MarginRight = 85f;
        document.MarginTop = 0f;
        document.MarginBottom = 50f;

        // Add a paragraph. It will be affected by the document margins.
        var p = document.InsertParagraph("This is a paragraph from a document with a left margin of 85, a right margin of 85, a top margin of 0 and a bottom margin of 50.");

        document.Save();
        Console.WriteLine( "\tCreated: Margins.docx\n" );
      }
    }

    #endregion
  }
}
