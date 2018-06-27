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
using System.Linq;

namespace Xceed.Words.NET.Examples
{
  public class HyperlinkSample
  {
    #region Private Members

    private const string HyperlinkSampleOutputDirectory = Program.SampleDirectory + @"Hyperlink\Output\";

    #endregion

    #region Constructors

    static HyperlinkSample()
    {
      if( !Directory.Exists( HyperlinkSample.HyperlinkSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( HyperlinkSample.HyperlinkSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Insert/Add/Remove hyperlinks from paragraphs. 
    /// </summary>
    public static void Hyperlinks()
    {
      Console.WriteLine( "\tHyperlinks()" );

      // Create a document
      using( DocX document = DocX.Create( HyperlinkSample.HyperlinkSampleOutputDirectory + @"Hyperlinks.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Insert/Remove Hyperlinks" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add an Hyperlink into this document.
        var h = document.AddHyperlink( "google", new Uri( "http://www.google.com" ) );

        // Add a paragraph.
        var p = document.InsertParagraph( "The  hyperlink has been inserted in this paragraph." );
        // insert an hyperlink at specific index in this paragraph.
        p.InsertHyperlink( h, 4 );
        p.SpacingAfter( 40d );

        // Get the first hyperlink in the document.
        var hyperlink = document.Hyperlinks.FirstOrDefault();
        if( hyperlink != null )
        {
          // Modify its text and Uri.
          hyperlink.Text = "xceed";
          hyperlink.Uri = new Uri( "http://www.xceed.com/" );
        }

        // Add an Hyperlink to this document.
        var h2 = document.AddHyperlink( "xceed", new Uri( "http://www.xceed.com/" ) );
        // Add a paragraph.
        var p2 = document.InsertParagraph( "A formatted hyperlink has been added at the end of this paragraph: " );
        // Append an hyperlink to a paragraph.
        p2.AppendHyperlink( h2 ).Color( Color.Blue ).UnderlineStyle( UnderlineStyle.singleLine );
        p2.Append( "." ).SpacingAfter( 40d );

        // Create a bookmark anchor.
        var bookmarkAnchor = "bookmarkAnchor";
        // Add an Hyperlink to this document pointing to a bookmark anchor.
        var h3 = document.AddHyperlink( "Special Data", bookmarkAnchor );
        // Add a paragraph.
        var p3 = document.InsertParagraph( "An hyperlink pointing to a bookmark of this Document has been added at the end of this paragraph: " );
        // Append an hyperlink to a paragraph.
        p3.AppendHyperlink( h3 ).Color( Color.Red ).UnderlineStyle( UnderlineStyle.singleLine );
        p3.Append( "." ).SpacingAfter( 40d );

        // Add an Hyperlink to this document.
        var h4 = document.AddHyperlink( "microsoft", new Uri( "http://www.microsoft.com" ) );
        // Add a paragraph
        var p4 = document.InsertParagraph( "The hyperlink from this paragraph has been removed. " );
        // Append an hyperlink to a paragraph.
        p4.AppendHyperlink( h4 ).Color( Color.Green ).UnderlineStyle( UnderlineStyle.singleLine ).Italic();

        // Remove the first hyperlink of paragraph 4.
        p4.RemoveHyperlink( 0 );

        // Add a paragraph.
        var p5 = document.InsertParagraph( "This is a paragraph containing a" );
        // Add a bookmark into the paragraph by setting its bookmark anchor.
        p5.AppendBookmark( bookmarkAnchor );
        p5.Append( " bookmark " );
        p5.Append( "referenced by a hyperlink defined in an earlier paragraph." );
        p5.SpacingBefore( 200d );

        document.Save();
        Console.WriteLine( "\tCreated: Hyperlinks.docx\n" );
      }
    }

    #endregion
  }
}
