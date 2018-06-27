/***************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2017 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  *************************************************************************************/
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;

namespace Xceed.Words.NET.Examples
{
  public class ParagraphSample
  {
    #region Private Members

    private static Dictionary<string, string> _replacePatterns = new Dictionary<string, string>()
    {
        { "COST", "$13.95" },
    };

    private const string ParagraphSampleOutputDirectory = Program.SampleDirectory + @"Paragraph\Output\";

    #endregion

    #region Constructors

    static ParagraphSample()
    {
      if( !Directory.Exists( ParagraphSample.ParagraphSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( ParagraphSample.ParagraphSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Create a document with formatted paragraphs.
    /// </summary>
    public static void SimpleFormattedParagraphs()
    {
      Console.WriteLine( "\tSimpleFormattedParagraphs()" );

      // Create a new document.
      using( DocX document = DocX.Create( ParagraphSample.ParagraphSampleOutputDirectory + @"SimpleFormattedParagraphs.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Formatted paragraphs" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert a Paragraph into this document.
        var p = document.InsertParagraph();

        // Append some text and add formatting.
        p.Append( "This is a simple formatted red bold paragraph" )
        .Font( new Font( "Arial" ) )
        .FontSize( 25 )
        .Color( Color.Red )
        .Bold()
        .Append( " containing a blue italic text." ).Font( new Font( "Times New Roman" ) ).Color( Color.Blue ).Italic()
        .SpacingAfter( 40 );

        // Insert another Paragraph into this document.
        var p2 = document.InsertParagraph();

        // Append some text and add formatting.
        p2.Append( "This is a formatted paragraph using spacing," )
        .Font( new Font( "Courier New" ) )
        .FontSize( 10 )
        .Italic()
        .Spacing( 5 )
        .Append( "highlight" ).Highlight( Highlight.yellow ).UnderlineColor( Color.Blue ).CapsStyle( CapsStyle.caps )
        .Append( " and strike through." ).StrikeThrough( StrikeThrough.strike )
        .SpacingAfter( 40 );

        // Insert another Paragraph into this document.
        var p3 = document.InsertParagraph();

        // Append some text with 2 TabStopPositions.
        p3.InsertTabStopPosition( Alignment.center, 216f, TabStopPositionLeader.dot )
        .InsertTabStopPosition( Alignment.right, 432f, TabStopPositionLeader.dot )
        .Append( "Text with TabStopPositions on Left\tMiddle\tand Right" )
        .SpacingAfter( 40 );

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: SimpleFormattedParagraphs.docx\n" );
      }
    }

    /// <summary>
    /// Create a document and add a paragraph with all its lines on a single page.
    /// </summary>
    public static void ForceParagraphOnSinglePage()
    {
      Console.WriteLine( "\tForceParagraphOnSinglePage()" );

      // Create a new document.
      using( DocX document = DocX.Create( ParagraphSample.ParagraphSampleOutputDirectory + @"ForceParagraphOnSinglePage.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Prevent paragraph split" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a Paragraph that will appear on 1st page.
        var p = document.InsertParagraph( "This is a paragraph on first page.\nLine2\nLine3\nLine4\nLine5\nLine6\nLine7\nLine8\nLine9\nLine10\nLine11\nLine12\nLine13\nLine14\nLine15\nLine16\nLine17\nLine18\nLine19\nLine20\nLine21\nLine22\nLine23\nLine24\nLine25\n" );
        p.FontSize(15).SpacingAfter( 30 );

        // Create a Paragraph where all its lines will appear on a same page.
        var p2 = document.InsertParagraph( "This is a paragraph where all its lines are on the same page. The paragraph does not split on 2 pages.\nLine2\nLine3\nLine4\nLine5\nLine6\nLine7\nLine8\nLine9\nLine10" );
        p2.SpacingAfter( 30 );

        // Indicate that all the paragraph's lines will be on the same page
        p2.KeepLinesTogether();

        // Create a Paragraph that will appear on 2nd page.
        var p3 = document.InsertParagraph( "This is a paragraph on second page.\nLine2\nLine3\nLine4\nLine5\nLine6\nLine7\nLine8\nLine9\nLine10" );       

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: ForceParagraphOnSinglePage.docx\n" );
      }
    }

    /// <summary>
    /// Create a document and add a paragraph with all its lines on the same page as the next paragraph.
    /// </summary>
    public static void ForceMultiParagraphsOnSinglePage()
    {
      Console.WriteLine( "\tForceMultiParagraphsOnSinglePage()" );

      // Create a new document.
      using( DocX document = DocX.Create( ParagraphSample.ParagraphSampleOutputDirectory + @"ForceMultiParagraphsOnSinglePage.docx" ) )
      {
        // Add a title.
        document.InsertParagraph( "Keeps Paragraphs on same page" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a Paragraph that will appear on 1st page.
        var p = document.InsertParagraph( "This is a paragraph on first page.\nLine2\nLine3\nLine4\nLine5\nLine6\nLine7\nLine8\nLine9\nLine10\nLine11\nLine12\nLine13\nLine14\nLine15\nLine16\nLine17\nLine18\nLine19\nLine20\nLine21\nLine22\n" );
        p.FontSize( 15 ).SpacingAfter( 30 );

        // Create a Paragraph where all its lines will appear on a same page as the next paragraph.
        var p2 = document.InsertParagraph( "This is a paragraph where all its lines are on the same page as the next paragraph.\nLine2\nLine3\nLine4\nLine5\nLine6\nLine7\nLine8\nLine9\nLine10" );
        p2.SpacingAfter( 30 );        

        // Indicate that this paragraph will be on the same page as the next paragraph.
        p2.KeepWithNextParagraph();

        // Create a Paragraph that will appear on 2nd page.
        var p3 = document.InsertParagraph( "This is a paragraph on second page.\nLine2\nLine3\nLine4\nLine5\nLine6\nLine7\nLine8\nLine9\nLine10" );

        // Indicate that all this paragraph's lines will be on the same page.
        p3.KeepLinesTogether();

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: ForceMultiParagraphsOnSinglePage.docx\n" );
      }
    }

    /// <summary>
    /// Create a document and insert, remove and replace texts.
    /// </summary>
    public static void TextActions()
    {
      Console.WriteLine( "\tTextActions()" );

      // Create a new document.
      using( DocX document = DocX.Create( ParagraphSample.ParagraphSampleOutputDirectory + @"TextActions.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Insert, remove and replace text" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a paragraph and insert text.
        var p1 = document.InsertParagraph( "In this paragraph we insert a comma, a colon and " );
        // Add a comma at index 17. 
        p1.InsertText( 17, "," );
        // Add a colon at index of character 'd'. 
        p1.InsertText( p1.Text.IndexOf( "d" ) + 1, ": " );
        // Add "a name" at the end of p1.Text.
        p1.InsertText( p1.Text.Length, "a name." );

        p1.SpacingAfter( 30 );

        // Create a paragraph and insert text.
        var p2 = document.InsertParagraph( "In this paragraph, we remove a mistaken word and a comma." );
        // Remove the word "mistaken".
        p2.RemoveText( 31, 9 );
        // Remove the comma sign.
        p2.RemoveText( p2.Text.IndexOf( "," ), 1 );

        p2.SpacingAfter( 30 );

        // Create a paragraph and insert text.
        var p3 = document.InsertParagraph( "In this paragraph, we replace an complex word with an easier one and spaces with hyphens." );
        // Replace the "complex" word with "easy" word.
        p3.ReplaceText( "complex", "easy" );
        // Replace the spaces with tabs
        p3.ReplaceText( " ", "--" );

        p3.SpacingAfter( 30 );

        // Create a paragraph and insert text.
        var p4 = document.InsertParagraph( "In this paragraph, we replace a word by using a handler: <COST>." );
        // Replace "<COST>" with "$13.95" using an handler
        p4.ReplaceText( "<(.*?)>", ReplaceTextHandler, false, RegexOptions.IgnoreCase, null, new Formatting() );

        p4.SpacingAfter( 30 );

        // Insert another Paragraph into this document.
        var p5 = document.InsertParagraph();

        // Append some text with track changes
        p5.Append( "This is a paragraph where tracking of modifications is used." );
        p5.ReplaceText( "modifications", "changes", true );

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: TextActions.docx\n" );
      }
    }

    /// <summary>
    /// Set different Heading type for a Paragraph.
    /// </summary>
    public static void Heading()
    {
      Console.WriteLine( "\tHeading()" );

      // Create a document.
      using( DocX document = DocX.Create( ParagraphSample.ParagraphSampleOutputDirectory + @"Heading.docx" ) )
      {
        // Add a title.
        document.InsertParagraph( "Heading types" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        var headingTypes = Enum.GetValues( typeof( HeadingType ) );

        foreach( HeadingType heading in headingTypes )
        {
          // Set a text containing the current Heading type.
          var text = string.Format( "This Paragraph is using \"{0}\" heading type.", heading.EnumDescription() );
          // Add a paragraph.
          var p = document.InsertParagraph().AppendLine( text );
          // Set the paragraph's heading type.
          p.Heading( heading );
        }

        document.Save();
        Console.WriteLine( "\tCreated: Heading.docx\n" );
      }
    }

    #endregion

    #region Private Methods

    private static string ReplaceTextHandler( string findStr )
    {
      if( _replacePatterns.ContainsKey( findStr ) )
      {
        return _replacePatterns[ findStr ];
      }
      return findStr;
    }

    #endregion
  }
}
