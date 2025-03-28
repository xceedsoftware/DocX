﻿/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Paragraph Sample Application
Copyright (c) 2009-2025 - Xceed Software Inc.

This application demonstrates how to create, format and position paragraphs
when using the API from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Xceed.Document.NET;
using Xceed.Drawing;

namespace Xceed.Words.NET.Examples
{
  public class ParagraphSample
  {
    #region Private Members

    private static Dictionary<string, string> _replacePatterns = new Dictionary<string, string>()
    {
        { "COST", "$13.95" },
    };

    private const string ParagraphSampleResourcesDirectory = Program.SampleDirectory + @"Paragraph\Resources\";
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

    public static void SimpleFormattedParagraphs()
    {
      Console.WriteLine( "\tSimpleFormattedParagraphs()" );

      // Create a new document.
      using( var document = DocX.Create( ParagraphSample.ParagraphSampleOutputDirectory + @"SimpleFormattedParagraphs.docx" ) )
      {
        document.SetDefaultFont( new Xceed.Document.NET.Font( "Arial" ), 15d, Color.Green );
        document.PageBackground = Color.LightGray;
        document.PageBorders = new Borders( new Border( BorderStyle.Tcbs_double, BorderSize.five, 20f, Color.Blue ) );

        // Add a title
        document.InsertParagraph( "Formatted paragraphs" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert a Paragraph into this document.
        var p = document.InsertParagraph();

        // Append some text and add formatting.
        p.Append( "This is a simple formatted red bold paragraph" )
        .Font( new Xceed.Document.NET.Font( "Arial" ) )
        .FontSize( 25 )
        .Color( Color.Red )
        .Bold()
        .Append( " containing a blue italic text." ).Font( new Xceed.Document.NET.Font( "Times New Roman" ) ).Color( Color.Blue ).Italic()
        .SpacingAfter( 40 );

        // Insert another Paragraph into this document.
        var p2 = document.InsertParagraph();

        // Append some text and add formatting.
        p2.Append( "This is a formatted paragraph using spacing, line spacing, " )
        .Font( new Xceed.Document.NET.Font( "Courier New" ) )
        .FontSize( 10 )
        .Italic()
        .Spacing( 5 )
        .SpacingLine( 22 )
        .Append( "highlight" ).Highlight( Highlight.yellow ).UnderlineColor( Color.Blue ).CapsStyle( CapsStyle.caps )
        .Append( " and strike through." ).StrikeThrough( StrikeThrough.strike )
        .SpacingAfter( 40 );

        // Insert another Paragraph into this document.
        var p3 = document.InsertParagraph();

        // Append some text with 2 TabStopPositions.
        p3.InsertTabStopPosition( Alignment.center, 216f, TabStopPositionLeader.dot )
        .InsertTabStopPosition( Alignment.right, 432f, TabStopPositionLeader.dot )
        .Append( "Text with TabStopPositions on Left\tMiddle\tand Right" )
        .FontSize( 11d )
        .SpacingAfter( 40 );

        // Insert another Paragraph into this document.
        var p4 = document.InsertParagraph();
        p4.Append( "This document is using an Arial green default font of size 15. It's also using a double blue page borders and light gray page background." )
          .SpacingAfter( 40 );

        // Append a Paragraph with Shading pattern format.
        var formatting = new Formatting()
        {
          ShadingPattern = new ShadingPattern()
          {
            Fill = Color.Blue,
            Style = PatternStyle.DkUpDiagonal,
            StyleColor = Color.White
          }
        };

        // Insert another Paragraph into this document.
        var p5 = document.InsertParagraph();

        // Append some text and add formatting to the text with shading pattern.
        p5.Append( "This is a simple formatted text with shading pattern", formatting );

        var shadingPattern = new ShadingPattern()
        {
          Fill = Color.Red,
          Style = PatternStyle.DkUpDiagonal,
          StyleColor = Color.White
        };

        // Insert another Paragraph into this document.
        var p6 = document.InsertParagraph();

        // Append some text and add formatting to the paragraph with shading pattern.
        p6.Append( "This is a simple formatted paragraph with shading pattern." )
          .ShadingPattern( shadingPattern, ShadingType.Paragraph )
          .SpacingAfter( 40 );

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: SimpleFormattedParagraphs.docx\n" );
      }
    }

    public static void StyleParagraphs()
    {
      Console.WriteLine( "\tStyleParagraphs()" );

      // Load a template document.
      using( var templateDoc = DocX.Load( ParagraphSample.ParagraphSampleResourcesDirectory + @"TemplateStyledParagraph.docx" ) )
      {
        // Retrieve the wanted styles from templateDoc.
        var styleId1 = DocX.GetParagraphStyleIdFromStyleName( templateDoc, "Heading 1" );  // styleName must match the style name from templateDoc.
        var styleId2 = DocX.GetParagraphStyleIdFromStyleName( templateDoc, "CustomStyle" );// styleName must match the style name from templateDoc.

        // Insert a paragraph with the Heading 1 style.
        var p = templateDoc.InsertParagraph();
        p.InsertText( "This paragraph is using an Heading 1 style." );
        p.StyleId = styleId1;
        p.AppendLine();

        // Insert a paragraph with the CustomStyle style.
        var p2 = templateDoc.InsertParagraph();
        p2.InsertText( "This paragraph is using a custom style." );
        p2.StyleId = styleId2;

        // Save the updated document as StyleParagraphs.docx.
        templateDoc.SaveAs( ParagraphSample.ParagraphSampleOutputDirectory + @"StyleParagraphs.docx" );

        Console.WriteLine( "\tStyleParagraphs.docx\n" );
      }
    }

    public static void ForceParagraphOnSinglePage()
    {
      Console.WriteLine( "\tForceParagraphOnSinglePage()" );

      // Create a new document.
      using( var document = DocX.Create( ParagraphSample.ParagraphSampleOutputDirectory + @"ForceParagraphOnSinglePage.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Prevent paragraph split" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a Paragraph that will appear on 1st page.
        var p = document.InsertParagraph( "This is a paragraph on first page.\nLine2\nLine3\nLine4\nLine5\nLine6\nLine7\nLine8\nLine9\nLine10\nLine11\nLine12\nLine13\nLine14\nLine15\nLine16\nLine17\nLine18\nLine19\nLine20\nLine21\nLine22\nLine23\nLine24\nLine25\n" );
        p.FontSize( 15 ).SpacingAfter( 30 );

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

    public static void ForceMultiParagraphsOnSinglePage()
    {
      Console.WriteLine( "\tForceMultiParagraphsOnSinglePage()" );

      // Create a new document.
      using( var document = DocX.Create( ParagraphSample.ParagraphSampleOutputDirectory + @"ForceMultiParagraphsOnSinglePage.docx" ) )
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

    public static void TextActions()
    {
      Console.WriteLine( "\tTextActions()" );

      // Create a new document.
      using( var document = DocX.Create( ParagraphSample.ParagraphSampleOutputDirectory + @"TextActions.docx" ) )
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
        var stringReplaceTextOptions = new StringReplaceTextOptions() { SearchValue = "complex", NewValue = "easy" };
        p3.ReplaceText( stringReplaceTextOptions );
        // Replace the spaces with hyphens.
        stringReplaceTextOptions = new StringReplaceTextOptions() { SearchValue = " ", NewValue = "--" };
        p3.ReplaceText( stringReplaceTextOptions );

        p3.SpacingAfter( 30 );

        // Create a paragraph and insert text.
        var p4 = document.InsertParagraph( "In this paragraph, we replace a word by using a handler: <COST>." );
        // Replace "<COST>" with "$13.95" using an handler
        var functionReplaceTextOptions = new FunctionReplaceTextOptions()
        {
          FindPattern = "<(.*?)>",
          RegexMatchHandler = ParagraphSample.ReplaceTextHandler,
          RegExOptions = RegexOptions.IgnoreCase
        };
        p4.ReplaceText( functionReplaceTextOptions );

        p4.SpacingAfter( 30 );

        // Insert another Paragraph into this document.
        var p5 = document.InsertParagraph();

        // Append some text with track changes
        p5.Append( "This is a paragraph where tracking of modifications is used." );
        stringReplaceTextOptions = new StringReplaceTextOptions() { SearchValue = "modifications", NewValue = "changes", TrackChanges = true };
        p5.ReplaceText( stringReplaceTextOptions );

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: TextActions.docx\n" );
      }
    }

    public static void Heading()
    {
      Console.WriteLine( "\tHeading()" );

      // Create a document.
      using( var document = DocX.Create( ParagraphSample.ParagraphSampleOutputDirectory + @"Heading.docx" ) )
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

    public static void AddObjectsFromOtherDocument()
    {
      Console.WriteLine( "\tAddObjectsFromOtherDocument()" );

      // Load a template document.
      using( var templateDoc = DocX.Load( ParagraphSample.ParagraphSampleResourcesDirectory + @"Template.docx" ) )
      {
        // Create a document.
        using( var document = DocX.Create( ParagraphSample.ParagraphSampleOutputDirectory + @"AddObjectsFromOtherDocument.docx" ) )
        {
          // Add a title.
          document.InsertParagraph( "Adding objects from another document" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

          // Insert a Paragraph into this document.
          var p = document.InsertParagraph();

          // Append some text and add formatting.
          p.Append( "This is a simple paragraph containing an image added from another document:" )
          .Font( new Xceed.Document.NET.Font( "Arial" ) )
          .FontSize( 15 )
          .SpacingAfter( 40d );

          // Get the image from the other document.
          var pictureToAdd = templateDoc.Pictures.FirstOrDefault();
          // Add the image in the new document.
          var newImage = document.AddImage( pictureToAdd.Stream );
          var newPicture = newImage.CreatePicture( pictureToAdd.Height, pictureToAdd.Width );
          p.AppendPicture( newPicture );

          // Insert a Paragraph into this document.
          var p2 = document.InsertParagraph();

          // Append some text and add formatting.
          p2.Append( "This is a simple paragraph added from another document, keeping its formatting:" )
          .Font( new Xceed.Document.NET.Font( "Arial" ) )
          .FontSize( 15 );

          // Get the paragraph from the other document.
          var paragraphToAdd = templateDoc.Paragraphs.FirstOrDefault( x => x.Text.Contains( "Main" ) );
          // Add the paragraph in the new document.
          p2.InsertParagraphAfterSelf( paragraphToAdd ).SpacingAfter( 40d );

          // Insert a Paragraph into this document.
          var p3 = document.InsertParagraph();

          // Append some text and add formatting.
          p3.Append( "This is a table added from another document, keeping its formatting:" )
          .Font( new Xceed.Document.NET.Font( "Arial" ) )
          .FontSize( 15 );

          // Get the table from the other document.
          var tableToAdd = templateDoc.Tables.FirstOrDefault();
          // Add the table in the new document.
          p3.InsertTableAfterSelf( tableToAdd );

          document.Save();
        }
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
