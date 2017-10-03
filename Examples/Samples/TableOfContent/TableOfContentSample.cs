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
  public class TableOfContentSample
  {
    #region Private Members

    private const string TableOfContentSampleOutputDirectory = Program.SampleDirectory + @"TableOfContent\Output\";

    #endregion

    #region Constructors

    static TableOfContentSample()
    {
      if( !Directory.Exists( TableOfContentSample.TableOfContentSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( TableOfContentSample.TableOfContentSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Add a Table of content to a document.
    /// </summary>
    public static void InsertTableOfContent()
    {
      Console.WriteLine( "\tInsertTableOfContent()" );

      // Creates a document
      using( DocX document = DocX.Create( TableOfContentSample.TableOfContentSampleOutputDirectory + @"InsertTableOfContent.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Insert Table of content" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert a table of content with a page break.
        document.InsertTableOfContents( "Teams", TableOfContentsSwitches.O | TableOfContentsSwitches.U | TableOfContentsSwitches.Z | TableOfContentsSwitches.H );
        document.InsertSectionPageBreak();

        // Create a paragraph and fill it in method AddTeams().
        var p = document.InsertParagraph();
        var rosters = TableOfContentSample.AddTeams( p );
        document.InsertParagraph( rosters );

        document.Save();
        Console.WriteLine( "\tCreated: InsertTableOfContent.docx\n" );
      }
    }

    /// <summary>
    /// Add a Table of content to a document by inserting it just before a reference paragraph.
    /// </summary>
    public static void InsertTableOfContentWithReference()
    {
      Console.WriteLine( "\tInsertTableOfContentWithReference()" );

      // Create a document.
      using( DocX document = DocX.Create( TableOfContentSample.TableOfContentSampleOutputDirectory + @"InsertTableOfContentWithReference.docx" ) )
      {
        // Add a title.
        document.InsertParagraph( "Insert Table of content with reference" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add an intro paragraph.
        var intro = document.InsertParagraph( "This page will show the team rosters of the American League East Division." );
        intro.SpacingAfter( 150d );

        // Create a paragraph and fill it in method AddTeams().
        var p = document.InsertParagraph();
        var rosters = TableOfContentSample.AddTeams( p );
        document.InsertParagraph( rosters );

        // Insert a table of content with a page break just before the paragraph p.
        document.InsertTableOfContents( p, 
                                        "Teams",
                                        TableOfContentsSwitches.O | TableOfContentsSwitches.U | TableOfContentsSwitches.Z | TableOfContentsSwitches.H,
                                        "Heading4" );

        document.Save();
        Console.WriteLine( "\tCreated: InsertTableOfContentWithReference.docx\n" );
      }
    }

    #endregion

    #region Private Methods

    private static Paragraph AddTeams( Paragraph paragraph )
    {
      // Add a title paragraph.
      var title = paragraph.InsertParagraphAfterSelf( "Team Rosters" ).Bold().FontSize( 20 ).SpacingAfter( 50d );
      title.Alignment = Alignment.center;

      // Add the content paragraphs and set a style for the Table of Content to recognize them.
      var p = title.InsertParagraphAfterSelf( "Boston Red Sox" ).Bold().FontSize( 15 ).SpacingAfter( 25d );
      p.StyleName = "Heading1";
      var p1 = p.InsertParagraphAfterSelf( "Tom Smith, P" )
                .AppendLine( "Mike Fitzgerald, C" )
                .AppendLine( "Tom Clancy, 1B" )
                .AppendLine( "Kevin Garnet, OF" ).SpacingAfter( 300d );

      var p2 = p1.InsertParagraphAfterSelf( "Tampa Rays" ).Bold().FontSize( 15 ).SpacingAfter( 25d );
      p2.StyleName = "Heading1";
      var p3 = p2.InsertParagraphAfterSelf( "Josh Hernandez, P" )
                 .AppendLine( "Jacob Trouba, C" )
                 .AppendLine( "Jesus Sanchez, 1B" )
                 .AppendLine( "Jose Ria, OF" ).SpacingAfter( 300d );

      var p4 = p3.InsertParagraphAfterSelf( "New York Yankees" ).Bold().FontSize( 15 ).SpacingAfter( 25d );
      p4.StyleName = "Heading1";
      var p5 = p4.InsertParagraphAfterSelf( "Derek Jones, P" )
                 .AppendLine( "Jose Riva, C" )
                 .AppendLine( "Bryan Smith, 1B" )
                 .AppendLine( "Carl Shattern, OF" ).SpacingAfter( 300d );

      var p6 = p5.InsertParagraphAfterSelf( "Baltimore Orioles" ).Bold().FontSize( 15 ).SpacingAfter( 25d );
      p6.StyleName = "Heading1";
      var p7 = p6.InsertParagraphAfterSelf( "Simon Delgar, P" )
                 .AppendLine( "Johnny Helpan, C" )
                 .AppendLine( "Miguel Danregados, 1B" )
                 .AppendLine( "Joe West, OF" ).SpacingAfter( 300d );

      var p8 = p7.InsertParagraphAfterSelf( "Toronto Blue Jays" ).Bold().FontSize( 15 ).SpacingAfter( 25d );
      p8.StyleName = "Heading1";
      var p9 = p8.InsertParagraphAfterSelf( "Samir Endoya, P" )
                 .AppendLine( "Steve Martin, C" )
                 .AppendLine( "Erik Young, 1B" )
                 .AppendLine( "Steve Martinek, OF" );

      return paragraph;
    }

    #endregion
  }
}
