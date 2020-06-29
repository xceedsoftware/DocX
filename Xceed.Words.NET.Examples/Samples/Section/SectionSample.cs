/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Section Sample Application
Copyright (c) 2009-2020 - Xceed Software Inc.

This application demonstrates how to insert sections when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.Drawing;
using System.IO;
using System.Linq;
using Xceed.Document.NET;

namespace Xceed.Words.NET.Examples
{
  public class SectionSample
  {
    #region Private Members

    private const string SectionSampleOutputDirectory = Program.SampleDirectory + @"Section\Output\";

    #endregion

    #region Constructors

    static SectionSample()
    {
      if( !Directory.Exists( SectionSample.SectionSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( SectionSample.SectionSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Create a document and insert Sections(with different footers) into it.
    /// </summary>
    public static void InsertSections()
    {
      Console.WriteLine( "\tInsertSections()" );

      // Create a document.
      using( var document = DocX.Create( SectionSample.SectionSampleOutputDirectory + @"InsertSections.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Inserting sections" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        document.DifferentOddAndEvenPages = true;

        // Section 1
        // Set Page parameters for section 1
        document.Sections[ 0 ].PageBorders = new Borders( new Border( BorderStyle.Tcbs_double, BorderSize.four, 5f, Color.Blue ) );
        // Set footers for section 1.
        document.Sections[ 0 ].AddFooters();
        document.Sections[ 0 ].DifferentFirstPage = true;
        var footers = document.Sections[ 0 ].Footers;
        footers.First.InsertParagraph( "This is the First page footer." );
        footers.Even.InsertParagraph( "This is the Even page footer." );
        footers.Odd.InsertParagraph( "This is the Odd page footer." );

        // Add paragraphs and page breaks in section 1.
        document.InsertParagraph( "FIRST" ).InsertPageBreakAfterSelf();
        document.InsertParagraph( "SECOND" ).InsertPageBreakAfterSelf();
        document.InsertParagraph( "THIRD" );

        // Add a section break as a page break to end section 1.
        // The new section properties will be based on last section properties.
        document.InsertSectionPageBreak();

        // Section 2
        // Set Page parameters for section 2
        document.Sections[ 1 ].PageBorders = new Borders( new Border( BorderStyle.Tcbs_none, BorderSize.one, 0f, Color.Transparent ) );
        document.Sections[ 1 ].PageWidth = 200f;
        document.Sections[ 1 ].PageHeight = 300f; 
        // Set footers for section 2.
        document.Sections[ 1 ].AddFooters();
        document.Sections[ 1 ].DifferentFirstPage = true;
        var footers2 = document.Sections[ 1 ].Footers;
        footers2.First.InsertParagraph( "This is the First page footer of Section 2." );
        footers2.Odd.InsertParagraph( "This is the Odd page footer of Section 2." );
        footers2.Even.InsertParagraph( "This is the Even page footer of Section 2." );

        // Add paragraphs and page breaks in section 2.
        document.InsertParagraph( "FOURTH" ).InsertPageBreakAfterSelf();
        document.InsertParagraph( "FIFTH" ).InsertPageBreakAfterSelf();
        document.InsertParagraph( "SIXTH" );

        // Add a section break as a page break to end section 2.
        // The new section properties will be based on last section properties.
        document.InsertSectionPageBreak();

        // Section 3
        // Set Page parameters for section 3
        document.Sections[ 2 ].PageWidth = 595f;
        document.Sections[ 2 ].PageHeight = 841f;
        document.Sections[ 2 ].MarginTop = 300f;
        document.Sections[ 2 ].MarginFooter = 120f;
        // Set footers for section 3.
        document.Sections[ 2 ].AddFooters();
        document.Sections[ 2 ].DifferentFirstPage = true;
        var footers3 = document.Sections[ 2 ].Footers;
        footers3.First.InsertParagraph( "This is the First page footer of Section 3." );
        footers3.Odd.InsertParagraph( "This is the Odd page footer of Section 3." );
        footers3.Even.InsertParagraph( "This is the Even page footer of Section 3." );

        // Add paragraphs and page breaks in section 3.
        document.InsertParagraph( "SEVENTH" ).InsertPageBreakAfterSelf();
        document.InsertParagraph( "EIGHTH" ).InsertPageBreakAfterSelf();
        document.InsertParagraph( "NINETH" );

        // Get the different sections.
        var sections = document.GetSections();

        // Add a paragraph to display the result of sections.
        var p = document.InsertParagraph( "This document contains " ).Append( sections.Count.ToString() ).Append( " Sections.\n" );
        p.SpacingBefore( 40d );
        // Display the paragraphs count per section from this document.
        for( int i = 0; i < sections.Count; ++i )
        {
          var section = sections[ i ];
          var paragraphs = section.SectionParagraphs;
          var nonEmptyParagraphs = paragraphs.Where( x => !string.IsNullOrEmpty( x.Text ) );
          p.Append( "Section " ).Append( (i + 1).ToString() ).Append( " has " ).Append( nonEmptyParagraphs.Count().ToString() ).Append( " non-empty paragraphs.\n" );
        }

        document.Save();
        Console.WriteLine( "\tCreated: InsertSections.docx\n" );
      }
    }

    public static void SetPageOrientations()
    {
      Console.WriteLine( "\tSetPageOrientations()" );

      // Create a document.
      using( var document = DocX.Create( SectionSample.SectionSampleOutputDirectory + @"SetPageOrientations.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Setting Pages Orientation" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Section 1
        // Set Page Orientation to Landscape.
        document.Sections[ 0 ].PageLayout.Orientation = Orientation.Landscape;

        // Add paragraphs in section 1.
        document.InsertParagraph( "This is the first page in Landscape format." );

        // Add a section break as a page break to end section 1.
        // The new section properties will be based on last section properties.
        document.InsertSectionPageBreak();

        // Section 2
        // Set Page Orientation to Portrait.
        document.Sections[ 1 ].PageLayout.Orientation = Orientation.Portrait;

        // Add paragraphs in section 2.
        document.InsertParagraph( "This is the second page in Portrait format." );

        // Add a section break as a page break to end section 2.
        // The new section properties will be based on last section properties.
        document.InsertSectionPageBreak();

        // Section 3
        // Set Page Orientation to Landscape.
        document.Sections[ 2 ].PageLayout.Orientation = Orientation.Landscape;

        // Add paragraphs in section 3.
        document.InsertParagraph( "This is the third page in Landscape format." );

        document.Save();
        Console.WriteLine( "\tCreated: SetPageOrientations.docx\n" );
      }
    }


    #endregion
  }
}
