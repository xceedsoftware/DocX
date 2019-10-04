/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Section Sample Application
Copyright (c) 2009-2018 - Xceed Software Inc.

This application demonstrates how to insert sections when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
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
    /// Create a document and insert Sections into it.
    /// </summary>
    public static void InsertSections()
    {
      Console.WriteLine( "\tInsertSections()" );

      // Create a document.
      using( var document = DocX.Create( SectionSample.SectionSampleOutputDirectory + @"InsertSections.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Inserting sections" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add 2 paragraphs
        document.InsertParagraph( "This is the first paragraph." );
        document.InsertParagraph( "This is the second paragraph." );
        // Add a paragraph and a section break.
        document.InsertSection();
        // Add a new paragraph
        document.InsertParagraph( "This is the third paragraph, in a new section." );
        // Add a paragraph and a page break.
        document.InsertSectionPageBreak();
        document.InsertParagraph( "This is the fourth paragraph, in a new section." );

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

    #endregion
  }
}
