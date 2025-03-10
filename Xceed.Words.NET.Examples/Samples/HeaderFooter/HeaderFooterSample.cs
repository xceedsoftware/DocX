/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Headers Footers Sample Application
Copyright (c) 2009-2025 - Xceed Software Inc.

This application demonstrates how to create headers and footers when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.IO;
using Xceed.Document.NET;

namespace Xceed.Words.NET.Examples
{
  public class HeaderFooterSample
  {
    #region Private Members

    private const string HeaderFooterSampleOutputDirectory = Program.SampleDirectory + @"HeaderFooter\Output\";

    #endregion

    #region Constructors

    static HeaderFooterSample()
    {
      if( !Directory.Exists( HeaderFooterSample.HeaderFooterSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( HeaderFooterSample.HeaderFooterSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    public static void HeadersFooters()
    {
      Console.WriteLine( "\tHeadersFooters()" );

      // Create a document.
      using( var document = DocX.Create( HeaderFooterSample.HeaderFooterSampleOutputDirectory + @"HeadersFooters.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Headers and Footers" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert a Paragraph in the first page of the document.
        var p1 = document.InsertParagraph("This is the ").Append( "first").Bold().Append(" page Content.");
        p1.SpacingBefore( 70d );
        p1.InsertPageBreakAfterSelf();

        // Insert a Paragraph in the second page of the document.
        var p2 = document.InsertParagraph( "This is the " ).Append( "second" ).Bold().Append( " page Content." );
        p2.InsertPageBreakAfterSelf();

        // Insert a Paragraph in the third page of the document.
        var p3 = document.InsertParagraph( "This is the " ).Append( "third" ).Bold().Append( " page Content." );
        p3.InsertPageBreakAfterSelf();

        // Insert a Paragraph in the third page of the document.
        var p4 = document.InsertParagraph( "This is the " ).Append( "fourth" ).Bold().Append( " page Content." );

        // Add Headers and Footers to the document.
        document.AddHeaders();
        document.AddFooters();

        // Force the first page to have a different Header and Footer.
        document.DifferentFirstPage = true;

        // Force odd & even pages to have different Headers and Footers.
        document.DifferentOddAndEvenPages = true;

        // Insert a Paragraph into the first Header.
        document.Headers.First.InsertParagraph("This is the ").Append("first").Bold().Append(" page header");

        // Insert a Paragraph and Page number into the first Footer.
        document.Footers.First.InsertParagraph( "This is the " ).Append( "first" ).Bold().Append( " page footer" );
        document.Footers.First.InsertParagraph( "Page " ).AppendPageNumber( PageNumberFormat.normal ).Append( " of " ).AppendPageCount( PageNumberFormat.normal );

        // Insert a Paragraph into the even Header.
        document.Headers.Even.InsertParagraph( "This is an " ).Append( "even" ).Bold().Append( " page header" );

        // Insert a Paragraph and Page number into the even Footer.
        document.Footers.Even.InsertParagraph( "This is an " ).Append( "even" ).Bold().Append( " page footer" );
        document.Footers.Even.InsertParagraph( "Page " ).AppendPageNumber( PageNumberFormat.normal ).Append( " of " ).AppendPageCount( PageNumberFormat.normal );

        // Insert a Paragraph into the odd Header.
        document.Headers.Odd.InsertParagraph( "This is an " ).Append( "odd" ).Bold().Append( " page header" );

        // Insert a Paragraph and Page number into the odd Footer.
        document.Footers.Odd.InsertParagraph( "This is an " ).Append( "odd" ).Bold().Append( " page footer" );
        document.Footers.Odd.InsertParagraph( "Page " ).AppendPageNumber( PageNumberFormat.normal ).Append( " of " ).AppendPageCount( PageNumberFormat.normal );

        document.Save();
        Console.WriteLine( "\tCreated: HeadersFooters.docx\n" );
      }
    }

    #endregion
  }
}
