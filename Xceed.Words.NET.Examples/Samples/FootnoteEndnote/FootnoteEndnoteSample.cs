/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Headers Footers Sample Application
Copyright (c) 2009-2022 - Xceed Software Inc.

This application demonstrates how to create footnotes and endnotes when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.Drawing;
using System.IO;
using Xceed.Document.NET;
using Xceed.Words.NET.Examples;

namespace Xceed.Words.NET.Example
{
  public class FootnoteEndnoteSample
  {
    #region Private Members

    private const string FootnoteEndnoteSampleResourcesDirectory = Program.SampleDirectory + @"FootnoteEndnote\Resources\";
    private const string FootnoteEndnoteSampleOutputDirectory = Program.SampleDirectory + @"FootnoteEndnote\Output\";

    #endregion

    #region Constructors

    static FootnoteEndnoteSample()
    {
      if( !Directory.Exists( FootnoteEndnoteSample.FootnoteEndnoteSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( FootnoteEndnoteSample.FootnoteEndnoteSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Add footnotes to a document.
    /// </summary>
    public static void AddFootnotes()
    {
      Console.WriteLine( "\tAddFootnotes()" );

      // Create a document.
      using( var document = DocX.Create( FootnoteEndnoteSample.FootnoteEndnoteSampleOutputDirectory + @"AddFootnotes.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Add Footnotes" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a standard footnote.
        var footnote1 = document.AddFootnote( "Montreal in the province of Quebec, Canada." );
        // Create a formatted footnote.
        var footnote2 = document.AddFootnote( "More than 100,000 satisfied customers.", new Formatting() { Bold = true, Size = 15d, FontColor = Color.Red } );
        // Create an hyperlink footnote.
        var footnote3 = document.AddFootnote( document.AddHyperlink( "Xceed Web site"
                                              , new Uri( "http://www.xceed.com" )
                                              ) );

        // Insert a Paragraph into this document.
        var p = document.InsertParagraph().SpacingBefore( 70d );

        // Append some text in a paragraph and add 3 footnotes.
        p.Append( "Xceed is an international software component development partner for technology leaders worldwide. Founded in the mid-90's with very humble aspirations in Montreal" )
         .AppendNote( footnote1 )
         .Append( ", Xceed has retained a garage start-up feel from the inside but with the capability of delivering world class finished products for our clients" )
         .AppendNote( footnote2 )
         .Append( " and partners. You can read more at Xceed.com." )
         .AppendNote( footnote3 );

        document.Save();
        Console.WriteLine( "\tCreated: AddFootnotes.docx\n" );
      }
    }

    /// <summary>
    /// Add custom footnotes to a document.
    /// </summary>
    public static void AddCustomFootnotes()
    {
      Console.WriteLine( "\tAddCustomFootnotes()" );

      // Create a document.
      using( var document = DocX.Create( FootnoteEndnoteSample.FootnoteEndnoteSampleOutputDirectory + @"AddCustomFootnotes.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Add custom Footnotes" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Modify the footnotes format to use upper letters instead of the default numbers.
        document.Sections[0].FootnoteProperties = new NoteProperties() { NumberFormat = NoteNumberFormat.upperLetter };

        // Create a standard footnote and update its paragraph.
        var footnote1 = document.AddFootnote( "docx documents are the saved files from Microsoft Word." );
        footnote1.Paragraphs[0].Append( " They were introduced in 2007." );

        // Create a standard footnote and use a specific symbol.
        var footnote2 = document.AddFootnote( "You can do the same actions as in Microsoft Word." );
        footnote2.CustomMark = new Symbol() { Font = new Xceed.Document.NET.Font( "Symbol" ), Code = 197 };

        // Create a standard footnote.
        var footnote3 = document.AddFootnote( "PDF files: short for portable document format files, are one of the most commonly used file types." );

        // Insert a Paragraph into this document.
        var p = document.InsertParagraph().SpacingBefore( 70d );

        // Append some text in a paragraph and add 3 footnotes.
        p.Append( "With its easy to use API, Xceed Words for .NET lets your application create new Microsoft Word .docx" )
         .AppendNote( footnote1 )
         .Append( " or PDF documents, or modify existing .docx documents. It gives you complete control" )
         .AppendNote( footnote2 )
         .Append( " over all content in a Word document, and lets you add or remove all commonly used element types, such as paragraphs, bulleted or numbered lists, images, tables, charts, headers and footers, sections, bookmarks, and more. Create PDF" )
         .AppendNote( footnote3 )
         .Append( " documents using the same API for creating Word documents." );

        document.Save();
        Console.WriteLine( "\tCreated: AddCustomFootnotes.docx\n" );
      }
    }

    /// <summary>
    /// Add endnotes to a document.
    /// </summary>
    public static void AddEndnotes()
    {
      Console.WriteLine( "\tAddEndnotes()" );

      // Create a document.
      using( var document = DocX.Create( FootnoteEndnoteSample.FootnoteEndnoteSampleOutputDirectory + @"AddEndnotes.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Add Endnotes" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a standard endnote.
        var endnote1 = document.AddEndnote( "WPF Components like DataGrid or Toolkit or a JS DataGrid component." );

        // Create a formatted endnote.
        var endnote2 = document.AddEndnote( "The calendar year contains 12 months.", new Formatting() { Bold = true, Size = 15d, FontColor = Color.Blue } );

        // Create a picture endnote.
        var image = document.AddImage( FootnoteEndnoteSample.FootnoteEndnoteSampleResourcesDirectory + @"xceed.png" );
        var endnote3 = document.AddEndnote( image.CreatePicture( 26f, 50f ) );

        // Insert a Paragraph into this document.
        var p1 = document.InsertParagraph( "What we do\n\n" ).FontSize( 15d ).Color( Color.Orange ).SpacingBefore( 70d );

        // Append some text in a first paragraph and append 1 endnote (with formatting for the number) to it.
        p1.Append( "We help developers save time. We help project managers deliver faster. We help businesses save money and time to market.\n\nHow do we do this ? Well, we provide comprehensive UI components" )
          .AppendNote( endnote1, new Formatting() { Bold = true, FontColor = Color.Red, Size = 15d } )
          .Append( " that allow developers to focus on innovation and their business requirements... and less time on the design. Xceed has been able to fill in the gap for developers seeking extra functionalities" +
                   " that are just not found in the out-of-the box solutions. Providing these components means they don't have to build them. Which means they are paid to do the stuff that really matters for the business " +
                   "(and to them too!). This is how we save companies loads of time and money. And for developers, well, they don't waste time on tedious stuff: they can do much more with less development!" )
          .InsertPageBreakAfterSelf();

        // Insert a Paragraph into this document.
        var p2 = document.InsertParagraph( "How we do it\n\n" ).FontSize( 15d ).Color( Color.Orange ).SpacingBefore( 70d );

        // Append some text in a 2nd paragraph and append 2 endnotes to it.
        p2.Append( "Over the years" )
          .AppendNote( endnote2 )
          .Append( " , we have listened. We have worked with client feedback to continuously improve the reliability and quality of our products. But we have also taken the time to understand your " +
                   "business. Your business has standards, and so does ours. We would not ask anything less from our teams than to meet and exceed your expectations so that you can be successful in your mission.\n\nXceed" )
          .AppendNote( endnote3 )
          .Append( " too have challenges in compliance and regulation, accounting, budgeting, project deadlines, policies and all that jazz. So, when it comes to working with you, rest assured that we strive for your" +
                   " satisfaction. We aim to deliver products with extensive functionality, a minimum of bugs and top technical support." );

        document.Save();
        Console.WriteLine( "\tCreated: AddEndnotes.docx\n" );
      }
    }

    #endregion
  }
}
