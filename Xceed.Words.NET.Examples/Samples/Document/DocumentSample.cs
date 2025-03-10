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
Xceed Words for .NET – Xceed.Words.NET.Examples – Document Sample Application
Copyright (c) 2009-2025 - Xceed Software Inc.

This application demonstrates how to modify the content of a document when using the API 
from the Xceed Words for .NET.

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
  public class DocumentSample
  {
    #region Private Members

    private static Dictionary<string, string> _replacePatterns = new Dictionary<string, string>()
    {
        { "OPPONENT", "Pittsburgh Penguins" },
        { "GAME_TIME", "7:30pm" },
        { "GAME_NUMBER", "161" },
        { "DATE", "October 18 2016" },
    };

    private const string DocumentSampleResourcesDirectory = Program.SampleDirectory + @"Document\Resources\";
    private const string DocumentSampleOutputDirectory = Program.SampleDirectory + @"Document\Output\";

    #endregion

    #region Constructors

    static DocumentSample()
    {
      if( !Directory.Exists( DocumentSample.DocumentSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( DocumentSample.DocumentSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    public static void ReplaceTextWithText()
    {
      Console.WriteLine( "\tReplaceTextWithText()" );

      // Load a document.
      using( var document = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"ReplaceText.docx" ) )
      {
        // Check if some of the replace patterns are used in the loaded document.
        if( document.FindUniqueByPattern( @"<[\w \=]{4,}>", RegexOptions.IgnoreCase ).Count > 0 )
        {
          // Do the replacement of all the found tags and with green bold strings.
          var replaceTextOptions = new FunctionReplaceTextOptions()
          {
            FindPattern = "<(.*?)>",
            RegexMatchHandler = DocumentSample.ReplaceFunc,
            RegExOptions = RegexOptions.IgnoreCase,
            NewFormatting = new Formatting() { Bold = true, FontColor = Color.Green }
          };
          document.ReplaceText( replaceTextOptions );

          // Save this document to disk.
          document.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"ReplacedText.docx" );
          Console.WriteLine( "\tCreated: ReplacedTextWithText.docx\n" );
        }
      }
    }

    public static void ReplaceTextWithMultipleTexts()
    {


      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void ReplaceTextWithObjects()
    {
      Console.WriteLine( "\tReplaceTextWithObjects()" );

      // Load a document.
      using( var document = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"ReplaceTextWithObjects.docx" ) )
      {
        // Create the image from disk and set its size.
        var image = document.AddImage( DocumentSample.DocumentSampleResourcesDirectory + @"2018.jpg" );
        var picture = image.CreatePicture( 175f, 325f );

        // Do the replacement of all the found tags with the specified image and ignore the case when searching for the tags.
        document.ReplaceTextWithObject( new ObjectReplaceTextOptions() { SearchValue = "<yEaR_IMAGE>", NewObject = picture, RegExOptions = RegexOptions.IgnoreCase } );

        // Create the hyperlink.
        var hyperlink = document.AddHyperlink( "(ref)", new Uri( "https://en.wikipedia.org/wiki/New_Year" ) );
        // Do the replacement of all the found tags with the specified hyperlink.
        document.ReplaceTextWithObject( new ObjectReplaceTextOptions() { SearchValue = "<year_link>", NewObject = hyperlink } );

        // Add a Table into the document and sets its values.
        var t = document.AddTable( 1, 2 );
        t.Design = TableDesign.DarkListAccent4;
        t.AutoFit = AutoFit.Window;
        t.Rows[ 0 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "xceed.com" );
        t.Rows[ 0 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "@copyright 2025" );
        document.ReplaceTextWithObject( new ObjectReplaceTextOptions() { SearchValue = "<year_table>", NewObject = t } );

        // Save this document to disk.
        document.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"ReplacedTextWithObjects.docx" );
        Console.WriteLine( "\tCreated: ReplacedTextWithObjects.docx\n" );
      }
    }

    public static void AddCustomProperties()
    {
      Console.WriteLine( "\tAddCustomProperties()" );

      // Create a new document.
      using( var document = DocX.Create( DocumentSample.DocumentSampleOutputDirectory + @"AddCustomProperties.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Adding Custom Properties to a document" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        //Add custom properties to document.
        document.AddCustomProperty( new CustomProperty( "CompanyName", "Xceed Software inc." ) );
        document.AddCustomProperty( new CustomProperty( "Product", "Xceed Words for .NET" ) );
        document.AddCustomProperty( new CustomProperty( "Address", "3141 Taschereau, Greenfield Park" ) );
        document.AddCustomProperty( new CustomProperty( "Date", DateTime.Now ) );

        // Add a paragraph displaying the number of custom properties.
        var p = document.InsertParagraph( "This document contains " ).Append( document.CustomProperties.Count.ToString() ).Append( " Custom Properties :" );
        p.SpacingAfter( 30 );

        // Display each propertie's name and value.
        foreach( var prop in document.CustomProperties )
        {
          document.InsertParagraph( prop.Value.Name ).Append( " = " ).Append( prop.Value.Value.ToString() ).AppendLine();
        }

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: AddCustomProperties.docx\n" );
      }
    }

    public static void ApplyTemplate()
    {
      Console.WriteLine( "\tApplyTemplate()" );

      // Create a new document.
      using( var document = DocX.Create( DocumentSample.DocumentSampleOutputDirectory + @"ApplyTemplate.docx" ) )
      {
        // The path to a template document,
        var templatePath = DocumentSample.DocumentSampleResourcesDirectory + @"Template.docx";

        document.DifferentOddAndEvenPages = true;

        // Apply a template to the document based on a path.
        document.ApplyTemplate( templatePath );

        // Add a paragraph at the end of the template.
        document.InsertParagraph( "This paragraph is not part of the template." ).FontSize( 15d ).UnderlineStyle( UnderlineStyle.singleLine ).SpacingBefore( 50d );

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: ApplyTemplate.docx\n" );
      }
    }

    public static void AppendDocument()
    {
      Console.WriteLine( "\tAppendDocument()" );

      // Load the first document.
      using( var document1 = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"First.docx" ) )
      {
        // Load the second document.
        using( var document2 = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"Second.docx" ) )
        {
          // Add a title
          document1.InsertParagraph( 0, "Append Document", false ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

          // Insert a document at the end of another document.
          // When true, document is added at the end. When false, document is added at beginning.
          document1.InsertDocument( document2, true, true );


          // Save this document to disk.
          document1.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"AppendDocument.docx" );
          Console.WriteLine( "\tCreated: AppendDocument.docx\n" );
        }
      }
    }

    public static void InsertDocument()
    {




      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void LoadDocumentWithFilename()
    {
      using( var doc = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"First.docx" ) )
      {
        // Add a title
        doc.InsertParagraph( 0, "Load Document with File name", false ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert a Paragraph into this document.
        var p = doc.InsertParagraph();

        // Append some text and add formatting.
        p.Append( "A small paragraph was added." );

        doc.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"LoadDocumentWithFilename.docx" );
      }
    }

    public static void LoadDocumentWithStream()
    {
      using( var fs = new FileStream( DocumentSample.DocumentSampleResourcesDirectory + @"First.docx", FileMode.Open, FileAccess.Read, FileShare.Read ) )
      {
        using( var doc = DocX.Load( fs ) )
        {
          // Add a title
          doc.InsertParagraph( 0, "Load Document with Stream", false ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

          // Insert a Paragraph into this document.
          var p = doc.InsertParagraph();

          // Append some text and add formatting.
          p.Append( "A small paragraph was added." );

          doc.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"LoadDocumentWithStream.docx" );
        }
      }
    }

    public static void LoadDocumentWithStringUrl()
    {
      using( var doc = DocX.Load( "https://calibre-ebook.com/downloads/demos/demo.docx" ) )
      {
        // Add a title
        doc.InsertParagraph( 0, "Load Document with string Url", false ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert a Paragraph into this document.
        var p = doc.InsertParagraph();

        // Append some text and add formatting.
        p.Append( "A small paragraph was added." );

        doc.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"LoadDocumentWithUrl.docx" );
      }
    }

    public static void AddHtml()
    {










      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void AddRtf()
    {





      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void AddHtmlFromFile()
    {



      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void AddRtfFromFile()
    {



      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void ReplaceTextWithHtml()
    {





 	      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    #endregion

    #region Private Methods

    private static string ReplaceFunc( string findStr )
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
