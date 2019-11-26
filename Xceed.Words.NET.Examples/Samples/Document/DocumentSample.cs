/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Document Sample Application
Copyright (c) 2009-2018 - Xceed Software Inc.

This application demonstrates how to modify the content of a document when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Xceed.Document.NET;

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

    /// <summary>
    /// Load a document and replace texts following a replace pattern.
    /// </summary>
    public static void ReplaceText()
    {
      Console.WriteLine( "\tReplaceText()" );

      // Load a document.
      using( var document = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"ReplaceText.docx" ) )
      {
        // Check if some of the replace patterns are used in the loaded document.
        if( document.FindUniqueByPattern( @"<[\w \=]{4,}>", RegexOptions.IgnoreCase ).Count > 0 )
        {
          // Do the replacement of all the found tags and with green bold strings.
          document.ReplaceText( "<(.*?)>", DocumentSample.ReplaceFunc, false, RegexOptions.IgnoreCase, new Formatting() { Bold = true, FontColor = System.Drawing.Color.Green } );

          // Save this document to disk.
          document.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"ReplacedText.docx" );
          Console.WriteLine( "\tCreated: ReplacedText.docx\n" );
        }
      }
    }

    /// <summary>
    /// Add custom properties to a document.
    /// </summary>
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
        var p = document.InsertParagraph( "This document contains " ).Append( document.CustomProperties.Count.ToString() ).Append(" Custom Properties :");
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

    /// <summary>
    /// Add a template to a document.
    /// </summary>
    public static void ApplyTemplate()
    {
      Console.WriteLine( "\tApplyTemplate()" );

      // Create a new document.
      using( var document = DocX.Create( DocumentSample.DocumentSampleOutputDirectory + @"ApplyTemplate.docx" ) )
      {
        // The path to a template document,
        var templatePath = DocumentSample.DocumentSampleResourcesDirectory + @"Template.docx";

        // Apply a template to the document based on a path.
        document.ApplyTemplate( templatePath );

        // Add a paragraph at the end of the template.
        document.InsertParagraph( "This paragraph is not part of the template." ).FontSize( 15d ).UnderlineStyle(UnderlineStyle.singleLine).SpacingBefore(50d);

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: ApplyTemplate.docx\n" );
      }
    }

    /// <summary>
    /// Insert a document at the end of another document.
    /// </summary>
    public static void AppendDocument()
    {
      Console.WriteLine( "\tAppendDocument()" );

      // Load the first document.
      using( var document1 = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"First.docx" ) )
      {
        // Load the second document.
        using( var document2 = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"Second.docx" ) )
        {
          // Insert a document at the end of another document.
          // When true, document is added at the end. When false, document is added at beginning.
          document1.InsertDocument( document2, true );

          // Save this document to disk.
          document1.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"AppendDocument.docx" );
          Console.WriteLine( "\tCreated: AppendDocument.docx\n" );
        }
      }
    }

    public static void LoadDocumentWithFilename()
    {
      using ( var doc = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"First.docx" ) )
      {
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
        // Insert a Paragraph into this document.
        var p = doc.InsertParagraph();

        // Append some text and add formatting.
        p.Append( "A small paragraph was added." );

        doc.SaveAs( DocumentSample.DocumentSampleOutputDirectory + @"LoadDocumentWithUrl.docx" );
      }
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
