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
using System.IO;
using System.Text.RegularExpressions;

namespace Xceed.Words.NET.Examples
{
  public class DocumentSample
  {
    #region Private Members

    private static Dictionary<string, string> _replacePatterns = new Dictionary<string, string>()
    {
        { "OPPONENT", "Pittsburgh Penguins" },
        { "GAME_TIME", "19h30" },
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
      using( DocX document = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"ReplaceText.docx" ) )
      {
        // Check if all the replace patterns are used in the loaded document.
        if( document.FindUniqueByPattern( @"<[\w \=]{4,}>", RegexOptions.IgnoreCase ).Count == _replacePatterns.Count )
        {
          // Do the replacement
          for( int i = 0; i < _replacePatterns.Count; ++i )
          {
            document.ReplaceText( "<(.*?)>", DocumentSample.ReplaceFunc, false, RegexOptions.IgnoreCase, null, new Formatting() );
          }

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
      using( DocX document = DocX.Create( DocumentSample.DocumentSampleOutputDirectory + @"AddCustomProperties.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Adding Custom Properties to a document" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        //Add custom properties to document.
        document.AddCustomProperty( new CustomProperty( "CompanyName", "Xceed Software inc." ) );
        document.AddCustomProperty( new CustomProperty( "Product", "Xceed Words for .NET" ) );
        document.AddCustomProperty( new CustomProperty( "Address", "10 Boul. de Mortagne" ) );
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
      using( DocX document = DocX.Create( DocumentSample.DocumentSampleOutputDirectory + @"ApplyTemplate.docx" ) )
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
      using( DocX document1 = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"First.docx" ) )
      {
        // Load the second document.
        using( DocX document2 = DocX.Load( DocumentSample.DocumentSampleResourcesDirectory + @"Second.docx" ) )
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
