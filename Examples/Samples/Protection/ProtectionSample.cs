/***************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2017 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  *************************************************************************************/
using System;
using System.Drawing;
using System.IO;

namespace Xceed.Words.NET.Examples
{
  public class ProtectionSample
  {
    #region Private Members

    private const string ProtectionSampleOutputDirectory = Program.SampleDirectory + @"Protection\Output\";

    #endregion

    #region Constructors

    static ProtectionSample()
    {
      if( !Directory.Exists( ProtectionSample.ProtectionSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( ProtectionSample.ProtectionSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Create a read only document that can be edited by entering a valid password.
    /// </summary>
    public static void AddPasswordProtection()
    {
      Console.WriteLine( "\tAddPasswordProtection()" );

      // Create a new document.
      using( DocX document = DocX.Create( ProtectionSample.ProtectionSampleOutputDirectory + @"AddPasswordProtection.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Document protection using password" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert a Paragraph into this document.
        var p = document.InsertParagraph();

        // Append some text and add formatting.
        p.Append( "This document is protected and can only be edited by stopping its protection with a valid password(\"xceed\")." )
        .Font( new Font( "Arial" ) )
        .FontSize( 25 )
        .Color( Color.Blue )
        .Bold();

        // Set the document as read only and add a password to unlock it.
        document.AddPasswordProtection( EditRestrictions.readOnly, "xceed" );

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: AddPasswordProtection.docx\n" );
      }
    }

    /// <summary>
    /// Create a read only document that can be edited by stopping the protection.
    /// </summary>
    public static void AddProtection()
    {
      Console.WriteLine( "\tAddProtection()" );

      // Create a new document.
      using( DocX document = DocX.Create( ProtectionSample.ProtectionSampleOutputDirectory + @"AddProtection.docx" ) )
      {
        // Add a title.
        document.InsertParagraph( "Document protection not using password" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert a Paragraph into this document.
        var p = document.InsertParagraph();

        // Append some text and add formatting.
        p.Append( "This document is protected and can only be edited by stopping its protection." )
        .Font( new Font( "Arial" ) )
        .FontSize( 25 )
        .Color( Color.Red )
        .Bold();

        document.AddProtection( EditRestrictions.readOnly );

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: AddProtection.docx\n" );
      }
    }

    #endregion
  }
}
