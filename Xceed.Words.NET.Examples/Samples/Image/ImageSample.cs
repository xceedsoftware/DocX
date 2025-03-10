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
Xceed Words for .NET – Xceed.Words.NET.Examples – Image Sample Application
Copyright (c) 2009-2025 - Xceed Software Inc.

This application demonstrates how to create, copy or modify a picture when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xceed.Document.NET;

namespace Xceed.Words.NET.Examples
{
  public class ImageSample
  {
    #region Private Members

    private const string ImageSampleResourcesDirectory = Program.SampleDirectory + @"Image\Resources\";
    private const string ImageSampleOutputDirectory = Program.SampleDirectory + @"Image\Output\";

    #endregion

    #region Constructors

    static ImageSample()
    {
      if( !Directory.Exists( ImageSample.ImageSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( ImageSample.ImageSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    public static void AddPicture()
    {
      Console.WriteLine( "\tAddPicture()" );

      // Create a document.
      using( var document = DocX.Create( ImageSample.ImageSampleOutputDirectory + @"AddPicture.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Adding Pictures" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a simple image from disk.
        var image = document.AddImage( ImageSample.ImageSampleResourcesDirectory + @"balloon.jpg" );
        var picture = image.CreatePicture( 112.5f, 112.5f );
        var p = document.InsertParagraph( "- Here is a simple picture added from disk:\n" );
        p.AppendPicture( picture );

        // Insert incremental "Figure 1" under picture by Picture public method.
        picture.InsertCaptionAfterSelf( "Figure" );

        p.SpacingAfter( 40 );

        // Add a rotated image from disk and set some alpha( 0 to 1 ).
        var rotatedPicture = image.CreatePicture( 112f, 112f );
        rotatedPicture.Rotation = 25;

        var p2 = document.InsertParagraph( "- Here is the same picture added from disk, but rotated:\n" );
        p2.AppendPicture( rotatedPicture );

        // Insert incremental "Figure 2" under picture by Paragraph public method.
        p2 = p2.InsertCaptionAfterSelf( "Figure" );

        p2.SpacingAfter( 40 );

        // Add a simple image from a stream
        var streamImage = document.AddImage( new FileStream( ImageSample.ImageSampleResourcesDirectory + @"balloon.jpg", FileMode.Open, FileAccess.Read ) );
        var pictureStream = streamImage.CreatePicture( 112f, 112f );
        var p3 = document.InsertParagraph( "- Here is the same picture added from a stream:\n" );
        p3.AppendPicture( pictureStream );

        // Insert incremental "Figure 3" under picture by Paragraph public method.
        p3.InsertCaptionAfterSelf( "Figure" );

        document.Save();
        Console.WriteLine( "\tCreated: AddPicture.docx\n" );
      }
    }

    public static void AddPictureWithTextWrapping()
    {






      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void CopyPicture()
    {
      Console.WriteLine( "\tCopyPicture()" );

      // Create a document.
      using( var document = DocX.Create( ImageSample.ImageSampleOutputDirectory + @"CopyPicture.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Copying Pictures" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a paragraph containing an image.
        var image = document.AddImage( ImageSample.ImageSampleResourcesDirectory + @"balloon.jpg" );
        var picture = image.CreatePicture( 75f, 75f );
        var p = document.InsertParagraph( "This is the first paragraph. " );
        p.AppendPicture( picture );
        p.AppendLine("It contains an image added from disk.");
        p.SpacingAfter( 50 );

        // Add a second paragraph containing no image. 
        var p2 = document.InsertParagraph( "This is the second paragraph. " );
        p2.AppendLine( "It contains a copy of the image located in the first paragraph." ).AppendLine();

        // Extract the first Picture from the first Paragraph.
        var firstPicture = p.Pictures.FirstOrDefault();
        if( firstPicture != null )
        {
          // copy it at the end of the second paragraph.
          p2.AppendPicture( firstPicture );
        }

        document.Save();
        Console.WriteLine( "\tCreated: CopyPicture.docx\n" );
      }
    }

    #endregion
  }
}
