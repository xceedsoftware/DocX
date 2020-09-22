/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Image Sample Application
Copyright (c) 2009-2020 - Xceed Software Inc.

This application demonstrates how to create, copy or modify a picture when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
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

    /// <summary>
    /// Add a picture loaded from disk or stream to a document.
    /// </summary>
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
        // Insert incremental "Figure 1" under picture.
        p = p.InsertCaptionAfterSelf( "Figure" );
        p.SpacingAfter( 40 );

        // Add a rotated image from disk and set some alpha( 0 to 1 ).
        var rotatedPicture = image.CreatePicture( 112f, 112f );
        rotatedPicture.Rotation = 25;

        var p2 = document.InsertParagraph( "- Here is the same picture added from disk, but rotated:\n" );
        p2.AppendPicture( rotatedPicture );
        // Insert incremental "Figure 2" under picture.
        p2 = p2.InsertCaptionAfterSelf( "Figure" );
        p2.SpacingAfter( 40 );

        // Add a simple image from a stream
        var streamImage = document.AddImage( new FileStream( ImageSample.ImageSampleResourcesDirectory + @"balloon.jpg", FileMode.Open, FileAccess.Read ) );
        var pictureStream = streamImage.CreatePicture( 112f, 112f );
        var p3 = document.InsertParagraph( "- Here is the same picture added from a stream:\n" );
        p3.AppendPicture( pictureStream );
        // Insert incremental "Figure 3" under picture.
        p3.InsertCaptionAfterSelf( "Figure" );

        document.Save();
        Console.WriteLine( "\tCreated: AddPicture.docx\n" );
      }
    }

    public static void AddPictureWithTextWrapping()
    {






      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    /// <summary>
    /// Copy a picture from a paragraph.
    /// </summary>
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

    /// <summary>
    /// Modify an image from a document by writing text into it.
    /// </summary>
    public static void ModifyImage()
    {
      Console.WriteLine( "\tModifyImage()" );

      // Open the document Input.docx.
      using( var document = DocX.Load( ImageSample.ImageSampleResourcesDirectory + @"Input.docx" ) )
      {
        // Add a title
        document.InsertParagraph( 0, "Modifying Image by adding text/circle into the following image", false ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Get the first image in the document.
        var image = document.Images.FirstOrDefault();
        if( image != null )
        {
          // Create a bitmap from the image.
          Bitmap bitmap;
          using( var stream = image.GetStream( FileMode.Open, FileAccess.ReadWrite ) )
          {
            bitmap = new Bitmap( stream );
          }
          // Get the graphic from the bitmap to be able to draw in it.
          var graphic = Graphics.FromImage( bitmap );
          if( graphic != null )
          {
            // Draw a string with a specific font, font size and color at (0,10) from top left of the image.
            graphic.DrawString( "@copyright", new System.Drawing.Font( "Arial Bold", 12 ), Brushes.Red, new PointF( 0f, 10f ) );
            // Draw a blue circle of 10x10 at (30, 5) from the top left of the image.
            graphic.FillEllipse( Brushes.Blue, 30, 5, 10, 10 );

            // Save this Bitmap back into the document using a Create\Write stream.
            bitmap.Save( image.GetStream( FileMode.Create, FileAccess.Write ), ImageFormat.Png );
          }
        }

        document.SaveAs( ImageSample.ImageSampleOutputDirectory + @"ModifyImage.docx" );
        Console.WriteLine( "\tCreated: ModifyImage.docx\n" );
      }
    }

    #endregion
  }
}
