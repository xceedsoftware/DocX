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
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

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
      using( DocX document = DocX.Create( ImageSample.ImageSampleOutputDirectory + @"AddPicture.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Adding Pictures" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a simple image from disk.
        var image = document.AddImage( ImageSample.ImageSampleResourcesDirectory + @"balloon.jpg" );
        var picture = image.CreatePicture( 150, 150 );
        var p = document.InsertParagraph( "Here is a simple picture added from disk:" );
        p.AppendPicture( picture );
        p.SpacingAfter( 30 );

        // Add a rotated image from disk.
        var rotatedPicture = image.CreatePicture( 150, 150 );
        rotatedPicture.Rotation = 25;

        var p2 = document.InsertParagraph( "Here is the same picture added from disk, but rotated:" );
        p2.AppendPicture( rotatedPicture );
        p2.SpacingAfter( 30 );

        // Add a simple image from a stream
        var streamImage = document.AddImage( new FileStream( ImageSample.ImageSampleResourcesDirectory + @"balloon.jpg", FileMode.Open, FileAccess.Read ) );
        var pictureStream = streamImage.CreatePicture( 150, 150 );
        var p3 = document.InsertParagraph( "Here is the same picture added from a stream:" );
        p3.AppendPicture( pictureStream );

        document.Save();
        Console.WriteLine( "\tCreated: AddPicture.docx\n" );
      }
    }

    /// <summary>
    /// Copy a picture from a paragraph.
    /// </summary>
    public static void CopyPicture()
    {
      Console.WriteLine( "\tCopyPicture()" );

      // Create a document.
      using( DocX document = DocX.Create( ImageSample.ImageSampleOutputDirectory + @"CopyPicture.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Copying Pictures" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a paragraph containing an image.
        var image = document.AddImage( ImageSample.ImageSampleResourcesDirectory + @"balloon.jpg" );
        var picture = image.CreatePicture( 100, 100 );
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
      using( DocX document = DocX.Load( ImageSample.ImageSampleResourcesDirectory + @"Input.docx" ) )
      {
        // Add a title
        document.InsertParagraph( 0, "Modifying Image by adding text/circle into the following image", false ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Get the first image in the document.
        var image = document.Images.FirstOrDefault();
        if( image != null )
        {
          // Create a bitmap from the image.
          var bitmap = new Bitmap( image.GetStream( FileMode.Open, FileAccess.ReadWrite ) );
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
