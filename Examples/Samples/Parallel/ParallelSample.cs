/***************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2017 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  *************************************************************************************/
using System;
using System.IO;
using System.Threading.Tasks;
using System.Linq;

namespace Xceed.Words.NET.Examples
{
  public class ParallelSample
  {
    #region Private Members

    private const string ParallelSampleResourcesDirectory = Program.SampleDirectory + @"Parallel\Resources\";
    private const string ParallelSampleOutputDirectory = Program.SampleDirectory + @"Parallel\Output\";

    #endregion

    #region Constructors

    static ParallelSample()
    {
      if( !Directory.Exists( ParallelSample.ParallelSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( ParallelSample.ParallelSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// For each of the documents in the folder 'Parallel\Resources\',
    /// Replace the string "Apple" with the string "Potato" and replace the "Apple" image by a "Potato" image.
    /// Do this in Parrallel accross many CPU cores.
    /// </summary>
    public static void DoParallelActions()
    {
      Console.WriteLine( "\tDoParallelActions()" );

      // Get the docx files from the Resources directory.
      var inputDir = new DirectoryInfo( ParallelSample.ParallelSampleResourcesDirectory );
      var inputFiles = inputDir.GetFiles( "*.docx" );

      // Loop through each document and do actions on them.
      Parallel.ForEach( inputFiles, f => ParallelSample.Action( f ) );
    }

    private static void Action( FileInfo file )
    {
      // Load the document.
      using( DocX document = DocX.Load( file.FullName ) )
      {
        // Replace texts in this document.
        document.ReplaceText( "Apples", "Potatoes" );
        document.ReplaceText( "An Apple", "A Potato" );

        // create the new image
        var newImage = document.AddImage( ParallelSample.ParallelSampleResourcesDirectory + @"potato.jpg" );

        // Look in each paragraph and remove its first image to replace it with the new one.
        foreach( var p in document.Paragraphs )
        {
          var oldPicture = p.Pictures.FirstOrDefault();
          if( oldPicture != null )
          {
            oldPicture.Remove();
            p.AppendPicture( newImage.CreatePicture( 150, 150 ) );
          }
        }

        document.SaveAs( ParallelSample.ParallelSampleOutputDirectory + "Output" + file.Name );
        Console.WriteLine( "\tCreated: Output" + file.Name + ".docx\n" );
      }
    }

    #endregion
  }
}
