/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Parallel Sample Application
Copyright (c) 2009-2020 - Xceed Software Inc.

This application demonstrates how to do parallel actions when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
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
      using( var document = DocX.Load( file.FullName ) )
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
            p.AppendPicture( newImage.CreatePicture( 112f, 112f ) );
          }
        }

        document.SaveAs( ParallelSample.ParallelSampleOutputDirectory + "Output" + file.Name );
        Console.WriteLine( "\tCreated: Output" + file.Name + ".docx\n" );
      }
    }

    #endregion
  }
}
