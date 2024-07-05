/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2024 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Parallel Sample Application
Copyright (c) 2009-2024 - Xceed Software Inc.

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
using Xceed.Document.NET;

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
        document.ReplaceText( new StringReplaceTextOptions() { SearchValue = "Apples", NewValue = "Potatoes" } );
        document.ReplaceText( new StringReplaceTextOptions() { SearchValue = "An Apple", NewValue = "A Potato" } );

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
