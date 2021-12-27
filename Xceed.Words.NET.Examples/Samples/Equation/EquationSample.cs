﻿/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Equation Sample Application
Copyright (c) 2009-2021 - Xceed Software Inc.

This application demonstrates how to insert an equation when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.Drawing;
using System.IO;
using Xceed.Document.NET;

namespace Xceed.Words.NET.Examples
{
  public class EquationSample
  {
    #region Private Members

    private static string EquationSampleOutputDirectory = Program.SampleDirectory + @"Equation\Output\";

    #endregion

    #region Constructors

    static EquationSample()
    {
      if( !Directory.Exists( EquationSample.EquationSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( EquationSample.EquationSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Create a document and add Equations in it.
    /// </summary>
    public static void InsertEquation()
    {
      Console.WriteLine( "\tEquationSample()" );

      // Create a document.
      using( var document = DocX.Create( EquationSample.EquationSampleOutputDirectory + @"EquationSample.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Inserting Equations" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        document.InsertParagraph( "A Linear equation : " );
        // Insert first Equation in this document.
        document.InsertEquation( "y = mx + b", Alignment.left ).SpacingAfter( 30d );

        document.InsertParagraph( "A Quadratic equation : " );
        // Insert second Equation in this document and add formatting.
        document.InsertEquation( "x = ( -b \u00B1 \u221A(b\u00B2 - 4ac))/2a" ).FontSize( 18 ).Color( Color.Blue );

        document.Save();
        Console.WriteLine( "\tCreated: EquationSample.docx\n" );
      }
    }

    #endregion
  }
}
