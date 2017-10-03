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
  public class EquationSample
  {
    #region Private Members

    private const string EquationSampleOutputDirectory = Program.SampleDirectory + @"Equation\Output\";

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
      using( DocX document = DocX.Create( EquationSample.EquationSampleOutputDirectory + @"EquationSample.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Inserting Equations" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        document.InsertParagraph( "A Linear equation : " );
        // Insert first Equation in this document.
        document.InsertEquation( "y = mx + b" ).SpacingAfter( 30d );

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
