/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Section Sample Application
Copyright (c) 2009-2020 - Xceed Software Inc.

This application demonstrates how to insert sections when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using Xceed.Document.NET;
using System.Linq;

namespace Xceed.Words.NET.Examples
{
  public class ShapeSample
  {
    #region Private Members

    private const string ShapeSampleOutputDirectory = Program.SampleDirectory + @"Shape\Output\";

    #endregion

    #region Constructors

    static ShapeSample()
    {
      if( !Directory.Exists( ShapeSample.ShapeSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( ShapeSample.ShapeSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Create a document and insert shapes and paragraphs into it.
    /// </summary>
    public static void AddShape()
    {







      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    /// <summary>
    /// Create a document and insert wrapping shapes and paragraphs into it.
    /// </summary>
    public static void AddShapeWithTextWrapping()
    {






      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    /// <summary>
    /// Create a document and insert a TextBox and paragraphs into it.
    /// </summary>
    public static void AddTextBox()
    {





      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    /// <summary>
    /// Create a document and insert wrapping textboxes and paragraphs into it.
    /// </summary>
    public static void AddTextBoxWithTextWrapping()
    {






      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    #endregion
  }
}
