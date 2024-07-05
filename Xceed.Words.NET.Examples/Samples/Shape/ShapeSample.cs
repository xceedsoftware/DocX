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
Xceed Words for .NET – Xceed.Words.NET.Examples – Section Sample Application
Copyright (c) 2009-2024 - Xceed Software Inc.

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

    public static void AddShape()
    {







      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void AddShapeWithTextWrapping()
    {






      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void AddTextBox()
    {





      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void AddTextBoxWithTextWrapping()
    {






      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    #endregion
  }
}
