/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2026 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Margin Sample Application
Copyright (c) 2009-2026 - Xceed Software Inc.

This application demonstrates how to use margins, indentations and directions
when using the API from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.IO;
using Xceed.Document.NET;

namespace Xceed.Words.NET.Examples
{
  public class WatermarkSample
  {
    #region Private Members

    private const string WatermarkSampleOutputDirectory = Program.SampleDirectory + @"Watermark\Output\";
    internal const string WatermarkSampleResourcesDirectory = Program.SampleDirectory + @"Watermark\Resources\";

    #endregion

    #region Constructors

    static WatermarkSample()
    {
      if( !Directory.Exists( WatermarkSample.WatermarkSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( WatermarkSample.WatermarkSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods


    public static void AddPngPictureWatermark()
    {


      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.

    }

    public static void AddJpgPictureWatermark()
    {


      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.

    }

    public static void AddJpgPictureWatermarkOptions()
    {




      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.

    }

    public static void AddJpgPictureWatermarkOddHeader()
    {












      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.

    }

    public static void AddTextWatermark()
    {



      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.

    }

    public static void AddTextWatermarkDiagonal()
    {



      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.

    }

    public static void AddTextWatermarkFont()
    {



      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.

    }

    public static void AddTextWatermarkColor()
    {



      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.

    }

    public static void AddMultipleSections()
    {










      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.

    }


    public static void LoadTextWatermark()
    {



      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.

    }

    public static void LoadPictureWatermark()
    {


      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.

    }

    public static void LoadAndRemoveWatermark()
    {



      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.

    }
#endregion
  }
}
