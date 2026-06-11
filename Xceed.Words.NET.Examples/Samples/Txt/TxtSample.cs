/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2026 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System;
using System.IO;

namespace Xceed.Words.NET.Examples
{
  public class TxtSample
  {

    private const string TxtSampleResourcesDirectory = Program.SampleDirectory + @"Txt\Resources\";
    private const string TxtSampleOutputDirectory = Program.SampleDirectory + @"Txt\Output\";

    #region Constructors

    static TxtSample()
    {
      if( !Directory.Exists( TxtSample.TxtSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( TxtSample.TxtSampleOutputDirectory );
      }
    }

    #endregion

    public static void CreateTxt()
    {


      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void SaveToTxt()
    {

      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void LoadToTxt()
    {


      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }
  }
}
