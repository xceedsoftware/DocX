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
Xceed Words for .NET – Xceed.Words.NET.Examples – Hyphenation Sample Application
Copyright (c) 2009-2024 - Xceed Software Inc.

This application demonstrates how to add and update text hyphenation when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.IO;
using Xceed.Document.NET;

namespace Xceed.Words.NET.Examples
{
  public class HyphenationSample
  {
    #region Private Members

    private const string HyphenationSampleResourceDirectory = Program.SampleDirectory + @"Hyphenation\Resources\";
    private const string HyphenationSampleOutputDirectory = Program.SampleDirectory + @"Hyphenation\Output\";

    #endregion

    #region Constructors

    static HyphenationSample()
    {
      if( !Directory.Exists( HyphenationSample.HyphenationSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( HyphenationSample.HyphenationSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    public static void CreateHyphenation()
    {






      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void UpdateHyphenation()
    {


      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    #endregion
  }
}
