/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – DigitalSignature Sample Application
Copyright (c) 2009-2025 - Xceed Software Inc.

This application demonstrates how to digitally sign a document when using the API 
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
  public class DigitalSignatureSample
  {
    #region Private Members

    private const string DigitalSignatureSampleOutputDirectory = Program.SampleDirectory + @"DigitalSignature\Output\";
    private const string DigitalSignatureSampleResourcesDirectory = Program.SampleDirectory + @"DigitalSignature\Resources\";

    #endregion

    #region Constructors

    static DigitalSignatureSample()
    {
      if( !Directory.Exists( DigitalSignatureSample.DigitalSignatureSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( DigitalSignatureSample.DigitalSignatureSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    public static void SignWithSignatureLine()
    {























      // This option is available in .NET Framework and when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void SignWithoutSignatureLine()
    {






      // This option is available in .NET Framework and when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void VerifySignatures()
    {






      // This option is available in .NET Framework and when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void RemoveSignatures()
    {




      // This option is available in .NET Framework and when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    public static void RemoveSignatureLines()
    {



      // This option is available in .NET Framework and when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    #endregion
  }
}
