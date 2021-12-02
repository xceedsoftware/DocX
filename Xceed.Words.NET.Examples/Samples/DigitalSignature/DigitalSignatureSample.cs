/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – DigitalSignature Sample Application
Copyright (c) 2009-2021 - Xceed Software Inc.

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

    /// <summary>
    /// Create a document, add 2 SignatureLines and digitally sign it.
    /// </summary>
    public static void SignWithSignatureLine()
    {























      // This option is available in .NET Framework and when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    /// <summary>
    /// Create a document and sign it without SignatureLines.
    /// </summary>
    public static void SignWithoutSignatureLine()
    {






      // This option is available in .NET Framework and when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    /// <summary>
    /// Load a document and verify the signatures and SignatureLines validity.
    /// </summary>
    public static void VerifySignatures()
    {






      // This option is available in .NET Framework and when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    /// <summary>
    /// Load a document and remove all the signatures, but keep the SignatureLines.
    /// </summary>
    public static void RemoveSignatures()
    {




      // This option is available in .NET Framework and when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    /// <summary>
    /// Load a document and remove all the signatureLines.
    /// </summary>
    public static void RemoveSignatureLines()
    {



      // This option is available in .NET Framework and when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    #endregion
  }
}
