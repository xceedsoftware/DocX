﻿/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace Xceed.Document.NET
{
  public class DocProperty : DocumentElement
  {

    #region Internal Members

    internal Regex _extractName = new Regex( @"DOCPROPERTY  (?<name>.*)  " );

    #endregion

    #region Public Properties

    public string Name
    {
      get;
      private set;
    }

    #endregion

    #region Constructors

    internal DocProperty( Document document, XElement xml ) : base( document, xml )
    {
      var instr = Xml.Attribute( XName.Get( "instr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main" ) ).Value;
      this.Name = _extractName.Match( instr.Trim() ).Groups[ "name" ].Value;
    }

    #endregion
  }
}
