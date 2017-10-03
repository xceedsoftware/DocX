/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace Xceed.Words.NET
{
  /// <summary>
  /// Represents a field of type document property. This field displays the value stored in a custom property.
  /// </summary>
  public class DocProperty : DocXElement
  {

    #region Internal Members

    internal Regex _extractName = new Regex( @"DOCPROPERTY  (?<name>.*)  " );

    #endregion

    #region Public Properties

    /// <summary>
    /// The custom property to display.
    /// </summary>
    public string Name
    {
      get;
      private set;
    }

    #endregion

    #region Constructors

    internal DocProperty( DocX document, XElement xml ) : base( document, xml )
    {
      var instr = Xml.Attribute( XName.Get( "instr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main" ) ).Value;
      this.Name = _extractName.Match( instr.Trim() ).Groups[ "name" ].Value;
    }

    #endregion
  }
}
