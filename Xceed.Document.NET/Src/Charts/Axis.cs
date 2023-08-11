/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2023 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace Xceed.Document.NET
{

  /// <summary>
  /// Axis base class
  /// </summary>
  public abstract class Axis
  {
    #region Private properties


    #endregion

    #region Public Properties

    /// <summary>
    /// ID of this Axis 
    /// </summary>
    public string Id
    {
      get
      {
        return Xml.Element( XName.Get( "axId", Document.c.NamespaceName ) ).Attribute( XName.Get( "val" ) ).Value;
      }
    }

    /// <summary>
    /// Return true if this axis is visible
    /// </summary>
    public bool IsVisible
    {
      get
      {
        return Xml.Element( XName.Get( "delete", Document.c.NamespaceName ) ).Attribute( XName.Get( "val" ) ).Value == "0";
      }
      set
      {
        if( value )
          Xml.Element( XName.Get( "delete", Document.c.NamespaceName ) ).Attribute( XName.Get( "val" ) ).Value = "0";
        else
          Xml.Element( XName.Get( "delete", Document.c.NamespaceName ) ).Attribute( XName.Get( "val" ) ).Value = "1";
      }
    }





































    #endregion














    #region Internal Properties

    /// <summary>
    /// Axis xml element
    /// </summary>
    internal XElement Xml
    {
      get; set;
    }

    #endregion

    #region Constructors

    internal Axis( XElement xml )
    {
      Xml = xml;
    }

    public Axis( String id )
    {
    }

    #endregion
  }




















































}
