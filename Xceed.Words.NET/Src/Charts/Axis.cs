/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System;
using System.Xml.Linq;

namespace Xceed.Words.NET
{
  /// <summary>
  /// Axis base class
  /// </summary>
  public abstract class Axis
  {
    #region Public Properties

    /// <summary>
    /// ID of this Axis 
    /// </summary>
    public String Id
    {
      get
      {
        return Xml.Element( XName.Get( "axId", DocX.c.NamespaceName ) ).Attribute( XName.Get( "val" ) ).Value;
      }
    }

    /// <summary>
    /// Return true if this axis is visible
    /// </summary>
    public Boolean IsVisible
    {
      get
      {
        return Xml.Element( XName.Get( "delete", DocX.c.NamespaceName ) ).Attribute( XName.Get( "val" ) ).Value == "0";
      }
      set
      {
        if( value )
          Xml.Element( XName.Get( "delete", DocX.c.NamespaceName ) ).Attribute( XName.Get( "val" ) ).Value = "0";
        else
          Xml.Element( XName.Get( "delete", DocX.c.NamespaceName ) ).Attribute( XName.Get( "val" ) ).Value = "1";
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

  /// <summary>
  /// Represents Category Axes
  /// </summary>
  public class CategoryAxis : Axis
  {
    internal CategoryAxis( XElement xml )
        : base( xml )
    {
    }

    public CategoryAxis( String id )
        : base( id )
    {
      Xml = XElement.Parse( String.Format(
        @"<c:catAx xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""> 
                <c:axId val=""{0}""/>
                <c:scaling>
                  <c:orientation val=""minMax""/>
                </c:scaling>
                <c:delete val=""0""/>
                <c:axPos val=""b""/>
                <c:majorTickMark val=""out""/>
                <c:minorTickMark val=""none""/>
                <c:tickLblPos val=""nextTo""/>
                <c:crossAx val=""154227840""/>
                <c:crosses val=""autoZero""/>
                <c:auto val=""1""/>
                <c:lblAlgn val=""ctr""/>
                <c:lblOffset val=""100""/>
                <c:noMultiLvlLbl val=""0""/>
              </c:catAx>", id ) );
    }
  }

  /// <summary>
  /// Represents Values Axes
  /// </summary>
  public class ValueAxis : Axis
  {
    internal ValueAxis( XElement xml )
        : base( xml )
    {
    }

    public ValueAxis( String id )
        : base( id )
    {
      Xml = XElement.Parse( String.Format(
        @"<c:valAx xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                <c:axId val=""{0}""/>
                <c:scaling>
                  <c:orientation val=""minMax""/>
                </c:scaling>
                <c:delete val=""0""/>
                <c:axPos val=""l""/>
                <c:numFmt sourceLinked=""0"" formatCode=""General""/>
                <c:majorGridlines/>
                <c:majorTickMark val=""out""/>
                <c:minorTickMark val=""none""/>
                <c:tickLblPos val=""nextTo""/>
                <c:crossAx val=""148921728""/>
                <c:crosses val=""autoZero""/>
                <c:crossBetween val=""between""/>
              </c:valAx>", id ) );
    }
  }
}
