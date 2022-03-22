/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2020 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System;
using System.Xml.Linq;

namespace Xceed.Document.NET
{
  public class ValueAxis : Axis
  {
    /// <summary>
    /// Represents Values Axes
    /// </summary>
    internal ValueAxis(XElement xml)
        : base(xml)
    {
    }

    public ValueAxis(String id)
        : base(id)
    {
      Xml = XElement.Parse(String.Format(
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
              </c:valAx>", id));
    }













  }
}
