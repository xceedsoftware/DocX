/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Xml.Linq;

namespace Xceed.Words.NET
{
  internal static class Extensions
  {
    internal static string ToHex( this Color source )
    {
      byte red = source.R;
      byte green = source.G;
      byte blue = source.B;

      string redHex = red.ToString( "X" );
      if( redHex.Length < 2 )
        redHex = "0" + redHex;

      string blueHex = blue.ToString( "X" );
      if( blueHex.Length < 2 )
        blueHex = "0" + blueHex;

      string greenHex = green.ToString( "X" );
      if( greenHex.Length < 2 )
        greenHex = "0" + greenHex;

      return string.Format( "{0}{1}{2}", redHex, greenHex, blueHex );
    }

    public static void Flatten( this XElement e, XName name, List<XElement> flat )
    {
      // Add this element (without its children) to the flat list.
      XElement clone = CloneElement( e );
      clone.Elements().Remove();

      // Filter elements using XName.
      if( clone.Name == name )
        flat.Add( clone );

      // Process the children.
      if( e.HasElements )
        foreach( XElement elem in e.Elements( name ) ) // Filter elements using XName
          elem.Flatten( name, flat );
    }

    public static string GetAttribute( this XElement el, XName name, string defaultValue = "" )
    {
      var attribute = el.Attribute( name );
      if( attribute != null )
        return attribute.Value;

      return defaultValue;
    }

    /// <summary>
    /// Sets margin for all the pages in a DocX document in inches.
    /// </summary>
    /// <param name="document"></param>
    /// <param name="top">Margin from the top. -1 for no change.</param>
    /// <param name="bottom">Margin from the bottom. -1 for no change.</param>
    /// <param name="right">Margin from the right. -1 for no change.</param>
    /// <param name="left">Margin from the left. -1 for no change.</param>
    public static void SetMargin( this DocX document, float top, float bottom, float right, float left )
    {
      var xNameSpace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
      var tempElement = document.PageLayout.Xml.Descendants( xNameSpace + "pgMar" );
      var multiplier = 1440;

      foreach( var item in tempElement )
      {
        if( top != -1 )
        {
          item.SetAttributeValue( xNameSpace + "top", multiplier * top );
        }
        if( bottom != -1 )
        {
          item.SetAttributeValue( xNameSpace + "bottom", multiplier * bottom );
        }
        if( right != -1 )
        {
          item.SetAttributeValue( xNameSpace + "right", multiplier * right );
        }
        if( left != -1 )
        {
          item.SetAttributeValue( xNameSpace + "left", multiplier * left );
        }
      }
    }

    private static XElement CloneElement( XElement element )
    {
      return new XElement( element.Name,
          element.Attributes(),
          element.Nodes().Select( n =>
           {
            XElement e = n as XElement;
            if( e != null )
              return CloneElement( e );
            return n;
          }
          )
      );
    }
  }
}
