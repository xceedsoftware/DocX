/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2024 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;

namespace Xceed.Document.NET
{
  internal static class XElementHelpers
  {
    internal static T GetValueToEnum<T>( XElement element )
    {
      if( element == null )
        throw new ArgumentNullException( "element" );

      var value = element.Attribute( XName.Get( "val" ) ).Value;
      foreach( T e in Enum.GetValues( typeof( T ) ) )
      {
        var fi = typeof( T ).GetField( e.ToString() );
        if( fi.GetCustomAttributes( typeof( XmlNameAttribute ), false ).Count() == 0 )
          throw new Exception( String.Format( "Attribute 'XmlNameAttribute' is not assigned to {0} fields!", typeof( T ).Name ) );
        var a = ( XmlNameAttribute )fi.GetCustomAttributes( typeof( XmlNameAttribute ), false ).First();
        if( a.XmlName == value )
          return e;
      }
      throw new ArgumentException( "Invalid element value!" );
    }

    internal static void SetValueFromEnum<T>( XElement element, T value )
    {
      if( element == null )
        throw new ArgumentNullException( "element" );
      element.Attribute( XName.Get( "val" ) ).Value = GetXmlNameFromEnum<T>( value );
    }

    internal static String GetXmlNameFromEnum<T>( T value )
    {
      if( value == null )
        throw new ArgumentNullException( "value" );

      var fi = typeof( T ).GetField( value.ToString() );
      if( fi.GetCustomAttributes( typeof( XmlNameAttribute ), false ).Count() == 0 )
        throw new Exception( String.Format( "Attribute 'XmlNameAttribute' is not assigned to {0} fields!", typeof( T ).Name ) );
      var a = ( XmlNameAttribute )fi.GetCustomAttributes( typeof( XmlNameAttribute ), false ).First();
      return a.XmlName;
    }
  }

  [AttributeUsage( AttributeTargets.Field, Inherited = false, AllowMultiple = false )]
  internal sealed class XmlNameAttribute : Attribute
  {
    public String XmlName
    {
      get; private set;
    }

    public XmlNameAttribute( String xmlName )
    {
      XmlName = xmlName;
    }
  }
}
