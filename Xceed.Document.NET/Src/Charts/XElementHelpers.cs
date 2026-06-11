/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2026 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System;
using System.Linq;
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
        var a = (XmlNameAttribute)fi.GetCustomAttributes( typeof( XmlNameAttribute ), false ).First();
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

    internal static T GetValueToEnum<T>( XElement element, string attributeName )
    {
      if( element == null )
        throw new ArgumentNullException( nameof( element ) );
      if( string.IsNullOrWhiteSpace( attributeName ) )
        throw new ArgumentNullException( nameof( attributeName ) );

      var attribute = element.Attribute( XName.Get( attributeName ) );
      if( attribute == null )
        throw new ArgumentException( $"Attribute '{attributeName}' not found in element." );

      var value = attribute.Value;
      foreach( T e in Enum.GetValues( typeof( T ) ) )
      {
        var fi = typeof( T ).GetField( e.ToString() );
        var attr = fi.GetCustomAttributes( typeof( XmlNameAttribute ), false ).FirstOrDefault();
        if( attr == null )
          throw new Exception( $"Attribute 'XmlNameAttribute' is not assigned to {typeof( T ).Name}.{e}." );

        var xmlName = ( (XmlNameAttribute)attr ).XmlName;
        if( xmlName == value )
          return e;
      }

      throw new ArgumentException( $"Invalid value '{value}' for enum '{typeof( T ).Name}'." );
    }

    internal static void SetValueFromEnum<T>( XElement element, T value, string attributeName )
    {
      if( element == null )
        throw new ArgumentNullException( nameof( element ) );
      if( string.IsNullOrWhiteSpace( attributeName ) )
        throw new ArgumentNullException( nameof( attributeName ) );

      var attribute = element.Attribute( XName.Get( attributeName ) );
      if( attribute == null )
        element.SetAttributeValue( XName.Get( attributeName ), GetXmlNameFromEnum( value ) );
      else
        attribute.Value = GetXmlNameFromEnum( value );
    }

    internal static string GetXmlNameFromEnum<T>( T value )
    {
      if( object.Equals( value, null ) )
        throw new ArgumentNullException( "value" );

      // Handle Nullable<TEnum>
      var enumType = Nullable.GetUnderlyingType( typeof( T ) ) ?? typeof( T );
      if( !enumType.IsEnum )
        throw new ArgumentException( "T must be an enum or Nullable<enum>", "value" );

      // Get the declared name
      var name = Enum.GetName( enumType, value );
      if( name == null )
        throw new ArgumentException( "Invalid enum value.", "value" );

      var fi = enumType.GetField( name );
      if( fi == null )
        throw new ArgumentException( "Enum field not found.", "value" );

      var attrs = fi.GetCustomAttributes( typeof( XmlNameAttribute ), false );
      if( attrs == null || attrs.Length == 0 )
        throw new Exception(
            string.Format( "Attribute 'XmlNameAttribute' is not assigned to {0}.{1}!", enumType.Name, name ) );

      return ( (XmlNameAttribute)attrs[ 0 ] ).XmlName;
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
