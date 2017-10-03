/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;

namespace Xceed.Words.NET
{
  internal static class XElementHelpers
  {
    /// <summary>
    /// Get value from XElement and convert it to enum
    /// </summary>
    /// <typeparam name="T">Enum type</typeparam>        
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

    /// <summary>
    /// Convert value to xml string and set it into XElement
    /// </summary>
    /// <typeparam name="T">Enum type</typeparam> 
    internal static void SetValueFromEnum<T>( XElement element, T value )
    {
      if( element == null )
        throw new ArgumentNullException( "element" );
      element.Attribute( XName.Get( "val" ) ).Value = GetXmlNameFromEnum<T>( value );
    }

    /// <summary>
    /// Return xml string for this value
    /// </summary>
    /// <typeparam name="T">Enum type</typeparam> 
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

  /// <summary>
  /// This attribute applied to enum's fields for definition their's real xml names in DocX file.
  /// </summary>
  /// <example>
  /// public enum MyEnum
  /// {
  ///    [XmlName("one")] // This means, that xml element has 'val="one"'
  ///    ValueOne,
  ///    [XmlName("two")] // This means, that xml element has 'val="two"'
  ///    ValueTwo
  /// }
  /// </example>
  [AttributeUsage( AttributeTargets.Field, Inherited = false, AllowMultiple = false )]
  internal sealed class XmlNameAttribute : Attribute
  {
    /// <summary>
    /// Real xml name
    /// </summary>
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
