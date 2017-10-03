/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System;
using System.Reflection;
using System.ComponentModel;

namespace Xceed.Words.NET
{
  public static class ExtensionsHeadings
  {

    public static Paragraph Heading( this Paragraph paragraph, HeadingType headingType )
    {
      var description = headingType.EnumDescription();
      paragraph.StyleName = description;
      return paragraph;
    }

    public static string EnumDescription( this Enum enumValue )
    {
      if( (enumValue == null) || (enumValue.ToString() == "0") )
        return string.Empty;

      var enumInfo = enumValue.GetType().GetField( enumValue.ToString() );
      var enumAttributes = ( DescriptionAttribute[] )enumInfo.GetCustomAttributes( typeof( DescriptionAttribute ), false );

      if( enumAttributes.Length > 0 )
        return enumAttributes[ 0 ].Description;
      else
        return enumValue.ToString();
    }

    public static bool HasFlag( this Enum variable, Enum value )
    {
      if( variable == null )
        return false;

      if( value == null )
        throw new ArgumentNullException( "value" );

      if( !Enum.IsDefined( variable.GetType(), value ) )
        throw new ArgumentException( string.Format( "Enumeration type mismatch.  The flag is of type '{0}', was expecting '{1}'.", value.GetType(), variable.GetType() ) );

      var num = Convert.ToUInt64( value );
      return ( ( Convert.ToUInt64( variable ) & num ) == num );
    }
  }
}
