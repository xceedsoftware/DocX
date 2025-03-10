/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System;

namespace Xceed.Document.NET
{
  public class CustomProperty
  {
    #region Public Properties

    public string Name
    {
      get;
      private set;
    }

    public object Value
    {
      get;
      private set;
    }

    #endregion

    #region Internal Properties

    internal string Type
    {
      get;
      private set;
    }

    internal Formatting Formatting
    {
      get;
      set;
    }

    #endregion

    #region Constructors

    public CustomProperty( string name, string value, Formatting formatting = null ) 
      : this( name, "lpwstr", value, formatting )
    {
    }

    public CustomProperty( string name, int value, Formatting formatting = null )
      : this( name, "i4", value, formatting )
    {
    }

    public CustomProperty( string name, double value, Formatting formatting = null ) 
      : this( name, "r8", value, formatting )
    {
    }

    public CustomProperty( string name, DateTime value, Formatting formatting = null )
      : this( name, "filetime", value.ToUniversalTime(), formatting )
    {
    }

    public CustomProperty( string name, bool value, Formatting formatting = null ) 
      : this( name, "bool", value, formatting )
    {
    }

    internal CustomProperty( string name, string type, string value, Formatting formatting = null )
    {
      object realValue;
      switch( type )
      {
        case "lpwstr":
          {
            realValue = value;
            break;
          }

        case "i4":
          {
            realValue = int.Parse( value, System.Globalization.CultureInfo.InvariantCulture );
            break;
          }

        case "r8":
          {
            realValue = Double.Parse( value, System.Globalization.CultureInfo.InvariantCulture );
            break;
          }

        case "filetime":
          {
            realValue = DateTime.Parse( value, System.Globalization.CultureInfo.InvariantCulture );
            break;
          }

        case "bool":
          {
            realValue = ( value == "0" )
                        ? false
                        : ( value == "1" ) ? true : bool.Parse( value );
            break;
          }

        default:
          throw new Exception();
      }

      this.Name = name;
      this.Type = type;
      this.Value = realValue;
      this.Formatting = formatting;
    }

    private CustomProperty( string name, string type, object value, Formatting formatting = null )
    {

      this.Name = name;
      this.Type = type;
      this.Value = value;
      this.Formatting = formatting;
    }

    #endregion
  }
}
