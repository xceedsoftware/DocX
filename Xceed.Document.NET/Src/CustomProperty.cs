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

namespace Xceed.Document.NET
{
  public class CustomProperty
  {
    #region Public Properties

    /// <summary>
    /// The name of this CustomProperty.
    /// </summary>
    public string Name
    {
      get;
      private set;
    }

    /// <summary>
    /// The value of this CustomProperty.
    /// </summary>
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

    /// <summary>
    /// Create a new CustomProperty to hold a string.
    /// </summary>
    /// <param name="name">The name of this CustomProperty.</param>
    /// <param name="value">The value of this CustomProperty.</param>
    public CustomProperty( string name, string value, Formatting formatting = null ) 
      : this( name, "lpwstr", value, formatting )
    {
    }

    /// <summary>
    /// Create a new CustomProperty to hold an int.
    /// </summary>
    /// <param name="name">The name of this CustomProperty.</param>
    /// <param name="value">The value of this CustomProperty.</param>
    public CustomProperty( string name, int value, Formatting formatting = null )
      : this( name, "i4", value, formatting )
    {
    }

    /// <summary>
    /// Create a new CustomProperty to hold a double.
    /// </summary>
    /// <param name="name">The name of this CustomProperty.</param>
    /// <param name="value">The value of this CustomProperty.</param>
    public CustomProperty( string name, double value, Formatting formatting = null ) 
      : this( name, "r8", value, formatting )
    {
    }

    /// <summary>
    /// Create a new CustomProperty to hold a DateTime.
    /// </summary>
    /// <param name="name">The name of this CustomProperty.</param>
    /// <param name="value">The value of this CustomProperty.</param>
    public CustomProperty( string name, DateTime value, Formatting formatting = null )
      : this( name, "filetime", value.ToUniversalTime(), formatting )
    {
    }

    /// <summary>
    /// Create a new CustomProperty to hold a bool.
    /// </summary>
    /// <param name="name">The name of this CustomProperty.</param>
    /// <param name="value">The value of this CustomProperty.</param>
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
            realValue = bool.Parse( value );
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
