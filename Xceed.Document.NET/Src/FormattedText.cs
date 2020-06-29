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
  public class FormattedText : IComparable
  {
    #region Public Members

    public int index;
    public string text;

    #endregion

    #region Private members

    private Formatting _formatting;
    private bool _setInitialFormatting = true;

    #endregion

    #region Public Properties

    public Formatting formatting
    {
      get
      {
        return _formatting;
      }
      set
      {
        _formatting = value;
        if( _setInitialFormatting )
        {
          this.InitialFormatting = ( value != null ) ? value.Clone() : null;
        }
      }
    }

    #endregion

    #region Internal Properties

    internal Formatting InitialFormatting
    {
      get;
      private set;
    }

    #endregion

    #region Constructors

    public FormattedText()
    {
    }

    #endregion

    #region Public Methods

    public int CompareTo( object obj )
    {
      FormattedText other = ( FormattedText )obj;
      FormattedText tf = this;

      if( other.formatting == null || tf.formatting == null )
        return -1;

      return tf.formatting.CompareTo( other.formatting );
    }

    #endregion

    #region Internal Methods

    internal void InternalModifyFormatting( Formatting f )
    {
      _setInitialFormatting = false;
      this.formatting = f;
      _setInitialFormatting = true;
    }

    #endregion
  }
}
