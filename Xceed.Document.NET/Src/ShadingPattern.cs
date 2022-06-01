/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2022 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System.ComponentModel;
using System.Drawing;

namespace Xceed.Document.NET
{
  public class ShadingPattern : INotifyPropertyChanged
  {
    #region Private Members

    private Color _fill;
    private PatternStyle _style;
    private Color _styleColor;

    #endregion

    #region Public Properties

    public Color Fill
    {
      get
      {
        return _fill;
      }
      set
      {
        _fill = value;
        OnPropertyChanged( "Fill" );
      }
    }

    public PatternStyle Style
    {
      get
      {
        return _style;
      }
      set
      {
        _style = value;
        OnPropertyChanged( "Style" );
      }
    }

    public Color StyleColor
    {
      get
      {
        return _styleColor;
      }
      set
      {
        _styleColor = value;
        OnPropertyChanged( "StyleColor" );
      }
    }

    #endregion


    #region INotifyPropertyChanged

    public event PropertyChangedEventHandler PropertyChanged;
    protected void OnPropertyChanged( string propertyName )
    {
      if( PropertyChanged != null )
      {
        PropertyChanged( this, new PropertyChangedEventArgs( propertyName ) );
      }
    }

    #endregion 
  }
}
