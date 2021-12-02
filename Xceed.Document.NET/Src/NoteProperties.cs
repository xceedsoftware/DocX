/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2020 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System.ComponentModel;

namespace Xceed.Document.NET
{
  public class NoteProperties : INotifyPropertyChanged
  {
    #region Private Members

    private NoteNumberFormat _numberFormat = NoteNumberFormat.number;
    private int _numberStart = 1;

    #endregion

    #region Public Properties

    public NoteNumberFormat NumberFormat
    {
      get
      {
        return _numberFormat;
      }
      set
      {
        _numberFormat = value;
        OnPropertyChanged( "NumberFormat" );
      }
    }

    public int NumberStart
    {
      get
      {
        return _numberStart;
      }
      set
      {
        _numberStart = value;
        OnPropertyChanged( "NumberStart" );
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
