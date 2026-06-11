/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2026 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System.ComponentModel;

namespace Xceed.Document.NET
{
  public class TableLook : INotifyPropertyChanged
  {
    #region Private members

    private bool _firstRow;
    private bool _lastRow;
    private bool _firstColumn;
    private bool _lastColumn;
    private bool _noHorizontalBanding;
    private bool _noVerticalBanding;

    #endregion

    #region Public Properties

    public bool FirstRow
    {
      get
      {
        return _firstRow;
      }
      set
      {
        _firstRow = value;
        OnPropertyChanged( "FirstRow" );
      }
    }

    public bool LastRow
    {
      get
      {
        return _lastRow;
      }
      set
      {
        _lastRow = value;
        OnPropertyChanged( "LastRow" );
      }
    }

    public bool FirstColumn
    {
      get
      {
        return _firstColumn;
      }
      set
      {
        _firstColumn = value;
        OnPropertyChanged( "FirstColumn" );
      }
    }

    public bool LastColumn
    {
      get
      {
        return _lastColumn;
      }
      set
      {
        _lastColumn = value;
        OnPropertyChanged( "LastColumn" );
      }
    }

    public bool NoHorizontalBanding
    {
      get
      {
        return _noHorizontalBanding;
      }
      set
      {
        _noHorizontalBanding = value;
        OnPropertyChanged( "NoHorizontalBanding" );
      }
    }

    public bool NoVerticalBanding
    {
      get
      {
        return _noVerticalBanding;
      }
      set
      {
        _noVerticalBanding = value;
        OnPropertyChanged( "NoVerticalBanding" );
      }
    }

    #endregion

    #region Constructors

    public TableLook()
    {
    }

    public TableLook( bool firstRow, bool lastRow, bool firstColumn, bool lastColumn, bool noHorizontalBanding, bool noVerticalBanding )
    {
      this.FirstRow = firstRow;
      this.LastRow = lastRow;
      this.FirstColumn = firstColumn;
      this.LastColumn = lastColumn;
      this.NoHorizontalBanding = noHorizontalBanding;
      this.NoVerticalBanding = noVerticalBanding;
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
