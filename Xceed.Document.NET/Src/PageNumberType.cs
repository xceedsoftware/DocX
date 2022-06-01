/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2022 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


using System;
using System.ComponentModel;

namespace Xceed.Document.NET
{
  public class PageNumberType : INotifyPropertyChanged
  {
    #region Private Members

    private int? _pageNumberStart;
    private int _chapterStyle;
    private NumberingFormat _pageNumberFormat;
    private ChapterSeperator _chapterNumberSeperator;

    #endregion

    #region Public Properties

    #region PageNumberStart

    public int? PageNumberStart
    {
      get
      {
        return _pageNumberStart;
      }
      set
      {
        _pageNumberStart = value;
        OnPropertyChanged("PageNumberStart");
      }
    }

    #endregion

    #region Chapter Style

    public int ChapterStyle
    {
      get
      {
        return _chapterStyle;
      }

      set
      {
        if (value <= 9 && value > 0)
        {
          _chapterStyle = value;
          OnPropertyChanged("ChapterStyle");
        }
        else
        {
          throw new Exception("The index number cannot be less than 0 and over 9.");
        }
      }
    }

    #endregion

    #region PageNumberFormat

    public NumberingFormat PageNumberFormat
    {
      get
      {
        return _pageNumberFormat;
      }

      set
      {
        _pageNumberFormat = value;
        OnPropertyChanged("PageNumberFormat");
      }
    }

    #endregion

    #region ChapterNumberSeperator

    public ChapterSeperator ChapterNumberSeperator
    {
      get
      {
        return _chapterNumberSeperator;
      }

      set
      {
        _chapterNumberSeperator = value;
        OnPropertyChanged("ChapterNumberSeperator");
      }
    }

    #endregion

    #endregion

    #region Constructors

    public PageNumberType()
    {
      _pageNumberStart = null;
      _chapterStyle = 1;
      _pageNumberFormat = NumberingFormat.decimalNormal;
      _chapterNumberSeperator = ChapterSeperator.hyphen;
    }

    #endregion

    #region INotifyPropertyChanged

    public event PropertyChangedEventHandler PropertyChanged;
    protected void OnPropertyChanged(string propertyName)
    {
      if (PropertyChanged != null)
      {
        PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
      }
    }

    #endregion
  }
}
