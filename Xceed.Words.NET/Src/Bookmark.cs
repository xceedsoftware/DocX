/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

namespace Xceed.Words.NET
{
  public class Bookmark
  {
    #region Public Properties

    public string Name
    {
      get; set;
    }
    public Paragraph Paragraph
    {
      get; set;
    }

    #endregion

    #region Constructors

    public Bookmark()
    {
    }

    #endregion

    #region Public Methods

    public void SetText( string text )
    {
      this.Paragraph.ReplaceAtBookmark( text, this.Name );
    }

    #endregion
  }
}
