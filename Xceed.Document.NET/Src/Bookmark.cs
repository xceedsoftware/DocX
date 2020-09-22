/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2020 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


namespace Xceed.Document.NET
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

    public void SetText( string text, Formatting formatting = null )
    {
      this.Paragraph.ReplaceAtBookmark( text, this.Name, formatting );
    }

    public void Remove()
    {
      this.Paragraph.RemoveBookmark( this.Name );
    }

    #endregion
  }
}
