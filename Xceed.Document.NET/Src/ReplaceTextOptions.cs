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
using System.Text.RegularExpressions;

namespace Xceed.Document.NET
{
  public abstract class ReplaceTextOptionsBase
  {
    internal ReplaceTextOptionsBase()
    {
      this.EndIndex = -1;
      this.FormattingToMatchOptions = MatchFormattingOptions.SubsetMatch;
      this.RegExOptions = RegexOptions.None;
      this.RemoveEmptyParagraph = true;
      this.StartIndex = -1;
    }

    public int EndIndex { get; set; }

    public Formatting FormattingToMatch { get; set; }

    public MatchFormattingOptions FormattingToMatchOptions { get; set; }

    public RegexOptions RegExOptions { get; set; }

    public bool RemoveEmptyParagraph { get; set; }

    public int StartIndex { get; set; }

    public bool StopAfterOneReplacement { get; set; }

    public bool TrackChanges { get; set; }
  }

  public class StringReplaceTextOptions : ReplaceTextOptionsBase
  {
    public StringReplaceTextOptions()
      : base()
    {
      this.EscapeRegEx = true;
    }

    public bool EscapeRegEx { get; set; }

    public Formatting NewFormatting { get; set; }

    public string NewValue { get; set; }

    public string SearchValue { get; set; }

    public bool UseRegExSubstitutions { get; set; }
  }

  public class FunctionReplaceTextOptions : ReplaceTextOptionsBase
  {
    public FunctionReplaceTextOptions()
      : base()
    {
    }

    public string FindPattern { get; set; }

    public Formatting NewFormatting { get; set; }

    public Func<string, string> RegexMatchHandler { get; set; }
  }

  public class ObjectReplaceTextOptions : ReplaceTextOptionsBase
  {
    public ObjectReplaceTextOptions()
      : base()
    {
      this.EscapeRegEx = true;
    }

    public DocumentElement NewObject { get; set; }

    public bool EscapeRegEx { get; set; }

    public string SearchValue { get; set; }
  }
}
