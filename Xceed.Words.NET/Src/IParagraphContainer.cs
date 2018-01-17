﻿/*************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2016 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  ***********************************************************************************/

using System.Collections.ObjectModel;

namespace Xceed.Words.NET
{
    public interface IParagraphContainer
    {
        ReadOnlyCollection<Paragraph> Paragraphs
        {
            get;
        }
    }
}
