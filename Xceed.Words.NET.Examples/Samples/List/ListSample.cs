/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – List Sample Application
Copyright (c) 2009-2020 - Xceed Software Inc.

This application demonstrates how to add lists when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.IO;
using Xceed.Document.NET;

namespace Xceed.Words.NET.Examples
{
  public class ListSample
  {
    #region Private Members

    private const string ListSampleResourceDirectory = Program.SampleDirectory + @"List\Resources\";
    private const string ListSampleOutputDirectory = Program.SampleDirectory + @"List\Output\";

    #endregion

    #region Constructors

    static ListSample()
    {
      if( !Directory.Exists( ListSample.ListSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( ListSample.ListSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Create a numbered and a bulleted lists with different listItem's levels.
    /// </summary>
    public static void AddList()
    {
      Console.WriteLine( "\tAddList()" );

      // Create a document.
      using( var document = DocX.Create( ListSample.ListSampleOutputDirectory + @"AddList.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Adding lists into a document" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a numbered list where the first ListItem is starting with number 1.
        var numberedList = document.AddList( "Berries", 0, ListItemType.Numbered, 1 );
        // Add Sub-items(level 1) to the preceding ListItem.
        document.AddListItem( numberedList, "Strawberries", 1 );
        document.AddListItem( numberedList, "Blueberries", 1 );
        document.AddListItem( numberedList, "Raspberries", 1 );
        // Add an item (level 0)
        document.AddListItem( numberedList, "Banana" );
        // Add an item (level 0)
        document.AddListItem( numberedList, "Apple" );
        // Add Sub-items(level 1) to the preceding ListItem.
        document.AddListItem( numberedList, "Red", 1 );
        document.AddListItem( numberedList, "Green", 1 );
        document.AddListItem( numberedList, "Yellow", 1 );

        // Add a bulleted list with its first item.
        var bulletedList = document.AddList( "Canada", 0, ListItemType.Bulleted);
        // Add Sub-items(level 1) to the preceding ListItem.
        document.AddListItem( bulletedList, "Toronto", 1 );
        document.AddListItem( bulletedList, "Montreal", 1 );
        // Add an item (level 0)
        document.AddListItem( bulletedList, "Brazil" );
        // Add an item (level 0)
        document.AddListItem( bulletedList, "USA" );
        // Add Sub-items(level 1) to the preceding ListItem.
        document.AddListItem( bulletedList, "New York", 1 );
        // Add Sub-items(level 2) to the preceding ListItem.
        document.AddListItem( bulletedList, "Brooklyn", 2 );
        document.AddListItem( bulletedList, "Manhattan", 2 );
        document.AddListItem( bulletedList, "Los Angeles", 1 );
        document.AddListItem( bulletedList, "Miami", 1 );
        // Add an item (level 0)
        document.AddListItem( bulletedList, "France" );
        // Add Sub-items(level 1) to the preceding ListItem.
        document.AddListItem( bulletedList, "Paris", 1 );

        // Insert the lists into the document.
        document.InsertParagraph( "This is a Numbered List:\n" );
        document.InsertList( numberedList );
        document.InsertParagraph().SpacingAfter( 40d );
        document.InsertParagraph( "This is a Bulleted List:\n" );
        document.InsertList( bulletedList, new Xceed.Document.NET.Font( "Cooper Black"), 15 );

        document.Save();
        Console.WriteLine( "\tCreated: AddList.docx\n" );
      }
    }

    /// <summary>
    /// Clone a list and modify some ListItems.
    /// </summary>
    public static void CloneLists()
    {



      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }

    #endregion
  }
}
