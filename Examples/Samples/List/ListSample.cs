/***************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2017 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  *************************************************************************************/

using System;
using System.Drawing;
using System.IO;

namespace Xceed.Words.NET.Examples
{
  public class ListSample
  {
    #region Private Members

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
      using( DocX document = DocX.Create( ListSample.ListSampleOutputDirectory + @"AddList.docx" ) )
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
        document.InsertList( bulletedList, new Font("Cooper Black"), 15 );

        document.Save();
        Console.WriteLine( "\tCreated: AddList.docx\n" );
      }
    }

    #endregion
  }
}
