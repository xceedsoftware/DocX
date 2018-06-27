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
using System.Linq;

namespace Xceed.Words.NET.Examples
{
  public class TableSample
  {
    #region Private Members

    private static Random rand = new Random();

    private const string TableSampleResourcesDirectory = Program.SampleDirectory + @"Table\Resources\";
    private const string TableSampleOutputDirectory = Program.SampleDirectory + @"Table\Output\";

    #endregion

    #region Constructors

    static TableSample()
    {
      if( !Directory.Exists( TableSample.TableSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( TableSample.TableSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Create a table, insert rows, image and replace text.
    /// </summary>
    public static void InsertRowAndImageTable()
    {
      Console.WriteLine( "\tInsertRowAndImageTable()" );

      // Create a document.
      using( DocX document = DocX.Create( TableSample.TableSampleOutputDirectory + @"InsertRowAndImageTable.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Inserting table" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add a Table into the document and sets its values.
        var t = document.AddTable( 5, 2 );
        t.Design = TableDesign.ColorfulListAccent1;
        t.Alignment = Alignment.center;
        t.Rows[ 0 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "Mike" );
        t.Rows[ 0 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "65" );
        t.Rows[ 1 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "Kevin" );
        t.Rows[ 1 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "62" );
        t.Rows[ 2 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "Carl" );
        t.Rows[ 2 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "60" );
        t.Rows[ 3 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "Michael" );
        t.Rows[ 3 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "59" );
        t.Rows[ 4 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "Shawn" );
        t.Rows[ 4 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "57" );

        // Add a row at the end of the table and sets its values.
        var r = t.InsertRow();
        r.Cells[ 0 ].Paragraphs[ 0 ].Append( "Mario" );
        r.Cells[ 1 ].Paragraphs[ 0 ].Append( "54" );

        // Add a row at the end of the table which is a copy of another row, and sets its values.
        var newPlayer = t.InsertRow( t.Rows[ 2 ] );
        newPlayer.ReplaceText( "Carl", "Max" );
        newPlayer.ReplaceText( "60", "50" );

        // Add an image into the document.    
        var image = document.AddImage( TableSample.TableSampleResourcesDirectory + @"logo_xceed.png" );
        // Create a picture from image.
        var picture = image.CreatePicture( 25, 100 );

        // Calculate totals points from second column in table.
        var totalPts = 0;
        foreach( var row in t.Rows )
        {
          totalPts += int.Parse( row.Cells[ 1 ].Paragraphs[ 0 ].Text );
        }

        // Add a row at the end of the table and sets its values.
        var totalRow = t.InsertRow();
        totalRow.Cells[ 0 ].Paragraphs[ 0 ].Append( "Total for " ).AppendPicture( picture );
        totalRow.Cells[ 1 ].Paragraphs[ 0 ].Append( totalPts.ToString() );
        totalRow.Cells[ 1 ].VerticalAlignment = VerticalAlignment.Center;

        // Insert a new Paragraph into the document.
        var p = document.InsertParagraph( "Xceed Top Players Points:" );
        p.SpacingAfter( 40d );

        // Insert the Table after the Paragraph.
        p.InsertTableAfterSelf( t );

        document.Save();
        Console.WriteLine( "\tCreated: InsertRowAndImageTable.docx\n" );
      }
    }

    /// <summary>
    /// Create a table and set the text direction of each cell.
    /// </summary>
    public static void TextDirectionTable()
    {
      Console.WriteLine( "\tTextDirectionTable()" );

      // Create a document.
      using( DocX document = DocX.Create( TableSample.TableSampleOutputDirectory + @"TextDirectionTable.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Text Direction of Table's cells" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a table.
        var table = document.AddTable( 2, 3 );
        table.Design = TableDesign.ColorfulList;

        // Set the table's values.
        table.Rows[ 0 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "First" );
        table.Rows[ 0 ].Cells[ 0 ].TextDirection = TextDirection.btLr;
        table.Rows[ 0 ].Cells[ 0 ].Paragraphs[ 0 ].Spacing( 5d );
        table.Rows[ 0 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "Second" );
        table.Rows[ 0 ].Cells[ 1 ].TextDirection = TextDirection.right;
        table.Rows[ 0 ].Cells[ 2 ].Paragraphs[ 0 ].Append( "Third" );
        table.Rows[ 0 ].Cells[ 2 ].Paragraphs[ 0 ].Spacing( 5d );
        table.Rows[ 0 ].Cells[ 2 ].TextDirection = TextDirection.btLr;
        table.Rows[ 1 ].Cells[ 0 ].Paragraphs[ 0 ].Append( "Fourth" );
        table.Rows[ 1 ].Cells[ 0 ].TextDirection = TextDirection.btLr;
        table.Rows[ 1 ].Cells[ 0 ].Paragraphs[ 0 ].Spacing( 5d );
        table.Rows[ 1 ].Cells[ 1 ].Paragraphs[ 0 ].Append( "Fifth" );
        table.Rows[ 1 ].Cells[ 2 ].Paragraphs[ 0 ].Append( "Sixth" ).Color( Color.White );
        table.Rows[ 1 ].Cells[ 2 ].TextDirection = TextDirection.btLr;
        // Last cell have a green background
        table.Rows[ 1 ].Cells[ 2 ].FillColor = Color.Green;
        table.Rows[ 1 ].Cells[ 2 ].Paragraphs[ 0 ].Spacing( 5d );

        // Set the table's column width.
        table.SetWidths( new float[] { 200, 300, 100 } );

        // Add the table into the document.
        document.InsertTable( table );

        document.Save();
        Console.WriteLine( "\tCreated: TextDirectionTable.docx\n" );
      }
    }

    /// <summary>
    /// Load a document, gets its table and replace the default row with updated copies of it.
    /// </summary>
    public static void CreateRowsFromTemplate()
    {
      Console.WriteLine( "\tCreateRowsFromTemplate()" );

      // Load a document
      using( DocX document = DocX.Load( TableSample.TableSampleResourcesDirectory + @"DocumentWithTemplateTable.docx" ) )
      {
        // get the table with caption "GROCERY_LIST" from the document.
        var groceryListTable = document.Tables.FirstOrDefault( t => t.TableCaption == "GROCERY_LIST" );
        if( groceryListTable == null )
        {
          Console.WriteLine( "\tError, couldn't find table with caption GROCERY_LIST in current document." );
        }
        else
        {
          if( groceryListTable.RowCount > 1 )
          {
            // Get the row pattern of the second row.
            var rowPattern = groceryListTable.Rows[ 1 ];

            // Add items (rows) to the grocery list.
            TableSample.AddItemToTable( groceryListTable, rowPattern, "Banana" );
            TableSample.AddItemToTable( groceryListTable, rowPattern, "Strawberry" );
            TableSample.AddItemToTable( groceryListTable, rowPattern, "Chicken" );
            TableSample.AddItemToTable( groceryListTable, rowPattern, "Bread" );
            TableSample.AddItemToTable( groceryListTable, rowPattern, "Eggs" );
            TableSample.AddItemToTable( groceryListTable, rowPattern, "Salad" );

            // Remove the pattern row.
            rowPattern.Remove();
          }
        }

        document.SaveAs( TableSample.TableSampleOutputDirectory + @"CreateTableFromTemplate.docx" );
        Console.WriteLine( "\tCreated: CreateTableFromTemplate.docx\n" );
      }
    }

    /// <summary>
    /// Add a Table in a document where its columns will have a specific width. In addition,
    /// the left margin of the row cells will be removed for all rows except the first.
    /// Finally, a blank border will be set for the table's top and bottom borders.
    /// </summary>
    public static void ColumnsWidth()
    {
      Console.WriteLine( "\tColumnsWidth()" );

      // Create a document
      using( DocX document = DocX.Create( TableSample.TableSampleOutputDirectory + @"ColumnsWidth.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Columns width" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Insert a title paragraph.
        var p = document.InsertParagraph( "In the following table, the cell's left margin has been removed for rows 2-6 as well as the top/bottom table's borders." ).Bold();
        p.Alignment = Alignment.center;
        p.SpacingAfter( 40d );

        // Add a table in a document of 1 row and 3 columns.
        var columnWidths = new float[] { 100f, 300f, 200f };
        var t = document.InsertTable( 1, columnWidths.Length );

        // Set the table's column width and background 
        t.SetWidths( columnWidths );
        t.Design = TableDesign.TableGrid;
        t.AutoFit = AutoFit.Contents;

        var row = t.Rows.First();

        // Fill in the columns of the first row in the table.
        for( int i = 0; i < row.Cells.Count; ++i )
        {
          row.Cells[i].Paragraphs.First().Append( "Data " + i );
        }

        // Add rows in the table.
        for( int i = 0; i < 5; i++ )
        {
          var newRow = t.InsertRow();

          // Fill in the columns of the new rows.
          for( int j = 0; j < newRow.Cells.Count; ++j )
          {
            var newCell = newRow.Cells[ j ];
            newCell.Paragraphs.First().Append( "Data " + i );
            // Remove the left margin of the new cells.
            newCell.MarginLeft = 0;
          }
        }

        // Set a blank border for the table's top/bottom borders.
        var blankBorder = new Border( BorderStyle.Tcbs_none, 0, 0, Color.White );
        t.SetBorder( TableBorderType.Bottom, blankBorder );
        t.SetBorder( TableBorderType.Top, blankBorder );

        document.Save();
        Console.WriteLine( "\tCreated: ColumnsWidth.docx\n" );
      }
    }

    /// <summary>
    /// Add a table and merged some cells. Individual cells can also be removed by shifting their right neighbors to the left.
    /// </summary>
    public static void MergeCells()
    {
      Console.WriteLine( "\tMergeCells()" );

      // Create a document.
      using( DocX document = DocX.Create( TableSample.TableSampleOutputDirectory + @"MergeCells.docx" ) )
      {
        // Add a title.
        document.InsertParagraph( "Merge and delete cells" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Add A table .               
        var t = document.AddTable( 3, 2 );
        t.Design = TableDesign.TableGrid;

        var t1 = document.InsertTable( t );

        // Add 4 columns in the table.
        t1.InsertColumn( t1.ColumnCount, true );
        t1.InsertColumn( t1.ColumnCount, true );
        t1.InsertColumn( t1.ColumnCount, true );
        t1.InsertColumn( t1.ColumnCount, true );

        // Merged Cells 1 to 4 in first row of the table.
        t1.Rows[ 0 ].MergeCells( 1, 4 );

        // Merged the last 2 Cells in the second row of the table.
        var columnCount = t1.Rows[ 1 ].ColumnCount;
        t1.Rows[ 1 ].MergeCells( columnCount - 2, columnCount - 1 );

        // Add text in each cell of the table.
        foreach( var r in t1.Rows )
        {
          for( int i = 0; i < r.Cells.Count; ++i )
          {
            var c = r.Cells[ i ];
            c.Paragraphs[ 0 ].InsertText( "Column " + i );
            c.Paragraphs[ 0 ].Alignment = Alignment.center;
          }
        }

        // Delete the second cell from the third row and shift the cells on its right by 1 to the left.
        t1.DeleteAndShiftCellsLeft( 2, 1 );

        document.Save();
        Console.WriteLine( "\tCreated: MergeCells.docx\n" );
      }
    }

    #endregion

    #region Private Methods

    private static void AddItemToTable( Table table, Row rowPattern, string productName )
    {
      // Gets a random unit price and quantity.
      var unitPrice = Math.Round( rand.NextDouble(), 2 );
      var unitQuantity = rand.Next( 1, 10 );

      // Insert a copy of the rowPattern at the last index in the table.
      var newItem = table.InsertRow( rowPattern, table.RowCount - 1 );

      // Replace the default values of the newly inserted row.
      newItem.ReplaceText( "%PRODUCT_NAME%", productName );
      newItem.ReplaceText( "%PRODUCT_UNITPRICE%", "$ " + unitPrice.ToString( "N2" ) );
      newItem.ReplaceText( "%PRODUCT_QUANTITY%", unitQuantity.ToString() );
      newItem.ReplaceText( "%PRODUCT_TOTALPRICE%", "$ " + ( unitPrice * unitQuantity ).ToString( "N2" ) );
    }

    #endregion
  }
}
