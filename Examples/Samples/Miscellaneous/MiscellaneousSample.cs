/***************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2017 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  *************************************************************************************/
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;

namespace Xceed.Words.NET.Examples
{
  public class MiscellaneousSample
  {
    #region Private Members

    private const string MiscellaneousSampleResourcesDirectory = Program.SampleDirectory + @"Miscellaneous\Resources\";
    private const string MiscellaneousSampleOutputDirectory = Program.SampleDirectory + @"Miscellaneous\Output\";

    #endregion

    #region Constructors

    static MiscellaneousSample()
    {
      if( !Directory.Exists( MiscellaneousSample.MiscellaneousSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( MiscellaneousSample.MiscellaneousSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Create a file and add a picture, a table, an hyperlink, paragraphs, a bulleted list and a numbered list.
    /// </summary>
    public static void CreateRecipe()
    {
      Console.WriteLine( "\tCreateRecipe()" );

      // Create a new document.
      using( DocX document = DocX.Create( MiscellaneousSample.MiscellaneousSampleOutputDirectory + @"CreateRecipe.docx" ) )
      {
        // Create a rotated picture from existing image.
        var image = document.AddImage( MiscellaneousSample.MiscellaneousSampleResourcesDirectory + @"cupcake.png" );
        var picture = image.CreatePicture();
        picture.Rotation = 20;

        // Create an hyperlink.
        var hyperlink = document.AddHyperlink( "Food.com", new Uri( "http://www.food.com/recipe/simple-vanilla-cupcakes-178370" ) );

        // Create a bulleted list for the ingredients.
        var bulletsList = document.AddList( "2 cups of flour", 0, ListItemType.Bulleted );
        document.AddListItem( bulletsList, "3⁄4 cup of sugar" );
        document.AddListItem( bulletsList, "1⁄2 cup of butter" );
        document.AddListItem( bulletsList, "2 eggs" );
        document.AddListItem( bulletsList, "1 cup of milk" );
        document.AddListItem( bulletsList, "2 teaspoons of baking powder" );
        document.AddListItem( bulletsList, "1⁄2 teaspoon of salt" );
        document.AddListItem( bulletsList, "1 teaspoon of vanilla essence" );

        // Create a table for text and the picture.
        var table = document.AddTable( 1, 2 );
        table.Design = TableDesign.LightListAccent3;
        table.AutoFit = AutoFit.Window;
        table.Rows[ 0 ].Cells[ 0 ].Paragraphs[ 0 ].AppendLine().AppendLine().Append( "Simple Vanilla Cupcakes Recipe" ).FontSize( 20 ).Font( new Font( "Comic Sans MS" ) );
        table.Rows[ 0 ].Cells[ 1 ].Paragraphs[ 0 ].AppendPicture( picture );

        // Create a numbered list for the directions.
        var recipeList = document.AddList( "Preheat oven to 375F and fill muffin cups with papers.", 0, ListItemType.Numbered, 1 );
        document.AddListItem( recipeList, "Mix butter and sugar until light and fluffy." );
        document.AddListItem( recipeList, "Beat in the eggs, one at a time.", 1 );
        document.AddListItem( recipeList, "Add the flour, baking powder and salt, alternate with milk and beat well." );
        document.AddListItem( recipeList, "Add in vanilla.", 1 );
        document.AddListItem( recipeList, "Divide in the pans and bake for 18 minutes." );
        document.AddListItem( recipeList, "Let cool 5 minutes an eat.", 1 );

        // Insert the data in page.
        document.InsertTable( table );
        var paragraph = document.InsertParagraph();
        paragraph.AppendLine();
        paragraph.AppendLine();
        paragraph.AppendLine( "Ingredients" ).FontSize( 15 ).Bold().Color(Color.BlueViolet);
        document.InsertList( bulletsList );
        var paragraph2 = document.InsertParagraph();
        paragraph2.AppendLine();
        paragraph2.AppendLine( "Directions" ).FontSize( 15 ).Bold().Color( Color.BlueViolet );
        document.InsertList( recipeList );
        var paragraph3 = document.InsertParagraph();
        paragraph3.AppendLine();
        paragraph3.AppendLine( "Reference: " ).AppendHyperlink( hyperlink ).Color( Color.Blue ).UnderlineColor( Color.Blue ).Append( "." );

        // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: CreateRecipe.docx\n" );
      }
    }

    /// <summary>
    /// Create a document and add headers/footers with tables and pictures, paragraphs and charts.
    /// </summary>
    public static void CompanyReport()
    {
      Console.WriteLine( "\tCompanyReport()" );

      // Create a new document.
      using( DocX document = DocX.Create( MiscellaneousSample.MiscellaneousSampleOutputDirectory + @"CompanyReport.docx" ) )
      {
        // Add headers and footers.
        document.AddHeaders();
        document.AddFooters();

        // Define the pages header's picture in a Table. Odd and even pages will have the same headers.
        var oddHeader = document.Headers.Odd;
        var headerFirstTable = oddHeader.InsertTable( 1, 2 );
        headerFirstTable.Design = TableDesign.ColorfulGrid;
        headerFirstTable.AutoFit = AutoFit.Window;
        var upperLeftParagraph = oddHeader.Tables[ 0 ].Rows[ 0 ].Cells[ 0 ].Paragraphs[ 0 ];
        var logo = document.AddImage( MiscellaneousSample.MiscellaneousSampleResourcesDirectory + @"Phone.png" );
        upperLeftParagraph.AppendPicture( logo.CreatePicture( 30, 100 ) );
        upperLeftParagraph.Alignment = Alignment.left;

        // Define the pages header's text in a Table. Odd and even pages will have the same footers.
        var upperRightParagraph = oddHeader.Tables[ 0 ].Rows[ 0 ].Cells[ 1 ].Paragraphs[ 0 ];
        upperRightParagraph.Append( "Toms Telecom Annual report" ).Color( Color.White );
        upperRightParagraph.SpacingBefore( 5d );
        upperRightParagraph.Alignment = Alignment.right;

        // Define the pages footer's picture in a Table.
        var oddFooter = document.Footers.Odd;
        var footerFirstTable = oddFooter.InsertTable( 1, 2 );
        footerFirstTable.Design = TableDesign.ColorfulGrid;
        footerFirstTable.AutoFit = AutoFit.Window;
        var lowerRightParagraph = oddFooter.Tables[ 0 ].Rows[ 0 ].Cells[ 1 ].Paragraphs[ 0 ];
        lowerRightParagraph.AppendPicture( logo.CreatePicture( 30, 100 ) );
        lowerRightParagraph.Alignment = Alignment.right;

        // Define the pages footer's text in a Table
        var lowerLeftParagraph = oddFooter.Tables[ 0 ].Rows[ 0 ].Cells[ 0 ].Paragraphs[ 0 ];
        lowerLeftParagraph.Append( "Toms Telecom 2016" ).Color( Color.White );
        lowerLeftParagraph.SpacingBefore( 5d );

        // Define Data in first page : a Paragraph.
        var paragraph = document.InsertParagraph();
        paragraph.AppendLine( "Toms Telecom Annual report\n2016" ).Bold().FontSize( 35 ).SpacingBefore( 150d );
        paragraph.Alignment = Alignment.center;
        paragraph.InsertPageBreakAfterSelf();

        // Define Data in second page : a Bar Chart.
        document.InsertParagraph("").SpacingAfter( 150d );
        var barChart = new BarChart();
        var sales = CompanyData.CreateSales();
        var salesSeries = new Series( "Sales Per Month" );
        salesSeries.Color = Color.GreenYellow;
        salesSeries.Bind( sales, "Month", "Sales" );
        barChart.AddSeries( salesSeries );        
        document.InsertChart( barChart );
        document.InsertParagraph("Sales were 11% greater in 2016 compared to 2015, with the usual drop during spring time.").SpacingBefore(35d).InsertPageBreakAfterSelf();

        // Define Data in third page : a Line Chart.
        document.InsertParagraph( "" ).SpacingAfter( 150d );
        var lineChart = new LineChart();
        var calls = CompanyData.CreateCallNumber();
        var callSeries = new Series( "Call Number Per Month" );
        callSeries.Bind( calls, "Month", "Calls" );
        lineChart.AddSeries( callSeries );
        document.InsertChart( lineChart );
        document.InsertParagraph( "The number of calls received was much lower in 2016 compared to 2015, by 31%. Winter is still the busiest time of year." ).SpacingBefore( 35d );

       // Save this document to disk.
        document.Save();
        Console.WriteLine( "\tCreated: CompanyReport.docx\n" );
      }
    }


    /// <summary>
    /// Load a document containing a templated invoice with Custom properties, tables, paragraphs and a picture. 
    /// Set the values of those Custom properties, modify the picture, and fill the details table.
    /// </summary>
    public static void CreateInvoice()
    {
      Console.WriteLine( "\tCreateInvoice()" );

      // Load the templated invoice.
      var templateDoc = DocX.Load( MiscellaneousSample.MiscellaneousSampleResourcesDirectory + @"TemplateInvoice.docx" );
      if( templateDoc != null )
      {
        // Create the invoice from the templated invoice.
        var invoice = MiscellaneousSample.CreateInvoiceFromTemplate( templateDoc );

        invoice.SaveAs( MiscellaneousSample.MiscellaneousSampleOutputDirectory + @"CreateInvoice.docx" );
        Console.WriteLine( "\tCreated: CreateInvoice.docx\n" );
      }
    }
    #endregion

    #region Private Methods

    private static DocX CreateInvoiceFromTemplate( DocX templateDoc )
    {
      // Fill in the document custom properties.
      templateDoc.AddCustomProperty( new CustomProperty( "InvoiceNumber", 1355 ) );
      templateDoc.AddCustomProperty( new CustomProperty( "InvoiceDate", new DateTime( 2016, 10, 15 ).ToShortDateString() ) );
      templateDoc.AddCustomProperty( new CustomProperty( "CompanyName", "Toms Telecoms" ) );
      templateDoc.AddCustomProperty( new CustomProperty( "CompanySlogan", "Always with you" ) );
      templateDoc.AddCustomProperty( new CustomProperty( "ClientName", "James Doh" ) );
      templateDoc.AddCustomProperty( new CustomProperty( "ClientStreetName", "123 Main street" ) );
      templateDoc.AddCustomProperty( new CustomProperty( "ClientCity", "Springfield, Ohio" ) );
      templateDoc.AddCustomProperty( new CustomProperty( "ClientZipCode", "54789" ) );
      templateDoc.AddCustomProperty( new CustomProperty( "ClientPhone", "438-585-9636" ) );
      templateDoc.AddCustomProperty( new CustomProperty( "ClientMail", "abc@gmail.com" ) );
      templateDoc.AddCustomProperty( new CustomProperty( "CompanyStreetName", "1458 Thompson Road" ) );
      templateDoc.AddCustomProperty( new CustomProperty( "CompanyCity", "Los Angeles, California" ) );
      templateDoc.AddCustomProperty( new CustomProperty( "CompanyZipCode", "90210" ) );
      templateDoc.AddCustomProperty( new CustomProperty( "CompanyPhone", "1-965-434-5786" ) );
      templateDoc.AddCustomProperty( new CustomProperty( "CompanySupport", "support@tomstelecoms.com" ) );

      // Remove the default logo and add the new one.
      var paragraphWithDefaultLogo = MiscellaneousSample.GetParagraphContainingPicture( templateDoc );
      if( paragraphWithDefaultLogo != null )
      {
        paragraphWithDefaultLogo.Pictures.First().Remove();
        var newLogo = templateDoc.AddImage( MiscellaneousSample.MiscellaneousSampleResourcesDirectory + @"Phone.png" );
        paragraphWithDefaultLogo.AppendPicture( newLogo.CreatePicture( 60, 180 ) );
      }

      // Fill the details table.
      MiscellaneousSample.FillDetailsTable( ref templateDoc );

      return templateDoc;
    }

    private static Paragraph GetParagraphContainingPicture( DocX doc )
    {
      foreach( var p in doc.Paragraphs )
      {
        var picture = p.Pictures.FirstOrDefault();
        if( picture != null )
        {
          return p;
        }
      }

      return null;
    }

    private static void FillDetailsTable( ref DocX templateDoc )
    {
      // The datas that will fill the details table.
      var datas = MiscellaneousSample.GetDetailsData();

      // T he table from the templated invoice.
      var detailsTable = templateDoc.Tables.LastOrDefault();
      if( detailsTable == null )
        return;

      // Remove all rows of the details table, except the header one.
      while( detailsTable.Rows.Count > 1 )
      {
        detailsTable.RemoveRow();
      }

      // Loop through each data rows and use them to add new rows in the detailsTable.
      foreach( DataRow data in datas.Rows )
      {
        var newRow = detailsTable.InsertRow();
        newRow.Cells.First().InsertParagraph( data.ItemArray.First().ToString() );
        newRow.Cells.Last().InsertParagraph( data.ItemArray.Last().ToString() );
      }

      // Calculate the total amount.
      var amountStrings = detailsTable.Rows.Select( r => r.Cells.Last().Paragraphs.Last().Text.Remove(0,1) ).ToList();
      amountStrings.RemoveAt( 0 ); // remove the header
      var totalAmount = amountStrings.Select( s => double.Parse( s ) ).Sum();

      // Add a Total row in the details table.
      var totalRow = detailsTable.InsertRow();
      totalRow.Cells.First().InsertParagraph( "TOTAL:" );
      totalRow.Cells.Last().InsertParagraph( string.Format( "${0}", totalAmount ) );
    }

    private static DataTable GetDetailsData()
    {
      // Create Data to fill the invoice details table.
      var dataTable = new DataTable();
      dataTable.Columns.AddRange( new DataColumn[] { new DataColumn( "Description" ), new DataColumn( "Amount" ) } );

      dataTable.Rows.Add( "Explorer 8698HD Terminal", "$149.95" );
      dataTable.Rows.Add( "MultiSwitch TV Connector", "$24.95" );
      dataTable.Rows.Add( "50 feets cable wires", "$22.49" );
      dataTable.Rows.Add( "Transit A449 Phone Modem", "$59.95" );
      dataTable.Rows.Add( "Toms Wi-Fi router", "$79.95" );
      dataTable.Rows.Add( "Toms Protect2000 Antivirus software", "$39.95" );
      dataTable.Rows.Add( "Installation (3h30)", "$154.49" );

      return dataTable;
    }

    #endregion

    #region Private Classes

    private class CompanyData
    {
      public string Month
      {
        get;
        set;
      }

      public int Sales
      {
        get;
        set;
      }

      public int Calls
      {
        get;
        set;
      }

      internal static List<CompanyData> CreateSales()
      {
        var sales = new List<CompanyData>();
        sales.Add( new CompanyData() { Month = "Jan", Sales = 2500 } );
        sales.Add( new CompanyData() { Month = "Fev", Sales = 3000 } );
        sales.Add( new CompanyData() { Month = "Mar", Sales = 2850 } );
        sales.Add( new CompanyData() { Month = "Apr", Sales = 1050 } );
        sales.Add( new CompanyData() { Month = "May", Sales = 1200 } );
        sales.Add( new CompanyData() { Month = "Jun", Sales = 2900 } );
        sales.Add( new CompanyData() { Month = "Jul", Sales = 3450 } );
        sales.Add( new CompanyData() { Month = "Aug", Sales = 3800 } );
        sales.Add( new CompanyData() { Month = "Sep", Sales = 2900 } );
        sales.Add( new CompanyData() { Month = "Oct", Sales = 2600 } );
        sales.Add( new CompanyData() { Month = "Nov", Sales = 3000 } );
        sales.Add( new CompanyData() { Month = "Dec", Sales = 2500 } );
        return sales;
      }

      internal static List<CompanyData> CreateCallNumber()
      {
        var calls = new List<CompanyData>();
        calls.Add( new CompanyData() { Month = "Jan", Calls = 1200 } );
        calls.Add( new CompanyData() { Month = "Fev", Calls = 1400 } );
        calls.Add( new CompanyData() { Month = "Mar", Calls = 400 } );
        calls.Add( new CompanyData() { Month = "Apr", Calls = 50 } );
        calls.Add( new CompanyData() { Month = "May", Calls = 220 } );
        calls.Add( new CompanyData() { Month = "Jun", Calls = 400 } );
        calls.Add( new CompanyData() { Month = "Jul", Calls = 880 } );
        calls.Add( new CompanyData() { Month = "Aug", Calls = 220 } );
        calls.Add( new CompanyData() { Month = "Sep", Calls = 550 } );
        calls.Add( new CompanyData() { Month = "Oct", Calls = 790 } );
        calls.Add( new CompanyData() { Month = "Nov", Calls = 990 } );
        calls.Add( new CompanyData() { Month = "Dec", Calls = 1300 } );
        return calls;
      }
    }

    #endregion
  }
}
