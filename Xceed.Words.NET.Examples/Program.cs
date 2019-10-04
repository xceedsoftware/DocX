/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Sample Application
Copyright (c) 2009-2018 - Xceed Software Inc.

This application demonstrates how to use the different features when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/
using System;
using System.Collections.Generic;
using System.Threading;

namespace Xceed.Words.NET.Examples
{
  public class Program
  {
    internal const string SampleDirectory = @"..\..\Samples\";

    private static void Main( string[] args )
    {

      Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo( "en-US" );

      //Paragraphs      
      ParagraphSample.SimpleFormattedParagraphs();
      ParagraphSample.ForceParagraphOnSinglePage();
      ParagraphSample.ForceMultiParagraphsOnSinglePage();
      ParagraphSample.TextActions();
      ParagraphSample.Heading();

      //Document
      DocumentSample.AddCustomProperties();
      DocumentSample.ReplaceText();
      DocumentSample.ApplyTemplate();
      DocumentSample.AppendDocument();

      //Images
      ImageSample.AddPicture();
      ImageSample.AddPictureWithTextWrapping();
      ImageSample.CopyPicture();
      ImageSample.ModifyImage();

      // Indentation / Direction / Margins
      MarginSample.SetDirection();
      MarginSample.Indentation();
      MarginSample.Margins();

      //Header/Footers
      HeaderFooterSample.HeadersFooters();

      //Tables
      TableSample.InsertRowAndImageTable();
      TableSample.TextDirectionTable();
      TableSample.CreateRowsFromTemplate();
      TableSample.ColumnsWidth();
      TableSample.MergeCells();

      //Hyperlink
      HyperlinkSample.Hyperlinks();

      //Section
      SectionSample.InsertSections();

      //Lists
      ListSample.AddList();

      //Equations
      EquationSample.InsertEquation();

      //Bookmarks
      BookmarkSample.InsertBookmarks();
      BookmarkSample.ReplaceText();

      //Charts
      ChartSample.BarChart();
      ChartSample.LineChart();
      ChartSample.PieChart();
      ChartSample.Chart3D();

      //Tale of Content
      TableOfContentSample.InsertTableOfContent();
      TableOfContentSample.InsertTableOfContentWithReference();

      //Lines
      LineSample.InsertHorizontalLine();

      //Protection
      ProtectionSample.AddPasswordProtection();
      ProtectionSample.AddProtection();

      //Parallel  
      ParallelSample.DoParallelActions();

      //Others
      MiscellaneousSample.CreateRecipe();
      MiscellaneousSample.CompanyReport();
      MiscellaneousSample.CreateInvoice();

      //PDF  
      PdfSample.ConvertToPDF();

      Console.WriteLine( "\nPress any key to exit." );
      Console.ReadKey();
    }

    #region Charts

    private class ChartData
    {
      public String Mounth
      {
        get; set;
      }
      public Double Money
      {
        get; set;
      }

      public static List<ChartData> CreateCompanyList1()
      {
        List<ChartData> company1 = new List<ChartData>();
        company1.Add( new ChartData() { Mounth = "January", Money = 100 } );
        company1.Add( new ChartData() { Mounth = "February", Money = 120 } );
        company1.Add( new ChartData() { Mounth = "March", Money = 140 } );
        return company1;
      }

      public static List<ChartData> CreateCompanyList2()
      {
        List<ChartData> company2 = new List<ChartData>();
        company2.Add( new ChartData() { Mounth = "January", Money = 80 } );
        company2.Add( new ChartData() { Mounth = "February", Money = 160 } );
        company2.Add( new ChartData() { Mounth = "March", Money = 130 } );
        return company2;
      }
    }

    #endregion
  }
}
