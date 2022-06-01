/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Sample Application
Copyright (c) 2009-2022 - Xceed Software Inc.

This application demonstrates how to use the different features when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading;
using Xceed.Words.NET.Example;

namespace Xceed.Words.NET.Examples
{
  public class Program
  {
#if NETCORE || NET5
    internal const string SampleDirectory = @"..\..\..\Samples\";
#else
    internal const string SampleDirectory = @"..\..\Samples\";
#endif

    private static void Main( string[] args )
    {

      var version = Assembly.GetExecutingAssembly().GetName().Version;
      var versionNumber = version.Major + "." + version.Minor;
      Console.WriteLine( "\nRunning Examples of Xceed Words for .NET version " + versionNumber + ".\n" );

      //Paragraphs
      ParagraphSample.SimpleFormattedParagraphs();
      ParagraphSample.StyleParagraphs();
      ParagraphSample.ForceParagraphOnSinglePage();
      ParagraphSample.ForceMultiParagraphsOnSinglePage();
      ParagraphSample.TextActions();
      ParagraphSample.Heading();
      ParagraphSample.AddObjectsFromOtherDocument();
      ParagraphSample.AddHtml();
      ParagraphSample.AddRtf();

      //Document
      DocumentSample.AddCustomProperties();
      DocumentSample.ReplaceTextWithText();
      DocumentSample.ReplaceTextWithObjects();
      DocumentSample.ApplyTemplate();
      DocumentSample.AppendDocument();
      DocumentSample.LoadDocumentWithFilename();
      DocumentSample.LoadDocumentWithStream();
      DocumentSample.LoadDocumentWithStringUrl();
      DocumentSample.AddHtmlFromFile();
      DocumentSample.AddRtfFromFile();
      DocumentSample.InsertDocument();

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
      TableSample.CloneTable();
      TableSample.AddTableWithTextWrapping();
      TableSample.TextDirectionTable();
      TableSample.CreateRowsFromTemplate();
      TableSample.ColumnsWidth();
      TableSample.MergeCells();
      TableSample.ShadingPattern();

      //Hyperlink
      HyperlinkSample.Hyperlinks();

      //Section
      SectionSample.InsertSections();
      SectionSample.SetPageOrientations();

      //Lists
      ListSample.AddList();
      ListSample.AddCustomNumberedList();
      ListSample.AddCustomBulletedList();
      ListSample.AddChapterList();
      ListSample.CloneLists();
      ListSample.ModifyList();

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
      ChartSample.ModifyChartData();
      ChartSample.AddChartWithTextWrapping();

      //Tale of Content
      TableOfContentSample.InsertTableOfContent();
      TableOfContentSample.InsertTableOfContentWithReference();
      TableOfContentSample.UpdateTableOfContent();

      //Lines
      LineSample.InsertHorizontalLine();

      //Protection
      ProtectionSample.AddPasswordProtection();
      ProtectionSample.AddProtection();
      ProtectionSample.ChangePasswordProtection();

      //Parallel  
      ParallelSample.DoParallelActions();

      //Others
      MiscellaneousSample.CreateRecipe();
      MiscellaneousSample.CompanyReport();
      MiscellaneousSample.CreateInvoice();
      MiscellaneousSample.MailMerge();

      //PDF  
      PdfSample.ConvertToPDFWithUninstalledFont();
      PdfSample.ConvertToPDF();

      //Shape
      ShapeSample.AddShape();
      ShapeSample.AddShapeWithTextWrapping();
      ShapeSample.AddTextBox();
      ShapeSample.AddTextBoxWithTextWrapping();

      //CheckBox
      CheckBoxSample.ModifyCheckBox();
      CheckBoxSample.AddCheckBox();

      //Hyphenation
      HyphenationSample.CreateHyphenation();
      HyphenationSample.UpdateHyphenation();

      //Footnotes Endnotes
      FootnoteEndnoteSample.AddFootnotes();
      FootnoteEndnoteSample.AddCustomFootnotes();
      FootnoteEndnoteSample.AddEndnotes();

      //Digital Signature
      DigitalSignatureSample.SignWithSignatureLine();
      DigitalSignatureSample.SignWithoutSignatureLine();
      DigitalSignatureSample.VerifySignatures();
      DigitalSignatureSample.RemoveSignatures();
      DigitalSignatureSample.RemoveSignatureLines();

      Console.WriteLine( "\nDone running Examples of Xceed Words for .NET version " + versionNumber + ".\n" );
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
