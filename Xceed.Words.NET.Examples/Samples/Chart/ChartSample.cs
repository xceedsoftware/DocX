/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Chart Sample Application
Copyright (c) 2009-2020 - Xceed Software Inc.

This application demonstrates how to create a chart when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Xceed.Document.NET;

namespace Xceed.Words.NET.Examples
{
  public class ChartSample
  {
    #region Private Members

    private const string ChartSampleOutputDirectory = Program.SampleDirectory + @"Chart\Output\";
    private const string ChartSampleResourceDirectory = Program.SampleDirectory + @"Chart\Resources\";

    #endregion

    #region Constructors

    static ChartSample()
    {
      if( !Directory.Exists( ChartSample.ChartSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( ChartSample.ChartSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Add a Bar chart to a document.
    /// </summary>
    public static void BarChart()
    {
      Console.WriteLine( "\tBarChart()" );

      // Creates a document
      using( var document = DocX.Create( ChartSample.ChartSampleOutputDirectory + @"BarChart.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Bar Chart" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a bar chart.
        var c = new BarChart();
        c.AddLegend( ChartLegendPosition.Left, false );
        c.BarDirection = BarDirection.Bar;
        c.BarGrouping = BarGrouping.Standard;
        c.GapWidth = 200;

        // Create the data.
        var canada = ChartData.CreateCanadaExpenses();
        var usa = ChartData.CreateUSAExpenses();
        var brazil = ChartData.CreateBrazilExpenses();

        // Create and add series
        var s1 = new Series( "Brazil" );
        s1.Color = Color.GreenYellow;
        s1.Bind( brazil, "Category", "Expenses" );
        c.AddSeries( s1 );

        var s2 = new Series( "USA" );
        s2.Color = Color.LightBlue;
        s2.Bind( usa, "Category", "Expenses" );
        c.AddSeries( s2 );

        var s3 = new Series( "Canada" );
        s3.Color = Color.Gray;
        s3.Bind( canada, "Category", "Expenses" );
        c.AddSeries( s3 );

        // Insert the chart into the document.
        document.InsertParagraph( "Expenses(M$) for selected categories per country" ).FontSize( 15 ).SpacingAfter( 10d );
        document.InsertChart( c, 350f, 550f );

        document.Save();
        Console.WriteLine( "\tCreated: BarChart.docx\n" );
      }
    }

    /// <summary>
    /// Add a Line chart to a document.
    /// </summary>
    public static void LineChart()
    {
      Console.WriteLine( "\tLineChartt()" );

      // Creates a document
      using( var document = DocX.Create( ChartSample.ChartSampleOutputDirectory + @"LineChart.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Line Chart" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a line chart.
        var c = new LineChart();
        c.AddLegend( ChartLegendPosition.Left, false );

        // Create the data.
        var canada = ChartData.CreateCanadaExpenses();
        var usa = ChartData.CreateUSAExpenses();
        var brazil = ChartData.CreateBrazilExpenses();

        // Create and add series
        var s1 = new Series( "Brazil" );
        s1.Color = Color.Yellow;
        s1.Bind( brazil, "Category", "Expenses" );
        c.AddSeries( s1 );

        var s2 = new Series( "USA" );
        s2.Color = Color.Blue;
        s2.Bind( usa, "Category", "Expenses" );
        c.AddSeries( s2 );

        var s3 = new Series( "Canada" );
        s3.Color = Color.Red;
        s3.Bind( canada, "Category", "Expenses" );
        c.AddSeries( s3 );

        // Insert chart into document
        document.InsertParagraph( "Expenses(M$) for selected categories per country" ).FontSize( 15 ).SpacingAfter( 10d );
        document.InsertChart( c );

        document.Save();
        Console.WriteLine( "\tCreated: LineChart.docx\n" );
      }
    }

    /// <summary>
    /// Add a Pie chart to a document.
    /// </summary>
    public static void PieChart()
    {
      Console.WriteLine( "\tPieChart()" );

      // Creates a document
      using( var document = DocX.Create( ChartSample.ChartSampleOutputDirectory + @"PieChart.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "Pie Chart" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a pie chart.
        var c = new PieChart();
        c.AddLegend( ChartLegendPosition.Left, false );

        // Create the data.
        var brazil = ChartData.CreateBrazilExpenses();

        // Create and add series
        var s1 = new Series( "Brazil" );
        s1.Bind( brazil, "Category", "Expenses" );
        c.AddSeries( s1 );

        // Insert chart into document
        document.InsertParagraph( "Expenses(M$) for selected categories in Brazil" ).FontSize( 15 ).SpacingAfter( 10d );
        document.InsertChart( c );

        document.Save();
        Console.WriteLine( "\tCreated: PieChart.docx\n" );
      }
    }

    /// <summary>
    /// Add a 3D bar chart to a document.
    /// </summary>
    /// 
    public static void Chart3D()
    {
      Console.WriteLine( "\tChart3D()" );

      // Creates a document
      using( var document = DocX.Create( ChartSample.ChartSampleOutputDirectory + @"3DChart.docx" ) )
      {
        // Add a title
        document.InsertParagraph( "3D Chart" ).FontSize( 15d ).SpacingAfter( 50d ).Alignment = Alignment.center;

        // Create a 3D Bar chart.
        var c = new BarChart();
        c.View3D = true;

        // Create the data.
        var brazil = ChartData.CreateBrazilExpenses();

        // Create and add series
        var s1 = new Series( "Brazil" );
        s1.Color = Color.GreenYellow;
        s1.Bind( brazil, "Category", "Expenses" );
        c.AddSeries( s1 );

        // Insert chart into document
        document.InsertParagraph( "Expenses(M$) for selected categories in Brazil" ).FontSize( 15 ).SpacingAfter( 10d );
        document.InsertChart( c );

        document.Save();
        Console.WriteLine( "\tCreated: 3DChart.docx\n" );
      }
    }

    public static void ModifyChartData()
    {








      // This option is available when you buy Xceed Words for .NET from https://xceed.com/xceed-words-for-net/.
    }
    #endregion
  }
}
