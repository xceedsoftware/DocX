/***************************************************************************************

   DocX – DocX is the community edition of Xceed Words for .NET

   Copyright (C) 2009-2017 Xceed Software Inc.

   This program is provided to you under the terms of the Microsoft Public
   License (Ms-PL) as published at http://wpftoolkit.codeplex.com/license 

   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/

  *************************************************************************************/
using System.Collections.Generic;

namespace Xceed.Words.NET.Examples
{
  internal class ChartData
  {
    public string Category
    {
      get;
      set;
    }
    public double Expenses
    {
      get;
      set;
    }

    public static List<ChartData> CreateCanadaExpenses()
    {
      var canada = new List<ChartData>();
      canada.Add( new ChartData() { Category = "Food", Expenses = 100 } );
      canada.Add( new ChartData() { Category = "Housing", Expenses = 120 } );
      canada.Add( new ChartData() { Category = "Transportation", Expenses = 140 } );
      canada.Add( new ChartData() { Category = "Health Care", Expenses = 150 } );
      return canada;
    }

    public static List<ChartData> CreateUSAExpenses()
    {
      var usa = new List<ChartData>();
      usa.Add( new ChartData() { Category = "Food", Expenses = 200 } );
      usa.Add( new ChartData() { Category = "Housing", Expenses = 150 } );
      usa.Add( new ChartData() { Category = "Transportation", Expenses = 110 } );
      usa.Add( new ChartData() { Category = "Health Care", Expenses = 100 } );
      return usa;
    }

    public static List<ChartData> CreateBrazilExpenses()
    {
      var brazil = new List<ChartData>();
      brazil.Add( new ChartData() { Category = "Food", Expenses = 125 } );
      brazil.Add( new ChartData() { Category = "Housing", Expenses = 80 } );
      brazil.Add( new ChartData() { Category = "Transportation", Expenses = 110 } );
      brazil.Add( new ChartData() { Category = "Health Care", Expenses = 60 } );
      return brazil;
    }
  }
}
