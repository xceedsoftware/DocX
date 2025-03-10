/***************************************************************************************
 
   DocX – DocX is the community edition of Xceed Words for .NET
 
   Copyright (C) 2009-2025 Xceed Software Inc.
 
   This program is provided to you under the terms of the XCEED SOFTWARE, INC.
   COMMUNITY LICENSE AGREEMENT (for non-commercial use) as published at 
   https://github.com/xceedsoftware/DocX/blob/master/license.md
 
   For more features and fast professional support,
   pick up Xceed Words for .NET at https://xceed.com/xceed-words-for-net/
 
  *************************************************************************************/


/***************************************************************************************
Xceed Words for .NET – Xceed.Words.NET.Examples – Chart Sample Application
Copyright (c) 2009-2025 - Xceed Software Inc.

This application demonstrates how to create a chart when using the API 
from the Xceed Words for .NET.

This file is part of Xceed Words for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
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
