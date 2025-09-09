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

    public double Budget
    {
      get;
      set;
    }

    public static List<ChartData> CreateCanadaExpenses()
    {
      var canada = new List<ChartData>
      {
        new ChartData() { Category = "Food", Expenses = 100, Budget = 115 },
        new ChartData() { Category = "Housing", Expenses = 120, Budget = 135 },
        new ChartData() { Category = "Transportation", Expenses = 140, Budget = 140 },
        new ChartData() { Category = "Health Care", Expenses = 150, Budget = 165 }
      };
      return canada;
    }

    public static List<ChartData> CreateUSAExpenses()
    {
      var usa = new List<ChartData>
      {
        new ChartData() { Category = "Food", Expenses = 200, Budget = 210 },
        new ChartData() { Category = "Housing", Expenses = 150, Budget = 160 },
        new ChartData() { Category = "Transportation", Expenses = 110, Budget = 125 },
        new ChartData() { Category = "Health Care", Expenses = 100, Budget = 120 }
      };
      return usa;
    }

    public static List<ChartData> CreateBrazilExpenses()
    {
      var brazil = new List<ChartData>
      {
        new ChartData() { Category = "Food", Expenses = 125, Budget = 155 },
        new ChartData() { Category = "Housing", Expenses = 80, Budget = 120 },
        new ChartData() { Category = "Transportation", Expenses = 110, Budget = 200 },
        new ChartData() { Category = "Health Care", Expenses = 60, Budget = 64 }
      };
      return brazil;
    }
  }

  internal class BubbleData
  {
    public double XValue { get; set; }
    public double YValue { get; set; }
    public double BubbleSize { get; set; }

    public static Dictionary<string, List<BubbleData>> GenerateBubbleChartData()
    {
      return new Dictionary<string, List<BubbleData>>()
        {
            { "Fresh Fruits", GenerateFreshFruitsData() },
            { "Fresh Vegetables", GenerateFreshVegetablesData() },
            { "Nuts", GenerateNutsData() },
            { "Tofu", GenerateTofuData() }
        };
    }

    private static List<BubbleData> GenerateFreshFruitsData()
    {
      return new List<BubbleData>
        {
            new BubbleData { XValue = 1.2, YValue = 30.5, BubbleSize = 1200 },
            new BubbleData { XValue = 2.4, YValue = 25.7, BubbleSize = 300 },
            new BubbleData { XValue = 3.6, YValue = 50.8, BubbleSize = 5000 },
            new BubbleData { XValue = 4.5, YValue = 45.0, BubbleSize = 400 },
            new BubbleData { XValue = 5.0, YValue = 60.5, BubbleSize = 750 },
            new BubbleData { XValue = 6.8, YValue = 38.4, BubbleSize = 1800 },
            new BubbleData { XValue = 7.3, YValue = 20.5, BubbleSize = 100 },
            new BubbleData { XValue = 8.1, YValue = 70.0, BubbleSize = 6000 },
            new BubbleData { XValue = 9.4, YValue = 55.2, BubbleSize = 220 },
            new BubbleData { XValue = 10.0, YValue = 48.3, BubbleSize = 900 }
        };
    }

    private static List<BubbleData> GenerateFreshVegetablesData()
    {
      return new List<BubbleData>
        {
            new BubbleData { XValue = 1.1, YValue = 20.0, BubbleSize = 150 },
            new BubbleData { XValue = 2.6, YValue = 35.5, BubbleSize = 850 },
            new BubbleData { XValue = 3.3, YValue = 40.7, BubbleSize = 2300 },
            new BubbleData { XValue = 4.9, YValue = 55.1, BubbleSize = 600 },
            new BubbleData { XValue = 5.7, YValue = 29.9, BubbleSize = 950 },
            new BubbleData { XValue = 6.4, YValue = 63.5, BubbleSize = 1200 },
            new BubbleData { XValue = 7.1, YValue = 25.3, BubbleSize = 400 },
            new BubbleData { XValue = 8.6, YValue = 50.0, BubbleSize = 5300 },
            new BubbleData { XValue = 9.3, YValue = 38.7, BubbleSize = 1600 },
            new BubbleData { XValue = 10.0, YValue = 48.2, BubbleSize = 720 }
        };
    }

    private static List<BubbleData> GenerateNutsData()
    {
      return new List<BubbleData>
        {
            new BubbleData { XValue = 2.0, YValue = 45.5, BubbleSize = 500 },
            new BubbleData { XValue = 3.5, YValue = 32.8, BubbleSize = 1000 },
            new BubbleData { XValue = 4.1, YValue = 60.3, BubbleSize = 2200 },
            new BubbleData { XValue = 5.0, YValue = 75.6, BubbleSize = 3800 },
            new BubbleData { XValue = 6.2, YValue = 55.2, BubbleSize = 450 },
            new BubbleData { XValue = 7.8, YValue = 40.9, BubbleSize = 1200 },
            new BubbleData { XValue = 8.3, YValue = 63.7, BubbleSize = 5400 },
            new BubbleData { XValue = 9.0, YValue = 28.5, BubbleSize = 300 },
            new BubbleData { XValue = 10.4, YValue = 50.2, BubbleSize = 220 },
            new BubbleData { XValue = 11.1, YValue = 65.8, BubbleSize = 1700 }
        };
    }

    private static List<BubbleData> GenerateTofuData()
    {
      return new List<BubbleData>
        {
            new BubbleData { XValue = 1.5, YValue = 28.2, BubbleSize = 140 },
            new BubbleData { XValue = 2.9, YValue = 35.4, BubbleSize = 600 },
            new BubbleData { XValue = 3.7, YValue = 48.7, BubbleSize = 1800 },
            new BubbleData { XValue = 4.8, YValue = 60.1, BubbleSize = 2600 },
            new BubbleData { XValue = 5.6, YValue = 55.0, BubbleSize = 3200 },
            new BubbleData { XValue = 6.9, YValue = 38.6, BubbleSize = 550 },
            new BubbleData { XValue = 7.4, YValue = 48.5, BubbleSize = 1100 },
            new BubbleData { XValue = 8.0, YValue = 70.2, BubbleSize = 5000 },
            new BubbleData { XValue = 9.2, YValue = 40.4, BubbleSize = 850 },
            new BubbleData { XValue = 10.8, YValue = 60.8, BubbleSize = 180 }
        };
    }
  }

  internal class SunburstData
  {
    public string Point { get; set; }
    public double? Size { get; set; }
    public List<SunburstData> Children { get; set; } = new List<SunburstData>();

    public SunburstData( string point, double? size = null )
    {
      Point = point;
      Size = size;
    }

    public static List<SunburstData> CreateSunburstData()
    {
      var sunburstData = new List<SunburstData>
        {
            new SunburstData("Branch 1")
            {
                Children = new List<SunburstData>
                {
                    new SunburstData("Stem 1")
                    {
                        Children = new List<SunburstData>
                        {
                            new SunburstData("Leaf 1", 22),
                            new SunburstData("Leaf 2", 12),
                            new SunburstData("Leaf 3", 18),
                        }
                    },
                    new SunburstData("Stem 2")
                    {
                        Children = new List<SunburstData>
                        {
                            new SunburstData("Leaf 4", 87),
                            new SunburstData("Leaf 5", 88),
                        }
                    },
                    new SunburstData("Leaf 6", 17),
                    new SunburstData("Leaf 7", 14),
                }
            },
            new SunburstData("Branch 2")
            {
                Children = new List<SunburstData>
                {
                    new SunburstData("Stem 3")
                    {
                        Children = new List<SunburstData>
                        {
                            new SunburstData("Leaf 8", 25),
                        }
                    },
                    new SunburstData("Leaf 9", 16),
                    new SunburstData("Stem 4")
                    {
                        Children = new List<SunburstData>
                        {
                            new SunburstData("Leaf 10", 24),
                            new SunburstData("Leaf 11", 89),
                        }
                    },
                }
            },
            new SunburstData("Branch 3")
            {
                Children = new List<SunburstData>
                {
                    new SunburstData("Stem 5")
                    {
                        Children = new List<SunburstData>
                        {
                            new SunburstData("Leaf 12", 16),
                            new SunburstData("Leaf 13", 19),
                        }
                    },
                    new SunburstData("Stem 6")
                    {
                        Children = new List<SunburstData>
                        {
                            new SunburstData("Leaf 14", 86),
                            new SunburstData("Leaf 15", 23),
                        }
                    },
                    new SunburstData("Leaf 16", 21),
                }
            }
        };

      return sunburstData;
    }
  }

  internal class ScatterData
  {
    public double Temperature { get; set; }
    public double Humidity { get; set; }

    public static Dictionary<string, List<ScatterData>> GenerateScatterChartData()
    {
      return new Dictionary<string, List<ScatterData>>()
        {
            { "Desert Climate", GenerateDesertClimateData() },
            { "Tropical Climate", GenerateTropicalClimateData() },
            { "Temperate Climate", GenerateTemperateClimateData() },
            { "Polar Climate", GeneratePolarClimateData() }
        };
    }

    private static List<ScatterData> GenerateDesertClimateData()
    {
      return new List<ScatterData>
        {
            new ScatterData { Temperature = 40.0, Humidity = 10.2 },
            new ScatterData { Temperature = 38.5, Humidity = 9.8 },
            new ScatterData { Temperature = 42.3, Humidity = 7.4 },
            new ScatterData { Temperature = 39.1, Humidity = 11.5 },
            new ScatterData { Temperature = 41.7, Humidity = 8.2 },
            new ScatterData { Temperature = 43.0, Humidity = 6.8 },
            new ScatterData { Temperature = 37.9, Humidity = 12.0 },
            new ScatterData { Temperature = 45.5, Humidity = 5.9 },
            new ScatterData { Temperature = 40.6, Humidity = 10.0 },
            new ScatterData { Temperature = 39.8, Humidity = 9.3 }
        };
    }

    private static List<ScatterData> GenerateTropicalClimateData()
    {
      return new List<ScatterData>
        {
            new ScatterData { Temperature = 30.5, Humidity = 85.0 },
            new ScatterData { Temperature = 31.2, Humidity = 88.1 },
            new ScatterData { Temperature = 29.8, Humidity = 90.5 },
            new ScatterData { Temperature = 32.0, Humidity = 87.4 },
            new ScatterData { Temperature = 30.7, Humidity = 89.3 },
            new ScatterData { Temperature = 28.9, Humidity = 92.2 },
            new ScatterData { Temperature = 31.5, Humidity = 86.0 },
            new ScatterData { Temperature = 29.7, Humidity = 91.6 },
            new ScatterData { Temperature = 30.0, Humidity = 88.9 },
            new ScatterData { Temperature = 32.3, Humidity = 85.7 }
        };
    }

    private static List<ScatterData> GenerateTemperateClimateData()
    {
      return new List<ScatterData>
        {
            new ScatterData { Temperature = 22.0, Humidity = 60.3 },
            new ScatterData { Temperature = 19.5, Humidity = 65.0 },
            new ScatterData { Temperature = 21.3, Humidity = 63.2 },
            new ScatterData { Temperature = 18.8, Humidity = 68.1 },
            new ScatterData { Temperature = 20.4, Humidity = 62.4 },
            new ScatterData { Temperature = 23.0, Humidity = 58.6 },
            new ScatterData { Temperature = 17.9, Humidity = 70.0 },
            new ScatterData { Temperature = 24.1, Humidity = 57.2 },
            new ScatterData { Temperature = 19.7, Humidity = 64.5 },
            new ScatterData { Temperature = 21.9, Humidity = 61.8 }
        };
    }

    private static List<ScatterData> GeneratePolarClimateData()
    {
      return new List<ScatterData>
        {
            new ScatterData { Temperature = -15.2, Humidity = 40.0 },
            new ScatterData { Temperature = -20.5, Humidity = 35.5 },
            new ScatterData { Temperature = -18.3, Humidity = 42.1 },
            new ScatterData { Temperature = -22.7, Humidity = 33.4 },
            new ScatterData { Temperature = -16.0, Humidity = 38.9 },
            new ScatterData { Temperature = -19.8, Humidity = 36.2 },
            new ScatterData { Temperature = -21.4, Humidity = 34.7 },
            new ScatterData { Temperature = -17.5, Humidity = 39.0 },
            new ScatterData { Temperature = -23.1, Humidity = 32.5 },
            new ScatterData { Temperature = -14.8, Humidity = 41.3 }
        };
    }
  }

}
