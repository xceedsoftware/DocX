using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Novacode
{
    /// <summary>
    /// This element contains the 2-D pie series for this chart.
    /// 21.2.2.141 pieChart (Pie Charts)
    /// </summary>
    public class PieChart : Chart
    {
        protected override XElement CreateChartXml()
        {
            return XElement.Parse(
                @"<c:pieChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                  </c:pieChart>");
        }
    }
}
