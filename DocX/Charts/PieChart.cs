using System;
using System.Xml.Linq;

namespace Novacode
{
    /// <summary>
    /// This element contains the 2-D pie series for this chart.
    /// 21.2.2.141 pieChart (Pie Charts)
    /// </summary>
    public class PieChart : Chart
    {
        public override Boolean IsAxisExist { get { return false; } }
        public override Int16 MaxSeriesCount { get { return 1; } }

        protected override XElement CreateChartXml()
        {
            return XElement.Parse(
                @"<c:pieChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                  </c:pieChart>");
        }
    }
}
