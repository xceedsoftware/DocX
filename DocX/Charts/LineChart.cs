using System.Xml.Linq;

namespace Novacode
{
    /// <summary>
    /// This element contains the 2-D line chart series.
    /// 21.2.2.97 lineChart (Line Charts)
    /// </summary>
    public class LineChart: Chart
    {
        /// <summary>
        /// Specifies the kind of grouping for a column, line, or area chart.
        /// </summary>
        public Grouping Grouping
        {
            get
            {
                return XElementHelpers.GetValueToEnum<Grouping>(
                    ChartXml.Element(XName.Get("grouping", DocX.c.NamespaceName)));
            }
            set
            {
                XElementHelpers.SetValueFromEnum<Grouping>(
                    ChartXml.Element(XName.Get("grouping", DocX.c.NamespaceName)), value);
            }
        }

        protected override XElement CreateChartXml()
        {
            return XElement.Parse(
                @"<c:lineChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                    <c:grouping val=""standard""/>                    
                  </c:lineChart>");
        }
    }

    /// <summary>
    /// Specifies the kind of grouping for a column, line, or area chart.
    /// 21.2.2.76 grouping (Grouping)
    /// </summary>
    public enum Grouping
    {
        [XmlName("percentStacked")]
        PercentStacked,
        [XmlName("stacked")]
        Stacked,
        [XmlName("standard")]
        Standard
    }
}
