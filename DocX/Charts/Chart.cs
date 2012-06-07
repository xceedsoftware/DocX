using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Collections;
using System.Drawing;

namespace Novacode
{
    /// <summary>
    /// Represents every Chart in this document.
    /// </summary>
    public abstract class Chart
    {
        protected XElement ChartXml { get; private set; }
        protected XElement ChartRootXml { get; private set; }

        /// <summary>
        /// The xml representation of this chart
        /// </summary>
        public XDocument Xml { get; private set; }

        #region Series

        /// <summary>
        /// Chart's series
        /// </summary>
        public List<Series> Series
        {
            get
            {
                List<Series> series = new List<Series>();
                foreach (XElement element in ChartXml.Elements(XName.Get("ser", DocX.c.NamespaceName)))
                {
                    series.Add(new Series(element));
                }
                return series;
            }
        }

        /// <summary>
        /// Return maximum count of series
        /// </summary>
        public virtual Int16 MaxSeriesCount { get { return Int16.MaxValue; } }

        /// <summary>
        /// Add a new series to this chart
        /// </summary>
        public void AddSeries(Series series)
        {
            if (ChartXml.Elements(XName.Get("ser", DocX.c.NamespaceName)).Count() == MaxSeriesCount)
                throw new InvalidOperationException("Maximum series for this chart is" + MaxSeriesCount.ToString() + "and have exceeded!");
            ChartXml.Add(series.Xml);
        }

        #endregion

        #region Legend

        /// <summary>
        /// Chart's legend.
        /// If legend doesn't exist property is null.
        /// </summary>
        public ChartLegend Legend { get; private set; }

        /// <summary>
        /// Add standart legend to the chart.
        /// </summary>
        public void AddLegend()
        {
            AddLegend(ChartLegendPosition.Right, false);
        }

        /// <summary>
        /// Add a legend with parameters to the chart.
        /// </summary>
        public void AddLegend(ChartLegendPosition position, Boolean overlay)
        {
            if (Legend != null)
                RemoveLegend();
            Legend = new ChartLegend(position, overlay);
            ChartRootXml.Add(Legend.Xml);
        }

        /// <summary>
        /// Remove the legend from the chart.
        /// </summary>
        public void RemoveLegend()
        {
            Legend.Xml.Remove();
            Legend = null;
        }

        #endregion

        #region Axis

        /// <summary>
        /// Represents the category axis
        /// </summary>
        public CategoryAxis CategoryAxis { get; private set; }

        /// <summary>
        /// Represents the values axis
        /// </summary>
        public ValueAxis ValueAxis { get; private set; }

        /// <summary>
        /// Represents existing the axis
        /// </summary>
        public virtual Boolean IsAxisExist { get { return true; } }

        #endregion

        /// <summary>
        /// Get or set 3D view for this chart
        /// </summary>
        public Boolean View3D
        {
            get
            {
                return ChartXml.Name.LocalName.Contains("3D");
            }
            set
            {
                if (value)
                {
                    if (!View3D)
                    {
                        String currentName = ChartXml.Name.LocalName;
                        ChartXml.Name = XName.Get(currentName.Replace("Chart", "3DChart"), DocX.c.NamespaceName);
                    }
                }
                else
                {
                    if (View3D)
                    {
                        String currentName = ChartXml.Name.LocalName;
                        ChartXml.Name = XName.Get(currentName.Replace("3DChart", "Chart"), DocX.c.NamespaceName);
                    }
                }
            }
        }

        /// <summary>
        /// Specifies how blank cells shall be plotted on a chart
        /// </summary>
        public DisplayBlanksAs DisplayBlanksAs
        {
            get
            {
                return XElementHelpers.GetValueToEnum<DisplayBlanksAs>(
                    ChartRootXml.Element(XName.Get("dispBlanksAs", DocX.c.NamespaceName)));
            }
            set
            {
                XElementHelpers.SetValueFromEnum<DisplayBlanksAs>(
                    ChartRootXml.Element(XName.Get("dispBlanksAs", DocX.c.NamespaceName)), value);
            }
        }

        /// <summary>
        /// Create an Chart for this document
        /// </summary>        
        public Chart()
        {
            // Create global xml
            Xml = XDocument.Parse
                (@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                   <c:chartSpace xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">  
                       <c:roundedCorners val=""0""/>
                       <c:chart>
                           <c:autoTitleDeleted val=""0""/>
                           <c:plotVisOnly val=""1""/>
                           <c:dispBlanksAs val=""gap""/>
                           <c:showDLblsOverMax val=""0""/>
                       </c:chart>
                   </c:chartSpace>");

            // Create a real chart xml in an inheritor
            ChartXml = CreateChartXml();

            // Create result plotarea element
            XElement plotAreaXml = new XElement(
                XName.Get("plotArea", DocX.c.NamespaceName),
                new XElement(XName.Get("layout", DocX.c.NamespaceName)),
                ChartXml);

            // Set labels 
            XElement dLblsXml = XElement.Parse(
                @"<c:dLbls xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                    <c:showLegendKey val=""0""/>
                    <c:showVal val=""0""/>
                    <c:showCatName val=""0""/>
                    <c:showSerName val=""0""/>
                    <c:showPercent val=""0""/>
                    <c:showBubbleSize val=""0""/>
                    <c:showLeaderLines val=""1""/>
                </c:dLbls>");
            ChartXml.Add(dLblsXml);

            // if axes exists, create their
            if (IsAxisExist)
            {
                CategoryAxis = new CategoryAxis("148921728");
                ValueAxis = new ValueAxis("154227840");

                XElement axIDcatXml = XElement.Parse(String.Format(
                    @"<c:axId val=""{0}"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""/>", CategoryAxis.Id));
                XElement axIDvalXml = XElement.Parse(String.Format(
                    @"<c:axId val=""{0}"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""/>", ValueAxis.Id));

                ChartXml.Add(axIDcatXml);
                ChartXml.Add(axIDvalXml);

                plotAreaXml.Add(CategoryAxis.Xml);
                plotAreaXml.Add(ValueAxis.Xml);
            }

            ChartRootXml = Xml.Root.Element(XName.Get("chart", DocX.c.NamespaceName));
            ChartRootXml.Add(plotAreaXml);
        }

        /// <summary>
        /// An abstract method which creates the current chart xml
        /// </summary>
        protected abstract XElement CreateChartXml();
    }



    /// <summary>
    /// Represents a chart series
    /// </summary>
    public class Series
    {
        private XElement strCache;
        private XElement numCache;

        /// <summary>
        /// Series xml element
        /// </summary>
        internal XElement Xml { get; private set; }

        public Color Color
        {
            get
            {
                XElement colorElement = Xml.Element(XName.Get("spPr", DocX.c.NamespaceName));
                if (colorElement == null)
                    return Color.Transparent;
                else
                    return Color.FromArgb(Int32.Parse(
                        colorElement.Element(XName.Get("solidFill", DocX.a.NamespaceName)).Element(XName.Get("srgbClr", DocX.a.NamespaceName)).Attribute(XName.Get("val")).Value,
                        System.Globalization.NumberStyles.HexNumber));
            }
            set
            {
                XElement colorElement = Xml.Element(XName.Get("spPr", DocX.c.NamespaceName));
                if (colorElement != null)
                    colorElement.Remove();
                colorElement = new XElement(
                    XName.Get("spPr", DocX.c.NamespaceName),
                    new XElement(
                        XName.Get("solidFill", DocX.a.NamespaceName),
                        new XElement(
                            XName.Get("srgbClr", DocX.a.NamespaceName),
                            new XAttribute(XName.Get("val"), value.ToHex()))));
                Xml.Add(colorElement);
            }
        }

        internal Series(XElement xml)
        {
            Xml = xml;
            strCache = xml.Element(XName.Get("cat", DocX.c.NamespaceName)).Element(XName.Get("strRef", DocX.c.NamespaceName)).Element(XName.Get("strCache", DocX.c.NamespaceName));
            numCache = xml.Element(XName.Get("val", DocX.c.NamespaceName)).Element(XName.Get("numRef", DocX.c.NamespaceName)).Element(XName.Get("numCache", DocX.c.NamespaceName));
        }

        public Series(String name)
        {
            strCache = new XElement(XName.Get("strCache", DocX.c.NamespaceName));
            numCache = new XElement(XName.Get("numCache", DocX.c.NamespaceName));

            Xml = new XElement(
                XName.Get("ser", DocX.c.NamespaceName),
                new XElement(
                    XName.Get("tx", DocX.c.NamespaceName),
                    new XElement(
                        XName.Get("strRef", DocX.c.NamespaceName),
                        new XElement(
                            XName.Get("strCache", DocX.c.NamespaceName),
                            new XElement(
                                XName.Get("pt", DocX.c.NamespaceName),
                                new XAttribute(XName.Get("idx"), "0"),
                                new XElement(XName.Get("v", DocX.c.NamespaceName), name)
                                )))),
                new XElement(XName.Get("invertIfNegative", DocX.c.NamespaceName), "0"),
                    new XElement(
                        XName.Get("cat", DocX.c.NamespaceName),
                        new XElement(XName.Get("strRef", DocX.c.NamespaceName), strCache)),
                new XElement(
                        XName.Get("val", DocX.c.NamespaceName),
                        new XElement(XName.Get("numRef", DocX.c.NamespaceName), numCache))
                );
        }

        public void Bind(ICollection list, String categoryPropertyName, String valuePropertyName)
        {
            XElement ptCount = new XElement(XName.Get("ptCount", DocX.c.NamespaceName), new XAttribute(XName.Get("val"), list.Count));
            XElement formatCode = new XElement(XName.Get("formatCode", DocX.c.NamespaceName), "General");

            strCache.RemoveAll();
            numCache.RemoveAll();

            strCache.Add(ptCount);
            numCache.Add(ptCount);
            numCache.Add(formatCode);

            Int32 index = 0;
            XElement pt;
            foreach (var item in list)
            {
                pt = new XElement(XName.Get("pt", DocX.c.NamespaceName), new XAttribute(XName.Get("idx"), index),
                    new XElement(XName.Get("v", DocX.c.NamespaceName), item.GetType().GetProperty(categoryPropertyName).GetValue(item, null)));
                strCache.Add(pt);
                pt = new XElement(XName.Get("pt", DocX.c.NamespaceName), new XAttribute(XName.Get("idx"), index),
                    new XElement(XName.Get("v", DocX.c.NamespaceName), item.GetType().GetProperty(valuePropertyName).GetValue(item, null)));
                numCache.Add(pt);
                index++;
            }
        }

        public void Bind(IList categories, IList values)
        {
            if (categories.Count != values.Count)
                throw new ArgumentException("Categories count must equal to Values count");

            XElement ptCount = new XElement(XName.Get("ptCount", DocX.c.NamespaceName), new XAttribute(XName.Get("val"), categories.Count));
            XElement formatCode = new XElement(XName.Get("formatCode", DocX.c.NamespaceName), "General");

            strCache.RemoveAll();
            numCache.RemoveAll();

            strCache.Add(ptCount);
            numCache.Add(ptCount);
            numCache.Add(formatCode);

            XElement pt;
            for (int index = 0; index < categories.Count; index++)
            {
                pt = new XElement(XName.Get("pt", DocX.c.NamespaceName), new XAttribute(XName.Get("idx"), index),
                    new XElement(XName.Get("v", DocX.c.NamespaceName), categories[index].ToString()));
                strCache.Add(pt);
                pt = new XElement(XName.Get("pt", DocX.c.NamespaceName), new XAttribute(XName.Get("idx"), index),
                    new XElement(XName.Get("v", DocX.c.NamespaceName), values[index].ToString()));
                numCache.Add(pt);
            }
        }
    }

    /// <summary>
    /// Represents a chart legend
    /// More: http://msdn.microsoft.com/ru-ru/library/cc845123.aspx
    /// </summary>
    public class ChartLegend
    {
        /// <summary>
        /// Legend xml element
        /// </summary>
        internal XElement Xml { get; private set; }

        /// <summary>
        /// Specifies that other chart elements shall be allowed to overlap this chart element
        /// </summary>
        public Boolean Overlay
        {
            get { return Xml.Element(XName.Get("overlay", DocX.c.NamespaceName)).Attribute("val").Value == "1"; }
            set { Xml.Element(XName.Get("overlay", DocX.c.NamespaceName)).Attribute("val").Value = GetOverlayValue(value); }
        }

        /// <summary>
        /// Specifies the possible positions for a legend
        /// </summary>
        public ChartLegendPosition Position
        {
            get
            {
                return XElementHelpers.GetValueToEnum<ChartLegendPosition>(
                    Xml.Element(XName.Get("legendPos", DocX.c.NamespaceName)));
            }
            set
            {
                XElementHelpers.SetValueFromEnum<ChartLegendPosition>(
                    Xml.Element(XName.Get("legendPos", DocX.c.NamespaceName)), value);
            }
        }

        internal ChartLegend(ChartLegendPosition position, Boolean overlay)
        {
            Xml = new XElement(
                XName.Get("legend", DocX.c.NamespaceName),
                new XElement(XName.Get("legendPos", DocX.c.NamespaceName), new XAttribute("val", XElementHelpers.GetXmlNameFromEnum<ChartLegendPosition>(position))),
                new XElement(XName.Get("overlay", DocX.c.NamespaceName), new XAttribute("val", GetOverlayValue(overlay)))
                );
        }

        /// <summary>
        /// ECMA-376, page 3840
        /// 21.2.2.132 overlay (Overlay)
        /// </summary>
        private String GetOverlayValue(Boolean overlay)
        {
            if (overlay)
                return "1";
            else
                return "0";
        }
    }

    /// <summary>
    /// Specifies the possible positions for a legend.
    /// 21.2.3.24 ST_LegendPos (Legend Position)
    /// </summary>
    public enum ChartLegendPosition
    {
        [XmlName("t")]
        Top,
        [XmlName("b")]
        Bottom,
        [XmlName("l")]
        Left,
        [XmlName("r")]
        Right,
        [XmlName("tr")]
        TopRight
    }

    /// <summary>
    /// Specifies the possible ways to display blanks.
    /// 21.2.3.10 ST_DispBlanksAs (Display Blanks As)
    /// </summary>
    public enum DisplayBlanksAs
    {
        [XmlName("gap")]
        Gap,
        [XmlName("span")]
        Span,
        [XmlName("zero")]
        Zero
    }
}
