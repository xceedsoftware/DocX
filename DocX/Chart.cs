using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using System.IO.Packaging;

namespace Novacode
{
    /// <summary>
    /// Represents a Chart in this document.
    /// </summary>
    public class Chart
    {
        private XElement LegendXml;
        private XElement ChartXml;
        private XElement ChartRootXml;

        public XDocument Xml { get; private set; }

        public List<XElement> Series
        {
            get
            {
                List<XElement> series = new List<XElement>();
                GetSeriesRecursive(Xml.Root, series);
                return series;
            }
        }

        private void GetSeriesRecursive(XElement element, List<XElement> series)
        {
            if (element.Name.LocalName == "ser")
            {
                series.Add(element);
            }
            else
            {
                if (element.HasElements)
                    foreach (XElement e in element.Elements())
                        GetSeriesRecursive(e, series);
            }
        }

        private ChartLegend legend;
        public ChartLegend Legend { get { return legend; } }

        /// <summary>
        /// Create an Chart for this document
        /// </summary>        
        public Chart()
        {
            #region Create Bar Chart
            ChartXml = XElement.Parse(
   @"<c:barChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
        <c:barDir val=""col""/>
        <c:grouping val=""clustered""/>
        <c:ser>
          <c:idx val=""0""/>
          <c:order val=""0""/>
          <c:tx>
            <c:strRef>
              <c:f>Лист1!$B$1</c:f>
              <c:strCache>
                <c:ptCount val=""1""/>
                <c:pt idx=""0"">
                  <c:v>Ряд 1</c:v>
                </c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:invertIfNegative val=""0""/>
          <c:cat>
            <c:strRef>
              <c:f>Лист1!$A$2:$A$5</c:f>
              <c:strCache>
                <c:ptCount val=""4""/>
                <c:pt idx=""0"">
                  <c:v>Категория 1</c:v>
                </c:pt>
                <c:pt idx=""1"">
                  <c:v>Категория 2</c:v>
                </c:pt>
                <c:pt idx=""2"">
                  <c:v>Категория 3</c:v>
                </c:pt>
                <c:pt idx=""3"">
                  <c:v>Категория 4</c:v>
                </c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Лист1!$B$2:$B$5</c:f>
              <c:numCache>
                <c:formatCode>Основной</c:formatCode>
                <c:ptCount val=""4""/>
                <c:pt idx=""0"">
                  <c:v>4.3</c:v>
                </c:pt>
                <c:pt idx=""1"">
                  <c:v>2.5</c:v>
                </c:pt>
                <c:pt idx=""2"">
                  <c:v>3.5</c:v>
                </c:pt>
                <c:pt idx=""3"">
                  <c:v>4.5</c:v>
                </c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:ser>
          <c:idx val=""1""/>
          <c:order val=""1""/>
          <c:tx>
            <c:strRef>
              <c:f>Лист1!$C$1</c:f>
              <c:strCache>
                <c:ptCount val=""1""/>
                <c:pt idx=""0"">
                  <c:v>Ряд 2</c:v>
                </c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:invertIfNegative val=""0""/>
          <c:cat>
            <c:strRef>
              <c:f>Лист1!$A$2:$A$5</c:f>
              <c:strCache>
                <c:ptCount val=""4""/>
                <c:pt idx=""0"">
                  <c:v>Категория 1</c:v>
                </c:pt>
                <c:pt idx=""1"">
                  <c:v>Категория 2</c:v>
                </c:pt>
                <c:pt idx=""2"">
                  <c:v>Категория 3</c:v>
                </c:pt>
                <c:pt idx=""3"">
                  <c:v>Категория 4</c:v>
                </c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Лист1!$C$2:$C$5</c:f>
              <c:numCache>
                <c:formatCode>Основной</c:formatCode>
                <c:ptCount val=""4""/>
                <c:pt idx=""0"">
                  <c:v>2.4</c:v>
                </c:pt>
                <c:pt idx=""1"">
                  <c:v>4.4000000000000004</c:v>
                </c:pt>
                <c:pt idx=""2"">
                  <c:v>1.8</c:v>
                </c:pt>
                <c:pt idx=""3"">
                  <c:v>2.8</c:v>
                </c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:ser>
          <c:idx val=""2""/>
          <c:order val=""2""/>
          <c:tx>
            <c:strRef>
              <c:f>Лист1!$D$1</c:f>
              <c:strCache>
                <c:ptCount val=""1""/>
                <c:pt idx=""0"">
                  <c:v>Ряд 3</c:v>
                </c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:invertIfNegative val=""0""/>
          <c:cat>
            <c:strRef>
              <c:f>Лист1!$A$2:$A$5</c:f>
              <c:strCache>
                <c:ptCount val=""4""/>
                <c:pt idx=""0"">
                  <c:v>Категория 1</c:v>
                </c:pt>
                <c:pt idx=""1"">
                  <c:v>Категория 2</c:v>
                </c:pt>
                <c:pt idx=""2"">
                  <c:v>Категория 3</c:v>
                </c:pt>
                <c:pt idx=""3"">
                  <c:v>Категория 4</c:v>
                </c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Лист1!$D$2:$D$5</c:f>
              <c:numCache>
                <c:formatCode>Основной</c:formatCode>
                <c:ptCount val=""4""/>
                <c:pt idx=""0"">
                  <c:v>2</c:v>
                </c:pt>
                <c:pt idx=""1"">
                  <c:v>2</c:v>
                </c:pt>
                <c:pt idx=""2"">
                  <c:v>3</c:v>
                </c:pt>
                <c:pt idx=""3"">
                  <c:v>5</c:v>
                </c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:dLbls>
          <c:showLegendKey val=""0""/>
          <c:showVal val=""0""/>
          <c:showCatName val=""0""/>
          <c:showSerName val=""0""/>
          <c:showPercent val=""0""/>
          <c:showBubbleSize val=""0""/>
        </c:dLbls>
        <c:gapWidth val=""150""/>
        <c:axId val=""148921728""/>
        <c:axId val=""154227840""/>
      </c:barChart>");
            #endregion

            XElement catAx = XElement.Parse(
      @"<c:catAx xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""> 
        <c:axId val=""148921728""/>
        <c:scaling>
          <c:orientation val=""minMax""/>
        </c:scaling>
        <c:delete val=""0""/>
        <c:axPos val=""b""/>
        <c:majorTickMark val=""out""/>
        <c:minorTickMark val=""none""/>
        <c:tickLblPos val=""nextTo""/>
        <c:crossAx val=""154227840""/>
        <c:crosses val=""autoZero""/>
        <c:auto val=""1""/>
        <c:lblAlgn val=""ctr""/>
        <c:lblOffset val=""100""/>
        <c:noMultiLvlLbl val=""0""/>
      </c:catAx>");

            XElement valAx = XElement.Parse(
      @"<c:valAx xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
        <c:axId val=""154227840""/>
        <c:scaling>
          <c:orientation val=""minMax""/>
        </c:scaling>
        <c:delete val=""0""/>
        <c:axPos val=""l""/>
        <c:majorGridlines/>
        <c:numFmt formatCode=""Основной"" sourceLinked=""1""/>
        <c:majorTickMark val=""out""/>
        <c:minorTickMark val=""none""/>
        <c:tickLblPos val=""nextTo""/>
        <c:crossAx val=""148921728""/>
        <c:crosses val=""autoZero""/>
        <c:crossBetween val=""between""/>
      </c:valAx>");

            XElement PlotArea = new XElement(
                XName.Get("plotArea", DocX.c.NamespaceName),
                new XElement(XName.Get("layout", DocX.c.NamespaceName)),
                ChartXml, catAx, valAx);

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

            ChartRootXml = Xml.Root.Element(XName.Get("chart", DocX.c.NamespaceName));
            ChartRootXml.Add(PlotArea);
        }

        public void AddLegend()
        {
            AddLegend(ChartLegendPosition.Right, false);
        }

        public void AddLegend(ChartLegendPosition position, Boolean overlay)
        {
            if (legend != null)
                RemoveLegend();
            legend = new ChartLegend(position, overlay);
            ChartRootXml.Add(legend.Xml);
        }

        public void RemoveLegend()
        {
            legend.Xml.Remove();
            legend = null;
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
                switch (Xml.Element(XName.Get("legendPos", DocX.c.NamespaceName)).Attribute("val").Value)
                {
                    case "t": return ChartLegendPosition.Top;
                    case "b": return ChartLegendPosition.Bottom;
                    case "l": return ChartLegendPosition.Left;
                    case "r": return ChartLegendPosition.Right;
                    case "tr": return ChartLegendPosition.TopRight;
                    default: throw new NotImplementedException();
                }
            }
            set
            {
                Xml.Element(XName.Get("legendPos", DocX.c.NamespaceName)).Attribute("val").Value = GetPositionValue(value);
            }
        }

        internal ChartLegend(ChartLegendPosition position, Boolean overlay)
        {
            Xml = new XElement(
                XName.Get("legend", DocX.c.NamespaceName),
                new XElement(XName.Get("legendPos", DocX.c.NamespaceName), new XAttribute("val", GetPositionValue(position))),
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

        /// <summary>
        /// ECMA-376, page 3906
        /// 21.2.3.24 ST_LegendPos (Legend Position)
        /// </summary>
        private String GetPositionValue(ChartLegendPosition position)
        {
            switch (position)
            {
                case ChartLegendPosition.Top:
                    return "t";
                case ChartLegendPosition.Bottom:
                    return "b";
                case ChartLegendPosition.Left:
                    return "l";
                case ChartLegendPosition.Right:
                    return "r";
                case ChartLegendPosition.TopRight:
                    return "tr";
                default:
                    throw new NotImplementedException();
            }
        }
    }

    /// <summary>
    /// Specifies the possible positions for a legend
    /// </summary>
    public enum ChartLegendPosition
    {
        Top,
        Bottom,
        Left,
        Right,
        TopRight
    }
}
