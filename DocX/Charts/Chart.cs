using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using System.IO.Packaging;
using System.Collections;
using System.Drawing;
using System.Reflection;

namespace Novacode
{
    /// <summary>
    /// Represents every Chart in this document.
    /// </summary>
    public abstract class Chart
    {
        protected XElement ChartXml { get; private set; }
        protected XElement ChartRootXml { get; private set; }

        public XDocument Xml { get; private set; }

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

        public virtual void AddSeries(Series series)
        {
            ChartXml.Add(series.Xml);
        }

        private ChartLegend legend;
        public ChartLegend Legend { get { return legend; } }

        /// <summary>
        /// Specifies how blank cells shall be plotted on a chart
        /// </summary>
        public DisplayBlanksAs DisplayBlanksAs
        {
            get
            {
                String value = ChartRootXml.Element(XName.Get("dispBlanksAs", DocX.c.NamespaceName)).Attribute(XName.Get("val")).Value;
                switch (value)
                {
                    case "gap": return DisplayBlanksAs.Gap;
                    case "span": return DisplayBlanksAs.Span;
                    case "zero": return DisplayBlanksAs.Zero;
                    default: throw new NotImplementedException("This DisplayBlanksAsType was not implement!");
                }
            }
            set
            {
                String newValue;
                switch (value)
                {
                    case DisplayBlanksAs.Gap: newValue = "gap";
                        break;
                    case DisplayBlanksAs.Span: newValue = "span";
                        break;
                    case DisplayBlanksAs.Zero: newValue = "zero";
                        break;
                    default: throw new NotImplementedException("This DisplayBlanksAsType was not implement!");
                }
                ChartRootXml.Element(XName.Get("dispBlanksAs", DocX.c.NamespaceName)).Attribute(XName.Get("val")).Value = newValue;
            }
        }

        /// <summary>
        /// Create an Chart for this document
        /// </summary>        
        public Chart()
        {
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
            ChartRootXml.Add(CreatePlotArea());
        }

        private XElement CreatePlotArea()
        {
            XElement dLbls = XElement.Parse(
                @"<c:dLbls xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                      <c:showLegendKey val=""0""/>
                      <c:showVal val=""0""/>
                      <c:showCatName val=""0""/>
                      <c:showSerName val=""0""/>
                      <c:showPercent val=""0""/>
                      <c:showBubbleSize val=""0""/>
                    </c:dLbls>");
            XElement axIDcat = XElement.Parse(
                @"<c:axId val=""154227840"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""/>");
            XElement axIDval = XElement.Parse(
                @"<c:axId val=""148921728"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""/>");

            ChartXml = CreateChartXml();
            ChartXml.Add(dLbls);            
            ChartXml.Add(axIDval);
            ChartXml.Add(axIDcat);

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
                <c:numFmt sourceLinked=""0"" formatCode=""General""/>
                <c:majorGridlines/>
                <c:majorTickMark val=""out""/>
                <c:minorTickMark val=""none""/>
                <c:tickLblPos val=""nextTo""/>
                <c:crossAx val=""148921728""/>
                <c:crosses val=""autoZero""/>
                <c:crossBetween val=""between""/>
              </c:valAx>");

            return new XElement(
                XName.Get("plotArea", DocX.c.NamespaceName),
                new XElement(XName.Get("layout", DocX.c.NamespaceName)),
                ChartXml, catAx, valAx);
        }

        protected abstract XElement CreateChartXml();


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
    /// Specifies the possible positions for a legend.
    /// 21.2.3.24 ST_LegendPos (Legend Position)
    /// </summary>
    public enum ChartLegendPosition
    {
        Top,
        Bottom,
        Left,
        Right,
        TopRight
    }

    /// <summary>
    /// Specifies the possible ways to display blanks.
    /// 21.2.3.10 ST_DispBlanksAs (Display Blanks As)
    /// </summary>
    public enum DisplayBlanksAs
    {
        Gap,
        Span,
        Zero
    }    
}
