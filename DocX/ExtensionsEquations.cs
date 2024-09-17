using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Novacode;
using System.Xml.Linq;

namespace OMath
{
    public static class ExtensionsEquations
    {
        public static Novacode.Paragraph InsertEquation(this DocX doc, Equation equation)
        {
            Paragraph eqParagraph = doc.InsertEquation("");
            XElement xml = eqParagraph.Xml;
            XNamespace mathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            XElement omath = xml.Descendants(mathNamespace + "oMathPara").First();
            omath.Elements().Remove();
            omath.Add(equation.Xml);

            return eqParagraph;
        }
    }
}
