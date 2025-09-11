using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Novacode;

namespace OMath
{
	public class Equation
	{
		private static XNamespace mathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
		private static XNamespace wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";


		private XElement xml;
		public XElement Xml { get => xml; set => xml = value; }

		#region Constructors

		private Equation(XElement el)
		{
			Xml = new XElement(mathNamespace + "oMath", el);
		}
		public Equation()
		{
			Xml = new XElement(mathNamespace + "oMath");
		}		
		public Equation(string content)
		{
			Xml = new XElement(mathNamespace + "oMath", CreateLiteral(content));
		}

		#endregion


		#region Operators

		public static implicit operator Equation(string content)
		{
			Equation eq = new Equation();
			eq.AppendElement(CreateLiteral(content));
			return eq;
		}
		public static Equation operator +(Equation eq1, Equation eq2)
		{
			Equation eq = new Equation();
			eq.Append(eq1);
			eq.Append("+");
			eq.Append(eq2);
			return eq;
		}
		public static Equation operator -(Equation eq1, Equation eq2)
		{
			Equation eq = new Equation();
			eq.Append(eq1);
			eq.Append("-");
			eq.Append(eq2);
			return eq;
		}
		public static Equation operator *(Equation eq1, Equation eq2)
		{
			Equation eq = new Equation();
			eq.Append(eq1);
			eq.Append("*");
			eq.Append(eq2);
			return eq;
		}
		public static Equation operator /(Equation eq1, Equation eq2)
		{
			return Equation.Fraction(eq1, eq2);
		}

		#endregion

		#region Methods
		/// <summary>
		/// Append and xml element to the object xml
		/// </summary>
		/// <param name="content"></param>
		public void AppendElement(XElement content)
		{
			Xml.Add(content);
		}
		/// <summary>
		/// Append an equation to the current equation
		/// </summary>
		/// <param name="content">Equation to append</param>
		public void Append(Equation content)
		{
			Xml.Add(content.Xml.Elements());
		}
		#endregion

		#region Equation Creation
			private static XElement CreateLiteral(object literal)
			{
				XElement lit = new XElement(mathNamespace + "r");
				XElement rPr = new XElement(wordNamespace + "rPr");
				XElement rFonts = new XElement(wordNamespace + "rFonts");

				rFonts.Add(new XAttribute(wordNamespace + "ascii", "Cambria Math"));
				rFonts.Add(new XAttribute(wordNamespace + "hAnsi", "Cambria Math"));
				rPr.Add(rFonts);
				lit.Add(rPr);
				lit.Add(new XElement(mathNamespace + "t", literal));
				return lit;
			}

			private static XElement CreateBox(object content)
			{
				XElement borderBox = new XElement(mathNamespace + "borderBox");
				borderBox.Add(new XElement(mathNamespace + "e", content));
				return borderBox;
			}
		/// <summary>
		/// Creates a Box around the equation
		/// </summary>
		/// <param name="content">Content of the Box</param>
		/// <returns></returns>
		public static Equation Box(Equation content)
			{
				return new Equation(CreateBox(content.Xml.Elements()));
			}

			#region Parenthesis
			private static XElement CreateParenthesis(object content, char opening = '(', char closing = ')')
			{
				XElement parenthesis = new XElement(mathNamespace + "d");

				if (opening != '(' || closing != ')')
				{
					XElement dPr = new XElement(mathNamespace + "dPr");
					XElement begChr = new XElement(mathNamespace + "begChr", new XAttribute(mathNamespace + "val", opening.ToString()));
					XElement endChr = new XElement(mathNamespace + "endChr", new XAttribute(mathNamespace + "val", closing.ToString()));
					dPr.Add(begChr);
					dPr.Add(endChr);
					parenthesis.Add(dPr);
				}
				XElement _content = new XElement(mathNamespace + "e", content);
				parenthesis.Add(_content);

				return parenthesis;
			}
		/// <summary>
		/// Creates parenthesis, brackets and other types of enclousures
		/// </summary>
		/// <param name="content">Equation inside the enclousure</param>
		/// <param name="opening">Opening char for the enclousure. Default is '('</param>
		/// <param name="closing">Closing char for the enclousure. Default is ')'</param>
		/// <returns></returns>
		public static Equation Parenthesis(Equation content, char opening = '(', char closing = ')')
			{
				return new Equation(CreateParenthesis(content, opening, closing));
			}
			#endregion

			#region Root
			private static XElement CreateRoot(object radicand, object degree = null)
			{
				XElement rad = new XElement(mathNamespace + "rad");

				XElement deg = new XElement(mathNamespace + "deg");
				if (degree == null)
				{
					XElement radProp = new XElement(mathNamespace + "radPr");
					XElement degHide = new XElement(mathNamespace + "degHide");
					degHide.Add(new XAttribute(mathNamespace + "val", 1));
					radProp.Add(degHide);
					rad.Add(radProp);
				}
				else
				{
					deg.Add(degree);
				}
				rad.Add(deg);

				XElement content = new XElement(mathNamespace + "e", radicand);
				//content.Add(new XElement(mathNamespace + "r", new XElement(mathNamespace + "t", radicand)));
				rad.Add(content);

				return rad;
			}
		/// <summary>
		/// Creates a n degree root 
		/// </summary>
		/// <param name="radicand">Radicand of the root</param>
		/// <param name="degree">Degree of the root</param>
		/// <returns></returns>
		public static Equation Root(Equation radicand, Equation degree)
			{
				Equation eq = new Equation(CreateRoot(radicand.Xml.Elements(), degree.Xml.Elements()));
				return eq;
			}
		/// <summary>
		/// Creates a square root
		/// </summary>
		/// <param name="radicand">Content of the square root</param>
		/// <returns></returns>
		public static Equation Root(Equation radicand)
			{
				Equation eq = new Equation(CreateRoot(radicand.Xml.Elements()));
				return eq;
			}
			#endregion

			#region Fraction
			private static XElement CreateFraction(object num, object den)
			{
				XElement numerator = new XElement(mathNamespace + "num", num);

				XElement denominator = new XElement(mathNamespace + "den", den);

				XElement fraction = new XElement(mathNamespace + "f");
				fraction.Add(numerator);
				fraction.Add(denominator);
				return fraction;
			}
		/// <summary>
		/// Creates a fraction
		/// </summary>
		/// <param name="num">Numerator of the fraction</param>
		/// <param name="den">Denominator of the fraction</param>
		/// <returns></returns>
		public static Equation Fraction(Equation num, Equation den)
			{
				Equation eq = new Equation();
				eq.AppendElement(CreateFraction(num.Xml.Elements(), den.Xml.Elements()));
				return eq;
			}
			#endregion

			#region SuperSubscripts
			private static XElement CreateSuperscript(object content, object superior)
			{
				XElement sSup = new XElement(mathNamespace + "sSup");
				XElement e = new XElement(mathNamespace + "e", content);
				XElement sup = new XElement(mathNamespace + "sup", superior);
				sSup.Add(e);
				sSup.Add(sup);
				return sSup;
			}
		/// <summary>
		/// Creates a superscript
		/// </summary>
		/// <param name="content">Base element</param>
		/// <param name="superior">Expoent</param>
		/// <returns></returns>
		public static Equation Superscript(Equation content, Equation superior)
			{
				return new Equation(CreateSuperscript(content.Xml.Elements(), superior.Xml.Elements()));
			}

			private static XElement CreateSubscript(object content, object inferior)
			{
				XElement sSub = new XElement(mathNamespace + "sSub");
				XElement e = new XElement(mathNamespace + "e", content);
				XElement sub = new XElement(mathNamespace + "sub", inferior);
				sSub.Add(e);
				sSub.Add(sub);
				return sSub;
			}
		/// <summary>
		/// Creates a subscript
		/// </summary>
		/// <param name="content">Base element</param>
		/// <param name="inferior">subscript</param>
		/// <returns></returns>
		public static Equation Subscript(Equation content, Equation inferior)
			{
				return new Equation(CreateSubscript(content.Xml.Elements(), inferior.Xml.Elements()));
			}

			private static XElement CreateSubSuperscript(object content, object inferior, object superior)
			{
				XElement sSubSup = new XElement(mathNamespace + "sSubSup");
				XElement e = new XElement(mathNamespace + "e", content);
				XElement sub = new XElement(mathNamespace + "sub", inferior);
				XElement sup = new XElement(mathNamespace + "sup", superior);
				sSubSup.Add(e);
				sSubSup.Add(sub);
				sSubSup.Add(sup);
				return sSubSup;
			}
		/// <summary>
		/// Creates an element with a superscript and a subscript
		/// </summary>
		/// <param name="content">Base element</param>
		/// <param name="inferior">Subscript</param>
		/// <param name="superior">Expoent</param>
		/// <returns></returns>
		public static Equation SubSuperscript(Equation content, Equation inferior, Equation superior)
			{
				return new Equation(CreateSubSuperscript(content.Xml.Elements(),inferior.Xml.Elements(), superior.Xml.Elements()));
			}

			private static XElement CreatePrescript(object content, object inferior, object superior)
			{
				XElement sPre = new XElement(mathNamespace + "sPre");
				XElement e = new XElement(mathNamespace + "e", content);
				XElement sub = new XElement(mathNamespace + "sub", inferior);
				XElement sup = new XElement(mathNamespace + "sup", superior);
				sPre.Add(sub);
				sPre.Add(sup);
				sPre.Add(e);
				return sPre;
			}
		/// <summary>
		/// Creates an element with a superscrit and a subscript before the base element
		/// </summary>
		/// <param name="content">Base element</param>
		/// <param name="inferior">Subscript</param>
		/// <param name="superior">Superscript</param>
		/// <returns></returns>
		public static Equation Prescript(Equation content, Equation inferior, Equation superior)
			{
				return new Equation(CreateSubSuperscript(content.Xml.Elements(), inferior.Xml.Elements(), superior.Xml.Elements()));
			}
			#endregion

			#region Integral, Sums, Products
			public enum IntegralPosition { Top, Front };
			public enum IntegralType { Simple, Double, Triple, Line, Surface, Volume, Sum, Product, Union, Intersection, Disjunction, Conjunction }
			private static Dictionary<IntegralType, char> integralTypeDictionary = new Dictionary<IntegralType, char>()
			{
				{IntegralType.Simple,' ' },
				{IntegralType.Double,'∬' },
				{IntegralType.Triple,'∭' },
				{IntegralType.Line,'∮' },
				{IntegralType.Surface,'∯' },
				{IntegralType.Volume,'∰' },
				{IntegralType.Sum,'∑' },
				{IntegralType.Product,'∏' },
				{IntegralType.Union,'⋃' },
				{IntegralType.Intersection,'⋂' },
				{IntegralType.Disjunction,'⋁' },
				{IntegralType.Conjunction,'⋀' }
			};
			private static XElement CreateIntegral(object content, object inferior = null, object superior = null, string position = "undOvr", char chr = ' ')
			{
				XElement nary = new XElement(mathNamespace + "nary");
				XElement naryPr = new XElement(mathNamespace + "naryPr");
				if (chr != ' ')
				{
					naryPr.Add(
						new XElement(
							mathNamespace + "chr",
							new XAttribute(
								mathNamespace + "val",
								chr
							)));
				}
				naryPr.Add(
						new XElement(
							mathNamespace + "limLoc",
							new XAttribute(
								mathNamespace + "val",
								position
							)));
				if (inferior == null)
				{
					naryPr.Add(
					new XElement(
						mathNamespace + "subHide",
						new XAttribute(
							mathNamespace + "val",
							"1"
						)));
				
				}
				else
				{
					nary.Add(new XElement(mathNamespace + "sub", inferior));
				}
				if (superior == null)
				{
					naryPr.Add(
						new XElement(
							mathNamespace + "supHide",
							new XAttribute(
								mathNamespace + "val",
								"1"
							)));
				}
				else
				{
					nary.Add(new XElement(mathNamespace + "sup", superior));
				}
				nary.Add(naryPr);
				nary.Add(new XElement(mathNamespace + "e", content));
				return nary;
			}
		/// <summary>
		/// Creates integrals, sums, products, union, intersection, disjuntion or cojunction
		/// </summary>
		/// <param name="content">Content of the equation</param>
		/// <param name="inferior">Bottom limit. To hide, use null. Default is null</param>
		/// <param name="superior">Upper limit. To hide, use null. Default is null</param>
		/// <param name="position">Position of limits, could be Top for above and bellow the symbol, 
		/// or Front, to be in front of the symbol </param>
		/// <param name="it">Symbol for the equation. Default is Simple (Integral)</param>
		/// <returns></returns>
		public static Equation Integral(Equation content, Equation inferior=null, Equation superior=null, IntegralPosition position=IntegralPosition.Front, IntegralType it = IntegralType.Simple)
		{
			string pos = position == IntegralPosition.Top ? "undOvr" : "subSup";
			return new Equation(CreateIntegral(content.Xml.Elements(), inferior.Xml.Elements(), superior.Xml.Elements(), pos, integralTypeDictionary[it]));
		}
			#endregion

			#region Matrix
			private static XElement CreateMatrix(object[,] content)
			{
				XElement m = new XElement(mathNamespace + "m");
				int columns = content.GetLength(1);
				XElement mPr = new XElement(mathNamespace + "mPr",
					new XElement(mathNamespace+"mcs",
						new XElement(mathNamespace+"mc",
							new XElement(mathNamespace+"mcPr",
								new XElement(mathNamespace+"count",
									new XAttribute(mathNamespace+"val",columns)),
								new XElement(mathNamespace + "mcJc",
									new XAttribute(mathNamespace + "val", "center")
									)))));
				m.Add(mPr);
				for(int row = 0; row < content.GetLength(0); row++)
				{
					XElement mr = new XElement(mathNamespace + "mr");
					for(int col = 0; col < columns; col++)
					{
						mr.Add(new XElement(mathNamespace + "e", content[row, col]));
					}
					m.Add(mr);
				}
				return m;
			}
		/// <summary>
		/// Creates a matrix
		/// </summary>
		/// <param name="content">Bidimensional array with the elements of the matrix</param>
		/// <returns></returns>
		public static Equation Matrix(Equation[,] content)
			{
				object[,] mat = new object[content.GetLength(0), content.GetLength(1)];
				for(int r=0;r< content.GetLength(0); r++)
				{
					for (int c = 0; c < content.GetLength(1); c++)
					{
						mat[r, c] = content[r, c].Xml.Elements();
					}
				}
				return new Equation(CreateMatrix(mat));
			}
		#endregion

			#region Enfasis
			private static XElement CreateEnfasis(object content, char enf=' ')
			{
				XElement acc = new XElement(mathNamespace + "acc");
				if (enf!=' ')
				{
					XElement accPr = new XElement(mathNamespace + "accPr");
					accPr.Add(new XElement(mathNamespace + "chr", new XAttribute(mathNamespace + "val", enf)));
					acc.Add(accPr);
				}
				acc.Add(new XElement(mathNamespace + "e", content));			
				return acc;
			}
		/// <summary>
		/// Creates a element with a char above
		/// </summary>
		/// <param name="content">The base element</param>
		/// <param name="enf">The char above the base element</param>
		/// <returns></returns>
		public static Equation Enfasis(Equation content, char enf)
			{
				return new Equation(CreateEnfasis(content.Xml.Elements(), enf));
			}
			#endregion

			#region Limit
			public enum LimitPosition { Bottom, Top };
			private static XElement CreateLimit(object content, object limit, LimitPosition limitPosition=LimitPosition.Bottom)
			{
				string type = limitPosition == LimitPosition.Top ? "limUpp" : "limLow";
				XElement lim = new XElement(mathNamespace + type);
				lim.Add(new XElement(mathNamespace + "e", content));
				lim.Add(new XElement(mathNamespace + "lim", limit));
				return lim;
			}
		/// <summary>
		/// Creates an element with an element below in a smaller font
		/// </summary>
		/// <param name="content">Base element</param>
		/// <param name="limit">Smaller font element</param>
		/// <returns></returns>
		public static Equation Limit(Equation content, Equation limit)
			{
				return new Equation(CreateLimit(content.Xml.Elements(), limit.Xml.Elements()));
			}
		/// <summary>
		/// Creates an element with an element above or below in a smaller font
		/// </summary>
		/// <param name="content">Base element</param>
		/// <param name="limit">Smaller font element</param>
		/// <param name="limitPosition">Position of the smaller element</param>
		/// <returns></returns>
		public static Equation Limit(Equation content, Equation limit, LimitPosition limitPosition)
			{
				return new Equation(CreateLimit(content.Xml.Elements(), limit.Xml.Elements(),limitPosition));
			}
			#endregion

			#region Function
			private static XElement CreateFunction(XElement name, object argument)
			{
				XElement func = new XElement(mathNamespace + "func");
			
				foreach (var n in name.Descendants(mathNamespace + "r"))
				{
					XElement rPr = n.Element(mathNamespace + "rPr");
					if (rPr == null)
					{
						n.AddFirst(new XElement(mathNamespace + "rPr"));
						rPr = n.Element(mathNamespace + "rPr");
					}
					rPr.Add(new XElement(mathNamespace + "sty", new XAttribute(mathNamespace + "val", "p")));
				}
				func.Add(new XElement(mathNamespace + "fName", name));			
			
				func.Add(new XElement(mathNamespace + "e", argument));
				return func;
			}
			private static XElement CreateFunction(IEnumerable<XElement> name, object argument)
			{
				XElement func = new XElement(mathNamespace + "func");
				foreach(var nam in name)
				{
					foreach (var n in nam.Descendants(mathNamespace + "r"))
					{
						XElement rPr = n.Element(mathNamespace + "rPr");
						if (rPr == null)
						{
							n.AddFirst(new XElement(mathNamespace + "rPr"));
							rPr = n.Element(mathNamespace + "rPr");
						}
						rPr.Add(new XElement(mathNamespace + "sty", new XAttribute(mathNamespace + "val", "p")));
					}
				}
				func.Add(new XElement(mathNamespace + "fName", name));
				func.Add(new XElement(mathNamespace + "e", argument));
				return func;
			}
		/// <summary>
		/// Creates a function with a font name in normal font 
		/// and an argument with math style font in front of the name element
		/// </summary>
		/// <param name="name">Normal font element</param>
		/// <param name="argument">Math font element</param>
		/// <returns></returns>
		public static Equation Function(Equation name, Equation argument)
			{
				return new Equation(CreateFunction(name.Xml.Elements(), argument.Xml.Elements()));
			}
			#endregion
		
		#endregion
	}
}
