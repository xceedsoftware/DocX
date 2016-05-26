using System;
using System.Drawing;

using Foundation;
using UIKit;

namespace System.Drawing
{
	//	ColorTranslator

	public static class ColorTranslator
	{
		//	FromHtml

		public static Color FromHtml(string color, float alpha = 1.0f)
		{
			color = color.Replace ("#", "").Replace (" ", "").Trim ();

			if (alpha > 1.0f) 
			{
				alpha = 1.0f;
			} 

			if (alpha < 0.0f) 
			{
				alpha = 0.0f;
			}

			int A = 0, R = 0, G = 0, B = 0;

			switch (color.Length) 
			{
				case 3 : // #RGB
				{
					A = (int)(alpha * 255);

					R = Convert.ToInt32(string.Format("{0}{0}", color.Substring(0, 1)), 16);

					G = Convert.ToInt32(string.Format("{0}{0}", color.Substring(1, 1)), 16);

					B = Convert.ToInt32(string.Format("{0}{0}", color.Substring(2, 1)), 16);

					break;
				}

				case 4 : // #ARGB
				{
					A = Convert.ToInt32(string.Format("{0}{0}", color.Substring(0, 1)), 16);

					R = Convert.ToInt32(string.Format("{0}{0}", color.Substring(1, 1)), 16);

					G = Convert.ToInt32(string.Format("{0}{0}", color.Substring(2, 1)), 16);

					B = Convert.ToInt32(string.Format("{0}{0}", color.Substring(3, 1)), 16);

					break;
				}

				case 6 : // #RRGGBB
				{
					A = (int)(alpha * 255);

					R = Convert.ToInt32(color.Substring(0, 2), 16);

					G = Convert.ToInt32(color.Substring(2, 2), 16);

					B = Convert.ToInt32(color.Substring(4, 2), 16);

					break;
				}   

				case 8 : // #RRGGBB
				{
					A = Convert.ToInt32(color.Substring(0, 2), 16);

					R = Convert.ToInt32(color.Substring(2, 2), 16);

					G = Convert.ToInt32(color.Substring(4, 2), 16);

					B = Convert.ToInt32(color.Substring(6, 2), 16);

					break;
				}   
			}

			return Color.FromArgb (A, R, G, B);
		}
	}
}