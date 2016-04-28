using System;

namespace System.Drawing
{
	//	FontFamily

	public class FontFamily : MarshalByRefObject
	{
		//	properties

		public string Name { get; private set; }

		//	constructor

		public FontFamily (string name)
		{
			this.Name = name;
		}
	}
}