using System;
using System.Linq;

namespace SheetsPersist
{
	/// <summary>
	/// Use this to set font color and font weight for a particular column in a Google Sheet.
	/// </summary>
	[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
	public class StyleAttribute : Attribute
	{
		public StyleAttribute(string color = "", FontWeight fontWeight = FontWeight.Normal)
		{
			FontWeight = fontWeight;
			Color = color;
		}
		public string Color { get; set; }
		public FontWeight FontWeight { get; set; }
	}
}
