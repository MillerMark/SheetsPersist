using System;
using System.Linq;

namespace SheetsPersist
{
	[AttributeUsage(AttributeTargets.Class, Inherited = false)]
	public class HeaderRowAttribute : Attribute
	{
		public HeaderRowAttribute(string color, FontWeight fontWeight, RowFreezeOption rowFreezeOption)
		{
			RowFreezeOption = rowFreezeOption;
			FontWeight = fontWeight;
			Color = color;
		}

		public string Color { get; set; }
		public FontWeight FontWeight { get; set; }
		public RowFreezeOption RowFreezeOption { get; set; }
	}
}
