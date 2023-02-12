using System;
using System.Linq;

namespace SheetsPersist
{
	[AttributeUsage(AttributeTargets.Class, Inherited = false)]
	public class SheetAttribute : Attribute
	{
		public SheetAttribute(string sheetName, int frozenRowCount = 0, int frozenColumnCount = 0, ReadValuesAs readValuesAs = ReadValuesAs.Formatted, WriteValuesAs writeValuesAs = WriteValuesAs.UserEntered)
		{
			WriteValuesAs = writeValuesAs;
			ReadValuesAs = readValuesAs;
			FrozenColumnCount = frozenColumnCount;
			FrozenRowCount = frozenRowCount;
			SheetName = sheetName;
		}

		public string SheetName { get; set; }
		public int FrozenRowCount { get; set; }
		public int FrozenColumnCount { get; set; }
		public ReadValuesAs ReadValuesAs { get; set; }
		public WriteValuesAs WriteValuesAs { get; set; }
	}
}
