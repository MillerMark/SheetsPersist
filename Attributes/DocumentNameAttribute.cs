using System;
using System.Linq;

namespace SheetsPersist
{
	[AttributeUsage(AttributeTargets.Class, Inherited = false)]
	public class DocumentNameAttribute : Attribute
	{
		public DocumentNameAttribute(string sheetName)
		{
			DocumentName = sheetName;
		}

		public string DocumentName { get; set; }
	}
}
