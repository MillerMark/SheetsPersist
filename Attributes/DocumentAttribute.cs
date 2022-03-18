using System;
using System.Linq;

namespace SheetsPersist
{
	[AttributeUsage(AttributeTargets.Class, Inherited = false)]
	public class DocumentAttribute : Attribute
	{
		public DocumentAttribute(string documentName)
		{
			DocumentName = documentName;
		}

		public string DocumentName { get; set; }
	}
}
