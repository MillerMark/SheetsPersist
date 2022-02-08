using System;
using System.Linq;

namespace SheetsPersist
{
	public abstract class FormatAttribute : Attribute
	{
		public string Pattern { get; set; }
		public FormatAttribute(string pattern)
		{
			Pattern = pattern;
		}
	}
}
