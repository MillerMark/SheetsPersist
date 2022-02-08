using System;
using System.Linq;

namespace SheetsPersist
{
	public class FormatDateAttribute : FormatAttribute
	{
		public FormatDateAttribute(string pattern) : base(pattern)
		{

		}
	}
}
