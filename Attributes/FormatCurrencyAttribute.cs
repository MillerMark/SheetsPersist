using System;
using System.Linq;

namespace SheetsPersist
{
	public class FormatCurrencyAttribute : FormatAttribute
	{
		public FormatCurrencyAttribute(string pattern) : base(pattern)
		{

		}
	}
}
