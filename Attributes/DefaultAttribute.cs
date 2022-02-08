using System;
using System.Linq;

namespace SheetsPersist
{
	[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
	public class DefaultAttribute : Attribute
	{
		public DefaultAttribute(object defaultValue)
		{
			DefaultValue = defaultValue;
		}
		public object DefaultValue { get; set; }
	}
}
