using System;
using System.Linq;

namespace SheetsPersist
{
	public abstract class ColumnNameAttribute : Attribute
	{
		public string ColumnName { get; set; }

		public ColumnNameAttribute(string columnName = "")
		{
			ColumnName = columnName;
		}
	}
}
