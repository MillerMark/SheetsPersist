using System;
using System.Linq;

namespace SheetsPersist
{
	public enum WriteValuesAs
	{
		/// <summary>
		/// The values the transferred to the Sheet will not be parsed and will be stored as-is.
		/// </summary>
		Raw = 0,

		/// <summary>
		/// The values will be parsed as if the user typed them into the UI. Numbers will stay as numbers, 
		/// but strings may be converted to numbers, dates, etc. following the same rules that are applied 
		/// when entering text into a cell via the Google Sheets UI.
		/// </summary>
		UserEntered = 1,
	}
}
