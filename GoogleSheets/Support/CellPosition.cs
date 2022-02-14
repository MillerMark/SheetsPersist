using System;
using System.Linq;

namespace SheetsPersist
{
	public readonly struct CellPosition
	{
		readonly int column;
		readonly int row;
		public CellPosition(int column, int row)
		{
			this.column = column;
			this.row = row;
		}

		public int Column => column;

		public int Row => row;
	}
}

