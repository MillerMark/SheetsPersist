using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SheetsPersist
{
	public static partial class GoogleSheets
	{
		private static void AddNote(IList<Request> requests, int columnIndex, string comment, string documentId, string tabName)
		{
			Request commentRequest = GetUpdateRequestForCellInTopRow(columnIndex, documentId, tabName);
			commentRequest.UpdateCells.Rows[0].Values[0].Note = comment;
			commentRequest.UpdateCells.Fields = "note";

			requests.Add(commentRequest);
		}

		private static void AddColumnNotes(IList<Request> requests, string documentId, string tabName, MemberInfo[] serializableFields)
		{
			int columnIndex = 0;

			foreach (MemberInfo memberInfo in serializableFields)
			{
				NoteAttribute commentAttribute = memberInfo.GetCustomAttribute<NoteAttribute>();
				if (commentAttribute != null && !string.IsNullOrEmpty(commentAttribute.ColumnNote))
					AddNote(requests, columnIndex, commentAttribute.ColumnNote, documentId, tabName);

				columnIndex++;
			}
		}
	}
}

