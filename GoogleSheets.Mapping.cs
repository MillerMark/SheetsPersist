using System;
using System.Collections.Generic;
using System.Linq;

namespace SheetsPersist
{
	public static partial class GoogleSheets
	{
		static Dictionary<string, string> documentIDs = new Dictionary<string, string>();

		/// <summary>
		/// Registers the specified documentName with the specified documentID.
		/// </summary>
		/// <param name="documentName">The name of the document (must match the string passed to the DocumentName attribute).</param>
		/// <param name="documentID">The spreadsheet document ID (from the URL when the spreadsheet is open).</param>
		public static void RegisterDocumentID(string documentName, string documentID)
		{
			documentIDs[documentName] = documentID;
		}

		static Dictionary<string, List<string>> documentSheetMap = new Dictionary<string, List<string>>();

		static void Track(string docName, string sheetName)
		{
			if (!documentSheetMap.ContainsKey(docName))
				documentSheetMap.Add(docName, new List<string>());
			List<string> sheetNames = documentSheetMap[docName];
			if (!sheetNames.Contains(sheetName))
				sheetNames.Add(sheetName);
		}
	}
}

