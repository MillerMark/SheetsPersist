using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace SheetsPersist
{
	public static partial class GoogleSheets
	{
		static Dictionary<Type, object> messageThrottlers = new Dictionary<Type, object>();

		private static void ValidateDocumentAndSheetNames(string documentName, string sheetName, bool trackSheetIfMissing = false)
		{
			ValidateDocumentName(documentName);

			if (trackSheetIfMissing)
				Track(documentName, sheetName);
			else if (documentSheetMap[documentName].IndexOf(sheetName) < 0)
				throw new InvalidDataException($"{nameof(sheetName)} (\"{sheetName}\") not found!");
		}

		private static void ValidateDocumentName(string documentName)
		{
			if (string.IsNullOrEmpty(documentName))
				throw new InvalidDataException($"documentName is null or empty.");

			if (!documentIDs.ContainsKey(documentName))
				throw new InvalidDataException($"{nameof(documentName)} (\"{documentName}\") not found!");
		}

		private static string GetDocumentName<T>()
		{
			DocumentNameAttribute documentNameAttribute = GetDocumentAttributes(typeof(T));
			return documentNameAttribute?.DocumentName;
		}

		internal static void InternalAppendRows<T>(T[] instances, string sheetNameOverride = null, string documentNameOverride = null) where T : class
		{
			if (instances == null || instances.Length == 0)
				return;

			GetDocumentAndSheetAttributes(typeof(T), out DocumentNameAttribute documentNameAttribute, out SheetNameAttribute tabNameAttribute);

			string documentName;
			string tabName;
			if (documentNameOverride != null)
				documentName = documentNameOverride;
			else
				documentName = documentNameAttribute.DocumentName;

			if (sheetNameOverride != null)
				tabName = sheetNameOverride;
			else
				tabName = tabNameAttribute.SheetName;

			Track(documentName, tabName);
			ValidateDocumentAndSheetNames(documentName, tabName);
			string documentId = documentIDs[documentName];

			IList<IList<Object>> values = new List<IList<Object>>();


			MemberInfo[] serializableFields = GetSerializableFields<ColumnAttribute>(typeof(T));

			foreach (object instance in instances)
			{
				IList<Object> row = new List<Object>();
				foreach (MemberInfo memberInfo in serializableFields)
					row.Add(GetValue(instance, memberInfo));

				values.Add(row);
			}

			ExecuteAppendRows(documentId, tabName, values);
		}

		private static void ExecuteRequests(string documentId, IList<Request> requests)
		{
			if (requests != null && requests.Count > 0)
			{
				BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest();
				body.Requests = requests;

				SpreadsheetsResource.BatchUpdateRequest batchUpdateRequest = Service.Spreadsheets.BatchUpdate(body, documentId);
				if (batchUpdateRequest != null)
				{
					BatchUpdateSpreadsheetResponse response = batchUpdateRequest.Execute();
				}
			}
		}

		static int? GetSheetId(string documentId, string sheetName)
		{
			Spreadsheet document = Service.Spreadsheets.Get(documentId).Execute();

			foreach (Sheet sheet in document.Sheets)
				if (sheet.Properties.Title == sheetName)
					return sheet.Properties.SheetId;

			return null;
		}

		static void DeleteRowByIndexInSheet(string documentId, string sheetName, int rowIndex)
		{
			if (rowIndex < 0)
				return;

			int? sheetId = GetSheetId(documentId, sheetName);
			Request RequestBody = new Request()
			{
				DeleteDimension = new DeleteDimensionRequest()
				{
					Range = new DimensionRange()
					{
						SheetId = sheetId,
						Dimension = "ROWS",
						StartIndex = Convert.ToInt32(rowIndex),
						EndIndex = Convert.ToInt32(rowIndex + 1)
					}
				}
			};

			List<Request> RequestContainer = new List<Request>();
			RequestContainer.Add(RequestBody);

			BatchUpdateSpreadsheetRequest deleteRequest = new BatchUpdateSpreadsheetRequest();
			deleteRequest.Requests = RequestContainer;

			SpreadsheetsResource.BatchUpdateRequest batchUpdate = Service.Spreadsheets.BatchUpdate(deleteRequest, documentId);
			batchUpdate.Execute();
		}

		private static Request GetUpdateRequestForCellInTopRow(int columnIndex, string documentId, string tabName)
		{
			Request commentRequest = new Request();
			commentRequest.UpdateCells = new UpdateCellsRequest();
			commentRequest.UpdateCells.Range = new GridRange();
			commentRequest.UpdateCells.Range.SheetId = GetSheetId(documentId, tabName);
			commentRequest.UpdateCells.Range.StartRowIndex = 0;
			commentRequest.UpdateCells.Range.EndRowIndex = 1;
			commentRequest.UpdateCells.Range.StartColumnIndex = columnIndex;
			commentRequest.UpdateCells.Range.EndColumnIndex = columnIndex + 1;
			commentRequest.UpdateCells.Rows = new List<RowData>();
			commentRequest.UpdateCells.Rows.Add(new RowData());
			commentRequest.UpdateCells.Rows[0].Values = new List<CellData>();
			commentRequest.UpdateCells.Rows[0].Values.Add(new CellData());
			return commentRequest;
		}

		private static Request GetRepeatCellRequestForEntireColumn(int columnIndex, string documentId, string tabName)
		{
			Request commentRequest = new Request();
			commentRequest.RepeatCell = new RepeatCellRequest();
			commentRequest.RepeatCell.Range = new GridRange();
			commentRequest.RepeatCell.Range.SheetId = GetSheetId(documentId, tabName);
			commentRequest.RepeatCell.Range.StartRowIndex = 0;
			commentRequest.RepeatCell.Range.EndRowIndex = null;
			commentRequest.RepeatCell.Range.StartColumnIndex = columnIndex;
			commentRequest.RepeatCell.Range.EndColumnIndex = columnIndex + 1;
			return commentRequest;
		}

		private static void UpdateRow(string sheetName, object[] instances, string[] saveOnlyTheseMembers, List<string> headerRow, IList<IList<object>> allRows, MemberInfo[] serializableFields, int i, int rowIndex, BatchUpdateValuesRequest requestBody)
		{
			for (int j = 0; j < serializableFields.Length; j++)
			{
				MemberInfo memberInfo = serializableFields[j];

				if (!HasMember(saveOnlyTheseMembers, memberInfo.Name))
					continue;

				if (instances[i] is ITrackPropertyChanges trackPropertyChanges && trackPropertyChanges.ChangedProperties != null)
				{
					if (trackPropertyChanges.ChangedProperties.IndexOf(memberInfo.Name) < 0)
						continue;
				}

				int columnIndex = GetColumnIndex(headerRow, GetColumnName<ColumnAttribute>(memberInfo));

				string existingValue = null;
				if (columnIndex < allRows[rowIndex].Count)  // Some rows may have fewer columns because the Google Sheets engine will efficiently return only the columns holding data.
					existingValue = GetExistingValue(allRows, columnIndex, rowIndex);

				string value = GetValue(instances[i], memberInfo);

				if (existingValue == null /* New */ || value != existingValue /* Mod */)
				{
					ValueRange body = GetValueRange(sheetName, columnIndex, rowIndex, value);

					requestBody.Data.Add(body);
				}
			}
		}

		private static void Execute(SpreadsheetsResource.ValuesResource.UpdateRequest request)
		{
			request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
			try
			{
				UpdateValuesResponse response = request.Execute();
				if (response != null)
				{

				}
			}
			catch (Exception ex)
			{
				string msg = ex.Message;
				System.Diagnostics.Debugger.Break();
			}
		}

		private static void Execute(SpreadsheetsResource.ValuesResource.BatchUpdateRequest request)
		{
			try
			{
				BatchUpdateValuesResponse response = request.Execute();
				if (response != null)
				{

				}
			}
			catch (Exception ex)
			{
				string msg = ex.Message;
				System.Diagnostics.Debugger.Break();
			}
		}

		static void ExecuteAppendRows(string documentId, string tabName, IList<IList<object>> values)
		{
			SpreadsheetsResource.ValuesResource.AppendRequest request = GetAppendRequest(documentId, tabName, values);
			try
			{
				var response = request.Execute();
			}
			catch (Exception ex)
			{

			}
		}

		static SpreadsheetsResource.ValuesResource.AppendRequest GetAppendRequest(string documentId, string tabName, IList<IList<object>> values)
		{
			SpreadsheetsResource.ValuesResource.AppendRequest request =
												 Service.Spreadsheets.Values.Append(new ValueRange() { Values = values }, documentId, tabName);
			request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
			request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
			return request;
		}

		static void ExecuteInsertRows(string documentId, string tabName, IList<IList<object>> values, int rowStartIndex = 0)
		{
			// TODO: Untested. Test this.
			BatchUpdateValuesRequest requestBody = GetBatchUpdateRequest();
			InsertDimensionRequest insertDimensionRequest = new InsertDimensionRequest();
			insertDimensionRequest.InheritFromBefore = false;
			insertDimensionRequest.Range.StartIndex = rowStartIndex;
			insertDimensionRequest.Range.StartIndex = rowStartIndex + values.Count;
			insertDimensionRequest.Range.Dimension = "ROWS";
			SpreadsheetsResource.ValuesResource.AppendRequest appendRequest =
									 Service.Spreadsheets.Values.Append(new ValueRange() { Values = values }, documentId, tabName);
			appendRequest.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
			appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
			try
			{
				var response = appendRequest.Execute();
			}
			catch (Exception ex)
			{
				ExecuteBatchUpdate(documentId, requestBody);
			}
		}

		private static void GetDocumentAndSheetAttributes(Type instanceType, out DocumentNameAttribute documentNameAttribute, out SheetNameAttribute sheetNameAttribute)
		{
			documentNameAttribute = GetDocumentAttributes(instanceType);

			sheetNameAttribute = instanceType.GetCustomAttribute<SheetNameAttribute>();
			if (sheetNameAttribute == null)
				throw new InvalidDataException($"{instanceType.Name} needs to specify the \"{nameof(SheetNameAttribute)}\" attribute.");
		}

		private static DocumentNameAttribute GetDocumentAttributes(Type instanceType)
		{
			DocumentNameAttribute documentNameAttribute = instanceType.GetCustomAttribute<DocumentNameAttribute>();
			if (documentNameAttribute == null)
				throw new InvalidDataException($"{instanceType.Name} needs to specify the \"{nameof(DocumentNameAttribute)}\" attribute.");
			return documentNameAttribute;
		}


		private static void GetHeaderRow(string docName, string tabName, out List<string> headerRow, out IList<IList<object>> allRows)
		{
			allRows = GetCells(docName, tabName);
			Dictionary<int, string> headers = new Dictionary<int, string>();
			IList<object> headerRowObjects = allRows[0];
			headerRow = headerRowObjects.Select(x => x.ToString()).ToList();
		}

		static string GetColumnName<TColumnAttribute>(MemberInfo memberInfo) where TColumnAttribute : ColumnNameAttribute
		{
			TColumnAttribute customAttribute = memberInfo.GetCustomAttribute<TColumnAttribute>();
			if (customAttribute == null || string.IsNullOrEmpty(customAttribute.ColumnName))
				return memberInfo.Name;
			return customAttribute.ColumnName;
		}

		static int GetColumnIndex(List<string> headerRow, string indexColumnName)
		{
			for (int i = 0; i < headerRow.Count; i++)
				if (string.Compare(headerRow[i], indexColumnName, StringComparison.OrdinalIgnoreCase) == 0)
					return i;
			return -1;
		}


		static int GetInstanceRowIndex(object obj, MemberInfo[] indexFields, IList<IList<object>> allRows, List<string> headerRow)
		{
			if (indexFields == null || indexFields.Length == 0)
				throw new InvalidDataException($"indexFields must contain data!");

			for (int rowIndex = 1; rowIndex < allRows.Count; rowIndex++)
			{
				bool found = false;
				for (int i = 0; i < indexFields.Length; i++)
				{
					string indexColumnName = GetColumnName<IndexerAttribute>(indexFields[i]);
					int column = GetColumnIndex(headerRow, indexColumnName);
					if (column == -1)
						throw new Exception($"Header not found in sheet: {indexColumnName}!");
					IList<object> thisRow = allRows[rowIndex];
					if (column >= thisRow.Count)
					{
						found = false;
						break;
					}
					if (thisRow[column].ToString() == GetValue(obj, indexFields[i]))
						found = true;
					else
					{
						found = false;
						break;
					}
				}
				if (found)
					return rowIndex;
			}

			return -1;
		}

		static string GetColumnId(int column)
		{
			int secondDigit = column / 26;
			int firstDigit = column % 26;
			string secondDigitStr = "";
			if (secondDigit > 0)
				secondDigitStr = GetColumnId(secondDigit - 1);
			return secondDigitStr + ((char)((byte)firstDigit + 65)).ToString();
		}

		static string GetRange(int columnIndex, int rowIndex)
		{
			return $"{GetColumnId(columnIndex)}{rowIndex + 1}";
		}

		static string GetExistingValue(IList<IList<object>> allRows, int columnIndex, int rowIndex)
		{
			object obj = allRows[rowIndex][columnIndex];
			if (obj == null)
				return string.Empty;
			return obj.ToString();
		}

		static bool HasMember(string[] memberList, string memberName)
		{
			if (memberList == null)
				return true;

			return memberList.Any(filterMember => filterMember == memberName);
		}

		static SheetsService service;

		static void AddInstance(IList<IList<object>> allRows, List<string> headerRow, MemberInfo[] serializableFields, object instance)
		{
			object[] row = new object[headerRow.Count];
			foreach (MemberInfo memberInfo in serializableFields)
			{
				int columnIndex = GetColumnIndex(headerRow, GetColumnName<ColumnAttribute>(memberInfo));
				row[columnIndex] = GetValue(instance, memberInfo);
			}
			allRows.Add(row.ToList());
		}

		static int AddRow(string sheetName, MemberInfo[] indexFields, List<string> headerRow, MemberInfo[] serializableFields, IList<IList<object>> allRows, object instance, BatchUpdateValuesRequest requestBody)
		{
			foreach (MemberInfo memberInfo in indexFields)
			{
				int columnIndex = GetColumnIndex(headerRow, GetColumnName<ColumnAttribute>(memberInfo));
				int rowIndex = allRows.Count;
				ValueRange body = GetValueRange(sheetName, columnIndex, rowIndex, GetValue(instance, memberInfo));

				requestBody.Data.Add(body);
			}

			AddInstance(allRows, headerRow, serializableFields, instance);
			return allRows.Count - 1;
		}

		private static ValueRange GetValueRange(string sheetName, int columnIndex, int rowIndex, string value)
		{
			string range = GetRange(columnIndex, rowIndex);

			ValueRange body = new ValueRange();
			body.MajorDimension = "ROWS";
			body.Range = $"{sheetName}!{range}";
			body.Values = new List<IList<object>>();
			body.Values.Add(new List<object>());
			body.Values[0].Add(value);
			return body;
		}

		private static void ExecuteBatchUpdate(string documentId, BatchUpdateValuesRequest requestBody)
		{
			SpreadsheetsResource.ValuesResource.BatchUpdateRequest request = Service.Spreadsheets.Values.BatchUpdate(requestBody, documentId);
			Execute(request);
		}

		private static BatchUpdateValuesRequest GetBatchUpdateRequest()
		{
			BatchUpdateValuesRequest requestBody = new BatchUpdateValuesRequest();
			requestBody.Data = new List<ValueRange>();
			requestBody.ValueInputOption = "USER_ENTERED";
			return requestBody;
		}
	}
}

