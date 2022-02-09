using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Reflection;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Threading;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.Globalization;

namespace SheetsPersist
{
	public static partial class GoogleSheets
	{
		const string STR_TextFormat = "textFormat";
		const string STR_FrozenRowCount = "gridProperties.frozenRowCount";
		/// <summary>
		/// Gets a List of T objects from the spreadsheet specified T's DocumentName and SheetName attributes.
		/// </summary>
		/// <typeparam name="T">The type to retrieve.</typeparam>
		/// <returns>Returns a List of T objects found in the associated spreadsheet.</returns>
		/// <exception cref="InvalidDataException">Thrown if T is missing the DocumentName or SheetName attributes.</exception>
		public static List<T> Get<T>() where T : new()
		{
			DocumentNameAttribute documentNameAttribute = typeof(T).GetCustomAttribute<DocumentNameAttribute>();
			if (documentNameAttribute == null)
				throw new InvalidDataException($"{nameof(DocumentNameAttribute)} not found on (\"{typeof(T).Name}\").");
			SheetNameAttribute sheetAttribute = typeof(T).GetCustomAttribute<SheetNameAttribute>();
			if (sheetAttribute == null)
				throw new InvalidDataException($"{nameof(SheetNameAttribute)} not found on (\"{typeof(T).Name}\").");

			return Get<T>(documentNameAttribute.DocumentName, sheetAttribute.SheetName);
		}

		static IList<IList<object>> GetCells(string docName, string sheetName, string cellRange = "")
		{
			if (!documentIDs.ContainsKey(docName))
				throw new InvalidDataException($"docName (\"{docName}\") not found!");

			string documentId = documentIDs[docName];

			string range;
			if (string.IsNullOrEmpty(cellRange))
				range = sheetName;  // Reading the entire file.
			else
				range = $"{sheetName}!{cellRange}";

			SpreadsheetsResource.ValuesResource.GetRequest request = Service.Spreadsheets.Values.Get(documentId, range);
			ValueRange response = request.Execute();
			return response.Values;
		}


		static T NewItem<T>(Dictionary<int, string> headers, IList<object> row) where T : new()
		{
			T result = new T();
			TransferValues(result, headers, row);
			return result;
		}

		/// <summary>
		/// Gets a List of T objects from the specified documentName and sheetName parameters.
		/// </summary>
		/// <typeparam name="T">The type to retrieve.</typeparam>
		/// <param name="documentName">The name of the document containing the data to get (must be registered with a call to 
		/// RegisterDocumentID).</param>
		/// <param name="sheetName">The name of the sheet (the tab) containing the data to get.</param>
		/// <returns>Returns a List of T objects found in the specified spreadsheet.</returns>
		public static List<T> Get<T>(string documentName, string sheetName) where T : new()
		{
			Track(documentName, sheetName);
			List<T> result = new List<T>();
			try
			{
				IList<IList<object>> allCells = GetCells(documentName, sheetName);
				Dictionary<int, string> headers = new Dictionary<int, string>();
				IList<object> headerRow = allCells[0];
				for (int i = 0; i < headerRow.Count; i++)
				{
					headers.Add(i, (string)headerRow[i]);
				}
				for (int row = 1; row < allCells.Count; row++)
				{
					result.Add(NewItem<T>(headers, allCells[row]));
				}
			}
			catch (Exception ex)
			{
				string title = $"{ex.GetType()} reading google sheet ({documentName} - {sheetName}): ";
				while (ex != null)
				{
					Console.WriteLine($"{title} \"{ex.Message}\"");
					ex = ex.InnerException;
					if (ex != null)
						title = $"Inner {ex.GetType()}: ";
				}
			}
			return result;
		}

		/// <summary>
		/// Saves changes in the specified instances to the specified document and sheet. If instances implements 
		/// <cref>ITrackPropertyChanges</cref>, then only changed properties will be saved.
		/// </summary>
		/// <param name="docName">The name of the spreadsheet document to save to (registered with a call to RegisterDocumentID).</param>
		/// <param name="sheetName">The name of the sheet (the tab in the document) to receive the saved data.</param>
		/// <param name="instances">The instances to save.</param>
		/// <param name="instanceType">The type of the instances to save.</param>
		/// <param name="saveOnlyTheseMembersStr">An optional comma-separated list of the names of the member properties to save.</param>
		public static void SaveChanges(string docName, string sheetName, object[] instances, Type instanceType, string saveOnlyTheseMembersStr = null)
		{
			if (instances == null || instances.Length == 0)
				return;

			ValidateDocumentAndSheetNames(docName, sheetName, true);

			string[] saveOnlyTheseMembers = null;
			if (saveOnlyTheseMembersStr != null)
				saveOnlyTheseMembers = saveOnlyTheseMembersStr.Split(',');

			string documentID = documentIDs[docName];
			List<string> headerRow;
			IList<IList<object>> allRows;
			GetHeaderRow(docName, sheetName, out headerRow, out allRows);

			MemberInfo[] indexFields = GetSerializableFields<IndexerAttribute>(instanceType);

			MemberInfo[] serializableFields = GetSerializableFields<ColumnAttribute>(instanceType);

			BatchUpdateValuesRequest requestBody = GetBatchUpdateRequest();
			for (int i = 0; i < instances.Length; i++)
			{
				int rowIndex = GetInstanceRowIndex(instances[i], indexFields, allRows, headerRow);
				if (rowIndex >= 0)
					UpdateRow(sheetName, instances, saveOnlyTheseMembers, headerRow, allRows, serializableFields, i, rowIndex, requestBody);
				else
					AddRow(sheetName, serializableFields, headerRow, serializableFields, allRows, instances[i], requestBody);
			}

			ExecuteBatchUpdate(documentID, requestBody);
		}

		/// <summary>
		/// Saves changes to the specified instances, writing values out to the associated spreadsheet (determined by the DocumentName and SheetName attributes on the instance type)
		/// </summary>
		/// <param name="instances">The instances to save. All instances in this list should be of the same type.</param>
		/// <param name="saveOnlyTheseMembersStr">An optional comma-separated list of the names of the member properties to save.</param>
		public static void SaveChanges(object[] instances, string saveOnlyTheseMembersStr = null)
		{
			if (instances == null || instances.Length == 0)
				return;
			Type instanceType = instances[0].GetType();

			GetDocumentAndSheetAttributes(instanceType, out DocumentNameAttribute documentNameAttribute, out SheetNameAttribute sheetNameAttribute);

			SaveChanges(documentNameAttribute.DocumentName, sheetNameAttribute.SheetName, instances, instanceType, saveOnlyTheseMembersStr);
		}

		/// <summary>
		/// Saves changes to the specified instance, writing values out to the associated spreadsheet (specified with the DocumentName and SheetName attributes)
		/// </summary>
		/// <param name="instance">The instance to save.</param>
		/// <param name="saveOnlyTheseMembersStr">An optional comma-separated list of the names of the member properties to save.</param>
		public static void SaveChanges(object instance, string saveOnlyTheseMembersStr = null)
		{
			object[] array = { instance };
			SaveChanges(array, saveOnlyTheseMembersStr);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="instance"></param>
		/// <param name="documentName">The name of the document to work with (must be registered with a call to 
		/// RegisterDocumentID).</param>
		/// <param name="sheetName"></param>
		static void DeleteRowInSheet(object instance, string documentName, string sheetName)
		{
			Type instanceType = instance.GetType();
			ValidateDocumentAndSheetNames(documentName, sheetName);
			string documentId = documentIDs[documentName];
			List<string> headerRow;
			IList<IList<object>> allRows;
			GetHeaderRow(documentName, sheetName, out headerRow, out allRows);

			MemberInfo[] indexFields = GetSerializableFields<IndexerAttribute>(instanceType);

			int rowIndex = GetInstanceRowIndex(instance, indexFields, allRows, headerRow);
			DeleteRowByIndexInSheet(documentId, sheetName, rowIndex);
		}

		public static void DeleteRow(object instance)
		{
			GetDocumentAndSheetAttributes(instance.GetType(), out DocumentNameAttribute documentNameAttribute, out SheetNameAttribute sheetNameAttribute);
			DeleteRowInSheet(instance, documentNameAttribute.DocumentName, sheetNameAttribute.SheetName);
		}

		public static void MakeSureSheetExists<T>(string sheetName)
		{
			if (!SheetExists<T>(sheetName))
				AddSheet<T>(sheetName);
		}

		public static bool SheetExists<T>(string sheetName)
		{
			return SheetExists(GetDocumentName<T>(), sheetName);
		}

		/// <summary>
		/// Returns true if the specified sheet exists on the specified document.
		/// </summary>
		/// <param name="documentName">The name of the document (must be registered with a call to 
		/// RegisterDocumentID).</param>
		/// <param name="sheetName">The name of the sheet (tab) to check.</param>
		/// <returns></returns>
		public static bool SheetExists(string documentName, string sheetName)
		{
			ValidateDocumentName(documentName);
			return GetSheetId(documentIDs[documentName], sheetName) != null;
		}

		public static void HexToColor(string hexStr, out float red, out float green, out float blue)
		{
			if (hexStr.IndexOf('#') != -1)
				hexStr = hexStr.Replace("#", "");

			red = int.Parse(hexStr.Substring(0, 2), NumberStyles.AllowHexSpecifier) / 255.0f;
			green = int.Parse(hexStr.Substring(2, 2), NumberStyles.AllowHexSpecifier) / 255.0f;
			blue = int.Parse(hexStr.Substring(4, 2), NumberStyles.AllowHexSpecifier) / 255.0f;
		}

		static void MakeSureTextFormatExists(CellData cellData)
		{
			if (cellData.UserEnteredFormat == null)
				cellData.UserEnteredFormat = new CellFormat();

			if (cellData.UserEnteredFormat.TextFormat == null)
				cellData.UserEnteredFormat.TextFormat = new TextFormat();
		}

		static void AddHeaderRowFormatting<T>(IList<Request> requests, string documentId, string sheetName)
		{
			HeaderRowAttribute headerRowAttribute = typeof(T).GetCustomAttribute<HeaderRowAttribute>();
			if (headerRowAttribute == null)
				return;

			CellData cellData = new CellData();

			string userEnteredFormatField = null;

			if (!string.IsNullOrWhiteSpace(headerRowAttribute.Color))
			{
				HexToColor(headerRowAttribute.Color, out float red, out float green, out float blue);
				MakeSureTextFormatExists(cellData);
				cellData.UserEnteredFormat.TextFormat.ForegroundColor = new Color();
				cellData.UserEnteredFormat.TextFormat.ForegroundColor.Red = red;
				cellData.UserEnteredFormat.TextFormat.ForegroundColor.Green = green;
				cellData.UserEnteredFormat.TextFormat.ForegroundColor.Blue = blue;
				cellData.UserEnteredFormat.TextFormat.ForegroundColor.Alpha = 1.0f;
				userEnteredFormatField = "userEnteredFormat.textFormat";
			}

			if (headerRowAttribute.FontWeight == FontWeight.Bold)
			{
				MakeSureTextFormatExists(cellData);
				cellData.UserEnteredFormat.TextFormat.Bold = true;
				userEnteredFormatField = "userEnteredFormat.textFormat";
			}
			

			Request headerRowRequest = GetRepeatCellRequestForTopRow(documentId, sheetName);

			if (headerRowAttribute.RowFreezeOption == RowFreezeOption.FreezeTopRow)
			{
				headerRowRequest.UpdateSheetProperties = new UpdateSheetPropertiesRequest();
				headerRowRequest.UpdateSheetProperties.Properties = new SheetProperties();
				headerRowRequest.UpdateSheetProperties.Properties.GridProperties = new GridProperties();
				headerRowRequest.UpdateSheetProperties.Properties.GridProperties.FrozenRowCount = 1;
				if (string.IsNullOrEmpty(userEnteredFormatField))
					userEnteredFormatField = STR_FrozenRowCount;
				else
					userEnteredFormatField += ";" + STR_FrozenRowCount;
			}
			
			headerRowRequest.RepeatCell.Cell = cellData;
			headerRowRequest.RepeatCell.Cell.UserEnteredFormat = new CellFormat();
			headerRowRequest.RepeatCell.Fields = userEnteredFormatField;

			requests.Add(headerRowRequest);
		}

		/// <summary>
		/// Adds a new specified sheet (tab) to the spreadsheet document associated with this element.
		/// </summary>
		/// <typeparam name="T">The type associated with the document to receive the new sheet. (must be registered with a call to 
		/// RegisterDocumentID).</typeparam>
		/// <param name="sheetName">The name of the sheet (the tab) to create.</param>
		public static void AddSheet<T>(string sheetName)
		{
			string documentName = GetDocumentName<T>();
			AddSheetRequest addSheetRequest = new AddSheetRequest();
			SheetProperties sheetProperties = new SheetProperties() { Title = sheetName };

			addSheetRequest.Properties = sheetProperties;

			BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();

			//Create requestList and set it on the batchUpdateSpreadsheetRequest
			List<Request> requestsList = new List<Request>();
			batchUpdateSpreadsheetRequest.Requests = requestsList;

			//Create a new request with containing the addSheetRequest and add it to the requestList
			Request request = new Request();
			request.AddSheet = addSheetRequest;
			requestsList.Add(request);

			//Add the requestList to the batchUpdateSpreadsheetRequest
			batchUpdateSpreadsheetRequest.Requests = requestsList;

			//Call the sheets API to execute the batchUpdate
			string documentId = documentIDs[documentName];
			Service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, documentId).Execute();

			// TODO: To optimize for performance, consider rolling this next ExecuteAppendRow call into the previous batch update.

			List<object> columns = new List<object>();
			MemberInfo[] serializableFields = GetSerializableFields<ColumnAttribute>(typeof(T));
			foreach (MemberInfo memberInfo in serializableFields)
			{
				ColumnAttribute columnAttribute = memberInfo.GetCustomAttribute<ColumnAttribute>();
				if (columnAttribute != null && !string.IsNullOrEmpty(columnAttribute.ColumnName))
					columns.Add(columnAttribute.ColumnName);
				else
					columns.Add(memberInfo.Name);
			}

			List<IList<object>> rows = new List<IList<object>>();
			rows.Add(columns);

			ExecuteAppendRows(documentId, sheetName, rows);

			IList<Request> requests = new List<Request>();
			AddColumnNotes(requests, documentId, sheetName, serializableFields);
			AddFormatting(requests, documentId, sheetName, serializableFields);
			AddHeaderRowFormatting<T>(requests, documentId, sheetName);

			ExecuteRequests(documentId, requests);
		}

		/// <summary>
		/// Appends a row representing the specified instance, writing its field values out to the spreadsheet. Make 
		/// sure your instance class has the [<cref>DocumentName</cref>] attribute (and that spreadsheet document has 
		/// been registered with <cref>RegisterSpreadsheetID</cref>), and add the [Column] attribute to any members that 
		/// you want to write out to the spreadsheet. Messages sent here are throttled so as not to exceed the Google 
		/// sheets per-minute usage limits. The time to wait between message bursts is determined by the 
		/// <cref>TimeBetweenThrottledUpdates</cref> property.
		/// </summary>
		/// <typeparam name="T">The type to append</typeparam>
		/// <param name="instance">The instance of the type containing the writable data (in its public fields and properties marked with the [Column] attribute).</param>
		/// <param name="sheetNameOverride">An optional override for the sheet (tab) name.</param>
		public static void AppendRow<T>(T instance, string sheetNameOverride = null) where T : class
		{
			if (!messageThrottlers.ContainsKey(typeof(T)))
				messageThrottlers[typeof(T)] = new MessageThrottler<T>(TimeBetweenThrottledUpdates);

			if (messageThrottlers[typeof(T)] is MessageThrottler<T> throttler)
				throttler.AppendRow(instance, sheetNameOverride);
		}

		/// <summary>
		/// Sends any throttled messages at once to GoogleSheets. Call when shutting down to make sure any messages in the 
		/// queue are written to the spreadsheet.
		/// </summary>
		public static void FlushAllMessages()
		{
			foreach (Type type in messageThrottlers.Keys)
			{
				Type throttlerType = messageThrottlers[type].GetType();
				MethodInfo methodDefinition = throttlerType.GetMethod(nameof(MessageThrottler<object>.FlushAllMessages), new Type[] { });
				methodDefinition.Invoke(messageThrottlers[type], new object[] { });
			}
		}

		public static TimeSpan TimeBetweenThrottledUpdates { get; set; } = TimeSpan.FromSeconds(5);
		public static SheetsService Service
		{
			get
			{
				if (service == null)
				{
					InitializeService(GetUserCredentials());
				}
				return service;
			}
		}
	}
}

