using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace SheetsPersist
{
	public static partial class GoogleSheets
	{
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
			AddHeaderRowFormatting(requests, documentId, sheetName, headerRowAttribute.Color, headerRowAttribute.FontWeight);
		}

		static void AddCellFormatting(IList<Request> requests, string documentId, string sheetName, CellPosition cellPosition, string color, FontWeight fontWeight)
		{
			GetCellData(color, fontWeight, out CellData cellData, out string userEnteredFormatField);
			if (cellData == null)
				return;

			Request headerRowRequest = GetRepeatCellRequest(documentId, sheetName, cellPosition);
			headerRowRequest.RepeatCell.Cell = cellData;
			headerRowRequest.RepeatCell.Fields = userEnteredFormatField;

			requests.Add(headerRowRequest);
		}

		static void AddHeaderColumnFormatting(IList<Request> requests, string documentId, string sheetName, MemberInfo[] serializableFields)
		{
			const int rowAfterHeader = 1;
			int cellColumn = 0;
			foreach (MemberInfo memberInfo in serializableFields)
			{
				StyleAttribute styleAttribute = memberInfo.GetCustomAttribute<StyleAttribute>();
				if (styleAttribute != null)
					AddCellFormatting(requests, documentId, sheetName, new CellPosition(cellColumn, rowAfterHeader), styleAttribute.Color, styleAttribute.FontWeight);
				cellColumn++;
			}
		}

		static void AddHeaderRowFormatting(string documentName, string sheetName, string color, FontWeight fontWeight)
		{
			string documentId = documentIDs[documentName];
			IList<Request> requests = new List<Request>();
			AddHeaderRowFormatting(requests, documentId, sheetName, color, fontWeight);
			ExecuteRequests(documentId, requests);
		}

		static void AddFormatting(IList<Request> requests, int columnIndex, string pattern, string type, string documentId, string sheetName)
		{
			Request formatRequest = GetRepeatCellRequestForEntireColumn(columnIndex, documentId, sheetName);
			formatRequest.RepeatCell.Cell = new CellData();
			formatRequest.RepeatCell.Cell.UserEnteredFormat = new CellFormat();
			formatRequest.RepeatCell.Cell.UserEnteredFormat.NumberFormat = new NumberFormat() { Pattern = pattern, Type = type };
			formatRequest.RepeatCell.Fields = "userEnteredFormat.numberFormat";

			requests.Add(formatRequest);
		}

		private static void AddFormatting(IList<Request> requests, string documentId, string sheetName, MemberInfo[] serializableFields)
		{
			int columnIndex = 0;

			foreach (MemberInfo memberInfo in serializableFields)
			{
				FormatNumberAttribute formatNumberAttribute = memberInfo.GetCustomAttribute<FormatNumberAttribute>();
				if (formatNumberAttribute != null && !string.IsNullOrEmpty(formatNumberAttribute.Pattern))
					AddFormatting(requests, columnIndex, formatNumberAttribute.Pattern, "NUMBER", documentId, sheetName);

				FormatDateAttribute formatDateAttribute = memberInfo.GetCustomAttribute<FormatDateAttribute>();
				if (formatDateAttribute != null && !string.IsNullOrEmpty(formatDateAttribute.Pattern))
					AddFormatting(requests, columnIndex, formatDateAttribute.Pattern, "DATE", documentId, sheetName);

				FormatCurrencyAttribute formatCurrencyAttribute = memberInfo.GetCustomAttribute<FormatCurrencyAttribute>();
				if (formatCurrencyAttribute != null && !string.IsNullOrEmpty(formatCurrencyAttribute.Pattern))
					AddFormatting(requests, columnIndex, formatCurrencyAttribute.Pattern, "CURRENCY", documentId, sheetName);

				columnIndex++;
			}
		}
	}
}

