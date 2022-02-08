using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SheetsPersist
{
	public static partial class GoogleSheets
	{
		private static void AddFormatting(IList<Request> requests, int columnIndex, string pattern, string type, string documentId, string tabName)
		{
			Request formatRequest = GetRepeatCellRequestForEntireColumn(columnIndex, documentId, tabName);
			// TODO: Figure out how to indicate the entire column
			formatRequest.RepeatCell.Cell = new CellData();
			formatRequest.RepeatCell.Cell.UserEnteredFormat = new CellFormat();
			formatRequest.RepeatCell.Cell.UserEnteredFormat.NumberFormat = new NumberFormat() { Pattern = pattern, Type = type };
			formatRequest.RepeatCell.Fields = "userEnteredFormat.numberFormat";

			requests.Add(formatRequest);
		}

		private static void AddFormatting(IList<Request> requests, string documentId, string tabName, MemberInfo[] serializableFields)
		{
			int columnIndex = 0;

			foreach (MemberInfo memberInfo in serializableFields)
			{
				FormatNumberAttribute formatNumberAttribute = memberInfo.GetCustomAttribute<FormatNumberAttribute>();
				if (formatNumberAttribute != null && !string.IsNullOrEmpty(formatNumberAttribute.Pattern))
					AddFormatting(requests, columnIndex, formatNumberAttribute.Pattern, "NUMBER", documentId, tabName);

				FormatDateAttribute formatDateAttribute = memberInfo.GetCustomAttribute<FormatDateAttribute>();
				if (formatDateAttribute != null && !string.IsNullOrEmpty(formatDateAttribute.Pattern))
					AddFormatting(requests, columnIndex, formatDateAttribute.Pattern, "DATE", documentId, tabName);

				FormatCurrencyAttribute formatCurrencyAttribute = memberInfo.GetCustomAttribute<FormatCurrencyAttribute>();
				if (formatCurrencyAttribute != null && !string.IsNullOrEmpty(formatCurrencyAttribute.Pattern))
					AddFormatting(requests, columnIndex, formatCurrencyAttribute.Pattern, "CURRENCY", documentId, tabName);

				columnIndex++;
			}
		}
	}
}

