using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SheetsPersist
{
	public static partial class GoogleSheets
	{
		static string GetValue(object obj, MemberInfo memberInfo)
		{
			if (memberInfo is PropertyInfo propInfo)
			{
				object value = propInfo.GetValue(obj);
				if (value == null)
					return null;
				return value.ToString();
			}

			if (memberInfo is FieldInfo fieldInfo)
				return fieldInfo.GetValue(obj).ToString();
			return null;
		}

		static MemberInfo[] GetSerializableFields<TAttribute>(Type instanceType) where TAttribute : Attribute
		{
			List<MemberInfo> memberInfo = new List<MemberInfo>();
			IEnumerable<PropertyInfo> properties = instanceType.GetProperties(BindingFlags.Public | BindingFlags.Instance).Where(x => x.GetCustomAttribute<TAttribute>() != null);
			IEnumerable<FieldInfo> fields = instanceType.GetFields(BindingFlags.Public | BindingFlags.Instance).Where(x => x.GetCustomAttribute<TAttribute>() != null);
			memberInfo.AddRange(properties);
			memberInfo.AddRange(fields);
			return memberInfo.ToArray();
		}

		static void SetEnumValue(PropertyInfo enumProperty, object instance, string valueStr)
		{
			int value = 0;
			if (!string.IsNullOrWhiteSpace(valueStr))
				try
				{
					string[] parts = valueStr.Split('|');
					foreach (string part in parts)
					{
						value += (int)Enum.Parse(enumProperty.PropertyType, part.Trim());
					}
				}
				// TODO: Consider re-throwing the error before publishing.
#pragma warning disable CS0168  // Used for diagnostics/debugging.
				catch (Exception ex)
				{
					System.Diagnostics.Debugger.Break();
					return;
				}
			enumProperty.SetValue(instance, value);
		}

		static void SetValue(PropertyInfo property, object instance, string value)
		{
			if (property.PropertyType.IsEnum)
				SetEnumValue(property, instance, value);
			else
				System.Diagnostics.Debugger.Break();
		}

		static PropertyInfo GetCorrespondingPropertyInfo(Type type, string header)
		{
			PropertyInfo[] properties = type.GetProperties();
			foreach (PropertyInfo propertyInfo in properties)
			{
				ColumnAttribute columnAttribute = propertyInfo.GetCustomAttribute<ColumnAttribute>();
				if (columnAttribute != null && columnAttribute.ColumnName == header)
					return propertyInfo;

			}
			return type.GetProperty(header);
		}

		static void TransferValues<T>(T instance, Dictionary<int, string> headers, IList<object> row) where T : new()
		{
			Type type = instance.GetType();

			for (int i = 0; i < row.Count; i++)  // There may be fewer rows than headers.
			{
				//! Not using ColumName here!!!
				PropertyInfo property = GetCorrespondingPropertyInfo(type, headers[i]);
				if (property == null)
					continue;

				string fullName = property.PropertyType.FullName;
				string value = (string)row[i];
				switch (fullName)
				{
					case "System.Int32":
						if (!int.TryParse(value, out int intValue))
							intValue = 0;
						property.SetValue(instance, intValue);
						break;
					case "System.Decimal":
						if (decimal.TryParse(value, out decimal decimalValue))
							property.SetValue(instance, decimalValue);
						else
						{
							System.Diagnostics.Debugger.Break();
						}
						break;
					case "System.Double":
						if (double.TryParse(value, out double doubleValue))
							property.SetValue(instance, doubleValue);
						else
						{
							// TODO: Consider specifying default values through attributes for given properties.
							property.SetValue(instance, 0);
						}
						break;
					case "System.String":
						property.SetValue(instance, row[i]);
						break;
					case "System.Boolean":
						string compareValue = value.ToLower().Trim();
						bool newValue = compareValue == "true" || compareValue == "x";
						property.SetValue(instance, newValue);
						break;
					case "System.DateTime":
						if (DateTime.TryParse(value, out DateTime date))
							property.SetValue(instance, date);
						break;
					default:
						SetValue(property, instance, value);
						break;
				}
			}

			SetDefaultsForEmptyCells(instance, headers, row, type);
		}

		static object GetDefaultValue(PropertyInfo property)
		{
			IEnumerable<DefaultAttribute> customAttributes = property.GetCustomAttributes<DefaultAttribute>();

			DefaultAttribute defaultAttribute = null;
			if (customAttributes.Any())
				defaultAttribute = customAttributes.First();
			if (defaultAttribute != null)
				return defaultAttribute.DefaultValue;

			switch (property.PropertyType.FullName)
			{
				case "System.Int32":
					return default(int);
				case "System.String":
					return string.Empty;
				case "System.Boolean":
					return default(bool);
				case "System.Decimal":
					return default(decimal);
				case "System.Double":
					return default(double);
				case "System.DateTime":
					return default(DateTime);
				default:
					if (property.PropertyType.BaseType.FullName == "System.Enum")
						return 0;
					else
						System.Diagnostics.Debugger.Break();
					break;
			}
			return null;
		}
		private static void SetDefaultsForEmptyCells<T>(T instance, Dictionary<int, string> headers, IList<object> row, Type type) where T : new()
		{
			for (int i = row.Count; i < headers.Count; i++)  // There may be fewer rows than headers.
			{
				PropertyInfo property = type.GetProperty(headers[i]);
				if (property == null)
					continue;

				object defaultValue = GetDefaultValue(property);
				property.SetValue(instance, defaultValue);
			}
		}
	}
}

