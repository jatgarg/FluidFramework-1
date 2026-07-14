// -----------------------------------------------------------------------------
// Helpers for Fluid handle JSON wire values:
//   {"type":"__fluid_handle__","url":"..."}
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.Json;

namespace Microsoft.Office.Web.Fluid
{
	public static class HandleWireFormat
	{
		public const string TypePropertyName = "type";
		public const string SerializedHandleTypeName = "__fluid_handle__";
		public const string UrlPropertyName = "url";

		public static bool IsHandleShape(JsonElement element)
		{
			return TryReadHandleUrl(element, out _);
		}

		public static object ReadHandleFromShape(JsonElement element, IFluidDataObjectRegistry? registry)
		{
			if (!TryReadHandleUrl(element, out string url))
			{
				throw new JsonException("Expected Fluid handle JSON object.");
			}

			return ResolveSerializedHandle(url, registry);
		}

		public static void WriteHandleShape(Utf8JsonWriter writer, string url)
		{
			ArgumentNullException.ThrowIfNull(writer);
			ArgumentNullException.ThrowIfNull(url);

			writer.WriteStartObject();
			writer.WriteString(TypePropertyName, SerializedHandleTypeName);
			writer.WriteString(UrlPropertyName, url);
			writer.WriteEndObject();
		}

		public static object? MakeHandlesSerializable(
			object? value,
			IFluidDataObjectRegistry? registry,
			string missingRegistryMessage)
		{
			if (value == null)
			{
				return null;
			}

			if (TryGetHandleUrl(value, registry, missingRegistryMessage, out string handleUrl))
			{
				return CreateSerializedHandleWireValue(handleUrl);
			}

			if (value is JsonElement)
			{
				return value;
			}

			if (value is IReadOnlyDictionary<string, object?> readOnlyDictionary)
			{
				Dictionary<string, object?> serializedDictionary = new Dictionary<string, object?>();
				foreach (KeyValuePair<string, object?> entry in readOnlyDictionary)
				{
					serializedDictionary[entry.Key] = MakeHandlesSerializable(entry.Value, registry, missingRegistryMessage);
				}

				return serializedDictionary;
			}

			if (value is IDictionary dictionary)
			{
				Dictionary<string, object?> serializedDictionary = new Dictionary<string, object?>();
				foreach (DictionaryEntry entry in dictionary)
				{
					if (entry.Key is string key)
					{
						serializedDictionary[key] = MakeHandlesSerializable(entry.Value, registry, missingRegistryMessage);
					}
				}

				return serializedDictionary;
			}

			if (value is IEnumerable enumerable && value is not string)
			{
				List<object?> serializedList = new List<object?>();
				foreach (object? item in enumerable)
				{
					serializedList.Add(MakeHandlesSerializable(item, registry, missingRegistryMessage));
				}

				return serializedList;
			}

			return value;
		}

		public static object? ReadJsonValueWithHandles(JsonElement element, IFluidDataObjectRegistry? registry)
		{
			return ConvertJsonElementWithHandles(element, registry, out bool _);
		}

		public static object? ResolveSerializedHandles(object? value, IFluidDataObjectRegistry? registry)
		{
			if (value == null)
			{
				return null;
			}

			if (value is SerializedFluidHandle serializedHandle)
			{
				return ResolveSerializedHandle(serializedHandle.Url, registry);
			}

			if (value is JsonElement element)
			{
				return ReadJsonValueWithHandles(element, registry);
			}

			if (value is IReadOnlyDictionary<string, object?> readOnlyDictionary)
			{
				Dictionary<string, object?> resolvedDictionary = new Dictionary<string, object?>();
				foreach (KeyValuePair<string, object?> entry in readOnlyDictionary)
				{
					resolvedDictionary[entry.Key] = ResolveSerializedHandles(entry.Value, registry);
				}

				return resolvedDictionary;
			}

			if (value is IDictionary dictionary)
			{
				Dictionary<string, object?> resolvedDictionary = new Dictionary<string, object?>();
				foreach (DictionaryEntry entry in dictionary)
				{
					if (entry.Key is string key)
					{
						resolvedDictionary[key] = ResolveSerializedHandles(entry.Value, registry);
					}
				}

				return resolvedDictionary;
			}

			if (value is IEnumerable enumerable && value is not string)
			{
				List<object?> resolvedList = new List<object?>();
				foreach (object? item in enumerable)
				{
					resolvedList.Add(ResolveSerializedHandles(item, registry));
				}

				return resolvedList;
			}

			return value;
		}

		public static bool TryReadHandleUrl(JsonElement element, out string url)
		{
			if (element.ValueKind == JsonValueKind.Object
				&& element.TryGetProperty(TypePropertyName, out JsonElement typeElement)
				&& typeElement.ValueKind == JsonValueKind.String
				&& typeElement.GetString() == SerializedHandleTypeName
				&& element.TryGetProperty(UrlPropertyName, out JsonElement urlElement)
				&& urlElement.ValueKind == JsonValueKind.String)
			{
				url = urlElement.GetString() ?? string.Empty;
				return true;
			}

			url = string.Empty;
			return false;
		}

		public static object ResolveSerializedHandle(string url, IFluidDataObjectRegistry? registry)
		{
			if (registry != null)
			{
				IFluidDataObject? dataObject = registry.FindDataObject(url);
				if (dataObject != null)
				{
					return dataObject;
				}
			}

			return new SerializedFluidHandle(url);
		}

		public static bool TryGetHandleUrl(
			object value,
			IFluidDataObjectRegistry? registry,
			string missingRegistryMessage,
			out string url)
		{
			if (value is SerializedFluidHandle serializedHandle)
			{
				url = serializedHandle.Url;
				return true;
			}

			if (value is IFluidDataObject dataObject)
			{
				if (registry != null)
				{
					url = registry.GetDataObjectUrl(dataObject);
					return true;
				}

				throw new OcsException(OcsGateErrorCode.InvalidOperation, missingRegistryMessage);
			}

			url = string.Empty;
			return false;
		}

		public static Dictionary<string, object?> CreateSerializedHandleWireValue(string url)
		{
			return new Dictionary<string, object?>()
			{
				[TypePropertyName] = SerializedHandleTypeName,
				[UrlPropertyName] = url,
			};
		}

		private static object? ConvertJsonElementWithHandles(JsonElement element, IFluidDataObjectRegistry? registry, out bool changed)
		{
			switch (element.ValueKind)
			{
				case JsonValueKind.Object:
					if (TryReadHandleUrl(element, out string url))
					{
						changed = true;
						return ResolveSerializedHandle(url, registry);
					}

					bool objectChanged = false;
					Dictionary<string, object?> objectValue = new Dictionary<string, object?>();
					foreach (JsonProperty property in element.EnumerateObject())
					{
						object? childValue = ConvertJsonElementWithHandles(property.Value, registry, out bool childChanged);
						objectChanged |= childChanged;
						objectValue[property.Name] = childValue;
					}

					changed = objectChanged;
					return objectChanged ? objectValue : element.Clone();

				case JsonValueKind.Array:
					bool arrayChanged = false;
					List<object?> arrayValue = new List<object?>();
					foreach (JsonElement item in element.EnumerateArray())
					{
						object? childValue = ConvertJsonElementWithHandles(item, registry, out bool childChanged);
						arrayChanged |= childChanged;
						arrayValue.Add(childValue);
					}

					changed = arrayChanged;
					return arrayChanged ? arrayValue : element.Clone();

				case JsonValueKind.Null:
				case JsonValueKind.Undefined:
					changed = false;
					return null;

				default:
					changed = false;
					return element.Clone();
			}
		}
	}
}
