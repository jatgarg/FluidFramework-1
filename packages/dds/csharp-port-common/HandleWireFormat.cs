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

		/// <summary>
		/// Name of the optional <c>payloadPending</c> property on <c>ISerializedHandle</c>.
		/// TS writes it only when <c>true</c>; omitted otherwise (matches
		/// <c>encodeHandleForSerialization</c> in <c>runtime-utils/src/handles.ts</c>).
		/// </summary>
		public const string PayloadPendingPropertyName = "payloadPending";

		public static bool IsHandleShape(JsonElement element)
		{
			return TryReadHandleUrl(element, out _, out _);
		}

		public static object ReadHandleFromShape(JsonElement element, IFluidDataObjectRegistry? registry)
		{
			if (!TryReadHandleUrl(element, out string url, out bool payloadPending))
			{
				throw new JsonException("Expected Fluid handle JSON object.");
			}

			return ResolveSerializedHandle(url, registry, payloadPending);
		}

		public static void WriteHandleShape(Utf8JsonWriter writer, string url)
		{
			WriteHandleShape(writer, url, payloadPending: false);
		}

		public static void WriteHandleShape(Utf8JsonWriter writer, string url, bool payloadPending)
		{
			ArgumentNullException.ThrowIfNull(writer);
			ArgumentNullException.ThrowIfNull(url);

			writer.WriteStartObject();
			writer.WriteString(TypePropertyName, SerializedHandleTypeName);
			writer.WriteString(UrlPropertyName, url);
			if (payloadPending)
			{
				writer.WriteBoolean(PayloadPendingPropertyName, true);
			}

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

			if (TryGetHandleUrl(value, registry, missingRegistryMessage, out string handleUrl, out bool payloadPending))
			{
				return CreateSerializedHandleWireValue(handleUrl, payloadPending);
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
				return ResolveSerializedHandle(serializedHandle.Url, registry, serializedHandle.PayloadPending);
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
			return TryReadHandleUrl(element, out url, out _);
		}

		public static bool TryReadHandleUrl(JsonElement element, out string url, out bool payloadPending)
		{
			if (element.ValueKind == JsonValueKind.Object
				&& element.TryGetProperty(TypePropertyName, out JsonElement typeElement)
				&& typeElement.ValueKind == JsonValueKind.String
				&& typeElement.GetString() == SerializedHandleTypeName
				&& element.TryGetProperty(UrlPropertyName, out JsonElement urlElement)
				&& urlElement.ValueKind == JsonValueKind.String)
			{
				url = urlElement.GetString() ?? string.Empty;
				payloadPending = element.TryGetProperty(PayloadPendingPropertyName, out JsonElement payloadPendingElement)
					&& payloadPendingElement.ValueKind == JsonValueKind.True;
				return true;
			}

			url = string.Empty;
			payloadPending = false;
			return false;
		}

		public static object ResolveSerializedHandle(string url, IFluidDataObjectRegistry? registry)
		{
			return ResolveSerializedHandle(url, registry, payloadPending: false);
		}

		public static object ResolveSerializedHandle(string url, IFluidDataObjectRegistry? registry, bool payloadPending)
		{
			// TS ref: serializer.ts constructs RemoteFluidObjectHandle with
			// `value.payloadPending === true` — the flag is preserved on the resolved
			// handle. A live IFluidDataObject has no place to carry payloadPending,
			// so only resolve non-pending handles to a live object. Pending handles
			// stay as SerializedFluidHandle so the flag round-trips correctly. This
			// also matches the semantic of a pending handle: the payload is not
			// ready yet, so dereferencing to a live object would be premature.
			if (!payloadPending && registry != null)
			{
				IFluidDataObject? dataObject = registry.FindDataObject(url);
				if (dataObject != null)
				{
					return dataObject;
				}
			}

			return new SerializedFluidHandle(url, payloadPending);
		}

		public static bool TryGetHandleUrl(
			object value,
			IFluidDataObjectRegistry? registry,
			string missingRegistryMessage,
			out string url)
		{
			return TryGetHandleUrl(value, registry, missingRegistryMessage, out url, out _);
		}

		public static bool TryGetHandleUrl(
			object value,
			IFluidDataObjectRegistry? registry,
			string missingRegistryMessage,
			out string url,
			out bool payloadPending)
		{
			if (value is SerializedFluidHandle serializedHandle)
			{
				url = serializedHandle.Url;
				payloadPending = serializedHandle.PayloadPending;
				return true;
			}

			if (value is IFluidDataObject dataObject)
			{
				if (registry != null)
				{
					url = registry.GetDataObjectUrl(dataObject);
					// Live registry-resolved handles have their payload already; TS omits
					// payloadPending in encodeHandleForSerialization's non-pending branch.
					payloadPending = false;
					return true;
				}

				throw new OcsException(OcsGateErrorCode.InvalidOperation, missingRegistryMessage);
			}

			url = string.Empty;
			payloadPending = false;
			return false;
		}

		public static Dictionary<string, object?> CreateSerializedHandleWireValue(string url)
		{
			return CreateSerializedHandleWireValue(url, payloadPending: false);
		}

		public static Dictionary<string, object?> CreateSerializedHandleWireValue(string url, bool payloadPending)
		{
			Dictionary<string, object?> wireValue = new Dictionary<string, object?>()
			{
				[TypePropertyName] = SerializedHandleTypeName,
				[UrlPropertyName] = url,
			};

			if (payloadPending)
			{
				wireValue[PayloadPendingPropertyName] = true;
			}

			return wireValue;
		}

		private static object? ConvertJsonElementWithHandles(JsonElement element, IFluidDataObjectRegistry? registry, out bool changed)
		{
			switch (element.ValueKind)
			{
				case JsonValueKind.Object:
					if (TryReadHandleUrl(element, out string url, out bool payloadPending))
					{
						changed = true;
						return ResolveSerializedHandle(url, registry, payloadPending);
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
