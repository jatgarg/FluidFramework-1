// -----------------------------------------------------------------------------
// System.Text.Json (de)serialization for DirectoryOperation and its subclasses.
// Wire format matches packages/dds/map/src/directory.ts (TS SharedDirectory).
// Part of the SharedDirectory C# feasibility port — Wave 2a.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace Microsoft.Office.Web.Fluid
{
	public static class DirectoryOpSerializer
	{
		private const string _typePropertyName = "type";
		private const string _pathPropertyName = "path";
		private const string _keyPropertyName = "key";
		private const string _valuePropertyName = "value";
		private const string _subdirNamePropertyName = "subdirName";

		private const string _setTypeName = "set";
		private const string _deleteTypeName = "delete";
		private const string _clearTypeName = "clear";
		private const string _createSubDirectoryTypeName = "createSubDirectory";
		private const string _deleteSubDirectoryTypeName = "deleteSubDirectory";
		private const string _serializedHandleTypeName = "__fluid_handle__";
		private const string _serializedHandleUrlPropertyName = "url";

		private static readonly JsonSerializerOptions _serializerOptions = new JsonSerializerOptions()
		{
			Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
		};

		private static readonly JsonWriterOptions _writerOptions = new JsonWriterOptions()
		{
			Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
		};

		public static string Serialize(DirectoryOperation op, IFluidDataObjectRegistry? registry = null)
		{
			ArgumentNullException.ThrowIfNull(op);

			using MemoryStream stream = new MemoryStream();
			using (Utf8JsonWriter writer = new Utf8JsonWriter(stream, _writerOptions))
			{
				WriteTo(writer, op, registry);
			}

			return Encoding.UTF8.GetString(stream.ToArray());
		}

		public static DirectoryOperation Deserialize(string json, IFluidDataObjectRegistry? registry = null)
		{
			ArgumentNullException.ThrowIfNull(json);

			return Deserialize(Encoding.UTF8.GetBytes(json), registry);
		}

		public static DirectoryOperation Deserialize(ReadOnlySpan<byte> jsonBytes, IFluidDataObjectRegistry? registry = null)
		{
			Utf8JsonReader reader = new Utf8JsonReader(jsonBytes);
			DirectoryOperation operation = ReadFrom(ref reader, registry);
			if (reader.Read())
			{
				throw new JsonException("Unexpected trailing JSON after directory operation.");
			}

			return operation;
		}

		public static void WriteTo(Utf8JsonWriter writer, DirectoryOperation op, IFluidDataObjectRegistry? registry = null)
		{
			ArgumentNullException.ThrowIfNull(writer);
			ArgumentNullException.ThrowIfNull(op);

			writer.WriteStartObject();

			switch (op)
			{
				case DirectorySetOperation setOperation:
					writer.WriteString(_typePropertyName, _setTypeName);
					writer.WriteString(_pathPropertyName, setOperation.Path ?? string.Empty);
					writer.WriteString(_keyPropertyName, setOperation.Key ?? string.Empty);
					writer.WritePropertyName(_valuePropertyName);
					WriteSerializableValue(writer, setOperation.Value, registry);
					break;

				case DirectoryDeleteOperation deleteOperation:
					writer.WriteString(_typePropertyName, _deleteTypeName);
					writer.WriteString(_pathPropertyName, deleteOperation.Path ?? string.Empty);
					writer.WriteString(_keyPropertyName, deleteOperation.Key ?? string.Empty);
					break;

				case DirectoryClearOperation clearOperation:
					writer.WriteString(_typePropertyName, _clearTypeName);
					writer.WriteString(_pathPropertyName, clearOperation.Path ?? string.Empty);
					break;

				case DirectoryCreateSubDirectoryOperation createSubDirectoryOperation:
					writer.WriteString(_typePropertyName, _createSubDirectoryTypeName);
					writer.WriteString(_pathPropertyName, createSubDirectoryOperation.Path ?? string.Empty);
					writer.WriteString(_subdirNamePropertyName, createSubDirectoryOperation.SubdirName ?? string.Empty);
					break;

				case DirectoryDeleteSubDirectoryOperation deleteSubDirectoryOperation:
					writer.WriteString(_typePropertyName, _deleteSubDirectoryTypeName);
					writer.WriteString(_pathPropertyName, deleteSubDirectoryOperation.Path ?? string.Empty);
					writer.WriteString(_subdirNamePropertyName, deleteSubDirectoryOperation.SubdirName ?? string.Empty);
					break;

				default:
					throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unknown directory op runtime type: {op.GetType().FullName}");
			}

			writer.WriteEndObject();
		}

		public static DirectoryOperation ReadFrom(ref Utf8JsonReader reader, IFluidDataObjectRegistry? registry = null)
		{
			if (reader.TokenType == JsonTokenType.None && !reader.Read())
			{
				throw new JsonException("Expected directory operation JSON object.");
			}

			if (reader.TokenType != JsonTokenType.StartObject)
			{
				throw new JsonException("Expected directory operation JSON object.");
			}

			string? typeString = null;
			string path = string.Empty;
			string key = string.Empty;
			string subdirName = string.Empty;
			SerializableValue? value = null;

			while (reader.Read())
			{
				if (reader.TokenType == JsonTokenType.EndObject)
				{
					return CreateOperation(typeString, path, key, value, subdirName);
				}

				if (reader.TokenType != JsonTokenType.PropertyName)
				{
					throw new JsonException("Expected directory operation property name.");
				}

				string propertyName = reader.GetString() ?? string.Empty;
				if (!reader.Read())
				{
					throw new JsonException($"Expected value for directory operation property '{propertyName}'.");
				}

				switch (propertyName)
				{
					case _typePropertyName:
						typeString = ReadString(ref reader, _typePropertyName);
						break;

					case _pathPropertyName:
						path = ReadString(ref reader, _pathPropertyName);
						break;

					case _keyPropertyName:
						key = ReadString(ref reader, _keyPropertyName);
						break;

					case _valuePropertyName:
						value = ReadSerializableValue(ref reader, registry);
						break;

					case _subdirNamePropertyName:
						subdirName = ReadString(ref reader, _subdirNamePropertyName);
						break;

					default:
						reader.Skip();
						break;
				}
			}

			throw new JsonException("Unexpected end of directory operation JSON object.");
		}

		private static DirectoryOperation CreateOperation(string? typeString, string path, string key, SerializableValue? value, string subdirName)
		{
			switch (typeString)
			{
				case _setTypeName:
					return new DirectorySetOperation()
					{
						Path = path,
						Key = key,
						Value = value ?? new SerializableValue(),
					};

				case _deleteTypeName:
					return new DirectoryDeleteOperation()
					{
						Path = path,
						Key = key,
					};

				case _clearTypeName:
					return new DirectoryClearOperation()
					{
						Path = path,
					};

				case _createSubDirectoryTypeName:
					return new DirectoryCreateSubDirectoryOperation()
					{
						Path = path,
						SubdirName = subdirName,
					};

				case _deleteSubDirectoryTypeName:
					return new DirectoryDeleteSubDirectoryOperation()
					{
						Path = path,
						SubdirName = subdirName,
					};

				default:
					throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unknown directory op type: {typeString}");
			}
		}

		private static void WriteSerializableValue(Utf8JsonWriter writer, SerializableValue value, IFluidDataObjectRegistry? registry)
		{
			ArgumentNullException.ThrowIfNull(value);

			writer.WriteStartObject();
			writer.WriteString(_typePropertyName, value.Type ?? string.Empty);
			writer.WritePropertyName(_valuePropertyName);
			JsonSerializer.Serialize(writer, MakeHandlesSerializable(value.Value, registry), _serializerOptions);
			writer.WriteEndObject();
		}

		private static SerializableValue ReadSerializableValue(ref Utf8JsonReader reader, IFluidDataObjectRegistry? registry)
		{
			if (reader.TokenType != JsonTokenType.StartObject)
			{
				throw new JsonException("Expected serializable value JSON object.");
			}

			string type = string.Empty;
			object? value = null;

			while (reader.Read())
			{
				if (reader.TokenType == JsonTokenType.EndObject)
				{
					return new SerializableValue()
					{
						Type = type,
						Value = value,
					};
				}

				if (reader.TokenType != JsonTokenType.PropertyName)
				{
					throw new JsonException("Expected serializable value property name.");
				}

				string propertyName = reader.GetString() ?? string.Empty;
				if (!reader.Read())
				{
					throw new JsonException($"Expected value for serializable value property '{propertyName}'.");
				}

				switch (propertyName)
				{
					case _typePropertyName:
						type = ReadString(ref reader, _typePropertyName);
						break;

					case _valuePropertyName:
						using (JsonDocument document = JsonDocument.ParseValue(ref reader))
						{
							value = ReadJsonValueWithHandles(document.RootElement, registry);
						}
						break;

					default:
						reader.Skip();
						break;
				}
			}

			throw new JsonException("Unexpected end of serializable value JSON object.");
		}

		internal static object? MakeHandlesSerializable(object? value, IFluidDataObjectRegistry? registry)
		{
			if (value == null)
			{
				return null;
			}

			if (TryGetHandleUrl(value, registry, out string handleUrl))
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
					serializedDictionary[entry.Key] = MakeHandlesSerializable(entry.Value, registry);
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
						serializedDictionary[key] = MakeHandlesSerializable(entry.Value, registry);
					}
				}

				return serializedDictionary;
			}

			if (value is IEnumerable enumerable && value is not string)
			{
				List<object?> serializedList = new List<object?>();
				foreach (object? item in enumerable)
				{
					serializedList.Add(MakeHandlesSerializable(item, registry));
				}

				return serializedList;
			}

			return value;
		}

		internal static object? ReadJsonValueWithHandles(JsonElement element, IFluidDataObjectRegistry? registry)
		{
			return ConvertJsonElementWithHandles(element, registry, out bool _);
		}

		internal static object? ResolveSerializedHandles(object? value, IFluidDataObjectRegistry? registry)
		{
			if (value == null)
			{
				return null;
			}

			if (value is SerializedFluidHandle serializedHandle)
			{
				return ResolveSerializedFluidHandle(serializedHandle.Url, registry);
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

		private static object? ConvertJsonElementWithHandles(JsonElement element, IFluidDataObjectRegistry? registry, out bool changed)
		{
			switch (element.ValueKind)
			{
				case JsonValueKind.Object:
					if (TryReadSerializedHandleUrl(element, out string? url))
					{
						changed = true;
						return ResolveSerializedFluidHandle(url, registry);
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

		private static bool TryReadSerializedHandleUrl(JsonElement element, out string url)
		{
			if (element.ValueKind == JsonValueKind.Object
				&& element.TryGetProperty(_typePropertyName, out JsonElement typeElement)
				&& typeElement.ValueKind == JsonValueKind.String
				&& typeElement.GetString() == _serializedHandleTypeName
				&& element.TryGetProperty(_serializedHandleUrlPropertyName, out JsonElement urlElement)
				&& urlElement.ValueKind == JsonValueKind.String)
			{
				url = urlElement.GetString() ?? string.Empty;
				return true;
			}

			url = string.Empty;
			return false;
		}

		private static object ResolveSerializedFluidHandle(string url, IFluidDataObjectRegistry? registry)
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

		private static bool TryGetHandleUrl(object value, IFluidDataObjectRegistry? registry, out string url)
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

				throw new OcsException(OcsGateErrorCode.InvalidOperation,
					"SharedDirectory requires an IFluidDataObjectRegistry to serialize handle values.");
			}

			url = string.Empty;
			return false;
		}

		private static Dictionary<string, object?> CreateSerializedHandleWireValue(string url)
		{
			return new Dictionary<string, object?>()
			{
				[_typePropertyName] = _serializedHandleTypeName,
				[_serializedHandleUrlPropertyName] = url,
			};
		}

		private static string ReadString(ref Utf8JsonReader reader, string propertyName)
		{
			if (reader.TokenType != JsonTokenType.String)
			{
				throw new JsonException($"Expected string value for property '{propertyName}'.");
			}

			return reader.GetString() ?? string.Empty;
		}
	}
}
