// -----------------------------------------------------------------------------
// System.Text.Json (de)serialization for DirectoryOperation and its subclasses.
// Wire format matches packages/dds/map/src/directory.ts (TS SharedDirectory).
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
		private const string _missingRegistryMessage = "SharedDirectory requires an IFluidDataObjectRegistry to serialize handle values.";

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

			// Boundary validation. Every TS directory op emits an absolute path
			// (packages/dds/map/src/directory.ts always constructs envelopes with
			// `path: this.absolutePath`) and concrete key / subdirName values. DTOs
			// default those fields to string.Empty for constructor convenience;
			// validate before writing so an uninitialized DTO can't silently produce
			// a wire message with empty required fields.
			ValidateRequiredWireFields(op);

			writer.WriteStartObject();

			switch (op)
			{
				case DirectorySetOperation setOperation:
					writer.WriteString(_typePropertyName, _setTypeName);
					writer.WriteString(_pathPropertyName, setOperation.Path);
					writer.WriteString(_keyPropertyName, setOperation.Key);
					writer.WritePropertyName(_valuePropertyName);
					WriteSerializableValue(writer, setOperation.Value, registry);
					break;

				case DirectoryDeleteOperation deleteOperation:
					writer.WriteString(_typePropertyName, _deleteTypeName);
					writer.WriteString(_pathPropertyName, deleteOperation.Path);
					writer.WriteString(_keyPropertyName, deleteOperation.Key);
					break;

				case DirectoryClearOperation clearOperation:
					writer.WriteString(_typePropertyName, _clearTypeName);
					writer.WriteString(_pathPropertyName, clearOperation.Path);
					break;

				case DirectoryCreateSubDirectoryOperation createSubDirectoryOperation:
					writer.WriteString(_typePropertyName, _createSubDirectoryTypeName);
					writer.WriteString(_pathPropertyName, createSubDirectoryOperation.Path);
					writer.WriteString(_subdirNamePropertyName, createSubDirectoryOperation.SubdirName);
					break;

				case DirectoryDeleteSubDirectoryOperation deleteSubDirectoryOperation:
					writer.WriteString(_typePropertyName, _deleteSubDirectoryTypeName);
					writer.WriteString(_pathPropertyName, deleteSubDirectoryOperation.Path);
					writer.WriteString(_subdirNamePropertyName, deleteSubDirectoryOperation.SubdirName);
					break;

				default:
					throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unknown directory op runtime type: {op.GetType().FullName}");
			}

			writer.WriteEndObject();
		}

		/// <summary>
		/// Rejects DTOs whose required wire fields are null. The root
		/// <c>Path</c> and <c>Value.Type</c> discriminant must be non-empty
		/// (both are TS wire invariants); empty-string keys and subdirectory
		/// names are accepted to match the public API contract.
		/// </summary>
		private static void ValidateRequiredWireFields(DirectoryOperation op)
		{
			if (string.IsNullOrEmpty(op.Path))
			{
				throw new OcsException(
					OcsGateErrorCode.InvalidOperation,
					$"Directory operation of type '{op.GetType().Name}' is missing required 'path' field.");
			}

			switch (op)
			{
				case DirectorySetOperation setOperation:
					if (setOperation.Key is null)
					{
						throw new OcsException(OcsGateErrorCode.InvalidOperation,
							"DirectorySetOperation is missing required 'key' field.");
					}

					if (setOperation.Value is null)
					{
						throw new OcsException(OcsGateErrorCode.InvalidOperation,
							"DirectorySetOperation is missing required 'value' field.");
					}

					if (string.IsNullOrEmpty(setOperation.Value.Type))
					{
						throw new OcsException(OcsGateErrorCode.InvalidOperation,
							"DirectorySetOperation.Value is missing required 'type' discriminant.");
					}

					break;

				case DirectoryDeleteOperation deleteOperation:
					if (deleteOperation.Key is null)
					{
						throw new OcsException(OcsGateErrorCode.InvalidOperation,
							"DirectoryDeleteOperation is missing required 'key' field.");
					}

					break;

				case DirectoryCreateSubDirectoryOperation createOperation:
					if (createOperation.SubdirName is null)
					{
						throw new OcsException(OcsGateErrorCode.InvalidOperation,
							"DirectoryCreateSubDirectoryOperation is missing required 'subdirName' field.");
					}

					break;

				case DirectoryDeleteSubDirectoryOperation deleteSubOperation:
					if (deleteSubOperation.SubdirName is null)
					{
						throw new OcsException(OcsGateErrorCode.InvalidOperation,
							"DirectoryDeleteSubDirectoryOperation is missing required 'subdirName' field.");
					}

					break;
			}
		}

		/// <summary>
		/// Reads a directory operation JSON object from <paramref name="reader"/>.
		/// </summary>
		/// <remarks>
		/// Uses the standard <see cref="Utf8JsonReader"/> streaming pattern:
		/// <c>StartObject → (PropertyName → value) × N → EndObject</c>. Each JSON
		/// property is visited at most once (property names are unique within a
		/// JSON object), so we accumulate the individual fields first and then
		/// dispatch to the correct op subclass in <see cref="CreateOperation"/>
		/// based on <c>typeString</c>. Unknown properties are skipped so the reader
		/// is tolerant of forward-compat additions on the wire.
		/// </remarks>
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
			string? path = null;
			string? key = null;
			string? subdirName = null;
			SerializableValue? value = null;
			// Tracks whether the JSON object contained a `value` property. We can't
			// just check `value != null`: `null` is a legal SerializableValue.Value
			// (TS `any` allows null/undefined, and `{"type":"Plain","value":null}` is a
			// valid wire shape). We need a separate flag so CreateOperation can throw
			// specifically for a missing required `value` on a `set` op, without
			// misidentifying "field present but null" as "field missing".
			bool valueSeen = false;

			while (reader.Read())
			{
				if (reader.TokenType == JsonTokenType.EndObject)
				{
					return CreateOperation(typeString, path, key, value, valueSeen, subdirName);
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
						valueSeen = true;
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

		private static DirectoryOperation CreateOperation(string? typeString, string? path, string? key, SerializableValue? value, bool valueSeen, string? subdirName)
		{
			if (typeString == null)
			{
				throw new OcsException(OcsGateErrorCode.UnknownOp, "Directory operation is missing required 'type' field.");
			}

			if (path == null)
			{
				throw new OcsException(OcsGateErrorCode.UnknownOp, $"Directory '{typeString}' operation is missing required 'path' field.");
			}

			switch (typeString)
			{
				case _setTypeName:
					if (key == null)
					{
						throw new OcsException(OcsGateErrorCode.UnknownOp, "Directory 'set' operation is missing required 'key' field.");
					}

					if (!valueSeen)
					{
						throw new OcsException(OcsGateErrorCode.UnknownOp, "Directory 'set' operation is missing required 'value' field.");
					}

					return new DirectorySetOperation()
					{
						Path = path,
						Key = key,
						// value != null is guaranteed once valueSeen is true; ReadSerializableValue never returns null.
						Value = value!,
					};

				case _deleteTypeName:
					if (key == null)
					{
						throw new OcsException(OcsGateErrorCode.UnknownOp, "Directory 'delete' operation is missing required 'key' field.");
					}

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
					if (subdirName == null)
					{
						throw new OcsException(OcsGateErrorCode.UnknownOp, "Directory 'createSubDirectory' operation is missing required 'subdirName' field.");
					}

					return new DirectoryCreateSubDirectoryOperation()
					{
						Path = path,
						SubdirName = subdirName,
					};

				case _deleteSubDirectoryTypeName:
					if (subdirName == null)
					{
						throw new OcsException(OcsGateErrorCode.UnknownOp, "Directory 'deleteSubDirectory' operation is missing required 'subdirName' field.");
					}

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

			// SerializableValue.Type is required on the wire (matches TS
			// ISerializableValue in internalInterfaces.ts). Boundary-validated here
			// so a hand-constructed SerializableValue with a default empty Type
			// never reaches the wire.
			if (string.IsNullOrEmpty(value.Type))
			{
				throw new OcsException(
					OcsGateErrorCode.InvalidOperation,
					"SerializableValue is missing required 'type' discriminant.");
			}

			writer.WriteStartObject();
			writer.WriteString(_typePropertyName, value.Type);
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

			// TS ref: localValues.ts, directory.ts.
			// TS branches only on `type === "Shared"`. Any other type (including
			// missing) falls through to the plain-value path and stores op.value.value
			// directly. Match TS runtime tolerance: default missing type to "Plain".
			string? type = null;
			object? value = null;

			while (reader.Read())
			{
				if (reader.TokenType == JsonTokenType.EndObject)
				{
					return new SerializableValue()
					{
						// Missing/absent type defaults to Plain so downstream sees a
						// well-formed value with a valid type discriminant.
						Type = type ?? ValueType.Plain,
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
			return HandleWireFormat.MakeHandlesSerializable(value, registry, _missingRegistryMessage);
		}

		internal static object? ReadJsonValueWithHandles(JsonElement element, IFluidDataObjectRegistry? registry)
		{
			return MaterializeJsonValue(element, registry);
		}

		internal static object? ResolveSerializedHandles(object? value, IFluidDataObjectRegistry? registry)
		{
			if (value == null)
			{
				return null;
			}

			if (value is SerializedFluidHandle serializedHandle)
			{
				return HandleWireFormat.ResolveSerializedHandle(serializedHandle.Url, registry);
			}

			if (value is JsonElement element)
			{
				return MaterializeJsonValue(element, registry);
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

		private static object? MaterializeJsonValue(JsonElement element, IFluidDataObjectRegistry? registry)
		{
			switch (element.ValueKind)
			{
				case JsonValueKind.Object:
					if (HandleWireFormat.TryReadHandleUrl(element, out string url, out bool payloadPending))
					{
						return HandleWireFormat.ResolveSerializedHandle(url, registry, payloadPending);
					}

					Dictionary<string, object?> objectValue = new Dictionary<string, object?>();
					foreach (JsonProperty property in element.EnumerateObject())
					{
						objectValue[property.Name] = MaterializeJsonValue(property.Value, registry);
					}

					return objectValue;

				case JsonValueKind.Array:
					List<object?> arrayValue = new List<object?>();
					foreach (JsonElement item in element.EnumerateArray())
					{
						arrayValue.Add(MaterializeJsonValue(item, registry));
					}

					return arrayValue;

				case JsonValueKind.String:
					return element.GetString();

				case JsonValueKind.Number:
					return element.GetDouble();

				case JsonValueKind.True:
					return true;

				case JsonValueKind.False:
					return false;

				case JsonValueKind.Null:
				case JsonValueKind.Undefined:
					return null;

				default:
					throw new JsonException($"Unsupported JSON value kind '{element.ValueKind}'.");
			}
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
