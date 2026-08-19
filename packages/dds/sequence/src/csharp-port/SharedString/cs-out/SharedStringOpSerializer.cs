// -----------------------------------------------------------------------------
// JSON (de)serialization for IMergeTreeOp (Insert / Remove / Annotate / Obliterate / Group / interval ops).
// Merge-tree op wire format matches packages/dds/merge-tree/src/ops.ts.
// Part of the SharedString C# feasibility port — Wave 6.
// POC scope: Insert + Remove + Annotate + basic Obliterate + Group + interval envelopes.
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
	public static class SharedStringOpSerializer
	{
		private const string _typePropertyName = "type";
		private const string _pos1PropertyName = "pos1";
		private const string _pos2PropertyName = "pos2";
		private const string _relativePos1PropertyName = "relativePos1";
		private const string _relativePos2PropertyName = "relativePos2";
		private const string _posPropertyName = "pos";
		private const string _beforePropertyName = "before";
		private const string _segPropertyName = "seg";
		private const string _textPropertyName = "text";
		private const string _markerPropertyName = "marker";
		private const string _refTypePropertyName = "refType";
		private const string _propsPropertyName = "props";
		private const string _opsPropertyName = "ops";
		private const string _intervalOpKindPropertyName = "intervalOpKind";
		private const string _collectionPropertyName = "collection";
		private const string _idPropertyName = "id";
		private const string _startPropertyName = "start";
		private const string _endPropertyName = "end";
		private const string _intervalTypePropertyName = "intervalType";
		private const string _stickinessPropertyName = "stickiness";
		private const string _startSidePropertyName = "startSide";
		private const string _endSidePropertyName = "endSide";
		private const string _keyPropertyName = "key";
		private const string _valuePropertyName = "value";
		private const string _opNamePropertyName = "opName";
		private const string _sequenceNumberPropertyName = "sequenceNumber";
		private const string _propertiesPropertyName = "properties";
		private const string _intervalIdPropertyName = "intervalId";
		private const string _referenceRangeLabelsPropertyName = "referenceRangeLabels";
		private const string _intervalMapOperationTypeName = "act";
		private const string _intervalAddOpName = "add";
		private const string _intervalChangeOpName = "change";
		private const string _intervalDeleteOpName = "delete";
		private const string _missingRegistryMessage = "SharedString requires an IFluidDataObjectRegistry to serialize handle values.";

		private static readonly JsonWriterOptions _writerOptions = new JsonWriterOptions()
		{
			Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
		};

		public static string Serialize(MergeTree.IMergeTreeOp op, IFluidDataObjectRegistry? registry = null, long? currentSequenceNumber = null)
		{
			ArgumentNullException.ThrowIfNull(op);

			using MemoryStream stream = new MemoryStream();
			using (Utf8JsonWriter writer = new Utf8JsonWriter(stream, _writerOptions))
			{
				WriteTo(writer, op, registry, currentSequenceNumber);
			}

			return Encoding.UTF8.GetString(stream.ToArray());
		}

		public static string SerializeIntervalCollectionOperation(MergeTree.IntervalOpMsg op, IFluidDataObjectRegistry? registry = null)
		{
			ArgumentNullException.ThrowIfNull(op);

			using MemoryStream stream = new MemoryStream();
			using (Utf8JsonWriter writer = new Utf8JsonWriter(stream, _writerOptions))
			{
				WriteIntervalCollectionMapOperation(writer, op, registry);
			}

			return Encoding.UTF8.GetString(stream.ToArray());
		}

		public static MergeTree.IMergeTreeOp Deserialize(string json, IFluidDataObjectRegistry? registry = null)
		{
			ArgumentNullException.ThrowIfNull(json);

			return Deserialize(Encoding.UTF8.GetBytes(json), registry);
		}

		public static MergeTree.IMergeTreeOp Deserialize(ReadOnlySpan<byte> jsonBytes, IFluidDataObjectRegistry? registry = null)
		{
			Utf8JsonReader reader = new Utf8JsonReader(jsonBytes);
			MergeTree.IMergeTreeOp operation = ReadFrom(ref reader, registry);
			if (reader.Read())
			{
				throw new JsonException("Unexpected trailing JSON after merge-tree operation.");
			}

			return operation;
		}

		public static void WriteTo(Utf8JsonWriter writer, MergeTree.IMergeTreeOp op, IFluidDataObjectRegistry? registry = null, long? currentSequenceNumber = null)
		{
			ArgumentNullException.ThrowIfNull(writer);
			ArgumentNullException.ThrowIfNull(op);

			// Interval ops go out as their own outer object; route before WriteStartObject.
			if (op is MergeTree.IntervalOpMsg intervalOp)
			{
				WriteIntervalCollectionMapOperation(writer, intervalOp, registry, currentSequenceNumber);
				return;
			}

			writer.WriteStartObject();

			switch (op.Type)
			{
				case MergeTree.MergeTreeDeltaType.Insert:
					if (op is not MergeTree.IMergeTreeInsertMsg insertOperation)
					{
						throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unknown merge-tree insert op runtime type: {op.GetType().FullName}");
					}

					writer.WriteNumber(_typePropertyName, (int)MergeTree.MergeTreeDeltaType.Insert);
					WritePositionOrRelative(
						writer,
						_pos1PropertyName,
						insertOperation.Pos1,
						_relativePos1PropertyName,
						insertOperation.RelativePos1,
						registry);
					WriteOptionalPositionOrRelative(
						writer,
						_pos2PropertyName,
						insertOperation.Pos2,
						_relativePos2PropertyName,
						insertOperation.RelativePos2,
						registry);
					writer.WritePropertyName(_segPropertyName);
					WriteSegmentSpec(writer, insertOperation.Seg, registry);
					break;

				case MergeTree.MergeTreeDeltaType.Remove:
					if (op is not MergeTree.IMergeTreeRemoveMsg removeOperation)
					{
						throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unknown merge-tree remove op runtime type: {op.GetType().FullName}");
					}

					writer.WriteNumber(_typePropertyName, (int)MergeTree.MergeTreeDeltaType.Remove);
					WritePositionOrRelative(
						writer,
						_pos1PropertyName,
						removeOperation.Pos1,
						_relativePos1PropertyName,
						removeOperation.RelativePos1,
						registry);
					WritePositionOrRelative(
						writer,
						_pos2PropertyName,
						removeOperation.Pos2,
						_relativePos2PropertyName,
						removeOperation.RelativePos2,
						registry);
					break;

				case MergeTree.MergeTreeDeltaType.Obliterate:
					if (op is not MergeTree.IMergeTreeObliterateMsg obliterateOperation)
					{
						throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unknown merge-tree obliterate op runtime type: {op.GetType().FullName}");
					}

					writer.WriteNumber(_typePropertyName, (int)MergeTree.MergeTreeDeltaType.Obliterate);
					writer.WriteNumber(_pos1PropertyName, RequireInt32(obliterateOperation.Pos1, _pos1PropertyName));
					writer.WriteNumber(_pos2PropertyName, RequireInt32(obliterateOperation.Pos2, _pos2PropertyName));
					break;

				case MergeTree.MergeTreeDeltaType.ObliterateSided:
					if (op is not MergeTree.IMergeTreeObliterateSidedMsg sidedObliterateOperation)
					{
						throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unknown merge-tree sided obliterate op runtime type: {op.GetType().FullName}");
					}

					writer.WriteNumber(_typePropertyName, (int)MergeTree.MergeTreeDeltaType.ObliterateSided);
					writer.WritePropertyName(_pos1PropertyName);
					WriteSequencePlace(writer, sidedObliterateOperation.Pos1);
					writer.WritePropertyName(_pos2PropertyName);
					WriteSequencePlace(writer, sidedObliterateOperation.Pos2);
					break;

				case MergeTree.MergeTreeDeltaType.Annotate:
					if (op is not MergeTree.MergeTreeAnnotateMsg annotateOperation)
					{
						throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unknown merge-tree annotate op runtime type: {op.GetType().FullName}");
					}

					writer.WriteNumber(_typePropertyName, (int)MergeTree.MergeTreeDeltaType.Annotate);
					WritePositionOrRelative(
						writer,
						_pos1PropertyName,
						annotateOperation.Pos1,
						_relativePos1PropertyName,
						annotateOperation.RelativePos1,
						registry);
					WritePositionOrRelative(
						writer,
						_pos2PropertyName,
						annotateOperation.Pos2,
						_relativePos2PropertyName,
						annotateOperation.RelativePos2,
						registry);
					writer.WritePropertyName(_propsPropertyName);
					if (annotateOperation.Props is null)
					{
						writer.WriteNullValue();
					}
					else
					{
						WritePropertySet(writer, annotateOperation.Props, registry);
					}

					break;

				case MergeTree.MergeTreeDeltaType.Group:
					if (op is not MergeTree.MergeTreeGroupMsg groupOperation)
					{
						throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unknown merge-tree group op runtime type: {op.GetType().FullName}");
					}

					writer.WriteNumber(_typePropertyName, (int)MergeTree.MergeTreeDeltaType.Group);
					writer.WritePropertyName(_opsPropertyName);
					writer.WriteStartArray();
					foreach (MergeTree.MergeTreeOp memberOperation in groupOperation.Ops)
					{
						if (!IsMergeTreeGroupMember(memberOperation))
						{
							throw new JsonException("Merge-tree group operations can only contain merge-tree delta operations.");
						}

						WriteTo(writer, memberOperation, registry);
					}

					writer.WriteEndArray();
					break;

				default:
					throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unknown merge-tree op type: {(int)op.Type}");
			}

			writer.WriteEndObject();
		}

		public static MergeTree.IMergeTreeOp ReadFrom(ref Utf8JsonReader reader, IFluidDataObjectRegistry? registry = null)
		{
			if (reader.TokenType == JsonTokenType.None && !reader.Read())
			{
				throw new JsonException("Expected merge-tree operation JSON object.");
			}

			if (reader.TokenType != JsonTokenType.StartObject)
			{
				throw new JsonException("Expected merge-tree operation JSON object.");
			}

			int? type = null;
			bool isIntervalMapOperation = false;
			int? pos1 = null;
			int? pos2 = null;
			object? relativePos1 = null;
			object? relativePos2 = null;
			MergeTree.SequencePlace? sidedPos1 = null;
			MergeTree.SequencePlace? sidedPos2 = null;
			object? seg = null;
			bool hasSeg = false;
			MergeTree.PropertySet? props = null;
			int? intervalOpKind = null;
			string collectionName = string.Empty;
			string intervalId = string.Empty;
			int? start = null;
			int? end = null;
			int? intervalType = null;
			int? stickiness = null;
			int? startSide = null;
			int? endSide = null;
			List<MergeTree.MergeTreeOp>? ops = null;
			string intervalMapKey = string.Empty;
			IntervalMapOperationValue? intervalMapOperationValue = null;

			while (reader.Read())
			{
				if (reader.TokenType == JsonTokenType.EndObject)
				{
					return CreateOperation(
						type,
						isIntervalMapOperation,
						pos1,
						pos2,
						relativePos1,
						relativePos2,
						sidedPos1,
						sidedPos2,
						seg,
						hasSeg,
						props,
						intervalOpKind,
						collectionName,
						intervalId,
						start,
						end,
						intervalType,
						stickiness,
						startSide,
						endSide,
						ops,
						intervalMapKey,
						intervalMapOperationValue);
				}

				if (reader.TokenType != JsonTokenType.PropertyName)
				{
					throw new JsonException("Expected merge-tree operation property name.");
				}

				string propertyName = reader.GetString() ?? string.Empty;
				if (!reader.Read())
				{
					throw new JsonException($"Expected value for merge-tree operation property '{propertyName}'.");
				}

				switch (propertyName)
				{
					case _typePropertyName:
						if (reader.TokenType == JsonTokenType.String
							&& string.Equals(reader.GetString(), _intervalMapOperationTypeName, StringComparison.Ordinal))
						{
							isIntervalMapOperation = true;
						}
						else
						{
							type = ReadOperationType(ref reader, _typePropertyName);
						}

						break;

					case _pos1PropertyName:
						if (reader.TokenType == JsonTokenType.StartObject)
						{
							sidedPos1 = ReadSequencePlace(ref reader, _pos1PropertyName);
						}
						else
						{
							pos1 = ReadInt32(ref reader, _pos1PropertyName);
						}

						break;

					case _pos2PropertyName:
						if (reader.TokenType == JsonTokenType.StartObject)
						{
							sidedPos2 = ReadSequencePlace(ref reader, _pos2PropertyName);
						}
						else
						{
							pos2 = ReadInt32(ref reader, _pos2PropertyName);
						}

						break;

					case _segPropertyName:
						seg = ReadSegmentSpec(ref reader, registry);
						hasSeg = true;
						break;

					case _propsPropertyName:
						props = reader.TokenType == JsonTokenType.Null ? null : ReadPropertySet(ref reader, _propsPropertyName, registry);
						break;

					case _relativePos1PropertyName:
						relativePos1 = ReadJsonValue(ref reader, registry);
						break;

					case _relativePos2PropertyName:
						relativePos2 = ReadJsonValue(ref reader, registry);
						break;

					case _opsPropertyName:
						ops = ReadGroupOps(ref reader, registry);
						break;

					case _intervalOpKindPropertyName:
						intervalOpKind = ReadInt32(ref reader, _intervalOpKindPropertyName);
						break;

					case _collectionPropertyName:
						collectionName = ReadString(ref reader, _collectionPropertyName);
						break;

					case _idPropertyName:
						intervalId = ReadString(ref reader, _idPropertyName);
						break;

					case _startPropertyName:
					{
						(int? pos, int? impliedSide) = ReadEndpointPositionOrSentinel(ref reader, _startPropertyName);
						start = pos;
						if (impliedSide.HasValue && !startSide.HasValue)
						{
							startSide = impliedSide;
						}

						break;
					}

					case _endPropertyName:
					{
						(int? pos, int? impliedSide) = ReadEndpointPositionOrSentinel(ref reader, _endPropertyName);
						end = pos;
						if (impliedSide.HasValue && !endSide.HasValue)
						{
							endSide = impliedSide;
						}

						break;
					}

					case _intervalTypePropertyName:
						intervalType = ReadInt32(ref reader, _intervalTypePropertyName);
						break;

					case _stickinessPropertyName:
						stickiness = ReadNullableInt32(ref reader, _stickinessPropertyName);
						break;

					case _startSidePropertyName:
						startSide = ReadNullableInt32(ref reader, _startSidePropertyName);
						break;

					case _endSidePropertyName:
						endSide = ReadNullableInt32(ref reader, _endSidePropertyName);
						break;

					case _keyPropertyName:
						intervalMapKey = ReadString(ref reader, _keyPropertyName);
						break;

					case _valuePropertyName:
						intervalMapOperationValue = ReadIntervalMapOperationValue(ref reader, registry);
						break;

					default:
						reader.Skip();
						break;
				}
			}

			throw new JsonException("Unexpected end of merge-tree operation JSON object.");
		}

		private static MergeTree.IMergeTreeOp CreateOperation(
			int? type,
			bool isIntervalMapOperation,
			int? pos1,
			int? pos2,
			object? relativePos1,
			object? relativePos2,
			MergeTree.SequencePlace? sidedPos1,
			MergeTree.SequencePlace? sidedPos2,
			object? seg,
			bool hasSeg,
			MergeTree.PropertySet? props,
			int? intervalOpKind,
			string collectionName,
			string intervalId,
			int? start,
			int? end,
			int? intervalType,
			int? stickiness,
			int? startSide,
			int? endSide,
			List<MergeTree.MergeTreeOp>? ops,
			string intervalMapKey,
			IntervalMapOperationValue? intervalMapOperationValue)
		{
			if (isIntervalMapOperation)
			{
				return CreateIntervalOperationFromMap(intervalMapKey, intervalMapOperationValue);
			}

			switch (type)
			{
				case (int)MergeTree.MergeTreeDeltaType.Insert:
					if (!hasSeg)
					{
						throw new JsonException("Expected merge-tree insert operation segment.");
					}

					return new MergeTree.MergeTreeInsertMsg()
					{
						Pos1 = pos1,
						RelativePos1 = relativePos1,
						Pos2 = pos2,
						RelativePos2 = relativePos2,
						Seg = seg,
					};

				case (int)MergeTree.MergeTreeDeltaType.Remove:
					RequirePositionOrRelative(pos1, relativePos1, _pos1PropertyName, _relativePos1PropertyName);
					RequirePositionOrRelative(pos2, relativePos2, _pos2PropertyName, _relativePos2PropertyName);
					return new MergeTree.MergeTreeRemoveMsg()
					{
						Pos1 = pos1,
						RelativePos1 = relativePos1,
						Pos2 = pos2,
						RelativePos2 = relativePos2,
					};

				case (int)MergeTree.MergeTreeDeltaType.Obliterate:
					return new MergeTree.MergeTreeObliterateMsg()
					{
						Pos1 = RequireInt32(pos1, _pos1PropertyName),
						Pos2 = RequireInt32(pos2, _pos2PropertyName),
					};

				case (int)MergeTree.MergeTreeDeltaType.ObliterateSided:
					return new MergeTree.MergeTreeObliterateSidedMsg()
					{
						Pos1 = RequireSequencePlace(sidedPos1, _pos1PropertyName),
						Pos2 = RequireSequencePlace(sidedPos2, _pos2PropertyName),
					};

				case (int)MergeTree.MergeTreeDeltaType.Annotate:
					RequirePositionOrRelative(pos1, relativePos1, _pos1PropertyName, _relativePos1PropertyName);
					RequirePositionOrRelative(pos2, relativePos2, _pos2PropertyName, _relativePos2PropertyName);
					return new MergeTree.MergeTreeAnnotateMsg()
					{
						Pos1 = pos1,
						RelativePos1 = relativePos1,
						Pos2 = pos2,
						RelativePos2 = relativePos2,
						Props = props,
					};

				case (int)MergeTree.MergeTreeDeltaType.Group:
					if (ops is null)
					{
						throw new JsonException("Expected merge-tree group operation ops array.");
					}

					MergeTree.MergeTreeGroupMsg groupOperation = new MergeTree.MergeTreeGroupMsg();
					groupOperation.Ops.AddRange(ops);
					return groupOperation;

				default:
					throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unknown merge-tree op type: {type?.ToString() ?? "<missing>"}");
			}
		}

		private static void WriteIntervalCollectionMapOperation(
			Utf8JsonWriter writer,
			MergeTree.IntervalOpMsg intervalOperation,
			IFluidDataObjectRegistry? registry,
			long? currentSequenceNumber = null)
		{
			writer.WriteStartObject();
			writer.WriteString(_typePropertyName, _intervalMapOperationTypeName);
			writer.WriteString(_keyPropertyName, intervalOperation.CollectionName);
			writer.WritePropertyName(_valuePropertyName);
			writer.WriteStartObject();

			switch (intervalOperation)
			{
				case MergeTree.IntervalAddOpMsg addOperation:
					writer.WriteString(_opNamePropertyName, _intervalAddOpName);
					writer.WritePropertyName(_valuePropertyName);
					writer.WriteStartObject();
					WriteIntervalMapPayloadHeader(
						writer,
						addOperation.IntervalType,
						currentSequenceNumber,
						addOperation.Stickiness,
						addOperation.StartSide,
						addOperation.EndSide);
					writer.WriteNumber(_startPropertyName, addOperation.Start);
					writer.WriteNumber(_endPropertyName, addOperation.End);
					writer.WritePropertyName(_propertiesPropertyName);
					WriteIntervalMapProperties(writer, addOperation.CollectionName, addOperation.IntervalId, addOperation.Props, registry);
					writer.WriteEndObject();
					break;

				case MergeTree.IntervalDeleteOpMsg deleteOperation:
					writer.WriteString(_opNamePropertyName, _intervalDeleteOpName);
					writer.WritePropertyName(_valuePropertyName);
					writer.WriteStartObject();
					WriteIntervalMapPayloadHeader(writer, Intervals.IntervalType.SlideOnRemove, currentSequenceNumber);
					writer.WritePropertyName(_propertiesPropertyName);
					WriteIntervalMapProperties(writer, deleteOperation.CollectionName, deleteOperation.IntervalId, null, registry);
					writer.WriteEndObject();
					break;

				case MergeTree.IntervalChangeOpMsg changeOperation:
					writer.WriteString(_opNamePropertyName, _intervalChangeOpName);
					writer.WritePropertyName(_valuePropertyName);
					writer.WriteStartObject();
					WriteIntervalMapPayloadHeader(
						writer,
						Intervals.IntervalType.SlideOnRemove,
						currentSequenceNumber,
						changeOperation.Stickiness,
						changeOperation.StartSide,
						changeOperation.EndSide);
					WriteNullableIntervalEndpoint(writer, _startPropertyName, changeOperation.Start);
					WriteNullableIntervalEndpoint(writer, _endPropertyName, changeOperation.End);
					writer.WritePropertyName(_propertiesPropertyName);
					// Pass Props so combined endpoint+property changes serialize
					// as a single op (matches TS intervalCollection.ts changeInterval).
					WriteIntervalMapProperties(writer, changeOperation.CollectionName, changeOperation.IntervalId, changeOperation.Props, registry);
					writer.WriteEndObject();
					break;

				case MergeTree.IntervalPropertyChangedOpMsg propertyChangedOperation:
					writer.WriteString(_opNamePropertyName, _intervalChangeOpName);
					writer.WritePropertyName(_valuePropertyName);
					writer.WriteStartObject();
					WriteIntervalMapPayloadHeader(writer, Intervals.IntervalType.SlideOnRemove, currentSequenceNumber);
					writer.WritePropertyName(_propertiesPropertyName);
					WriteIntervalMapProperties(
						writer,
						propertyChangedOperation.CollectionName,
						propertyChangedOperation.IntervalId,
						propertyChangedOperation.Props,
						registry);
					writer.WriteEndObject();
					break;

				default:
					throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unknown interval op runtime type: {intervalOperation.GetType().FullName}");
			}

			writer.WriteEndObject();
			writer.WriteEndObject();
		}

		private static void WriteIntervalMapPayloadHeader(
			Utf8JsonWriter writer,
			Intervals.IntervalType intervalType,
			long? currentSequenceNumber = null,
			Intervals.IntervalStickiness? stickiness = null,
			Intervals.Side? startSide = null,
			Intervals.Side? endSide = null)
		{
			// Serialized interval ops carry the client's current known sequence number;
			// reconnect/rebase logic reads it back.
			writer.WriteNumber(_sequenceNumberPropertyName, currentSequenceNumber ?? 0);
			writer.WriteNumber(_intervalTypePropertyName, (int)intervalType);
			if (stickiness.HasValue)
			{
				writer.WriteNumber(_stickinessPropertyName, (int)stickiness.Value);
			}

			if (startSide.HasValue)
			{
				writer.WriteNumber(_startSidePropertyName, (int)startSide.Value);
			}

			if (endSide.HasValue)
			{
				writer.WriteNumber(_endSidePropertyName, (int)endSide.Value);
			}
		}

		private static void WriteNullableIntervalEndpoint(Utf8JsonWriter writer, string propertyName, int? value)
		{
			if (value.HasValue)
			{
				writer.WriteNumber(propertyName, value.Value);
			}
		}

		private static void WriteIntervalMapProperties(
			Utf8JsonWriter writer,
			string collectionName,
			string intervalId,
			IReadOnlyDictionary<string, object?>? properties,
			IFluidDataObjectRegistry? registry)
		{
			writer.WriteStartObject();
			if (properties is not null)
			{
				foreach (KeyValuePair<string, object?> property in properties)
				{
					if (string.Equals(property.Key, _intervalIdPropertyName, StringComparison.Ordinal)
						|| string.Equals(property.Key, _referenceRangeLabelsPropertyName, StringComparison.Ordinal))
					{
						continue;
					}

					writer.WritePropertyName(property.Key);
					WriteJsonValue(writer, property.Value, registry);
				}
			}

			writer.WriteString(_intervalIdPropertyName, intervalId);
			writer.WritePropertyName(_referenceRangeLabelsPropertyName);
			writer.WriteStartArray();
			writer.WriteStringValue(collectionName);
			writer.WriteEndArray();
			writer.WriteEndObject();
		}

		private static void WriteSequencePlace(Utf8JsonWriter writer, MergeTree.SequencePlace place)
		{
			if (place is null)
			{
				throw new JsonException("Expected sided obliterate sequence place.");
			}

			writer.WriteStartObject();
			writer.WriteNumber(_posPropertyName, place.Position);
			writer.WriteBoolean(_beforePropertyName, place.Side == MergeTree.Side.Before);
			writer.WriteEndObject();
		}

		private static MergeTree.SequencePlace ReadSequencePlace(ref Utf8JsonReader reader, string propertyName)
		{
			if (reader.TokenType != JsonTokenType.StartObject)
			{
				throw new JsonException($"Expected sequence place object for property '{propertyName}'.");
			}

			int? pos = null;
			bool? before = null;
			while (reader.Read())
			{
				if (reader.TokenType == JsonTokenType.EndObject)
				{
					int position = RequireInt32(pos, $"{propertyName}.{_posPropertyName}");
					MergeTree.Side side = before == true
						? MergeTree.Side.Before
						: MergeTree.Side.After;
					if (position == -1 && side == MergeTree.Side.After)
					{
						return MergeTree.SequencePlace.Start;
					}

					if (position == -1 && side == MergeTree.Side.Before)
					{
						return MergeTree.SequencePlace.End;
					}

					return MergeTree.SequencePlace.At(position, side);
				}

				if (reader.TokenType != JsonTokenType.PropertyName)
				{
					throw new JsonException($"Expected sequence place property name for '{propertyName}'.");
				}

				string childPropertyName = reader.GetString() ?? string.Empty;
				if (!reader.Read())
				{
					throw new JsonException($"Expected value for sequence place property '{childPropertyName}'.");
				}

				switch (childPropertyName)
				{
					case _posPropertyName:
						pos = ReadInt32(ref reader, $"{propertyName}.{_posPropertyName}");
						break;

					case _beforePropertyName:
						before = ReadBoolean(ref reader, $"{propertyName}.{_beforePropertyName}");
						break;

					default:
						reader.Skip();
						break;
				}
			}

			throw new JsonException($"Unexpected end of sequence place object for property '{propertyName}'.");
		}

		private static List<MergeTree.MergeTreeOp> ReadGroupOps(ref Utf8JsonReader reader, IFluidDataObjectRegistry? registry)
		{
			if (reader.TokenType != JsonTokenType.StartArray)
			{
				throw new JsonException("Expected merge-tree group operation ops array.");
			}

			List<MergeTree.MergeTreeOp> ops = new List<MergeTree.MergeTreeOp>();
			while (reader.Read())
			{
				if (reader.TokenType == JsonTokenType.EndArray)
				{
					return ops;
				}

				MergeTree.IMergeTreeOp operation = ReadFrom(ref reader, registry);
				if (operation is not MergeTree.MergeTreeOp mergeTreeOperation)
				{
					throw new JsonException("Merge-tree group operations can only contain merge-tree delta operations.");
				}

				if (!IsMergeTreeGroupMember(mergeTreeOperation))
				{
					throw new JsonException("Merge-tree group operations can only contain merge-tree delta operations.");
				}

				ops.Add(mergeTreeOperation);
			}

			throw new JsonException("Unexpected end of merge-tree group operation ops array.");
		}

		private static bool IsMergeTreeGroupMember(MergeTree.MergeTreeOp op)
		{
			// Group ops carry merge-tree delta ops only; interval ops go out individually.
			return op.Type == MergeTree.MergeTreeDeltaType.Insert
				|| op.Type == MergeTree.MergeTreeDeltaType.Remove
				|| op.Type == MergeTree.MergeTreeDeltaType.Annotate
				|| op.Type == MergeTree.MergeTreeDeltaType.Obliterate
				|| op.Type == MergeTree.MergeTreeDeltaType.ObliterateSided
				|| op.Type == MergeTree.MergeTreeDeltaType.Group;
		}

		private static IntervalMapOperationValue ReadIntervalMapOperationValue(ref Utf8JsonReader reader, IFluidDataObjectRegistry? registry)
		{
			if (reader.TokenType != JsonTokenType.StartObject)
			{
				throw new JsonException("Expected interval map operation value object.");
			}

			string opName = string.Empty;
			IntervalMapPayload? payload = null;
			while (reader.Read())
			{
				if (reader.TokenType == JsonTokenType.EndObject)
				{
					return new IntervalMapOperationValue(opName, payload);
				}

				if (reader.TokenType != JsonTokenType.PropertyName)
				{
					throw new JsonException("Expected interval map operation property name.");
				}

				string propertyName = reader.GetString() ?? string.Empty;
				if (!reader.Read())
				{
					throw new JsonException($"Expected value for interval map operation property '{propertyName}'.");
				}

				switch (propertyName)
				{
					case _opNamePropertyName:
						opName = ReadString(ref reader, _opNamePropertyName);
						break;

					case _valuePropertyName:
						payload = ReadIntervalMapPayload(ref reader, registry);
						break;

					default:
						reader.Skip();
						break;
				}
			}

			throw new JsonException("Unexpected end of interval map operation value object.");
		}

		private static IntervalMapPayload ReadIntervalMapPayload(ref Utf8JsonReader reader, IFluidDataObjectRegistry? registry)
		{
			if (reader.TokenType != JsonTokenType.StartObject)
			{
				throw new JsonException("Expected interval map payload object.");
			}

			int? start = null;
			bool hasStart = false;
			int? end = null;
			bool hasEnd = false;
			int? intervalType = null;
			int? stickiness = null;
			int? startSide = null;
			int? endSide = null;
			MergeTree.PropertySet? properties = null;
			while (reader.Read())
			{
				if (reader.TokenType == JsonTokenType.EndObject)
				{
					return new IntervalMapPayload(
						start,
						hasStart,
						end,
						hasEnd,
						intervalType,
						stickiness,
						startSide,
						endSide,
						properties);
				}

				if (reader.TokenType != JsonTokenType.PropertyName)
				{
					throw new JsonException("Expected interval map payload property name.");
				}

				string propertyName = reader.GetString() ?? string.Empty;
				if (!reader.Read())
				{
					throw new JsonException($"Expected value for interval map payload property '{propertyName}'.");
				}

				switch (propertyName)
				{
					case _startPropertyName:
					{
						(int? pos, int? impliedSide) = ReadEndpointPositionOrSentinel(ref reader, _startPropertyName);
						start = pos;
						hasStart = true;
						if (impliedSide.HasValue && !startSide.HasValue)
						{
							startSide = impliedSide;
						}

						break;
					}

					case _endPropertyName:
					{
						(int? pos, int? impliedSide) = ReadEndpointPositionOrSentinel(ref reader, _endPropertyName);
						end = pos;
						hasEnd = true;
						if (impliedSide.HasValue && !endSide.HasValue)
						{
							endSide = impliedSide;
						}

						break;
					}

					case _intervalTypePropertyName:
						intervalType = ReadInt32(ref reader, _intervalTypePropertyName);
						break;

					case _stickinessPropertyName:
						stickiness = ReadNullableInt32(ref reader, _stickinessPropertyName);
						break;

					case _startSidePropertyName:
						startSide = ReadNullableInt32(ref reader, _startSidePropertyName);
						break;

					case _endSidePropertyName:
						endSide = ReadNullableInt32(ref reader, _endSidePropertyName);
						break;

					case _propertiesPropertyName:
						properties = reader.TokenType == JsonTokenType.Null ? null : ReadPropertySet(ref reader, _propertiesPropertyName, registry);
						break;

					case _sequenceNumberPropertyName:
						reader.Skip();
						break;

					default:
						reader.Skip();
						break;
				}
			}

			throw new JsonException("Unexpected end of interval map payload object.");
		}

		private static MergeTree.IMergeTreeOp CreateIntervalOperationFromMap(
			string collectionName,
			IntervalMapOperationValue? operationValue)
		{
			if (operationValue is null || operationValue.Payload is null)
			{
				throw new JsonException("Expected interval map operation value.");
			}

			IntervalMapPayload payload = operationValue.Payload;
			string intervalId = GetIntervalId(payload);
			MergeTree.PropertySet userProperties = ExtractIntervalUserProperties(payload.Properties);
			switch (operationValue.OpName)
			{
				case _intervalAddOpName:
					return new MergeTree.IntervalAddOpMsg()
					{
						CollectionName = collectionName,
						IntervalId = intervalId,
						Start = RequireInt32(payload.Start, _startPropertyName),
						End = RequireInt32(payload.End, _endPropertyName),
						IntervalType = payload.IntervalType.HasValue
							? (Intervals.IntervalType)payload.IntervalType.Value
							: Intervals.IntervalType.SlideOnRemove,
						Stickiness = payload.Stickiness.HasValue
							? (Intervals.IntervalStickiness)payload.Stickiness.Value
							: Intervals.IntervalStickiness.End,
						StartSide = payload.StartSide.HasValue
							? (Intervals.Side)payload.StartSide.Value
							: Intervals.Side.Before,
						EndSide = payload.EndSide.HasValue
							? (Intervals.Side)payload.EndSide.Value
							: Intervals.Side.Before,
						Props = userProperties,
					};

				case _intervalDeleteOpName:
					return new MergeTree.IntervalDeleteOpMsg()
					{
						CollectionName = collectionName,
						IntervalId = intervalId,
					};

				case _intervalChangeOpName when payload.HasStart || payload.HasEnd:
					return new MergeTree.IntervalChangeOpMsg()
					{
						CollectionName = collectionName,
						IntervalId = intervalId,
						Start = payload.Start,
						End = payload.End,
						Stickiness = payload.Stickiness.HasValue ? (Intervals.IntervalStickiness)payload.Stickiness.Value : (Intervals.IntervalStickiness?)null,
						StartSide = payload.StartSide.HasValue ? (Intervals.Side)payload.StartSide.Value : (Intervals.Side?)null,
						EndSide = payload.EndSide.HasValue ? (Intervals.Side)payload.EndSide.Value : (Intervals.Side?)null,
						// TS ref: packages/dds/sequence/src/intervalCollection.ts —
						// TS emits combined ops carrying both endpoints and user
						// props. Preserve any non-empty user props from the wire
						// so the receive path applies them alongside the endpoint
						// change.
						Props = userProperties.Count > 0 ? userProperties : null,
					};

				case _intervalChangeOpName:
					return new MergeTree.IntervalPropertyChangedOpMsg()
					{
						CollectionName = collectionName,
						IntervalId = intervalId,
						Props = userProperties,
					};

				default:
					throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unknown interval map op name: {operationValue.OpName}");
			}
		}

		private static string GetIntervalId(IntervalMapPayload payload)
		{
			if (payload.Properties is not null
				&& payload.Properties.TryGetValue(_intervalIdPropertyName, out object? intervalIdValue)
				&& intervalIdValue is string intervalId)
			{
				return intervalId;
			}

			if (payload.HasStart && payload.HasEnd)
			{
				return $"legacy{payload.Start}-{payload.End}";
			}

			throw new JsonException("Expected interval id in interval map operation properties.");
		}

		private static MergeTree.PropertySet ExtractIntervalUserProperties(MergeTree.PropertySet? properties)
		{
			MergeTree.PropertySet userProperties = new();
			if (properties is null)
			{
				return userProperties;
			}

			foreach (KeyValuePair<string, object?> property in properties)
			{
				if (string.Equals(property.Key, _intervalIdPropertyName, StringComparison.Ordinal)
					|| string.Equals(property.Key, _referenceRangeLabelsPropertyName, StringComparison.Ordinal))
				{
					continue;
				}

				userProperties[property.Key] = property.Value;
			}

			return userProperties;
		}

		private static void WritePositionOrRelative(
			Utf8JsonWriter writer,
			string positionPropertyName,
			int? position,
			string relativePositionPropertyName,
			object? relativePosition,
			IFluidDataObjectRegistry? registry)
		{
			if (!position.HasValue && relativePosition is null)
			{
				throw new JsonException($"Expected '{positionPropertyName}' or '{relativePositionPropertyName}' on merge-tree operation.");
			}

			WriteOptionalPositionOrRelative(
				writer,
				positionPropertyName,
				position,
				relativePositionPropertyName,
				relativePosition,
				registry);
		}

		private static void WriteOptionalPositionOrRelative(
			Utf8JsonWriter writer,
			string positionPropertyName,
			int? position,
			string relativePositionPropertyName,
			object? relativePosition,
			IFluidDataObjectRegistry? registry)
		{
			if (position.HasValue)
			{
				writer.WriteNumber(positionPropertyName, position.Value);
			}

			if (relativePosition is not null)
			{
				writer.WritePropertyName(relativePositionPropertyName);
				WriteJsonValue(writer, relativePosition, registry);
			}
		}

		private static void WriteSegmentSpec(Utf8JsonWriter writer, object? segmentSpec, IFluidDataObjectRegistry? registry)
		{
			if (segmentSpec is null)
			{
				throw new JsonException("Expected merge-tree insert operation segment.");
			}

			switch (segmentSpec)
			{
				case MergeTree.Marker marker:
					WriteSegmentSpec(writer, marker.ToJSONObject(), registry);
					break;

				case MergeTree.IJSONMarkerSegment markerSegmentSpec:
					WriteMarkerSegmentObject(writer, markerSegmentSpec.Marker.RefType, markerSegmentSpec.Props ?? markerSegmentSpec.Marker.Props, registry);
					break;

				case MergeTree.TextSegment textSegment:
					WriteSegmentSpec(writer, textSegment.ToJSONObject(), registry);
					break;

				case MergeTree.IJSONTextSegment textSegmentSpec:
					WriteTextSegmentObject(writer, textSegmentSpec.Text, textSegmentSpec.Props, registry);
					break;

				case string text:
					writer.WriteStringValue(text);
					break;

				case JsonElement jsonElement:
					WriteJsonElementSegmentSpec(writer, jsonElement);
					break;

				default:
					throw new NotSupportedException("Only TextSegment and Marker insert payloads are supported by the POC serializer.");
			}
		}

		private static void WriteJsonElementSegmentSpec(Utf8JsonWriter writer, JsonElement jsonElement)
		{
			if (jsonElement.ValueKind != JsonValueKind.String && jsonElement.ValueKind != JsonValueKind.Object)
			{
				throw new JsonException("Expected text segment JSON string or object.");
			}

			jsonElement.WriteTo(writer);
		}

		private static object ReadSegmentSpec(ref Utf8JsonReader reader, IFluidDataObjectRegistry? registry)
		{
			switch (reader.TokenType)
			{
				case JsonTokenType.String:
					string text = reader.GetString() ?? string.Empty;
					return MergeTree.TextSegment.FromJSONObject(text)?.ToJSONObject() ?? text;

				case JsonTokenType.StartObject:
					return ReadSegmentObject(ref reader, registry);

				default:
					throw new JsonException("Expected text or marker segment JSON string or object.");
			}
		}

		private static object ReadSegmentObject(ref Utf8JsonReader reader, IFluidDataObjectRegistry? registry)
		{
			if (reader.TokenType != JsonTokenType.StartObject)
			{
				throw new JsonException("Expected segment JSON object.");
			}

			string? text = null;
			bool hasText = false;
			MergeTree.JSONMarkerDef? marker = null;
			bool hasMarker = false;
			MergeTree.PropertySet? props = null;

			while (reader.Read())
			{
				if (reader.TokenType == JsonTokenType.EndObject)
				{
					if (hasText && !hasMarker)
					{
						if (text is null)
						{
							throw new JsonException("Expected text segment property 'text'.");
						}

						return new MergeTree.JSONTextSegment()
						{
							Text = text,
							Props = props,
						};
					}

					if (hasMarker && !hasText)
					{
						return new MergeTree.JSONMarkerSegment()
						{
							Marker = marker ?? new MergeTree.JSONMarkerDef(),
							Props = props ?? marker?.Props,
						};
					}

					throw new JsonException("Expected text or marker segment JSON object.");
				}

				if (reader.TokenType != JsonTokenType.PropertyName)
				{
					throw new JsonException("Expected segment property name.");
				}

				string propertyName = reader.GetString() ?? string.Empty;
				if (!reader.Read())
				{
					throw new JsonException($"Expected value for segment property '{propertyName}'.");
				}

				switch (propertyName)
				{
					case _textPropertyName:
						text = ReadString(ref reader, _textPropertyName);
						hasText = true;
						break;

					case _markerPropertyName:
						marker = ReadMarkerDef(ref reader, registry);
						hasMarker = true;
						break;

					case _propsPropertyName:
						props = reader.TokenType == JsonTokenType.Null ? null : ReadPropertySet(ref reader, _propsPropertyName, registry);
						break;

					default:
						reader.Skip();
						break;
				}
			}

			throw new JsonException("Unexpected end of segment JSON object.");
		}

		private static MergeTree.JSONTextSegment ReadTextSegmentObject(ref Utf8JsonReader reader, IFluidDataObjectRegistry? registry)
		{
			if (reader.TokenType != JsonTokenType.StartObject)
			{
				throw new JsonException("Expected text segment JSON object.");
			}

			string? text = null;
			MergeTree.PropertySet? props = null;

			while (reader.Read())
			{
				if (reader.TokenType == JsonTokenType.EndObject)
				{
					if (text is null)
					{
						throw new JsonException("Expected text segment property 'text'.");
					}

					return new MergeTree.JSONTextSegment()
					{
						Text = text,
						Props = props,
					};
				}

				if (reader.TokenType != JsonTokenType.PropertyName)
				{
					throw new JsonException("Expected text segment property name.");
				}

				string propertyName = reader.GetString() ?? string.Empty;
				if (!reader.Read())
				{
					throw new JsonException($"Expected value for text segment property '{propertyName}'.");
				}

				switch (propertyName)
				{
					case _textPropertyName:
						text = ReadString(ref reader, _textPropertyName);
						break;

					case _propsPropertyName:
						props = reader.TokenType == JsonTokenType.Null ? null : ReadPropertySet(ref reader, _propsPropertyName, registry);
						break;

					default:
						reader.Skip();
						break;
				}
			}

			throw new JsonException("Unexpected end of text segment JSON object.");
		}

		private static void WriteTextSegmentObject(Utf8JsonWriter writer, string text, MergeTree.PropertySet? props, IFluidDataObjectRegistry? registry)
		{
			writer.WriteStartObject();
			writer.WriteString(_textPropertyName, text);

			if (props is not null)
			{
				writer.WritePropertyName(_propsPropertyName);
				WritePropertySet(writer, props, registry);
			}

			writer.WriteEndObject();
		}

		private static void WriteMarkerSegmentObject(Utf8JsonWriter writer, MergeTree.ReferenceType refType, MergeTree.PropertySet? props, IFluidDataObjectRegistry? registry)
		{
			writer.WriteStartObject();
			writer.WritePropertyName(_markerPropertyName);
			writer.WriteStartObject();
			writer.WriteNumber(_refTypePropertyName, (int)refType);
			writer.WriteEndObject();

			if (props is not null)
			{
				writer.WritePropertyName(_propsPropertyName);
				WritePropertySet(writer, props, registry);
			}

			writer.WriteEndObject();
		}

		private static MergeTree.JSONMarkerDef ReadMarkerDef(ref Utf8JsonReader reader, IFluidDataObjectRegistry? registry)
		{
			if (reader.TokenType != JsonTokenType.StartObject)
			{
				throw new JsonException("Expected marker segment JSON object.");
			}

			MergeTree.ReferenceType refType = MergeTree.ReferenceType.Simple;
			MergeTree.PropertySet? props = null;

			while (reader.Read())
			{
				if (reader.TokenType == JsonTokenType.EndObject)
				{
					return new MergeTree.JSONMarkerDef()
					{
						RefType = refType,
						Props = props,
					};
				}

				if (reader.TokenType != JsonTokenType.PropertyName)
				{
					throw new JsonException("Expected marker segment property name.");
				}

				string propertyName = reader.GetString() ?? string.Empty;
				if (!reader.Read())
				{
					throw new JsonException($"Expected value for marker segment property '{propertyName}'.");
				}

				switch (propertyName)
				{
					case _refTypePropertyName:
						refType = (MergeTree.ReferenceType)ReadInt32(ref reader, _refTypePropertyName);
						break;

					case _propsPropertyName:
						props = reader.TokenType == JsonTokenType.Null ? null : ReadPropertySet(ref reader, _propsPropertyName, registry);
						break;

					default:
						reader.Skip();
						break;
				}
			}

			throw new JsonException("Unexpected end of marker segment JSON object.");
		}

		private static void WritePropertySet(Utf8JsonWriter writer, IReadOnlyDictionary<string, object?> properties, IFluidDataObjectRegistry? registry)
		{
			writer.WriteStartObject();

			foreach (KeyValuePair<string, object?> property in properties)
			{
				writer.WritePropertyName(property.Key);
				WriteJsonValue(writer, property.Value, registry);
			}

			writer.WriteEndObject();
		}

		private static MergeTree.PropertySet ReadPropertySet(ref Utf8JsonReader reader, string propertyName, IFluidDataObjectRegistry? registry)
		{
			if (reader.TokenType != JsonTokenType.StartObject)
			{
				throw new JsonException($"Expected object value for property '{propertyName}'.");
			}

			MergeTree.PropertySet properties = new MergeTree.PropertySet();

			while (reader.Read())
			{
				if (reader.TokenType == JsonTokenType.EndObject)
				{
					return properties;
				}

				if (reader.TokenType != JsonTokenType.PropertyName)
				{
					throw new JsonException($"Expected property name in '{propertyName}'.");
				}

				string key = reader.GetString() ?? string.Empty;
				if (!reader.Read())
				{
					throw new JsonException($"Expected value for property '{key}'.");
				}

				properties[key] = ReadJsonValue(ref reader, registry);
			}

			throw new JsonException($"Unexpected end of property '{propertyName}' object.");
		}

		private static void WriteJsonValue(Utf8JsonWriter writer, object? value, IFluidDataObjectRegistry? registry)
		{
			object? serializableValue = HandleWireFormat.MakeHandlesSerializable(value, registry, _missingRegistryMessage);
			switch (serializableValue)
			{
				case null:
					writer.WriteNullValue();
					break;

				case JsonElement jsonElement:
					jsonElement.WriteTo(writer);
					break;

				case string stringValue:
					writer.WriteStringValue(stringValue);
					break;

				case bool boolValue:
					writer.WriteBooleanValue(boolValue);
					break;

				case int intValue:
					writer.WriteNumberValue(intValue);
					break;

				case long longValue:
					writer.WriteNumberValue(longValue);
					break;

				case double doubleValue:
					writer.WriteNumberValue(doubleValue);
					break;

				case float floatValue:
					writer.WriteNumberValue(floatValue);
					break;

				case decimal decimalValue:
					writer.WriteNumberValue(decimalValue);
					break;

				// TS ref: packages/dds/merge-tree/src/properties.ts PropertySet —
				// TS treats all numeric properties as ordinary JSON numbers. CLR
				// small-integer types (short, byte, sbyte, ushort, uint, ulong)
				// represent the same value space; widen them to the appropriate
				// JSON number instead of throwing after the local tree has
				// already been mutated.
				case short shortValue:
					writer.WriteNumberValue(shortValue);
					break;

				case ushort ushortValue:
					writer.WriteNumberValue(ushortValue);
					break;

				case byte byteValue:
					writer.WriteNumberValue(byteValue);
					break;

				case sbyte sbyteValue:
					writer.WriteNumberValue(sbyteValue);
					break;

				case uint uintValue:
					writer.WriteNumberValue(uintValue);
					break;

				case ulong ulongValue:
					writer.WriteNumberValue(ulongValue);
					break;

				case IReadOnlyDictionary<string, object?> dictionary:
					WritePropertySet(writer, dictionary, registry);
					break;

				case IDictionary dictionary:
					writer.WriteStartObject();
					foreach (DictionaryEntry entry in dictionary)
					{
						if (entry.Key is not string key)
						{
							throw new NotSupportedException("JSON property dictionaries must use string keys.");
						}

						writer.WritePropertyName(key);
						WriteJsonValue(writer, entry.Value, registry);
					}

					writer.WriteEndObject();
					break;

				case IEnumerable values:
					writer.WriteStartArray();
					foreach (object? item in values)
					{
						WriteJsonValue(writer, item, registry);
					}

					writer.WriteEndArray();
					break;

				default:
					throw new NotSupportedException($"Unsupported JSON property value type: {serializableValue.GetType().FullName}");
			}
		}

		private static object? ReadJsonValue(ref Utf8JsonReader reader, IFluidDataObjectRegistry? registry)
		{
			switch (reader.TokenType)
			{
				case JsonTokenType.Null:
					return null;

				case JsonTokenType.String:
					return reader.GetString();

				case JsonTokenType.Number:
					if (reader.TryGetInt32(out int intValue))
					{
						return intValue;
					}

					if (reader.TryGetInt64(out long longValue))
					{
						return longValue;
					}

					return reader.GetDouble();

				case JsonTokenType.True:
					return true;

				case JsonTokenType.False:
					return false;

				case JsonTokenType.StartObject:
					return ReadJsonObjectValue(ref reader, registry);

				case JsonTokenType.StartArray:
					return ReadJsonArray(ref reader, registry);

				default:
					throw new JsonException("Unexpected JSON value.");
			}
		}

		private static object? ReadJsonObjectValue(ref Utf8JsonReader reader, IFluidDataObjectRegistry? registry)
		{
			MergeTree.PropertySet properties = ReadPropertySet(ref reader, "object", registry);
			if (TryReadHandleUrl(properties, out string url, out bool payloadPending))
			{
				return HandleWireFormat.ResolveSerializedHandle(url, registry, payloadPending);
			}

			// If the type marker is present but url is missing/invalid, this is
			// a malformed handle. TS identifies handles by the type marker alone
			// and would fail on dereference; reject here rather than silently
			// pass the object through as ordinary property data.
			if (properties.TryGetValue(HandleWireFormat.TypePropertyName, out object? typeMarker)
				&& typeMarker is string typeMarkerStr
				&& typeMarkerStr == HandleWireFormat.SerializedHandleTypeName)
			{
				throw new OcsException(
					OcsGateErrorCode.InvalidOperation,
					"Serialized Fluid handle is missing required 'url' property.");
			}

			return properties;
		}

		private static List<object?> ReadJsonArray(ref Utf8JsonReader reader, IFluidDataObjectRegistry? registry)
		{
			if (reader.TokenType != JsonTokenType.StartArray)
			{
				throw new JsonException("Expected JSON array.");
			}

			List<object?> values = new List<object?>();

			while (reader.Read())
			{
				if (reader.TokenType == JsonTokenType.EndArray)
				{
					return values;
				}

				values.Add(ReadJsonValue(ref reader, registry));
			}

			throw new JsonException("Unexpected end of JSON array.");
		}

		private static bool TryReadHandleUrl(IReadOnlyDictionary<string, object?> properties, out string url, out bool payloadPending)
		{
			if (properties.TryGetValue(HandleWireFormat.TypePropertyName, out object? typeValue)
				&& typeValue is string typeString
				&& typeString == HandleWireFormat.SerializedHandleTypeName
				&& properties.TryGetValue(HandleWireFormat.UrlPropertyName, out object? urlValue)
				&& urlValue is string urlString)
			{
				url = urlString;
				// payloadPending is optional and only ever set to true; any other shape is "not pending".
				payloadPending = properties.TryGetValue(HandleWireFormat.PayloadPendingPropertyName, out object? pendingValue)
					&& pendingValue is bool pendingBool
					&& pendingBool;
				return true;
			}

			url = string.Empty;
			payloadPending = false;
			return false;
		}

		private static int ReadOperationType(ref Utf8JsonReader reader, string propertyName)
		{
			if (reader.TokenType == JsonTokenType.Number && reader.TryGetInt32(out int value))
			{
				return value;
			}

			if (reader.TokenType == JsonTokenType.String)
			{
				string? typeName = reader.GetString();
				return typeName switch
				{
					"insert" => (int)MergeTree.MergeTreeDeltaType.Insert,
					"remove" => (int)MergeTree.MergeTreeDeltaType.Remove,
					"annotate" => (int)MergeTree.MergeTreeDeltaType.Annotate,
					"group" => (int)MergeTree.MergeTreeDeltaType.Group,
					"obliterate" => (int)MergeTree.MergeTreeDeltaType.Obliterate,
					"obliterateSided" => (int)MergeTree.MergeTreeDeltaType.ObliterateSided,
					_ => throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unknown merge-tree op type: {typeName ?? "<missing>"}"),
				};
			}

			throw new JsonException($"Expected int or string value for property '{propertyName}'.");
		}

		private static int ReadInt32(ref Utf8JsonReader reader, string propertyName)
		{
			if (reader.TokenType != JsonTokenType.Number || !reader.TryGetInt32(out int value))
			{
				throw new JsonException($"Expected int value for property '{propertyName}'.");
			}

			return value;
		}

		private static bool ReadBoolean(ref Utf8JsonReader reader, string propertyName)
		{
			if (reader.TokenType == JsonTokenType.True)
			{
				return true;
			}

			if (reader.TokenType == JsonTokenType.False)
			{
				return false;
			}

			throw new JsonException($"Expected bool value for property '{propertyName}'.");
		}

		private static int? ReadNullableInt32(ref Utf8JsonReader reader, string propertyName)
		{
			if (reader.TokenType == JsonTokenType.Null)
			{
				return null;
			}

			return ReadInt32(ref reader, propertyName);
		}

		// TS ref: packages/dds/merge-tree/src/sequencePlace.ts normalizePlace —
		// TS interval endpoints on the wire may be a number OR a string sentinel:
		//   "start" → { pos: -1, side: Side.After }   (points before position 0)
		//   "end"   → { pos: -1, side: Side.Before }  (points after last position)
		// The port previously accepted only numbers, throwing on valid TS wire
		// with sentinel endpoints. Reads either a number (via ReadNullableInt32)
		// or one of the two sentinel strings; on sentinel returns (-1, impliedSide).
		// impliedSide is null when the value was a plain number.
		private static (int? position, int? impliedSide) ReadEndpointPositionOrSentinel(ref Utf8JsonReader reader, string propertyName)
		{
			if (reader.TokenType == JsonTokenType.String)
			{
				string? sentinel = reader.GetString();
				if (sentinel == "start")
				{
					return (-1, (int)Intervals.Side.After);
				}

				if (sentinel == "end")
				{
					return (-1, (int)Intervals.Side.Before);
				}

				throw new JsonException($"Invalid endpoint sentinel '{sentinel}' for property '{propertyName}'.");
			}

			return (ReadNullableInt32(ref reader, propertyName), null);
		}

		private static int RequireInt32(int? value, string propertyName)
		{
			if (!value.HasValue)
			{
				throw new JsonException($"Expected int value for property '{propertyName}'.");
			}

			return value.Value;
		}

		private static void RequirePositionOrRelative(int? position, object? relativePosition, string positionPropertyName, string relativePositionPropertyName)
		{
			if (!position.HasValue && relativePosition is null)
			{
				throw new JsonException($"Expected '{positionPropertyName}' or '{relativePositionPropertyName}' on merge-tree operation.");
			}
		}

		private static bool RequireBoolean(bool? value, string propertyName)
		{
			if (!value.HasValue)
			{
				throw new JsonException($"Expected bool value for property '{propertyName}'.");
			}

			return value.Value;
		}

		private static MergeTree.SequencePlace RequireSequencePlace(MergeTree.SequencePlace? value, string propertyName)
		{
			if (value is null)
			{
				throw new JsonException($"Expected sequence place value for property '{propertyName}'.");
			}

			return value;
		}

		private static void RequireIntervalOpKind(int? actual, MergeTree.IntervalOpKind expected)
		{
			if (actual != (int)expected)
			{
				throw new OcsException(
					OcsGateErrorCode.UnknownOp,
					$"Unknown interval op kind for type {(int)expected}: {actual?.ToString() ?? "<missing>"}");
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

		private sealed record IntervalMapOperationValue(string OpName, IntervalMapPayload? Payload);

		private sealed record IntervalMapPayload(
			int? Start,
			bool HasStart,
			int? End,
			bool HasEnd,
			int? IntervalType,
			int? Stickiness,
			int? StartSide,
			int? EndSide,
			MergeTree.PropertySet? Properties);
	}
}
