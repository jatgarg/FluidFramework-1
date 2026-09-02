// -----------------------------------------------------------------------------
// Ported from packages/dds/merge-tree/src/mergeTreeNodes.ts + referencePositions.ts
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Microsoft.Office.Web.Fluid.MergeTree
{
    [Flags]
    public enum ReferenceType
    {
        Simple = 0x0,
        Tile = 0x1,
        RangeBegin = 0x10,
        RangeEnd = 0x20,
        SlideOnRemove = 0x40,
        StayOnRemove = 0x80,
        Transient = 0x100,
    }

    public interface IJSONMarkerSegment : IJSONSegment
    {
        JSONMarkerDef Marker { get; set; }

        PropertySet? Props { get; set; }
    }

    public sealed class JSONMarkerDef
    {
        [JsonPropertyName("refType")]
        public ReferenceType RefType { get; set; } = ReferenceType.Simple;

        [JsonPropertyName("props")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public PropertySet? Props { get; set; }
    }

    public sealed class JSONMarkerSegment : IJSONMarkerSegment
    {
        [JsonPropertyName("marker")]
        public JSONMarkerDef Marker { get; set; } = new();

        [JsonPropertyName("props")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public PropertySet? Props { get; set; }
    }

    public sealed class Marker : BaseSegment
    {
        public const string TypeName = "Marker";

        public const string reservedMarkerIdKey = "markerId";

        public const string reservedTileLabelsKey = "referenceTileLabels";

        public const string ReservedMarkerIdKey = reservedMarkerIdKey;

        public const string ReservedTileLabelsKey = reservedTileLabelsKey;

        public Marker(ReferenceType refType, PropertySet? props = null)
            : base(DeepCloneProperties(props))
        {
            RefType = refType;
            CachedLength = 1;
        }

        public override string Type => TypeName;

        public ReferenceType RefType { get; }

        public static bool Is(ISegment? segment)
        {
            return segment is Marker marker && marker.Type == TypeName;
        }

        public static Marker Make(ReferenceType refType, PropertySet? props = null)
        {
            return new Marker(refType, props);
        }

        public static Marker? FromJSONObject(object? spec)
        {
            switch (spec)
            {
                case Marker marker:
                    return marker.Clone();

                case IJSONMarkerSegment markerSpec:
                    return Make(markerSpec.Marker.RefType, markerSpec.Props ?? markerSpec.Marker.Props);

                case JsonElement jsonElement:
                    return FromJsonElement(jsonElement);

                default:
                    return null;
            }
        }

        public string? GetId()
        {
            return Properties is not null
                && Properties.TryGetValue(reservedMarkerIdKey, out object? id)
                ? id as string
                : null;
        }

        public bool HasTileLabel(string tileLabel)
        {
            if (!RefTypeIncludesFlag(ReferenceType.Tile) || tileLabel is null)
            {
                return false;
            }

            foreach (string label in GetTileLabels())
            {
                if (string.Equals(label, tileLabel, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        public bool RefTypeIncludesFlag(ReferenceType flags)
        {
            return (RefType & flags) != 0;
        }

        public IEnumerable<string> GetTileLabels()
        {
            if (Properties is null
                || !Properties.TryGetValue(reservedTileLabelsKey, out object? labels)
                || labels is null)
            {
                yield break;
            }

            if (labels is string singleLabel)
            {
                yield return singleLabel;
                yield break;
            }

            if (labels is JsonElement jsonElement)
            {
                if (jsonElement.ValueKind == JsonValueKind.String)
                {
                    yield return jsonElement.GetString() ?? string.Empty;
                }
                else if (jsonElement.ValueKind == JsonValueKind.Array)
                {
                    foreach (JsonElement item in jsonElement.EnumerateArray())
                    {
                        if (item.ValueKind == JsonValueKind.String)
                        {
                            yield return item.GetString() ?? string.Empty;
                        }
                    }
                }

                yield break;
            }

            if (labels is IEnumerable<string> stringLabels)
            {
                foreach (string label in stringLabels)
                {
                    yield return label;
                }

                yield break;
            }

            if (labels is IEnumerable enumerable)
            {
                foreach (object? item in enumerable)
                {
                    if (item is string label)
                    {
                        yield return label;
                    }
                }
            }
        }

        public override object ToJSONObject()
        {
            return new JSONMarkerSegment()
            {
                Marker = new JSONMarkerDef()
                {
                    RefType = RefType,
                },
                Props = DeepCloneProperties(Properties),
            };
        }

        public override Marker Clone()
        {
            Marker clone = Make(RefType, Properties);
            CopyMetadataTo(clone);
            clone.Properties = DeepCloneProperties(Properties);
            return clone;
        }

        public override bool CanAppend(ISegment segment)
        {
            return false;
        }

        public override void Append(ISegment segment)
        {
            throw new LoggingError("Can not append to marker.");
        }

        public override string ToString()
        {
            return "\u200E";
        }

        protected override BaseSegment? CreateSplitSegmentAt(int pos)
        {
            return null;
        }

        private static Marker? FromJsonElement(JsonElement jsonElement)
        {
            if (jsonElement.ValueKind != JsonValueKind.Object
                || !jsonElement.TryGetProperty("marker", out JsonElement markerElement))
            {
                return null;
            }

            ReferenceType refType = ReferenceType.Simple;
            PropertySet? props = null;
            if (markerElement.ValueKind == JsonValueKind.Object)
            {
                if (markerElement.TryGetProperty("refType", out JsonElement refTypeElement)
                    && refTypeElement.ValueKind == JsonValueKind.Number
                    && refTypeElement.TryGetInt32(out int refTypeValue))
                {
                    refType = (ReferenceType)refTypeValue;
                }

                if (markerElement.TryGetProperty("props", out JsonElement nestedPropsElement)
                    && nestedPropsElement.ValueKind != JsonValueKind.Null
                    && nestedPropsElement.ValueKind != JsonValueKind.Undefined)
                {
                    props = JsonElementToPropertySet(nestedPropsElement);
                }
            }

            if (jsonElement.TryGetProperty("props", out JsonElement propsElement)
                && propsElement.ValueKind != JsonValueKind.Null
                && propsElement.ValueKind != JsonValueKind.Undefined)
            {
                props = JsonElementToPropertySet(propsElement);
            }

            return Make(refType, props);
        }

        private static PropertySet? DeepCloneProperties(IReadOnlyDictionary<string, object?>? properties)
        {
            if (properties is null)
            {
                return null;
            }

            PropertySet clone = new();
            foreach (KeyValuePair<string, object?> property in properties)
            {
                clone[property.Key] = DeepCloneValue(property.Value);
            }

            return clone;
        }

        private static object? DeepCloneValue(object? value)
        {
            switch (value)
            {
                case null:
                    return null;

                case JsonElement jsonElement:
                    return JsonElementToObject(jsonElement);

                case string:
                case bool:
                case byte:
                case sbyte:
                case short:
                case ushort:
                case int:
                case uint:
                case long:
                case ulong:
                case float:
                case double:
                case decimal:
                    return value;

                case IReadOnlyDictionary<string, object?> dictionary:
                    return DeepCloneProperties(dictionary);

                case IDictionary nonGenericDictionary:
                    PropertySet dictionaryClone = new();
                    foreach (DictionaryEntry entry in nonGenericDictionary)
                    {
                        if (entry.Key is string key)
                        {
                            dictionaryClone[key] = DeepCloneValue(entry.Value);
                        }
                    }

                    return dictionaryClone;

                case IEnumerable enumerable:
                    List<object?> list = new();
                    foreach (object? item in enumerable)
                    {
                        list.Add(DeepCloneValue(item));
                    }

                    return list;

                default:
                    return value;
            }
        }

        private static PropertySet JsonElementToPropertySet(JsonElement jsonElement)
        {
            PropertySet properties = new();
            if (jsonElement.ValueKind != JsonValueKind.Object)
            {
                return properties;
            }

            foreach (JsonProperty property in jsonElement.EnumerateObject())
            {
                properties[property.Name] = JsonElementToObject(property.Value);
            }

            return properties;
        }

        private static object? JsonElementToObject(JsonElement jsonElement)
        {
            switch (jsonElement.ValueKind)
            {
                case JsonValueKind.Null:
                case JsonValueKind.Undefined:
                    return null;

                case JsonValueKind.String:
                    return jsonElement.GetString();

                case JsonValueKind.Number:
                    if (jsonElement.TryGetInt32(out int intValue))
                    {
                        return intValue;
                    }

                    if (jsonElement.TryGetInt64(out long longValue))
                    {
                        return longValue;
                    }

                    return jsonElement.GetDouble();

                case JsonValueKind.True:
                    return true;

                case JsonValueKind.False:
                    return false;

                case JsonValueKind.Object:
                    return JsonElementToPropertySet(jsonElement);

                case JsonValueKind.Array:
                    List<object?> values = new();
                    foreach (JsonElement item in jsonElement.EnumerateArray())
                    {
                        values.Add(JsonElementToObject(item));
                    }

                    return values;

                default:
                    return null;
            }
        }
    }
}
