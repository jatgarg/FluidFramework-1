// -----------------------------------------------------------------------------
// Ported from packages/dds/merge-tree/src/properties.ts
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections;
using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid.MergeTree
{
    /// <summary>
    /// A loosely-typed mapping from strings to any value.
    /// </summary>
    /// <remarks>
    /// Property sets are expected to be JSON-serializable.
    /// </remarks>
    public sealed class PropertySet : Dictionary<string, object?>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PropertySet" /> class.
        /// </summary>
        public PropertySet()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PropertySet" /> class from an existing map.
        /// </summary>
        /// <param name="dictionary">The properties to copy.</param>
        public PropertySet(IDictionary<string, object?> dictionary)
            : base(dictionary)
        {
        }
    }

    /// <summary>
    /// Utility methods for string-keyed property maps.
    /// </summary>
    public static class PropertyMap
    {
        /// <summary>
        /// Compares two property sets for equality.
        /// </summary>
        /// <param name="a">The first property set.</param>
        /// <param name="b">The second property set.</param>
        /// <returns><see langword="true" /> when the property sets contain the same keys and deeply equal values.</returns>
        public static bool MatchProperties(
            IReadOnlyDictionary<string, object?>? a,
            IReadOnlyDictionary<string, object?>? b)
        {
            int countA = a?.Count ?? 0;
            int countB = b?.Count ?? 0;
            if (countA != countB)
            {
                return false;
            }

            if (a is null)
            {
                return true;
            }

            foreach (KeyValuePair<string, object?> property in a)
            {
                if (b is null || !b.TryGetValue(property.Key, out object? valueB))
                {
                    return false;
                }

                if (!MatchPropertyValue(property.Value, valueB))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Adds properties from one map to another.
        /// </summary>
        /// <typeparam name="T">The property value type.</typeparam>
        /// <param name="baseMap">The map to update.</param>
        /// <param name="extension">The properties to apply.</param>
        /// <returns>The updated <paramref name="baseMap" />.</returns>
        /// <remarks>
        /// A <see langword="null" /> value in <paramref name="extension" /> removes that key from
        /// <paramref name="baseMap" />, matching the TypeScript port's deletion behavior.
        /// </remarks>
        public static Dictionary<string, T> Extend<T>(
            Dictionary<string, T> baseMap,
            IReadOnlyDictionary<string, T>? extension)
        {
            if (extension is not null)
            {
                foreach (KeyValuePair<string, T> property in extension)
                {
                    if (property.Value is null)
                    {
                        baseMap.Remove(property.Key);
                    }
                    else
                    {
                        baseMap[property.Key] = property.Value;
                    }
                }
            }

            return baseMap;
        }

        /// <summary>
        /// Clones properties in a given map into a new map.
        /// </summary>
        /// <typeparam name="T">The property value type.</typeparam>
        /// <param name="extension">The properties to clone.</param>
        /// <returns>A shallow clone of <paramref name="extension" />, or <see langword="null" /> when it is absent.</returns>
        public static Dictionary<string, T>? Clone<T>(IReadOnlyDictionary<string, T>? extension)
        {
            if (extension is null)
            {
                return null;
            }

            Dictionary<string, T> cloneMap = CreateMap<T>();
            return Extend(cloneMap, extension);
        }

        /// <summary>
        /// Clones properties in a given property set into a new property set.
        /// </summary>
        /// <param name="extension">The property set to clone.</param>
        /// <returns>A shallow clone of <paramref name="extension" />, or <see langword="null" /> when it is absent.</returns>
        public static PropertySet? ClonePropertySet(IReadOnlyDictionary<string, object?>? extension)
        {
            if (extension is null)
            {
                return null;
            }

            PropertySet cloneMap = new();
            Extend(cloneMap, extension);
            return cloneMap;
        }

        // Deep-clone a property set. Nested dictionaries and lists are
        // cloned recursively so a listener holding the clone cannot mutate
        // structures other consumers still reference. Scalars, strings,
        // and IFluidDataObject handles are shared by reference (effectively
        // immutable from a listener's perspective).
        public static PropertySet? DeepClonePropertySet(IReadOnlyDictionary<string, object?>? extension)
        {
            if (extension is null)
            {
                return null;
            }

            PropertySet cloneMap = new();
            foreach (KeyValuePair<string, object?> property in extension)
            {
                cloneMap[property.Key] = DeepCloneValue(property.Value);
            }

            return cloneMap;
        }

        private static object? DeepCloneValue(object? value)
        {
            switch (value)
            {
                case null:
                case string:
                case bool:
                case System.ValueType:
                    return value;
                case IReadOnlyDictionary<string, object?> nested:
                    return DeepClonePropertySet(nested);
                case IList list:
                    List<object?> clonedList = new(list.Count);
                    foreach (object? item in list)
                    {
                        clonedList.Add(DeepCloneValue(item));
                    }

                    return clonedList;
                default:
                    return value;
            }
        }

        /// <summary>
        /// Adds properties in one property set to another property set. If the property set being added to does not exist, creates one.
        /// </summary>
        /// <param name="oldProperties">The property set to update, or <see langword="null" /> to create one.</param>
        /// <param name="newProperties">The properties to add.</param>
        /// <returns>The updated property set.</returns>
        public static PropertySet AddProperties(
            PropertySet? oldProperties,
            IReadOnlyDictionary<string, object?> newProperties)
        {
            PropertySet properties = oldProperties ?? new PropertySet();
            Extend(properties, newProperties);
            return properties;
        }

        /// <summary>
        /// Replaces missing values in one map with values for the same key from another map.
        /// </summary>
        /// <typeparam name="T">The property value type.</typeparam>
        /// <param name="baseMap">The map to update.</param>
        /// <param name="extension">The fallback properties to apply.</param>
        /// <returns>The updated <paramref name="baseMap" />.</returns>
        public static Dictionary<string, T> ExtendIfUndefined<T>(
            Dictionary<string, T> baseMap,
            IReadOnlyDictionary<string, T>? extension)
        {
            if (extension is not null)
            {
                foreach (KeyValuePair<string, T> property in extension)
                {
                    if (!baseMap.ContainsKey(property.Key))
                    {
                        baseMap[property.Key] = property.Value;
                    }
                }
            }

            return baseMap;
        }

        /// <summary>
        /// Combines two maps into a new map.
        /// </summary>
        /// <typeparam name="T">The property value type.</typeparam>
        /// <param name="baseMap">The base properties to copy.</param>
        /// <param name="extension">The properties to apply over <paramref name="baseMap" />.</param>
        /// <returns>A new map containing the combined properties.</returns>
        public static Dictionary<string, T> Combine<T>(
            IReadOnlyDictionary<string, T>? baseMap,
            IReadOnlyDictionary<string, T>? extension)
        {
            Dictionary<string, T> combined = CreateMap<T>();
            Extend(combined, baseMap);
            Extend(combined, extension);
            return combined;
        }

        /// <summary>
        /// Creates a string-keyed map with good performance.
        /// </summary>
        /// <typeparam name="T">The property value type.</typeparam>
        /// <returns>A new map from strings to values of type <typeparamref name="T" />.</returns>
        public static Dictionary<string, T> CreateMap<T>()
        {
            return new Dictionary<string, T>();
        }

        private static bool MatchPropertyValue(object? valueA, object? valueB)
        {
            if (ReferenceEquals(valueA, valueB))
            {
                return true;
            }

            // Fast paths for common primitive/string values — avoid the expensive
            // dictionary/list shape probes below when neither operand needs them.
            if (valueA is null || valueB is null)
            {
                return valueA is null && valueB is null;
            }

            if (valueA is string || valueB is string)
            {
                return valueA is string && valueB is string && (string)valueA == (string)valueB;
            }

            // Wire-equivalent numeric values must compare equal regardless
            // of CLR runtime type. TS `===` on numbers compares values, not
            // types; int/long/double/JsonElement encoding the same JSON
            // number should not split delta ranges. Unify at `double`.
            if (TryReadNumeric(valueA, out double numericA) && TryReadNumeric(valueB, out double numericB))
            {
                return numericA == numericB;
            }

            if (valueA is System.ValueType && valueB is System.ValueType)
            {
                return valueA.Equals(valueB);
            }

            bool isDictionaryA = TryAsPropertyDictionary(valueA, out IReadOnlyDictionary<string, object?>? dictionaryA);
            bool isDictionaryB = TryAsPropertyDictionary(valueB, out IReadOnlyDictionary<string, object?>? dictionaryB);
            if (isDictionaryA || isDictionaryB)
            {
                return isDictionaryA && isDictionaryB && MatchProperties(dictionaryA, dictionaryB);
            }

            // TS matchProperties deep-compares array/object property values.
            // Object.is on different array references returns false, so JS
            // element-wise compares. Match that so equal-content lists
            // don't report unequal via CLR object-identity fallback.
            IList? listA = valueA as IList;
            IList? listB = valueB as IList;
            if (listA is not null || listB is not null)
            {
                return listA is not null && listB is not null && MatchPropertyLists(listA, listB);
            }

            return Equals(valueA, valueB);
        }

        private static bool TryReadNumeric(object value, out double numeric)
        {
            // Ordering doesn't affect correctness: each case is a type pattern
            // (case int i:, case long l:, etc.) that matches by exact runtime
            // type. Primitive numeric types are disjoint at runtime — a boxed
            // long never matches `case int i:` no matter where either sits.
            // The bool case is an explicit rejection: bool is IConvertible and
            // would silently coerce to 0/1 in some numeric contexts, but here
            // we want a hard "no" so downstream MatchProperties doesn't treat
            // `true` and `1` as equal.
            switch (value)
            {
                case bool:
                    numeric = 0;
                    return false;
                case byte b:
                    numeric = b;
                    return true;
                case sbyte sb:
                    numeric = sb;
                    return true;
                case short s:
                    numeric = s;
                    return true;
                case ushort us:
                    numeric = us;
                    return true;
                case int i:
                    numeric = i;
                    return true;
                case uint ui:
                    numeric = ui;
                    return true;
                case long l:
                    numeric = l;
                    return true;
                case ulong ul:
                    numeric = ul;
                    return true;
                case float f:
                    numeric = f;
                    return true;
                case double d:
                    numeric = d;
                    return true;
                case decimal dec:
                    numeric = (double)dec;
                    return true;
                case System.Text.Json.JsonElement je when je.ValueKind == System.Text.Json.JsonValueKind.Number:
                    if (je.TryGetDouble(out double val))
                    {
                        numeric = val;
                        return true;
                    }

                    numeric = 0;
                    return false;
                default:
                    numeric = 0;
                    return false;
            }
        }

        private static bool MatchPropertyLists(IList a, IList b)
        {
            if (a.Count != b.Count)
            {
                return false;
            }

            for (int i = 0; i < a.Count; i++)
            {
                if (!MatchPropertyValue(a[i], b[i]))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool TryAsPropertyDictionary(
            object? value,
            out IReadOnlyDictionary<string, object?>? dictionary)
        {
            if (value is IReadOnlyDictionary<string, object?> propertyDictionary)
            {
                dictionary = propertyDictionary;
                return true;
            }

            if (value is IDictionary nonGenericDictionary)
            {
                Dictionary<string, object?> converted = new();
                foreach (DictionaryEntry entry in nonGenericDictionary)
                {
                    if (entry.Key is not string key)
                    {
                        dictionary = null;
                        return false;
                    }

                    converted[key] = entry.Value;
                }

                dictionary = converted;
                return true;
            }

            dictionary = null;
            return false;
        }
    }
}
