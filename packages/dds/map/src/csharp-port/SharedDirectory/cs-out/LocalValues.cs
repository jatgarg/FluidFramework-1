// -----------------------------------------------------------------------------
// Ported from packages/dds/map/src/localValues.ts
// Part of the SharedDirectory C# feasibility port.
// Plain values and Fluid-handle values serialize via DirectoryOpSerializer.
// Legacy Shared-value migration remains stubbed.
// -----------------------------------------------------------------------------

#nullable enable

using System;

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// A local value to be stored in a container-type DDS.
	/// </summary>
	public interface ILocalValue
	{
		/// <summary>The in-memory value stored within.</summary>
		object? Value { get; }
	}

	/// <summary>Type-annotation values used in wire format ("Plain", "Shared").</summary>
	public static class ValueType
	{
		public const string Plain = "Plain";
		public const string Shared = "Shared"; // legacy (pre-handles) — see migrateIfSharedSerializable
	}

	public static class LocalValues
	{
		/// <summary>
		/// Convert a value to its serialized form, i.e. to be used in ops and summaries.
		/// </summary>
		public static SerializedValue SerializeValue(object? value, IFluidDataObjectRegistry? registry = null)
		{
			return new SerializedValue()
			{
				Type = ValueType.Plain,
				Value = DirectoryOpSerializer.MakeHandlesSerializable(value, registry),
			};
		}

		/// <summary>
		/// Very old versions of Fluid permitted a different type of stored value which represented
		/// a SharedObject held directly. This functionality has since been replaced with handles.
		/// If the passed value is one of those old values, mutate it in-place to a modern handle
		/// value; otherwise no-op.
		/// </summary>
		public static void MigrateIfSharedSerializable(SerializableValue serializable)
		{
			if (serializable.Type == ValueType.Shared)
			{
				// TODO(wave2): Build the legacy handle wire payload if WN needs pre-handle documents.
				throw new NotImplementedException(
					"LocalValues.MigrateIfSharedSerializable legacy Shared-value migration is a stub. Wave 2.");
			}
		}
	}
}
