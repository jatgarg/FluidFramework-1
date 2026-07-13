// -----------------------------------------------------------------------------
// Ported from packages/dds/map/src/interfaces.ts and internalInterfaces.ts
// Op discriminated union ported from the top of packages/dds/map/src/directory.ts
// Part of the SharedDirectory C# feasibility port.
// -----------------------------------------------------------------------------

using System;
using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid
{
	// -------------------------------------------------------------
	// Event args + delegates (match host FluidSharedMap.cs pattern)
	// -------------------------------------------------------------

	/// <summary>
	/// Type of valueChanged event parameter.
	/// </summary>
	public class ValueChangedEventArgs
	{
		/// <summary>
		/// The key storing the value that changed.
		/// </summary>
		public string Key { get; set; } = string.Empty;

		/// <summary>
		/// The value that was stored at the key prior to the change.
		/// </summary>
		public object? PreviousValue { get; set; }

		/// <summary>
		/// The absolute path to the IDirectory storing the key which changed.
		/// </summary>
		public string Path { get; set; } = string.Empty; // absolute path of the containing subdirectory

		/// <summary>
		/// True if this change was caused by a local mutation; false if it was applied from a remote op.
		/// </summary>
		public bool Local { get; set; }
	}

	/// <summary>
	/// Handles value changed events.
	/// </summary>
	/// <param name="sender">The event sender.</param>
	/// <param name="e">Information on the key that changed and its value prior to the change.</param>
	public delegate void ValueChangedEventHandler(object sender, ValueChangedEventArgs e);

	/// <summary>
	/// Type of subdirectory event parameter.
	/// </summary>
	public class SubDirectoryEventArgs
	{
		/// <summary>
		/// The name of the subdirectory that was created or deleted.
		/// </summary>
		public string SubdirName { get; set; } = string.Empty;

		/// <summary>
		/// The absolute path of the parent directory.
		/// </summary>
		public string ParentPath { get; set; } = string.Empty; // absolute path of parent

		/// <summary>
		/// True if this change was caused by a local mutation; false if it was applied from a remote op.
		/// </summary>
		public bool Local { get; set; }
	}

	/// <summary>
	/// Handles subdirectory created or deleted events.
	/// </summary>
	/// <param name="sender">The event sender.</param>
	/// <param name="e">Information on the subdirectory that was created or deleted.</param>
	public delegate void SubDirectoryEventHandler(object sender, SubDirectoryEventArgs e);

	// -------------------------------------------------------------
	// Public IDirectory / ISharedDirectory
	// -------------------------------------------------------------

	/// <summary>Interface describing actions on a directory.</summary>
	public interface IDirectory : IEnumerable<KeyValuePair<string, object?>>
	{
		/// <summary>The absolute path of the directory (posix-style, rooted at "/").</summary>
		string AbsolutePath { get; }

		// Key/value APIs (mirror TS Map<string, any>)
		/// <summary>
		/// Retrieves the value stored at the given key from the directory.
		/// </summary>
		/// <param name="key">Key to retrieve from.</param>
		/// <returns>The stored value, or null if the key is not set.</returns>
		object? Get(string key);

		/// <summary>
		/// Sets the value stored at key to the provided value.
		/// </summary>
		/// <param name="key">Key to set at.</param>
		/// <param name="value">Value to set.</param>
		/// <returns>The IDirectory itself.</returns>
		IDirectory Set(string key, object? value);

		/// <summary>
		/// Checks whether the directory has a value for the given key.
		/// </summary>
		/// <param name="key">Key to check.</param>
		/// <returns>True if the key exists; otherwise, false.</returns>
		bool Has(string key);

		/// <summary>
		/// Removes the specified element from the directory by its key.
		/// </summary>
		/// <param name="key">Key to delete.</param>
		/// <returns>True if an element existed and has been removed, or false if the element does not exist.</returns>
		bool Delete(string key);

		/// <summary>
		/// Removes all entries from the directory.
		/// </summary>
		void Clear();

		/// <summary>
		/// Gets the keys stored in the directory.
		/// </summary>
		IReadOnlyCollection<string> Keys { get; }

		/// <summary>
		/// Gets the values stored in the directory. Order is undefined.
		/// </summary>
		IReadOnlyCollection<object?> Values { get; }

		/// <summary>
		/// Returns an enumeration of key/value pairs stored in the directory.
		/// This is equivalent to iterating <see cref="IDirectory" /> directly
		/// via <c>foreach</c>.
		/// </summary>
		IEnumerable<KeyValuePair<string, object?>> Entries();

		/// <summary>
		/// Gets the number of key/value entries within the directory.
		/// </summary>
		int Count { get; }

		// Subdirectory APIs
		/// <summary>
		/// Creates an IDirectory child of this IDirectory, or retrieves the existing IDirectory child if one with the
		/// same name already exists.
		/// </summary>
		/// <param name="subdirName">Name of the new child directory to create.</param>
		/// <returns>The IDirectory child that was created or retrieved.</returns>
		IDirectory CreateSubDirectory(string subdirName);

		/// <summary>
		/// Gets an IDirectory child of this IDirectory, if it exists.
		/// </summary>
		/// <param name="subdirName">Name of the child directory to get.</param>
		/// <returns>The requested IDirectory, or null if it does not exist.</returns>
		IDirectory? GetSubDirectory(string subdirName);

		/// <summary>
		/// Checks whether this directory has a child directory with the given name.
		/// </summary>
		/// <param name="subdirName">Name of the child directory to check.</param>
		/// <returns>True if it exists, false otherwise.</returns>
		bool HasSubDirectory(string subdirName);

		/// <summary>
		/// Deletes an IDirectory child of this IDirectory, if it exists, along with all descendent keys and directories.
		/// </summary>
		/// <param name="subdirName">Name of the child directory to delete.</param>
		/// <returns>True if the IDirectory existed and was deleted, false if it did not exist.</returns>
		bool DeleteSubDirectory(string subdirName);

		/// <summary>
		/// Get the number of sub directory within the directory.
		/// </summary>
		/// <returns>The number of sub directory within a directory.</returns>
		int CountSubDirectory();

		/// <summary>Iterates child subdirectories in creation order.</summary>
		/// <returns>The IDirectory iterator.</returns>
		IEnumerable<KeyValuePair<string, IDirectory>> SubDirectories();

		// Path-based navigation (posix-style relative paths)
		/// <summary>
		/// Get an IDirectory within the directory, in order to use relative paths from that location.
		/// </summary>
		/// <param name="relativePath">Path of the IDirectory to get, relative to this IDirectory.</param>
		/// <returns>The requested IDirectory, or null if it does not exist.</returns>
		IDirectory? GetWorkingDirectory(string relativePath);

		// Events (delegate+event pattern per host FluidSharedMap convention)
		/// <summary>
		/// Emitted when a key is set or deleted.
		/// </summary>
		event ValueChangedEventHandler? OnValueChanged;

		/// <summary>
		/// Emitted when a subdirectory is created.
		/// </summary>
		event SubDirectoryEventHandler? OnSubDirectoryCreated;

		/// <summary>
		/// Emitted when a subdirectory is deleted.
		/// </summary>
		event SubDirectoryEventHandler? OnSubDirectoryDeleted;
	}

	/// <summary>
	/// Provides a hierarchical organization of map-like data structures as SubDirectories.
	/// The values stored within can be accessed like a map, and the hierarchy can be navigated using path syntax.
	/// SubDirectories can be retrieved for use as working directories.
	/// </summary>
	public interface ISharedDirectory : IDirectory, IFluidDataObject, IFluidDataObjectMessageHandler
	{
		// No additional members for the demo. Later waves may add: dispose,
		// load-from-snapshot, submit-op APIs.
	}

	// -------------------------------------------------------------
	// Serializable value types (from internalInterfaces.ts)
	// -------------------------------------------------------------

	/// <summary>
	/// The ready-for-serialization format of values contained in DDS contents.
	/// </summary>
	public class SerializableValue
	{
		/// <summary>
		/// A type annotation to help indicate how the value serializes.
		/// </summary>
		public string Type { get; set; } = string.Empty;

		/// <summary>
		/// The JSONable representation of the value.
		/// </summary>
		public object? Value { get; set; }
	}

	/// <summary>
	/// Serialized SerializableValue counterpart.
	/// </summary>
	public class SerializedValue
	{
		/// <summary>
		/// A type annotation to help indicate how the value serializes.
		/// </summary>
		public string Type { get; set; } = string.Empty;

		/// <summary>
		/// Snapshot representation of the value.
		/// </summary>
		/// <remarks>Will be null if the original value was undefined.</remarks>
		public object? Value { get; set; }
	}

	// -------------------------------------------------------------
	// Op discriminated union — port for later waves; declared now so the
	// op type surface is visible.
	// -------------------------------------------------------------

	/// <summary>
	/// String identifiers for directory operation types.
	/// </summary>
	public enum DirectoryOpType
	{
		/// <summary>
		/// Operation indicating a value should be set for a key.
		/// </summary>
		Set,

		/// <summary>
		/// Operation indicating a key should be deleted from the directory.
		/// </summary>
		Delete,

		/// <summary>
		/// Operation indicating the directory should be cleared.
		/// </summary>
		Clear,

		/// <summary>
		/// Operation indicating a subdirectory should be created.
		/// </summary>
		CreateSubDirectory,

		/// <summary>
		/// Operation indicating a subdirectory should be deleted.
		/// </summary>
		DeleteSubDirectory,
	}

	/// <summary>Base type for all directory ops. Concrete subclasses carry op-specific fields.</summary>
	public abstract class DirectoryOperation
	{
		/// <summary>
		/// String identifier of the operation type.
		/// </summary>
		public abstract DirectoryOpType Type { get; }

		/// <summary>
		/// Absolute path of the directory affected by the operation.
		/// </summary>
		public string Path { get; set; } = string.Empty;
	}

	/// <summary>
	/// Operation indicating a value should be set for a key.
	/// </summary>
	public sealed class DirectorySetOperation : DirectoryOperation
	{
		/// <summary>
		/// String identifier of the operation type.
		/// </summary>
		public override DirectoryOpType Type => DirectoryOpType.Set;

		/// <summary>
		/// Directory key being modified.
		/// </summary>
		public string Key { get; set; } = string.Empty;

		/// <summary>
		/// Value to be set on the key.
		/// </summary>
		public SerializableValue Value { get; set; } = new SerializableValue();
	}

	/// <summary>
	/// Operation indicating a key should be deleted from the directory.
	/// </summary>
	public sealed class DirectoryDeleteOperation : DirectoryOperation
	{
		/// <summary>
		/// String identifier of the operation type.
		/// </summary>
		public override DirectoryOpType Type => DirectoryOpType.Delete;

		/// <summary>
		/// Directory key being modified.
		/// </summary>
		public string Key { get; set; } = string.Empty;
	}

	/// <summary>
	/// Operation indicating the directory should be cleared.
	/// </summary>
	public sealed class DirectoryClearOperation : DirectoryOperation
	{
		/// <summary>
		/// String identifier of the operation type.
		/// </summary>
		public override DirectoryOpType Type => DirectoryOpType.Clear;
	}

	/// <summary>
	/// Operation indicating a subdirectory should be created.
	/// </summary>
	public sealed class DirectoryCreateSubDirectoryOperation : DirectoryOperation
	{
		/// <summary>
		/// String identifier of the operation type.
		/// </summary>
		public override DirectoryOpType Type => DirectoryOpType.CreateSubDirectory;

		/// <summary>
		/// Name of the new subdirectory.
		/// </summary>
		public string SubdirName { get; set; } = string.Empty;
	}

	/// <summary>
	/// Operation indicating a subdirectory should be deleted.
	/// </summary>
	public sealed class DirectoryDeleteSubDirectoryOperation : DirectoryOperation
	{
		/// <summary>
		/// String identifier of the operation type.
		/// </summary>
		public override DirectoryOpType Type => DirectoryOpType.DeleteSubDirectory;

		/// <summary>
		/// Name of the subdirectory to be deleted.
		/// </summary>
		public string SubdirName { get; set; } = string.Empty;
	}

	// -------------------------------------------------------------
	// Map ops (from internalInterfaces.ts) — for later; port declaratively.
	// -------------------------------------------------------------

	/// <summary>
	/// String identifiers for map operation types.
	/// </summary>
	public enum MapOpType
	{
		/// <summary>
		/// Operation indicating a value should be set for a key.
		/// </summary>
		Set,

		/// <summary>
		/// Operation indicating a key should be deleted from the map.
		/// </summary>
		Delete,

		/// <summary>
		/// Operation indicating the map should be cleared.
		/// </summary>
		Clear,
	}

	/// <summary>
	/// Base type for all map ops. Concrete subclasses carry op-specific fields.
	/// </summary>
	public abstract class MapOperation
	{
		/// <summary>
		/// String identifier of the operation type.
		/// </summary>
		public abstract MapOpType Type { get; }
	}

	/// <summary>
	/// Operation indicating a value should be set for a key.
	/// </summary>
	public sealed class MapSetOperation : MapOperation
	{
		/// <summary>
		/// String identifier of the operation type.
		/// </summary>
		public override MapOpType Type => MapOpType.Set;

		/// <summary>
		/// Map key being modified.
		/// </summary>
		public string Key { get; set; } = string.Empty;

		/// <summary>
		/// Value to be set on the key.
		/// </summary>
		public SerializableValue Value { get; set; } = new SerializableValue();
	}

	/// <summary>
	/// Operation indicating a key should be deleted from the map.
	/// </summary>
	public sealed class MapDeleteOperation : MapOperation
	{
		/// <summary>
		/// String identifier of the operation type.
		/// </summary>
		public override MapOpType Type => MapOpType.Delete;

		/// <summary>
		/// Map key being modified.
		/// </summary>
		public string Key { get; set; } = string.Empty;
	}

	/// <summary>
	/// Operation indicating the map should be cleared.
	/// </summary>
	public sealed class MapClearOperation : MapOperation
	{
		/// <summary>
		/// String identifier of the operation type.
		/// </summary>
		public override MapOpType Type => MapOpType.Clear;
	}
}
