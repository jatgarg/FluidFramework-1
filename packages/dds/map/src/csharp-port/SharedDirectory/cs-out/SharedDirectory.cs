// -----------------------------------------------------------------------------
// Ported from packages/dds/map/src/directory.ts (SharedDirectory class subset).
// Part of the SharedDirectory C# feasibility-demo port.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid
{
	public sealed class SharedDirectory : ISharedDirectory
	{
		private readonly string _id;
		private readonly SubDirectory _root;
		private readonly IFluidDataObjectSender? _sender;
		private readonly IFluidDataObjectRegistry? _registry;

		public SharedDirectory(string? id = null, IFluidDataObjectSender? sender = null, IFluidDataObjectRegistry? registry = null)
		{
			_id = id ?? FluidObjectId.CreateId();
			_sender = sender;
			_registry = registry;
			_root = new SubDirectory(this, parent: null, absolutePath: "/");
		}

		public string Id => _id;

		public string AbsolutePath => _root.AbsolutePath;

		internal SubDirectory RootDirectory => _root;

		internal IFluidDataObjectSender? Sender => _sender;

		internal IFluidDataObjectRegistry? Registry => _registry;

		public event ValueChangedEventHandler? OnValueChanged;
		public event SubDirectoryEventHandler? OnSubDirectoryCreated;
		public event SubDirectoryEventHandler? OnSubDirectoryDeleted;

		public object? Get(string key)
		{
			return _root.Get(key);
		}

		public IDirectory Set(string key, object? value)
		{
			_root.Set(key, value);
			return this;
		}

		public bool Has(string key)
		{
			return _root.Has(key);
		}

		public bool Delete(string key)
		{
			return _root.Delete(key);
		}

		public void Clear()
		{
			_root.Clear();
		}

		public IReadOnlyCollection<string> Keys => _root.Keys;

		public IReadOnlyCollection<object?> Values => _root.Values;

		public IEnumerable<KeyValuePair<string, object?>> Entries() => _root.Entries();

		public IEnumerator<KeyValuePair<string, object?>> GetEnumerator() => _root.GetEnumerator();

		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();

		public int Count => _root.Count;

		public IDirectory CreateSubDirectory(string subdirName)
		{
			return _root.CreateSubDirectory(subdirName);
		}

		public IDirectory? GetSubDirectory(string subdirName)
		{
			return _root.GetSubDirectory(subdirName);
		}

		public bool HasSubDirectory(string subdirName)
		{
			return _root.HasSubDirectory(subdirName);
		}

		public bool DeleteSubDirectory(string subdirName)
		{
			return _root.DeleteSubDirectory(subdirName);
		}

		public int CountSubDirectory()
		{
			return _root.CountSubDirectory();
		}

		public IEnumerable<KeyValuePair<string, IDirectory>> SubDirectories()
		{
			return _root.SubDirectories();
		}

		public IDirectory? GetWorkingDirectory(string relativePath)
		{
			return _root.GetWorkingDirectory(relativePath);
		}

		public void ProcessDataObjectOp(SequencedDocumentMessageDescriptor descriptor, string opJson)
		{
			DirectoryOperation op = DirectoryOpSerializer.Deserialize(opJson, _registry);
			ProcessDirectoryOperation(descriptor, op);
		}

		public void ProcessDataObjectAttach(SequencedDocumentMessageDescriptor descriptor, fluidDataStoreMessageAttach op)
		{
			// Server-side client does not act on attach; state comes from snapshot.
		}

		internal SequenceNumber SubmitDirectoryOp(DirectoryOperation op)
		{
			if (_sender == null)
			{
				return default;
			}

			return _sender.QueueDataObjectMessage(_id, GetOpTypeName(op), DirectoryOpSerializer.Serialize(op, _registry));
		}

		private void ProcessDirectoryOperation(SequencedDocumentMessageDescriptor descriptor, DirectoryOperation op)
		{
			SubDirectory target = ResolveSubDirectoryByPath(op.Path);

			if (descriptor.Origin == OpOrigin.Local)
			{
				switch (op)
				{
					case DirectorySetOperation setOp:
						target.ProcessAckForKey(setOp.Key, descriptor.Seq);
						break;

					case DirectoryDeleteOperation deleteOp:
						target.ProcessAckForKey(deleteOp.Key, descriptor.Seq);
						break;

					case DirectoryClearOperation:
						target.ProcessAckForClear(descriptor.Seq);
						break;

					case DirectoryCreateSubDirectoryOperation createOp:
						target.ProcessAckForSubdir(createOp.SubdirName, descriptor.Seq);
						break;

					case DirectoryDeleteSubDirectoryOperation deleteSubdirOp:
						target.ProcessAckForSubdir(deleteSubdirOp.SubdirName, descriptor.Seq);
						break;

					default:
						throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unhandled op type: {op.GetType().Name}");
				}

				return;
			}

			switch (op)
			{
				case DirectorySetOperation setOp:
					target.ApplyRemoteSet(setOp.Key, setOp.Value.Value);
					break;

				case DirectoryDeleteOperation deleteOp:
					target.ApplyRemoteDelete(deleteOp.Key);
					break;

				case DirectoryClearOperation:
					target.ApplyRemoteClear();
					break;

				case DirectoryCreateSubDirectoryOperation createOp:
					target.ApplyRemoteCreateSubDirectory(createOp.SubdirName);
					break;

				case DirectoryDeleteSubDirectoryOperation deleteSubdirOp:
					target.ApplyRemoteDeleteSubDirectory(deleteSubdirOp.SubdirName);
					break;

				default:
					throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unhandled op type: {op.GetType().Name}");
			}
		}

		/// <summary>
		/// Parses and populates this SharedDirectory from snapshot JSON. The directory must be empty
		/// (no keys, no subdirectories). No ops are emitted and no events fire — this is
		/// initial-state hydration, not a mutation.
		/// </summary>
		public void LoadFromSnapshot(string snapshotJson, Func<string, string>? blobResolver = null)
		{
			LoadFromSnapshot(DirectorySnapshotLoader.Load(snapshotJson, blobResolver, _registry));
		}

		/// <summary>
		/// Populates this SharedDirectory from a snapshot DTO. The directory must be empty
		/// (no keys, no subdirectories). No ops are emitted and no events fire — this is
		/// initial-state hydration, not a mutation.
		/// </summary>
		public void LoadFromSnapshot(DirectorySnapshotDto snapshot)
		{
			if (snapshot == null)
			{
				throw new ArgumentNullException(nameof(snapshot));
			}

			if (_root.Count > 0 || _root.CountSubDirectory() > 0)
			{
				throw new OcsException(OcsGateErrorCode.InvalidState,
					"SharedDirectory.LoadFromSnapshot: directory must be empty; loading into a non-empty directory is not supported in this wave.");
			}

			_root.PopulateFromSnapshot(snapshot);
		}

		private SubDirectory ResolveSubDirectoryByPath(string absolutePath)
		{
			if (absolutePath == "/" || absolutePath == string.Empty)
			{
				return _root;
			}

			string[] segments = absolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);
			SubDirectory cursor = _root;
			foreach (string segment in segments)
			{
				SubDirectory? child = cursor.GetSequencedSubDirectoryInternal(segment);
				if (child == null)
				{
					throw new OcsException(OcsGateErrorCode.InvalidOperation,
						$"SharedDirectory.ResolveSubDirectoryByPath: path '{absolutePath}' does not exist (missing segment '{segment}')");
				}

				cursor = child;
			}

			return cursor;
		}

		private static string GetOpTypeName(DirectoryOperation op)
		{
			return op switch
			{
				DirectorySetOperation => "set",
				DirectoryDeleteOperation => "delete",
				DirectoryClearOperation => "clear",
				DirectoryCreateSubDirectoryOperation => "createSubDirectory",
				DirectoryDeleteSubDirectoryOperation => "deleteSubDirectory",
				_ => throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unknown directory op runtime type: {op.GetType().FullName}"),
			};
		}

		internal void PropagateValueChanged(ValueChangedEventArgs args)
		{
			ValueChangedEventHandler? raiseEvent = OnValueChanged;
			if (raiseEvent != null)
			{
				raiseEvent(this, args);
			}
		}

		internal void PropagateSubDirectoryCreated(SubDirectoryEventArgs args)
		{
			SubDirectoryEventHandler? raiseEvent = OnSubDirectoryCreated;
			if (raiseEvent != null)
			{
				raiseEvent(this, args);
			}
		}

		internal void PropagateSubDirectoryDeleted(SubDirectoryEventArgs args)
		{
			SubDirectoryEventHandler? raiseEvent = OnSubDirectoryDeleted;
			if (raiseEvent != null)
			{
				raiseEvent(this, args);
			}
		}
	}
}
