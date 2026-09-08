// -----------------------------------------------------------------------------
// Ported from packages/dds/map/src/directory.ts (SubDirectory class subset).
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;

namespace Microsoft.Office.Web.Fluid
{
	internal sealed class SeqData
	{
		public SeqData(long seq, long clientSeq)
		{
			Seq = seq;
			ClientSeq = clientSeq;
		}

		public long Seq { get; set; }

		public long ClientSeq { get; set; }
	}

	internal sealed class SubDirectory : IDirectory
	{
		private const long _noPending = -1;
		internal const string DetachedClientId = "detached";

		private readonly object _lock = new object();
		private readonly SharedDirectory _root;
		private readonly SubDirectory? _parent;
		private readonly string _absolutePath;
		private readonly string _localName;
		private readonly Dictionary<string, object?> _storage;
		private readonly Dictionary<string, SubDirectory> _subdirs;
		private readonly List<string> _subdirOrder;
		private readonly List<PendingStorageEntry> _pendingStorageData = new List<PendingStorageEntry>();
		private readonly List<PendingSubDirectoryEntry> _pendingSubDirectoryData = new List<PendingSubDirectoryEntry>();
		private readonly Dictionary<long, object> _pendingByClientSequenceNumber = new Dictionary<long, object>();
		private long _localCreationSeq;

		private abstract class PendingStorageEntry
		{
			protected PendingStorageEntry(SubDirectory subdir)
			{
				Subdir = subdir;
			}

			public SubDirectory Subdir { get; }
		}

		private sealed class PendingKeySet
		{
			public PendingKeySet(object? value, PendingKeyLifetime lifetime, SubDirectory subdir)
			{
				Value = value;
				Lifetime = lifetime;
				Subdir = subdir;
			}

			public object? Value { get; }

			public PendingKeyLifetime Lifetime { get; }

			public SubDirectory Subdir { get; }

			public long ClientSequenceNumber { get; set; } = _noPending;
		}

		private sealed class PendingKeyLifetime : PendingStorageEntry
		{
			public PendingKeyLifetime(string key, SubDirectory subdir)
				: base(subdir)
			{
				Key = key;
			}

			public string Key { get; }

			public List<PendingKeySet> KeySets { get; } = new List<PendingKeySet>();
		}

		private sealed class PendingKeyDelete : PendingStorageEntry
		{
			public PendingKeyDelete(string key, SubDirectory subdir)
				: base(subdir)
			{
				Key = key;
			}

			public string Key { get; }

			public long ClientSequenceNumber { get; set; } = _noPending;
		}

		private sealed class PendingClear : PendingStorageEntry
		{
			public PendingClear(SubDirectory subdir)
				: base(subdir)
			{
			}

			public long ClientSequenceNumber { get; set; } = _noPending;
		}

		private abstract class PendingSubDirectoryEntry
		{
			protected PendingSubDirectoryEntry(string subdirName, SubDirectory parentSubdir)
			{
				SubdirName = subdirName;
				ParentSubdir = parentSubdir;
			}

			public string SubdirName { get; }

			public SubDirectory ParentSubdir { get; }

			public long ClientSequenceNumber { get; set; } = _noPending;
		}

		private sealed class PendingSubDirectoryCreate : PendingSubDirectoryEntry
		{
			public PendingSubDirectoryCreate(string subdirName, SubDirectory parentSubdir, SubDirectory subdir)
				: base(subdirName, parentSubdir)
			{
				Subdir = subdir;
			}

			public SubDirectory Subdir { get; }
		}

		private sealed class PendingSubDirectoryDelete : PendingSubDirectoryEntry
		{
			public PendingSubDirectoryDelete(string subdirName, SubDirectory parentSubdir, SubDirectory subdir)
				: base(subdirName, parentSubdir)
			{
				Subdir = subdir;
			}

			public SubDirectory Subdir { get; }
		}

		internal SubDirectory(
			SharedDirectory root,
			SubDirectory? parent,
			string absolutePath,
			string localName,
			SeqData seqData,
			IEnumerable<string>? clientIds)
		{
			_root = root;
			_parent = parent;
			_absolutePath = absolutePath;
			_localName = localName;
			SeqData = seqData;
			ClientIds = clientIds == null ? new HashSet<string>() : new HashSet<string>(clientIds);
			_storage = new Dictionary<string, object?>();
			_subdirs = new Dictionary<string, SubDirectory>();
			_subdirOrder = new List<string>();
		}

		public string AbsolutePath => _absolutePath;

		internal SeqData SeqData { get; }

		internal HashSet<string> ClientIds { get; }

		public event ValueChangedEventHandler? OnValueChanged;
		public event SubDirectoryEventHandler? OnSubDirectoryCreated;
		public event SubDirectoryEventHandler? OnSubDirectoryDeleted;

		public object? Get(string key)
		{
			if (key == null)
			{
				throw new ArgumentNullException(nameof(key));
			}

			lock (_lock)
			{
				return GetOptimisticValueNoLock(key);
			}
		}

		public IDirectory Set(string key, object? value)
		{
			if (key == null)
			{
				throw new ArgumentNullException(nameof(key));
			}

			ValueChangedEventArgs args;
			lock (_lock)
			{
				object? previousValue = GetOptimisticValueNoLock(key);
				if (_root.Sender == null)
				{
					_storage[key] = value;
				}
				else
				{
					PendingKeyLifetime lifetime = GetOrCreatePendingKeyLifetimeNoLock(key);
					PendingKeySet pendingKeySet = new PendingKeySet(value, lifetime, this);
					lifetime.KeySets.Add(pendingKeySet);
					SequenceNumber seq = SubmitSetOp(key, value);
					RegisterPendingAckNoLock(seq, pendingKeySet);
				}

				args = new ValueChangedEventArgs()
				{
					Key = key,
					PreviousValue = previousValue,
					Path = _absolutePath,
					Local = true,
				};
			}

			RaiseValueChanged(args);
			return this;
		}

		public bool Has(string key)
		{
			if (key == null)
			{
				throw new ArgumentNullException(nameof(key));
			}

			lock (_lock)
			{
				return OptimisticallyHasNoLock(key);
			}
		}

		public bool Delete(string key)
		{
			if (key == null)
			{
				throw new ArgumentNullException(nameof(key));
			}

			ValueChangedEventArgs? args = null;
			bool deleted;
			lock (_lock)
			{
				bool hadValue = OptimisticallyHasNoLock(key);
				object? previousValue = GetOptimisticValueNoLock(key);
				if (_root.Sender == null)
				{
					if (!_storage.Remove(key))
					{
						return false;
					}

					deleted = true;
				}
				else
				{
					PendingKeyDelete pendingKeyDelete = new PendingKeyDelete(key, this);
					_pendingStorageData.Add(pendingKeyDelete);
					SequenceNumber seq = SubmitDeleteOp(key);
					RegisterPendingAckNoLock(seq, pendingKeyDelete);
					deleted = true;
				}

				if (hadValue)
				{
					args = new ValueChangedEventArgs()
					{
						Key = key,
						PreviousValue = previousValue,
						Path = _absolutePath,
						Local = true,
					};
				}
			}

			if (args != null)
			{
				RaiseValueChanged(args);
			}

			return deleted;
		}

		public void Clear()
		{
			List<ValueChangedEventArgs> args;
			lock (_lock)
			{
				args = GetOptimisticEntriesNoLock().Select(
					entry => new ValueChangedEventArgs()
					{
						Key = entry.Key,
						PreviousValue = entry.Value,
						Path = _absolutePath,
						Local = true,
					}).ToList();

				if (_root.Sender == null)
				{
					_storage.Clear();
				}
				else
				{
					PendingClear pendingClear = new PendingClear(this);
					_pendingStorageData.Add(pendingClear);
					SequenceNumber seq = SubmitClearOp();
					RegisterPendingAckNoLock(seq, pendingClear);
				}
			}

			foreach (ValueChangedEventArgs arg in args)
			{
				RaiseValueChanged(arg);
			}
		}

		public IReadOnlyCollection<string> Keys
		{
			get
			{
				lock (_lock)
				{
					return GetOptimisticEntriesNoLock().Select(entry => entry.Key).ToList().AsReadOnly();
				}
			}
		}

		public IReadOnlyCollection<object?> Values
		{
			get
			{
				lock (_lock)
				{
					return GetOptimisticEntriesNoLock().Select(entry => entry.Value).ToList().AsReadOnly();
				}
			}
		}

		public IEnumerable<KeyValuePair<string, object?>> Entries()
		{
			lock (_lock)
			{
				return GetOptimisticEntriesNoLock();
			}
		}

		public IEnumerator<KeyValuePair<string, object?>> GetEnumerator()
		{
			return Entries().GetEnumerator();
		}

		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return GetEnumerator();
		}

		public int Count
		{
			get
			{
				lock (_lock)
				{
					return GetOptimisticEntriesNoLock().Count;
				}
			}
		}

		public IDirectory CreateSubDirectory(string subdirName)
		{
			ValidateSubDirectoryName(subdirName);

			SubDirectory subdir;
			SubDirectoryEventArgs args;
			lock (_lock)
			{
				SubDirectory? existingSubdir = GetOptimisticSubDirectoryNoLock(subdirName);
				if (existingSubdir != null)
				{
					return existingSubdir;
				}

				subdir = new SubDirectory(
					_root,
					this,
					MakeChildAbsolutePath(_absolutePath, subdirName),
					subdirName,
					CreateLocalSeqDataNoLock(),
					CreateLocalClientIdsNoLock());
				if (_root.Sender == null)
				{
					AddSequencedSubDirectoryNoLock(subdirName, subdir);
				}
				else
				{
					PendingSubDirectoryCreate pendingSubdirCreate = new PendingSubDirectoryCreate(subdirName, this, subdir);
					_pendingSubDirectoryData.Add(pendingSubdirCreate);
					SequenceNumber seq = SubmitCreateSubDirectoryOp(subdirName);
					RegisterPendingAckNoLock(seq, pendingSubdirCreate);
				}

				args = new SubDirectoryEventArgs()
				{
					SubdirName = subdirName,
					ParentPath = _absolutePath,
					Path = subdirName,
					Local = true,
				};
			}

			RaiseSubDirectoryCreated(args);
			return subdir;
		}

		public IDirectory? GetSubDirectory(string subdirName)
		{
			return GetSubDirectoryInternal(subdirName);
		}

		internal SubDirectory? GetSubDirectoryInternal(string subdirName)
		{
			if (subdirName == null)
			{
				throw new ArgumentNullException(nameof(subdirName));
			}

			lock (_lock)
			{
				return GetOptimisticSubDirectoryNoLock(subdirName);
			}
		}

		internal SubDirectory? GetSequencedSubDirectoryInternal(string subdirName)
		{
			if (subdirName == null)
			{
				throw new ArgumentNullException(nameof(subdirName));
			}

			lock (_lock)
			{
				return _subdirs.TryGetValue(subdirName, out SubDirectory? subdir) ? subdir : null;
			}
		}

		public bool HasSubDirectory(string subdirName)
		{
			if (subdirName == null)
			{
				throw new ArgumentNullException(nameof(subdirName));
			}

			lock (_lock)
			{
				return GetOptimisticSubDirectoryNoLock(subdirName) != null;
			}
		}

		public bool DeleteSubDirectory(string subdirName)
		{
			if (subdirName == null)
			{
				throw new ArgumentNullException(nameof(subdirName));
			}

			SubDirectoryEventArgs args;
			lock (_lock)
			{
				SubDirectory? previousSubdir = GetOptimisticSubDirectoryNoLock(subdirName);
				if (previousSubdir == null)
				{
					return false;
				}

				if (_root.Sender == null)
				{
					RemoveSequencedSubDirectoryNoLock(subdirName);
				}
				else
				{
					PendingSubDirectoryDelete pendingSubdirDelete = new PendingSubDirectoryDelete(subdirName, this, previousSubdir);
					_pendingSubDirectoryData.Add(pendingSubdirDelete);
					SequenceNumber seq = SubmitDeleteSubDirectoryOp(subdirName);
					RegisterPendingAckNoLock(seq, pendingSubdirDelete);
				}

				args = new SubDirectoryEventArgs()
				{
					SubdirName = subdirName,
					ParentPath = _absolutePath,
					Path = subdirName,
					Local = true,
				};
			}

			RaiseSubDirectoryDeleted(args);
			return true;
		}

		public int CountSubDirectory()
		{
			lock (_lock)
			{
				return GetOptimisticSubDirectoriesNoLock().Count;
			}
		}

		public IEnumerable<KeyValuePair<string, IDirectory>> SubDirectories()
		{
			lock (_lock)
			{
				return GetOptimisticSubDirectoriesNoLock()
					.Select(entry => new KeyValuePair<string, IDirectory>(entry.Key, entry.Value))
					.ToList();
			}
		}

		public IDirectory? GetWorkingDirectory(string relativePath)
		{
			if (relativePath == null)
			{
				throw new ArgumentNullException(nameof(relativePath));
			}

			SubDirectory current = _root.RootDirectory;
			foreach (string segment in GetAbsolutePathSegments(relativePath))
			{
				lock (current._lock)
				{
					SubDirectory? child = current.GetOptimisticSubDirectoryNoLock(segment);
					if (child == null)
					{
						return null;
					}

					current = child;
				}
			}

			return current;
		}

		internal bool ProcessAckForKey(string key, SequencedDocumentMessageDescriptor msg, SubDirectory? targetSubdir)
		{
			lock (_lock)
			{
				if (!IsMessageForCurrentInstanceOfSubDirectory(msg, targetSubdir))
				{
					return false;
				}

				SequenceNumber seq = msg.Seq;
				ValidateAckSequenceNoLock(seq, $"SubDirectory.ProcessAckForKey: client seq is missing for key '{key}' at path '{_absolutePath}'");
				if (!_pendingByClientSequenceNumber.TryGetValue(seq.clientSequenceNumber, out object? pending))
				{
					ThrowNoPendingKeyAckNoLock(key, seq.clientSequenceNumber);
				}

				if (pending is PendingKeySet pendingKeySet)
				{
					ProcessAckForKeySetNoLock(key, pendingKeySet, seq.clientSequenceNumber);
				}
				else if (pending is PendingKeyDelete pendingKeyDelete)
				{
					ProcessAckForKeyDeleteNoLock(key, pendingKeyDelete, seq.clientSequenceNumber);
				}
				else
				{
					throw new LoggingError(						$"SubDirectory.ProcessAckForKey: client seq {seq.clientSequenceNumber} does not reference a pending key entry at path '{_absolutePath}'");
				}

				_pendingByClientSequenceNumber.Remove(seq.clientSequenceNumber);
				return true;
			}
		}

		internal bool ProcessAckForSubdir(string subdirName, SequencedDocumentMessageDescriptor msg, SubDirectory? targetSubdir)
		{
			lock (_lock)
			{
				if (!IsMessageForCurrentInstanceOfSubDirectory(msg, targetSubdir))
				{
					return false;
				}

				SequenceNumber seq = msg.Seq;
				ValidateAckSequenceNoLock(seq, $"SubDirectory.ProcessAckForSubdir: client seq is missing for subdirectory '{subdirName}' at path '{_absolutePath}'");
				if (!_pendingByClientSequenceNumber.TryGetValue(seq.clientSequenceNumber, out object? pending))
				{
					ThrowNoPendingSubdirAckNoLock(subdirName, seq.clientSequenceNumber);
				}

				if (pending is not PendingSubDirectoryEntry pendingSubdir || pendingSubdir.SubdirName != subdirName)
				{
					throw new LoggingError(						$"SubDirectory.ProcessAckForSubdir: client seq {seq.clientSequenceNumber} does not reference pending subdirectory '{subdirName}' at path '{_absolutePath}'");
				}

				int pendingEntryIndex = FindFirstPendingSubDirectoryEntryIndexNoLock(subdirName);
				if (pendingEntryIndex < 0 || !ReferenceEquals(_pendingSubDirectoryData[pendingEntryIndex], pendingSubdir))
				{
					throw new LoggingError(						$"SubDirectory.ProcessAckForSubdir: ack for subdirectory '{subdirName}' is not the next pending entry (client seq {seq.clientSequenceNumber})");
				}

				_pendingSubDirectoryData.RemoveAt(pendingEntryIndex);
				if (pendingSubdir is PendingSubDirectoryCreate pendingCreate)
				{
					if (!_subdirs.ContainsKey(subdirName))
					{
						AddSequencedSubDirectoryNoLock(subdirName, pendingCreate.Subdir);
					}

					MarkCreatedSubDirectorySequencedNoLock(pendingCreate.Subdir, msg);
				}
				else
				{
					if (_subdirs.TryGetValue(subdirName, out SubDirectory? deletedSubdir))
					{
						deletedSubdir.ClearSubDirectorySequencedData();
					}

					RemoveSequencedSubDirectoryNoLock(subdirName);
				}

				_pendingByClientSequenceNumber.Remove(seq.clientSequenceNumber);
				return true;
			}
		}

		internal bool ProcessAckForClear(SequencedDocumentMessageDescriptor msg, SubDirectory? targetSubdir)
		{
			lock (_lock)
			{
				if (!IsMessageForCurrentInstanceOfSubDirectory(msg, targetSubdir))
				{
					return false;
				}

				SequenceNumber seq = msg.Seq;
				ValidateAckSequenceNoLock(seq, $"SubDirectory.ProcessAckForClear: client seq is missing at path '{_absolutePath}'");
				if (!_pendingByClientSequenceNumber.TryGetValue(seq.clientSequenceNumber, out object? pending))
				{
					ThrowNoPendingClearAckNoLock(seq.clientSequenceNumber);
				}

				if (pending is not PendingClear pendingClear)
				{
					throw new LoggingError(						$"SubDirectory.ProcessAckForClear: client seq {seq.clientSequenceNumber} does not reference a pending clear at path '{_absolutePath}'");
				}

				if (_pendingStorageData.Count == 0 || !ReferenceEquals(_pendingStorageData[0], pendingClear))
				{
					throw new LoggingError(						$"SubDirectory.ProcessAckForClear: clear ack is not the next pending storage entry at path '{_absolutePath}' (client seq {seq.clientSequenceNumber})");
				}

				_pendingStorageData.RemoveAt(0);
				_storage.Clear();
				_pendingByClientSequenceNumber.Remove(seq.clientSequenceNumber);
				return true;
			}
		}

		internal void ApplyRemoteSet(string key, object? value, SequencedDocumentMessageDescriptor msg)
		{
			ValueChangedEventArgs? args = null;
			lock (_lock)
			{
				if (!IsMessageForCurrentInstanceOfSubDirectory(msg, targetSubdir: null))
				{
					return;
				}

				_storage.TryGetValue(key, out object? previous);
				_storage[key] = value;
				if (!HasPendingStorageEntryForKeyOrClearNoLock(key))
				{
					args = new ValueChangedEventArgs()
					{
						Key = key,
						PreviousValue = previous,
						Path = _absolutePath,
						Local = false,
					};
				}
			}

			if (args != null)
			{
				RaiseValueChanged(args);
			}
		}

		internal void ApplyRemoteDelete(string key, SequencedDocumentMessageDescriptor msg)
		{
			ValueChangedEventArgs? args = null;
			lock (_lock)
			{
				if (!IsMessageForCurrentInstanceOfSubDirectory(msg, targetSubdir: null))
				{
					return;
				}

				// TS emits valueChanged with previousValue: undefined even when
				// the key isn't present locally, as long as no pending op
				// suppresses it. Remove(out) returns the prior value in a single
				// hash lookup and leaves `previous` as null when the key was absent.
				_storage.Remove(key, out object? previous);
				if (!HasPendingStorageEntryForKeyOrClearNoLock(key))
				{
					args = new ValueChangedEventArgs()
					{
						Key = key,
						PreviousValue = previous,
						Path = _absolutePath,
						Local = false,
					};
				}
			}

			if (args != null)
			{
				RaiseValueChanged(args);
			}
		}

		internal void ApplyRemoteClear(SequencedDocumentMessageDescriptor msg)
		{
			List<ValueChangedEventArgs> args = new List<ValueChangedEventArgs>();
			lock (_lock)
			{
				if (!IsMessageForCurrentInstanceOfSubDirectory(msg, targetSubdir: null))
				{
					return;
				}

				List<KeyValuePair<string, object?>> previousEntries = _storage.ToList();
				_storage.Clear();
				if (HasPendingClearNoLock())
				{
					return;
				}

				foreach (KeyValuePair<string, object?> entry in previousEntries)
				{
					if (HasPendingStorageEntryForKeyNoLock(entry.Key))
					{
						continue;
					}

					args.Add(new ValueChangedEventArgs()
					{
						Key = entry.Key,
						PreviousValue = entry.Value,
						Path = _absolutePath,
						Local = false,
					});
				}
			}

			foreach (ValueChangedEventArgs arg in args)
			{
				RaiseValueChanged(arg);
			}
		}

		internal void ApplyRemoteCreateSubDirectory(string subdirName, SequencedDocumentMessageDescriptor msg)
		{
			SubDirectoryEventArgs? args = null;
			lock (_lock)
			{
				if (!IsMessageForCurrentInstanceOfSubDirectory(msg, targetSubdir: null))
				{
					return;
				}

				bool hasPending = HasPendingSubDirectoryEntryNoLock(subdirName);
				SubDirectory? subdir = GetOptimisticSubDirectoryNoLock(subdirName);
				if (subdir == null)
				{
					subdir = new SubDirectory(
						_root,
						this,
						MakeChildAbsolutePath(_absolutePath, subdirName),
						subdirName,
						CreateRemoteSeqData(msg),
						CreateRemoteClientIds(msg));
				}
				else
				{
					MarkCreatedSubDirectorySequencedNoLock(subdir, msg);
				}

				if (!_subdirs.ContainsKey(subdirName))
				{
					AddSequencedSubDirectoryNoLock(subdirName, subdir);
					if (!hasPending)
					{
						args = new SubDirectoryEventArgs()
						{
							SubdirName = subdirName,
							ParentPath = _absolutePath,
							Path = subdirName,
							Local = false,
						};
					}
				}
			}

			if (args != null)
			{
				RaiseSubDirectoryCreated(args);
			}
		}

		internal void ApplyRemoteDeleteSubDirectory(string subdirName, SequencedDocumentMessageDescriptor msg)
		{
			SubDirectoryEventArgs? args = null;
			lock (_lock)
			{
				if (!IsMessageForCurrentInstanceOfSubDirectory(msg, targetSubdir: null))
				{
					return;
				}

				if (!_subdirs.ContainsKey(subdirName))
				{
					return;
				}

				SubDirectory deletedSubdir = _subdirs[subdirName];
				deletedSubdir.ClearSubDirectorySequencedData();
				RemoveSequencedSubDirectoryNoLock(subdirName);
				if (!HasPendingSubDirectoryDeleteNoLock(subdirName))
				{
					args = new SubDirectoryEventArgs()
					{
						SubdirName = subdirName,
						ParentPath = _absolutePath,
						Path = subdirName,
						Local = false,
					};
				}
			}

			if (args != null)
			{
				RaiseSubDirectoryDeleted(args);
			}
		}

		internal void PopulateFromSnapshot(DirectorySnapshotDto dto)
		{
			if (dto == null)
			{
				throw new ArgumentNullException(nameof(dto));
			}

			lock (_lock)
			{
				// Snapshot population intentionally merges into the existing tree so
				// blob-split fragments behave like TS's sequential populate(...) calls.
				foreach (KeyValuePair<string, SerializableValue> kvp in dto.Storage)
				{
					_storage[kvp.Key] = DirectoryOpSerializer.ResolveSerializedHandles(kvp.Value.Value, _root.Registry);
				}

				Dictionary<long, long> snapshotClientSeqByCreateSeq = new Dictionary<long, long>();
				foreach (KeyValuePair<string, DirectorySnapshotDto> kvp in dto.Subdirectories)
				{
					if (!_subdirs.TryGetValue(kvp.Key, out SubDirectory? child))
					{
						string childPath = MakeChildAbsolutePath(_absolutePath, kvp.Key);
						child = new SubDirectory(
							_root,
							this,
							childPath,
							kvp.Key,
							CreateSnapshotSeqDataNoLock(kvp.Value.CreateInfo, snapshotClientSeqByCreateSeq),
							CreateSnapshotClientIds(kvp.Value.CreateInfo));
						AddSequencedSubDirectoryNoLock(kvp.Key, child);
					}

					child.PopulateFromSnapshot(kvp.Value);
				}
			}
		}

		private SequenceNumber SubmitSetOp(string key, object? value)
		{
			DirectorySetOperation op = new DirectorySetOperation()
			{
				Path = _absolutePath,
				Key = key,
				Value = new SerializableValue()
				{
					Type = "Plain",
					Value = value,
				},
			};
			return _root.SubmitDirectoryOp(op);
		}

		private SequenceNumber SubmitDeleteOp(string key)
		{
			DirectoryDeleteOperation op = new DirectoryDeleteOperation()
			{
				Path = _absolutePath,
				Key = key,
			};
			return _root.SubmitDirectoryOp(op);
		}

		private SequenceNumber SubmitClearOp()
		{
			DirectoryClearOperation op = new DirectoryClearOperation()
			{
				Path = _absolutePath,
			};
			return _root.SubmitDirectoryOp(op);
		}

		private SequenceNumber SubmitCreateSubDirectoryOp(string subdirName)
		{
			DirectoryCreateSubDirectoryOperation op = new DirectoryCreateSubDirectoryOperation()
			{
				Path = _absolutePath,
				SubdirName = subdirName,
			};
			return _root.SubmitDirectoryOp(op);
		}

		private SequenceNumber SubmitDeleteSubDirectoryOp(string subdirName)
		{
			DirectoryDeleteSubDirectoryOperation op = new DirectoryDeleteSubDirectoryOperation()
			{
				Path = _absolutePath,
				SubdirName = subdirName,
			};
			return _root.SubmitDirectoryOp(op);
		}

		private void RegisterPendingAckNoLock(SequenceNumber seq, object pending)
		{
			if (!seq.HasClientSequenceNumber)
			{
				return;
			}

			long clientSequenceNumber = seq.clientSequenceNumber;
			switch (pending)
			{
				case PendingKeySet pendingKeySet:
					pendingKeySet.ClientSequenceNumber = clientSequenceNumber;
					break;

				case PendingKeyDelete pendingKeyDelete:
					pendingKeyDelete.ClientSequenceNumber = clientSequenceNumber;
					break;

				case PendingClear pendingClear:
					pendingClear.ClientSequenceNumber = clientSequenceNumber;
					break;

				case PendingSubDirectoryEntry pendingSubdir:
					pendingSubdir.ClientSequenceNumber = clientSequenceNumber;
					break;
			}

			_pendingByClientSequenceNumber.Add(clientSequenceNumber, pending);
			_root.RegisterPendingLocalOpSubDirectory(seq, this);
		}

		internal void ClearSubDirectorySequencedData()
		{
			lock (_lock)
			{
				foreach (SubDirectory subdir in _subdirs.Values.ToList())
				{
					subdir.ClearSubDirectorySequencedData();
				}

				SeqData.Seq = -1;
				SeqData.ClientSeq = -1;
				_storage.Clear();
				_subdirs.Clear();
				_subdirOrder.Clear();
				ClientIds.Clear();
				ClientIds.Add(GetLocalClientIdOrDetached());
			}
		}

		internal void ClearStalePendingEntry(SequenceNumber seq)
		{
			if (!seq.HasClientSequenceNumber)
			{
				return;
			}

			lock (_lock)
			{
				long clientSequenceNumber = seq.clientSequenceNumber;
				if (!_pendingByClientSequenceNumber.TryGetValue(clientSequenceNumber, out object? pending))
				{
					return;
				}

				if (pending is PendingKeySet pendingKeySet)
				{
					PendingKeyLifetime lifetime = pendingKeySet.Lifetime;
					lifetime.KeySets.Remove(pendingKeySet);
					if (lifetime.KeySets.Count == 0)
					{
						_pendingStorageData.Remove(lifetime);
					}
				}
				else if (pending is PendingStorageEntry pendingStorageEntry)
				{
					_pendingStorageData.Remove(pendingStorageEntry);
				}
				else if (pending is PendingSubDirectoryEntry pendingSubdirEntry)
				{
					_pendingSubDirectoryData.Remove(pendingSubdirEntry);
				}

				_pendingByClientSequenceNumber.Remove(clientSequenceNumber);
			}
		}

		private bool IsMessageForCurrentInstanceOfSubDirectory(SequencedDocumentMessageDescriptor msg, SubDirectory? targetSubdir)
		{
			return (targetSubdir == null || ReferenceEquals(targetSubdir, this))
				&& ((msg.HasClientId && ClientIds.Contains(msg.ClientId!))
					|| ClientIds.Contains(DetachedClientId)
					|| (SeqData.Seq != -1 && SeqData.Seq <= msg.RefSeq));
		}

		private SeqData CreateLocalSeqDataNoLock()
		{
			long seq = _root.Sender == null ? 0 : -1;
			return new SeqData(seq, ++_localCreationSeq);
		}

		private HashSet<string> CreateLocalClientIdsNoLock()
		{
			return new HashSet<string>() { GetLocalClientIdOrDetached() };
		}

		private SeqData CreateRemoteSeqData(SequencedDocumentMessageDescriptor msg)
		{
			long clientSeq = msg.Seq.HasClientSequenceNumber ? msg.Seq.clientSequenceNumber : -1;
			return new SeqData(msg.Seq.sequenceNumber, clientSeq);
		}

		private string GetLocalClientIdOrDetached()
		{
			return _root.Sender?.LocalClientId ?? DetachedClientId;
		}

		private IEnumerable<string> CreateRemoteClientIds(SequencedDocumentMessageDescriptor msg)
		{
			if (msg.HasClientId)
			{
				yield return msg.ClientId!;
			}
		}

		private SeqData CreateSnapshotSeqDataNoLock(
			DirectoryCreateInfo? createInfo,
			Dictionary<long, long> snapshotClientSeqByCreateSeq)
		{
			if (createInfo != null && createInfo.Csn > 0)
			{
				snapshotClientSeqByCreateSeq.TryGetValue(createInfo.Csn, out long fakeClientSeq);
				snapshotClientSeqByCreateSeq[createInfo.Csn] = fakeClientSeq + 1;
				return new SeqData(createInfo.Csn, fakeClientSeq);
			}

			return new SeqData(seq: 0, clientSeq: ++_localCreationSeq);
		}

		private IEnumerable<string>? CreateSnapshotClientIds(DirectoryCreateInfo? createInfo)
		{
			return createInfo?.CcIds;
		}

		private void MarkCreatedSubDirectorySequencedNoLock(SubDirectory subdir, SequencedDocumentMessageDescriptor msg)
		{
			if (msg.HasClientId)
			{
				subdir.ClientIds.Add(msg.ClientId!);
			}

			if (SeqData.Seq != -1 && SeqData.Seq <= msg.Seq.sequenceNumber && subdir.SeqData.Seq == -1)
			{
				subdir.SeqData.Seq = msg.Seq.sequenceNumber;
				// TS ref: directory.ts assigns clientSeq = clientSequenceNumber
				// unconditionally. The message's clientSeq drives SeqDataComparator
				// sibling ordering for grouped batches, so it must be preserved for
				// both local and remote origins.
				subdir.SeqData.ClientSeq = msg.Seq.HasClientSequenceNumber
					? msg.Seq.clientSequenceNumber
					: -1;
			}
		}

		private void ProcessAckForKeySetNoLock(string key, PendingKeySet pendingKeySet, long clientSequenceNumber)
		{
			PendingKeyLifetime lifetime = pendingKeySet.Lifetime;
			if (lifetime.Key != key)
			{
				throw new LoggingError(					$"SubDirectory.ProcessAckForKey: client seq {clientSequenceNumber} references key '{lifetime.Key}' instead of '{key}' at path '{_absolutePath}'");
			}

			int lifetimeIndex = _pendingStorageData.IndexOf(lifetime);
			int firstKeyEntryIndex = FindFirstPendingStorageEntryIndexForKeyNoLock(key);
			if (lifetimeIndex < 0 || lifetimeIndex != firstKeyEntryIndex || lifetime.KeySets.Count == 0 || !ReferenceEquals(lifetime.KeySets[0], pendingKeySet))
			{
				throw new LoggingError(					$"SubDirectory.ProcessAckForKey: set ack for key '{key}' is not the next pending set (client seq {clientSequenceNumber})");
			}

			lifetime.KeySets.RemoveAt(0);
			if (lifetime.KeySets.Count == 0)
			{
				_pendingStorageData.RemoveAt(lifetimeIndex);
			}

			_storage[key] = pendingKeySet.Value;
		}

		private void ProcessAckForKeyDeleteNoLock(string key, PendingKeyDelete pendingKeyDelete, long clientSequenceNumber)
		{
			if (pendingKeyDelete.Key != key)
			{
				throw new LoggingError(					$"SubDirectory.ProcessAckForKey: client seq {clientSequenceNumber} references key '{pendingKeyDelete.Key}' instead of '{key}' at path '{_absolutePath}'");
			}

			int pendingEntryIndex = _pendingStorageData.FindIndex(entry => ReferenceEquals(entry, pendingKeyDelete));
			int firstKeyEntryIndex = FindFirstPendingStorageEntryIndexForKeyNoLock(key);
			if (pendingEntryIndex < 0 || pendingEntryIndex != firstKeyEntryIndex)
			{
				throw new LoggingError(					$"SubDirectory.ProcessAckForKey: delete ack for key '{key}' is not the next pending key entry (client seq {clientSequenceNumber})");
			}

			_pendingStorageData.RemoveAt(pendingEntryIndex);
			_storage.Remove(key);
		}

		private void ValidateAckSequenceNoLock(SequenceNumber seq, string message)
		{
			if (!seq.HasClientSequenceNumber)
			{
				throw new LoggingError(message);
			}
		}

		private void ThrowNoPendingKeyAckNoLock(string key, long clientSequenceNumber)
		{
			if (HasPendingStorageEntryForKeyNoLock(key))
			{
				throw new LoggingError(					$"SubDirectory.ProcessAckForKey: ack does not match the pending entry for key '{key}' at path '{_absolutePath}' (client seq {clientSequenceNumber})");
			}

			throw new LoggingError(				$"SubDirectory.ProcessAckForKey: no pending entry for key '{key}' at path '{_absolutePath}' (client seq {clientSequenceNumber})");
		}

		private void ThrowNoPendingSubdirAckNoLock(string subdirName, long clientSequenceNumber)
		{
			if (HasPendingSubDirectoryEntryNoLock(subdirName))
			{
				throw new LoggingError(					$"SubDirectory.ProcessAckForSubdir: ack does not match the pending entry for subdirectory '{subdirName}' at path '{_absolutePath}' (client seq {clientSequenceNumber})");
			}

			throw new LoggingError(				$"SubDirectory.ProcessAckForSubdir: no pending entry for subdirectory '{subdirName}' at path '{_absolutePath}' (client seq {clientSequenceNumber})");
		}

		private void ThrowNoPendingClearAckNoLock(long clientSequenceNumber)
		{
			if (HasPendingClearNoLock())
			{
				throw new LoggingError(					$"SubDirectory.ProcessAckForClear: ack does not match the pending clear at path '{_absolutePath}' (client seq {clientSequenceNumber})");
			}

			throw new LoggingError(				$"SubDirectory.ProcessAckForClear: no pending clear at path '{_absolutePath}' (client seq {clientSequenceNumber})");
		}

		private object? GetOptimisticValueNoLock(string key)
		{
			PendingStorageEntry? latestPendingEntry = FindLatestPendingStorageEntryForKeyOrClearNoLock(key);
			if (latestPendingEntry == null)
			{
				_storage.TryGetValue(key, out object? value);
				return value;
			}

			if (latestPendingEntry is PendingKeyLifetime lifetime && lifetime.KeySets.Count > 0)
			{
				return lifetime.KeySets[lifetime.KeySets.Count - 1].Value;
			}

			return null;
		}

		private bool OptimisticallyHasNoLock(string key)
		{
			PendingStorageEntry? latestPendingEntry = FindLatestPendingStorageEntryForKeyOrClearNoLock(key);
			return latestPendingEntry == null ? _storage.ContainsKey(key) : latestPendingEntry is PendingKeyLifetime;
		}

		private List<KeyValuePair<string, object?>> GetOptimisticEntriesNoLock()
		{
			List<KeyValuePair<string, object?>> entries = new List<KeyValuePair<string, object?>>();
			foreach (KeyValuePair<string, object?> entry in _storage)
			{
				if (!HasPendingDeleteOrClearNoLock(entry.Key))
				{
					entries.Add(new KeyValuePair<string, object?>(entry.Key, GetOptimisticValueNoLock(entry.Key)));
				}
			}

			for (int i = 0; i < _pendingStorageData.Count; i++)
			{
				if (_pendingStorageData[i] is PendingKeyLifetime lifetime && lifetime.KeySets.Count > 0)
				{
					int mostRecentDeleteOrClearIndex = FindLastPendingDeleteOrClearIndexNoLock(lifetime.Key);
					if (i > mostRecentDeleteOrClearIndex && (!_storage.ContainsKey(lifetime.Key) || mostRecentDeleteOrClearIndex != -1))
					{
						PendingKeySet latestPendingValue = lifetime.KeySets[lifetime.KeySets.Count - 1];
						entries.Add(new KeyValuePair<string, object?>(lifetime.Key, latestPendingValue.Value));
					}
				}
			}

			return entries;
		}

		private PendingKeyLifetime GetOrCreatePendingKeyLifetimeNoLock(string key)
		{
			PendingStorageEntry? latestPendingEntry = FindLatestPendingStorageEntryForKeyOrClearNoLock(key);
			if (latestPendingEntry is PendingKeyLifetime lifetime)
			{
				return lifetime;
			}

			PendingKeyLifetime newLifetime = new PendingKeyLifetime(key, this);
			_pendingStorageData.Add(newLifetime);
			return newLifetime;
		}

		private PendingStorageEntry? FindLatestPendingStorageEntryForKeyOrClearNoLock(string key)
		{
			for (int i = _pendingStorageData.Count - 1; i >= 0; i--)
			{
				PendingStorageEntry entry = _pendingStorageData[i];
				if (entry is PendingClear || StorageEntryKeyEquals(entry, key))
				{
					return entry;
				}
			}

			return null;
		}

		private int FindLastPendingDeleteOrClearIndexNoLock(string key)
		{
			for (int i = _pendingStorageData.Count - 1; i >= 0; i--)
			{
				PendingStorageEntry entry = _pendingStorageData[i];
				if (entry is PendingClear || (entry is PendingKeyDelete pendingDelete && pendingDelete.Key == key))
				{
					return i;
				}
			}

			return -1;
		}

		private int FindFirstPendingStorageEntryIndexForKeyNoLock(string key)
		{
			return _pendingStorageData.FindIndex(entry => entry is not PendingClear && StorageEntryKeyEquals(entry, key));
		}

		private bool HasPendingStorageEntryForKeyOrClearNoLock(string key)
		{
			return _pendingStorageData.Any(entry => entry is PendingClear || StorageEntryKeyEquals(entry, key));
		}

		private bool HasPendingStorageEntryForKeyNoLock(string key)
		{
			return _pendingStorageData.Any(entry => StorageEntryKeyEquals(entry, key));
		}

		private bool HasPendingDeleteOrClearNoLock(string key)
		{
			return _pendingStorageData.Any(entry => entry is PendingClear || (entry is PendingKeyDelete pendingDelete && pendingDelete.Key == key));
		}

		private bool HasPendingClearNoLock()
		{
			return _pendingStorageData.Any(entry => entry is PendingClear);
		}

		private static bool StorageEntryKeyEquals(PendingStorageEntry entry, string key)
		{
			return entry switch
			{
				PendingKeyLifetime lifetime => lifetime.Key == key,
				PendingKeyDelete pendingDelete => pendingDelete.Key == key,
				_ => false,
			};
		}

		private SubDirectory? GetOptimisticSubDirectoryNoLock(string subdirName)
		{
			PendingSubDirectoryEntry? latestPendingEntry = FindLatestPendingSubDirectoryEntryNoLock(subdirName);
			if (latestPendingEntry == null)
			{
				return _subdirs.TryGetValue(subdirName, out SubDirectory? subdir) ? subdir : null;
			}

			return latestPendingEntry is PendingSubDirectoryCreate pendingCreate ? pendingCreate.Subdir : null;
		}

		private List<KeyValuePair<string, SubDirectory>> GetOptimisticSubDirectoriesNoLock()
		{
			List<KeyValuePair<string, SubDirectory>> subdirs = new List<KeyValuePair<string, SubDirectory>>();
			HashSet<string> sequencedSubdirNames = new HashSet<string>();
			foreach (string subdirName in _subdirOrder)
			{
				if (!_subdirs.ContainsKey(subdirName))
				{
					continue;
				}

				sequencedSubdirNames.Add(subdirName);
				SubDirectory? optimisticSubdir = GetOptimisticSubDirectoryNoLock(subdirName);
				if (optimisticSubdir != null)
				{
					subdirs.Add(new KeyValuePair<string, SubDirectory>(subdirName, optimisticSubdir));
				}
			}

			HashSet<string> pendingSubdirNames = new HashSet<string>();
			foreach (PendingSubDirectoryEntry entry in _pendingSubDirectoryData)
			{
				if (sequencedSubdirNames.Contains(entry.SubdirName) || !pendingSubdirNames.Add(entry.SubdirName))
				{
					continue;
				}

				SubDirectory? optimisticSubdir = GetOptimisticSubDirectoryNoLock(entry.SubdirName);
				if (optimisticSubdir != null)
				{
					subdirs.Add(new KeyValuePair<string, SubDirectory>(entry.SubdirName, optimisticSubdir));
				}
			}

			return subdirs.OrderBy(entry => entry.Value.SeqData, SeqDataComparer.Instance).ToList();
		}

		private sealed class SeqDataComparer : IComparer<SeqData>
		{
			public static readonly SeqDataComparer Instance = new SeqDataComparer();

			private SeqDataComparer()
			{
			}

			public int Compare(SeqData? a, SeqData? b)
			{
				if (ReferenceEquals(a, b))
				{
					return 0;
				}

				if (a == null)
				{
					return 1;
				}

				if (b == null)
				{
					return -1;
				}

				return SeqDataComparator(a, b);
			}
		}

		private static int SeqDataComparator(SeqData a, SeqData b)
		{
			if (IsAcknowledgedOrDetached(a))
			{
				if (IsAcknowledgedOrDetached(b))
				{
					return a.Seq == b.Seq ? a.ClientSeq.CompareTo(b.ClientSeq) : a.Seq.CompareTo(b.Seq);
				}
				else
				{
					return -1;
				}
			}
			else
			{
				if (IsAcknowledgedOrDetached(b))
				{
					return 1;
				}
				else
				{
					return a.Seq == b.Seq ? a.ClientSeq.CompareTo(b.ClientSeq) : a.Seq.CompareTo(b.Seq);
				}
			}
		}

		private static bool IsAcknowledgedOrDetached(SeqData seqData)
		{
			return seqData.Seq >= 0;
		}

		private PendingSubDirectoryEntry? FindLatestPendingSubDirectoryEntryNoLock(string subdirName)
		{
			for (int i = _pendingSubDirectoryData.Count - 1; i >= 0; i--)
			{
				PendingSubDirectoryEntry entry = _pendingSubDirectoryData[i];
				if (entry.SubdirName == subdirName)
				{
					return entry;
				}
			}

			return null;
		}

		private int FindFirstPendingSubDirectoryEntryIndexNoLock(string subdirName)
		{
			return _pendingSubDirectoryData.FindIndex(entry => entry.SubdirName == subdirName);
		}

		private bool HasPendingSubDirectoryEntryNoLock(string subdirName)
		{
			return _pendingSubDirectoryData.Any(entry => entry.SubdirName == subdirName);
		}

		private bool HasPendingSubDirectoryDeleteNoLock(string subdirName)
		{
			return _pendingSubDirectoryData.Any(entry => entry.SubdirName == subdirName && entry is PendingSubDirectoryDelete);
		}

		private void AddSequencedSubDirectoryNoLock(string subdirName, SubDirectory subdir)
		{
			_subdirs[subdirName] = subdir;
			if (!_subdirOrder.Contains(subdirName))
			{
				_subdirOrder.Add(subdirName);
			}
		}

		private bool RemoveSequencedSubDirectoryNoLock(string subdirName)
		{
			bool removed = _subdirs.Remove(subdirName);
			if (removed)
			{
				_subdirOrder.Remove(subdirName);
			}

			return removed;
		}

		private void RaiseValueChanged(ValueChangedEventArgs args)
		{
			// TS ref: directory.ts. TS emits `containedValueChanged` only on
			// the direct-container subdirectory (this.emit) and `valueChanged` directly
			// on the SharedDirectory (this.directory.emit) — never via bubbling through
			// subdirectory ancestors.
			ValueChangedEventHandler? raiseEvent = OnValueChanged;
			if (raiseEvent != null)
			{
				raiseEvent(this, args);
			}

			_root.PropagateValueChanged(args);
		}

		private void RaiseSubDirectoryCreated(SubDirectoryEventArgs args)
		{
			SubDirectoryEventHandler? raiseEvent = OnSubDirectoryCreated;
			if (raiseEvent != null)
			{
				raiseEvent(this, args);
			}

			if (_parent != null)
			{
				// TS ref: directory.ts uses
				//   this.emit("subDirectoryCreated", posix.join(subDirName, relativePath), ...)
				// when re-emitting a child's event, so each ancestor sees the joined
				// path relative to itself, not the raw name at the emission site.
				SubDirectoryEventArgs bubbledArgs = new SubDirectoryEventArgs()
				{
					SubdirName = args.SubdirName,
					ParentPath = args.ParentPath,
					Path = DirectoryPath.Join(GetLocalName(), args.Path),
					Local = args.Local,
				};
				_parent.RaiseSubDirectoryCreated(bubbledArgs);
				return;
			}

			_root.PropagateSubDirectoryCreated(args);
		}

		private void RaiseSubDirectoryDeleted(SubDirectoryEventArgs args)
		{
			SubDirectoryEventHandler? raiseEvent = OnSubDirectoryDeleted;
			if (raiseEvent != null)
			{
				raiseEvent(this, args);
			}

			if (_parent != null)
			{
				// TS ref: directory.ts — same posix.join semantics as
				// subDirectoryCreated above.
				SubDirectoryEventArgs bubbledArgs = new SubDirectoryEventArgs()
				{
					SubdirName = args.SubdirName,
					ParentPath = args.ParentPath,
					Path = DirectoryPath.Join(GetLocalName(), args.Path),
					Local = args.Local,
				};
				_parent.RaiseSubDirectoryDeleted(bubbledArgs);
				return;
			}

			_root.PropagateSubDirectoryDeleted(args);
		}

		/// <summary>
		/// Returns this subdirectory's name relative to its parent (the key
		/// under which it lives in the parent's map). Empty for the root.
		/// Used when bubbling subdirectory events so each ancestor sees a
		/// joined relative path — matching TS
		/// <c>posix.join(subDirName, relativePath)</c>. Storing the caller-
		/// supplied name is required because <see cref="_absolutePath"/> is
		/// posix-normalized, and normalizing subdirectory names such as
		/// <c>"."</c>, <c>".."</c>, and <c>""</c> would collapse away.
		/// </summary>
		private string GetLocalName() => _localName;

		private static void ValidateSubDirectoryName(string subdirName)
		{
			if (subdirName == null)
			{
				throw new ArgumentNullException(nameof(subdirName));
			}

			if (subdirName.IndexOf("/", StringComparison.Ordinal) >= 0)
			{
				throw new LoggingError($"SubDirectory name may not contain /");
			}
		}

		private List<string> GetAbsolutePathSegments(string relativePath)
		{
			List<string> segments = new List<string>();
			if (!relativePath.StartsWith("/", StringComparison.Ordinal))
			{
				AppendNormalizedSegments(segments, _absolutePath);
			}

			AppendNormalizedSegments(segments, relativePath);
			return segments;
		}

		private static void AppendNormalizedSegments(List<string> segments, string path)
		{
			string[] pathSegments = path.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
			foreach (string segment in pathSegments)
			{
				if (segment == ".")
				{
					continue;
				}

				if (segment == "..")
				{
					if (segments.Count > 0)
					{
						segments.RemoveAt(segments.Count - 1);
					}

					continue;
				}

				segments.Add(segment);
			}
		}

		private static string MakeChildAbsolutePath(string parentPath, string subdirName)
		{
			// TS ref: packages/dds/map/src/directory.ts uses posix.join(this.absolutePath, subdirName).
			// Literal concatenation would preserve '.' and '..' subdirectory names, so
			// CreateSubDirectory(".") would yield '/.' locally but TS would emit '/' on the wire.
			return DirectoryPath.Join(parentPath, subdirName);
		}
	}
}
