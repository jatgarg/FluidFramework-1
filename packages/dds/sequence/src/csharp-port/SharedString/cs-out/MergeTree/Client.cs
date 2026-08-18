// -----------------------------------------------------------------------------
// Ported from packages/dds/merge-tree/src/client.ts (subset for POC)
// Part of the SharedString C# feasibility port — Wave 4.
// 
// POC scope: pending-op tracking, op emission, remote-op reception with
// transformation (block-aware via MergeTree.GetContainingSegment(refSeq)).
// Skipped: sided obliterate, resubmit,
// attribution, short-client-id mapping.
// 
// CORRECTNESS NOTE: Two-client convergence is the primary correctness goal.
// See Wave 8 tests for validation.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;
using System.Text.Json;

using Microsoft.Office.Web.Fluid.Intervals;

namespace Microsoft.Office.Web.Fluid.MergeTree
{
	internal sealed class PendingOpEntry
	{
		public long ClientSeq { get; set; }

		public long LocalSeq { get; init; }

		public List<ISegment> Segments { get; init; } = new();

		public long RefSeq { get; set; }

		public IMergeTreeOp? Op { get; init; }

		public bool IsInsert { get; init; }

		public bool IsRemove { get; init; }

		public bool IsAnnotate { get; init; }

		public bool IsObliterate { get; init; }

		public int AnnotateStart { get; init; }

		public int AnnotateEnd { get; init; }

		public PropertySet? AnnotateProps { get; init; }

		public bool IsInterval { get; init; }

		public IntervalOpKind? IntervalOpKind { get; init; }

		public string? IntervalCollectionName { get; init; }

		public string? IntervalId { get; init; }

		public LocalReferencePosition? IntervalStartReference { get; init; }

		public LocalReferencePosition? IntervalEndReference { get; init; }

		public bool DropAfterRebase { get; set; }
	}

	internal sealed class RebasedSegmentRun
	{
		public int Start { get; set; }

		public int End { get; set; }

		public List<ISegment> Segments { get; } = new();
	}

	public sealed class MergeTreeDelta
	{
		public MergeTreeDeltaType Operation { get; init; }

		public IMergeTreeOp Op { get; init; } = null!;

		public IReadOnlyList<ISegment> Segments { get; init; } = Array.Empty<ISegment>();

		public IReadOnlyList<MergeTreeDeltaRange> Ranges { get; init; } = Array.Empty<MergeTreeDeltaRange>();

		public string ClientId { get; init; } = string.Empty;
	}

	public sealed class MergeTreeDeltaRange
	{
		public ISegment Segment { get; init; } = null!;

		public int Position { get; init; }

		public int Length { get; init; }
	}

	/// <summary>
	/// Client-side wrapper around <see cref="MergeTree" /> for local sequencing, pending-op tracking, and op application.
	/// </summary>
	public sealed class Client
	{
		private readonly List<PendingOpEntry> _pendingOps = new();
		private readonly Dictionary<string, IntervalCollection> _intervalCollections = new(StringComparer.Ordinal);
		private readonly Queue<IMergeTreeOp> _emittedOps = new();
		private long _clientSeq;

		/// <summary>
		/// Initializes a new instance of the <see cref="Client" /> class.
		/// </summary>
		/// <param name="clientId">The long client id for this client.</param>
		/// <param name="mergeTree">Optional merge tree to wrap.</param>
		public Client(string clientId, MergeTree? mergeTree = null)
		{
			if (string.IsNullOrEmpty(clientId))
			{
				throw new ArgumentException("Client id must be provided.", nameof(clientId));
			}

			ClientId = clientId;
			MergeTree = mergeTree ?? new MergeTree();
			CollabWindowCurrentSeq = MergeTree.CurrentSeq;
			CollabWindowMinSeq = MergeTree.MinSeq;
		}

		/// <summary>
		/// Gets the long client id for this client.
		/// </summary>
		public string ClientId { get; }

		/// <summary>
		/// Gets the current local client sequence number.
		/// </summary>
		public long CurrentClientSeq => _clientSeq;

		/// <summary>
		/// Gets the highest server-assigned sequence number observed by this client.
		/// </summary>
		public long CollabWindowCurrentSeq { get; private set; }

		/// <summary>
		/// Gets the current collaboration-window minimum sequence number.
		/// </summary>
		public long CollabWindowMinSeq { get; private set; }

		/// <summary>
		/// Gets the wrapped merge tree.
		/// </summary>
		internal MergeTree MergeTree { get; }

		/// <summary>
		/// Gets the marker id index for live markers in this client.
		/// </summary>
		public MarkerIndex Markers { get; } = new();

		/// <summary>
		/// Gets the number of local ops awaiting acknowledgement.
		/// </summary>
		internal int PendingOpCount => _pendingOps.Count;

		public IntervalCollection GetOrCreateIntervalCollection(string name, IntervalType endpointType)
		{
			if (string.IsNullOrEmpty(name))
			{
				throw new ArgumentException("Interval collection name must be provided.", nameof(name));
			}

			if (!_intervalCollections.TryGetValue(name, out IntervalCollection? collection))
			{
				collection = new IntervalCollection(name, endpointType, MergeTree, new ClientIntervalOpSender(this));
				_intervalCollections[name] = collection;
			}

			return collection;
		}

		public bool TryDequeueEmittedOp(out IMergeTreeOp? op)
		{
			if (_emittedOps.Count == 0)
			{
				op = null;
				return false;
			}

			op = _emittedOps.Dequeue();
			return true;
		}

		internal void AssociatePendingOpsWithClientSequence(IEnumerable<long> localSeqs, long clientSeq)
		{
			ArgumentNullException.ThrowIfNull(localSeqs);

			HashSet<long> localSeqSet = new(localSeqs);
			if (localSeqSet.Count == 0)
			{
				return;
			}

			foreach (PendingOpEntry pending in _pendingOps)
			{
				if (localSeqSet.Contains(pending.LocalSeq))
				{
					pending.ClientSeq = clientSeq;
				}
			}
		}

		internal void SetCollaborationWindow(long minSeq, long currentSeq, bool runZamboni = true)
		{
			CollabWindowCurrentSeq = currentSeq;
			MergeTree.CurrentSeq = currentSeq;
			bool minSeqAdvanced = SetMinimumSequenceNumber(minSeq);
			if (minSeqAdvanced && runZamboni)
			{
				MergeTree.ZamboniSegments();
			}
		}

		/// <summary>
		/// Inserts text locally and returns the outbound insert op.
		/// </summary>
		/// <param name="position">The current local-view position at which to insert.</param>
		/// <param name="text">The text to insert.</param>
		/// <param name="properties">Optional segment properties.</param>
		/// <returns>The outbound merge-tree insert op.</returns>
		public IMergeTreeInsertMsg InsertText(int position, string text, PropertySet? properties = null)
		{
			if (text is null)
			{
				throw new ArgumentNullException(nameof(text));
			}

			if (text.Length == 0)
			{
				return new MergeTreeInsertMsg()
				{
					Pos1 = position,
					Seg = string.Empty,
				};
			}

			long clientSeq = NextClientSeq();
			TextSegment segment = new(text, PropertyMap.ClonePropertySet(properties));

			MergeTree.InsertSegments(
				position,
				new ISegment[] { segment },
				MergeTree.UnassignedSequenceNumber,
				MergeTree.UnassignedSequenceNumber,
				ClientId);
			segment.LocalSeq = clientSeq;
			MergeTreeInsertMsg op = new()
			{
				Pos1 = position,
				Seg = segment.ToJSONObject(),
				ClientSeq = clientSeq,
			};

			_pendingOps.Add(new PendingOpEntry()
			{
				ClientSeq = clientSeq,
				LocalSeq = clientSeq,
				Segments = new List<ISegment>() { segment },
				RefSeq = CollabWindowCurrentSeq,
				Op = op,
				IsInsert = true,
			});

			return op;
		}

		/// <summary>
		/// Inserts a marker locally and returns the outbound insert op.
		/// </summary>
		/// <param name="position">The current local-view position at which to insert.</param>
		/// <param name="refType">The marker reference type.</param>
		/// <param name="properties">Optional marker properties.</param>
		/// <returns>The outbound merge-tree insert op.</returns>
		public IMergeTreeInsertMsg InsertMarker(int position, ReferenceType refType, PropertySet? properties = null)
		{
			long clientSeq = NextClientSeq();
			Marker segment = Marker.Make(refType, properties);

			MergeTree.InsertSegments(
				position,
				new ISegment[] { segment },
				MergeTree.UnassignedSequenceNumber,
				MergeTree.UnassignedSequenceNumber,
				ClientId);
			segment.LocalSeq = clientSeq;
			Markers.Add(segment);
			MergeTreeInsertMsg op = new()
			{
				Pos1 = position,
				Seg = segment.ToJSONObject(),
				ClientSeq = clientSeq,
			};

			_pendingOps.Add(new PendingOpEntry()
			{
				ClientSeq = clientSeq,
				LocalSeq = clientSeq,
				Segments = new List<ISegment>() { segment },
				RefSeq = CollabWindowCurrentSeq,
				Op = op,
				IsInsert = true,
			});

			return op;
		}

		/// <summary>
		/// Removes text locally and returns the outbound remove op.
		/// </summary>
		/// <param name="start">The inclusive current local-view start position.</param>
		/// <param name="end">The exclusive current local-view end position.</param>
		/// <returns>The outbound merge-tree remove op.</returns>
		public IMergeTreeRemoveMsg RemoveText(int start, int end)
		{
			if (end <= start)
			{
				throw new ArgumentOutOfRangeException(nameof(end), "Remove end must be greater than start.");
			}

			long clientSeq = NextClientSeq();
			HashSet<SetRemoveOperationStamp> previouslyPendingRemoved = CollectPendingRemovedStamps();

			MergeTree.MarkRangeRemoved(
				start,
				end,
				MergeTree.UnassignedSequenceNumber,
				MergeTree.UnassignedSequenceNumber,
				ClientId);

			List<ISegment> removedSegments = new();
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				for (int i = 0; i < segment.RemoveStamps.Count; i++)
				{
					if (segment.RemoveStamps[i] is SetRemoveOperationStamp stamp
						&& stamp.Seq == MergeTree.UnassignedSequenceNumber
						&& stamp.LocalSeq is null
						&& !previouslyPendingRemoved.Contains(stamp))
					{
						segment.RemoveStamps[i] = stamp with { LocalSeq = clientSeq };
						removedSegments.Add(segment);
						break;
					}
				}
			}

			RemoveMarkersFromIndex(removedSegments);
			MergeTreeRemoveMsg op = new()
			{
				Pos1 = start,
				Pos2 = end,
				ClientSeq = clientSeq,
			};

			_pendingOps.Add(new PendingOpEntry()
			{
				ClientSeq = clientSeq,
				LocalSeq = clientSeq,
				Segments = removedSegments,
				RefSeq = CollabWindowCurrentSeq,
				Op = op,
				IsRemove = true,
				IsInsert = false,
			});

			return op;
		}

		/// <summary>
		/// Obliterates text locally and returns the outbound obliterate op, or <see langword="null" /> for an empty range.
		/// </summary>
		/// <param name="start">The inclusive current local-view start position.</param>
		/// <param name="end">The exclusive current local-view end position.</param>
		/// <returns>The outbound merge-tree obliterate op, or <see langword="null" /> for a no-op.</returns>
		public IMergeTreeObliterateMsg? ObliterateRangeLocal(int start, int end)
		{
			if (end < start)
			{
				throw new ArgumentOutOfRangeException(nameof(end), "Obliterate end must be greater than or equal to start.");
			}

			if (end == start)
			{
				MergeTree.ObliterateRangeSided(
					SequencePlace.At(start, Side.Before),
					SequencePlace.At(end, Side.Before),
					MergeTree.UnassignedSequenceNumber,
					MergeTree.UnassignedSequenceNumber,
					ClientId);
				return null;
			}

			long clientSeq = NextClientSeq();
			HashSet<SliceRemoveOperationStamp> previouslyPendingObliterated = CollectPendingObliteratedStamps();

			MergeTree.ObliterateRangeSided(
				SequencePlace.At(start, Side.Before),
				SequencePlace.At(end, Side.Before),
				MergeTree.UnassignedSequenceNumber,
				MergeTree.UnassignedSequenceNumber,
				ClientId);

			List<ISegment> obliteratedSegments = new();
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				for (int i = 0; i < segment.RemoveStamps.Count; i++)
				{
					if (segment.RemoveStamps[i] is SliceRemoveOperationStamp stamp
						&& stamp.Seq == MergeTree.UnassignedSequenceNumber
						&& stamp.LocalSeq is null
						&& !previouslyPendingObliterated.Contains(stamp))
					{
						segment.RemoveStamps[i] = stamp with { LocalSeq = clientSeq };
						obliteratedSegments.Add(segment);
						break;
					}
				}
			}

			RemoveMarkersFromIndex(obliteratedSegments);
			MergeTreeObliterateMsg op = new()
			{
				Pos1 = start,
				Pos2 = end,
				ClientSeq = clientSeq,
			};

			_pendingOps.Add(new PendingOpEntry()
			{
				ClientSeq = clientSeq,
				LocalSeq = clientSeq,
				Segments = obliteratedSegments,
				RefSeq = CollabWindowCurrentSeq,
				Op = op,
				IsObliterate = true,
			});

			return op;
		}

		/// <summary>
		/// Obliterates text locally with explicit boundary sidedness and returns the outbound sided obliterate op.
		/// </summary>
		/// <param name="start">The start sequence place.</param>
		/// <param name="end">The end sequence place.</param>
		/// <returns>The outbound merge-tree sided obliterate op, or <see langword="null" /> for a no-op.</returns>
		public IMergeTreeObliterateSidedMsg? ObliterateRangeLocal(SequencePlace start, SequencePlace end)
		{
			if (start is null)
			{
				throw new ArgumentNullException(nameof(start));
			}

			if (end is null)
			{
				throw new ArgumentNullException(nameof(end));
			}

			int length = MergeTree.GetLength(MergeTree.UnassignedSequenceNumber, ClientId);
			int startPosition = NormalizeSequencePlacePosition(start, length);
			int endPosition = NormalizeSequencePlacePosition(end, length);
			if (endPosition < startPosition)
			{
				throw new ArgumentOutOfRangeException(nameof(end), "Obliterate end must be greater than or equal to start.");
			}

			if (endPosition == startPosition && start.Side == end.Side)
			{
				MergeTree.ObliterateRangeSided(
					start,
					end,
					MergeTree.UnassignedSequenceNumber,
					MergeTree.UnassignedSequenceNumber,
					ClientId);
				return null;
			}

			long clientSeq = NextClientSeq();
			HashSet<SliceRemoveOperationStamp> previouslyPendingObliterated = CollectPendingObliteratedStamps();

			MergeTree.ObliterateRangeSided(
				start,
				end,
				MergeTree.UnassignedSequenceNumber,
				MergeTree.UnassignedSequenceNumber,
				ClientId);

			List<ISegment> obliteratedSegments = new();
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				for (int i = 0; i < segment.RemoveStamps.Count; i++)
				{
					if (segment.RemoveStamps[i] is SliceRemoveOperationStamp stamp
						&& stamp.Seq == MergeTree.UnassignedSequenceNumber
						&& stamp.LocalSeq is null
						&& !previouslyPendingObliterated.Contains(stamp))
					{
						segment.RemoveStamps[i] = stamp with { LocalSeq = clientSeq };
						obliteratedSegments.Add(segment);
						break;
					}
				}
			}

			RemoveMarkersFromIndex(obliteratedSegments);
			MergeTreeObliterateSidedMsg op = new()
			{
				Pos1 = start.Clone(),
				Pos2 = end.Clone(),
				ClientSeq = clientSeq,
			};

			_pendingOps.Add(new PendingOpEntry()
			{
				ClientSeq = clientSeq,
				LocalSeq = clientSeq,
				Segments = obliteratedSegments,
				RefSeq = CollabWindowCurrentSeq,
				Op = op,
				IsObliterate = true,
			});

			return op;
		}

		/// <summary>
		/// Annotates text locally and returns the outbound annotate op.
		/// </summary>
		/// <param name="start">The inclusive current local-view start position.</param>
		/// <param name="end">The exclusive current local-view end position.</param>
		/// <param name="props">The properties to merge into the range.</param>
		/// <returns>The outbound merge-tree annotate op.</returns>
		public MergeTreeAnnotateMsg Annotate(int start, int end, PropertySet props)
		{
			if (props is null)
			{
				throw new ArgumentNullException(nameof(props));
			}

			if (end < start)
			{
				throw new ArgumentOutOfRangeException(nameof(end), "Annotate end must be greater than or equal to start.");
			}

			long clientSeq = NextClientSeq();
			PropertySet annotateProps = CloneAnnotateProps(props);
			MergeTree.AnnotateRange(
				start,
				end,
				annotateProps,
				MergeTree.UnassignedSequenceNumber,
				MergeTree.UnassignedSequenceNumber,
				ClientId);
			MergeTreeAnnotateMsg op = new()
			{
				Pos1 = start,
				Pos2 = end,
				Props = CloneAnnotateProps(annotateProps),
				ClientSeq = clientSeq,
			};
			List<ISegment> annotatedSegments = CollectVisibleSegmentsInRange(start, end);
			MarkPendingAnnotateSegments(annotatedSegments, clientSeq, annotateProps);

			_pendingOps.Add(new PendingOpEntry()
			{
				ClientSeq = clientSeq,
				LocalSeq = clientSeq,
				Segments = annotatedSegments,
				RefSeq = CollabWindowCurrentSeq,
				Op = op,
				IsAnnotate = true,
				AnnotateStart = start,
				AnnotateEnd = end,
				AnnotateProps = CloneAnnotateProps(annotateProps),
			});

			return op;
		}

		/// <summary>
		/// Applies a sequenced merge-tree op, handling both remote ops and acknowledgements of local ops.
		/// </summary>
		/// <param name="op">The merge-tree op to apply.</param>
		/// <param name="seq">The server-assigned sequence number.</param>
		/// <param name="refSeq">The reference sequence number the op positions are relative to.</param>
		/// <param name="clientId">The long client id that submitted the op.</param>
		/// <param name="clientSeq">The submitting client's local sequence number, when available.</param>
		public IReadOnlyList<MergeTreeDelta> ApplyOp(
			IMergeTreeOp op,
			long seq,
			long refSeq,
			string clientId,
			long? clientSeq = null,
			long? minimumSequenceNumber = null)
		{
			if (op is null)
			{
				throw new ArgumentNullException(nameof(op));
			}

			if (string.IsNullOrEmpty(clientId))
			{
				throw new ArgumentException("Client id must be provided.", nameof(clientId));
			}

			RecordSequence(seq);

			bool clientSeqIsLocal = !clientSeq.HasValue && op.ClientSeq.HasValue;
			long? effectiveClientSeq = clientSeq ?? op.ClientSeq;
			if (string.Equals(clientId, ClientId, StringComparison.Ordinal))
			{
				AcknowledgeLocalOp(op, seq, effectiveClientSeq, clientSeqIsLocal);
				ApplyOwnIntervalAck(op);
			}
			else
			{
				IReadOnlyList<MergeTreeDelta> deltas = ApplyRemoteOp(op, seq, refSeq, clientId);
				if (minimumSequenceNumber is long remoteMinSeq && SetMinimumSequenceNumber(remoteMinSeq))
				{
					MergeTree.ZamboniSegments();
				}

				return deltas;
			}

			if (minimumSequenceNumber is long minSeq && SetMinimumSequenceNumber(minSeq))
			{
				MergeTree.ZamboniSegments();
			}

			return Array.Empty<MergeTreeDelta>();
		}

		private IReadOnlyList<MergeTreeDelta> ApplyRemoteOp(IMergeTreeOp op, long seq, long refSeq, string clientId, long? perspectiveSeq = null)
		{
			switch (op.Type)
			{
				case MergeTreeDeltaType.Insert:
					return ApplyRemoteInsert((IMergeTreeInsertMsg)op, seq, refSeq, clientId, perspectiveSeq);

				case MergeTreeDeltaType.Remove:
					return ApplyRemoteRemove((IMergeTreeRemoveMsg)op, seq, refSeq, clientId, perspectiveSeq);

				case MergeTreeDeltaType.Obliterate:
					return ApplyRemoteObliterate((IMergeTreeObliterateMsg)op, seq, refSeq, clientId, perspectiveSeq);

				case MergeTreeDeltaType.ObliterateSided:
					return ApplyRemoteObliterateSided((IMergeTreeObliterateSidedMsg)op, seq, refSeq, clientId, perspectiveSeq);

				case MergeTreeDeltaType.Annotate:
					return ApplyRemoteAnnotate((MergeTreeAnnotateMsg)op, seq, refSeq, clientId, perspectiveSeq);

				case MergeTreeDeltaType.Group:
					return ApplyRemoteGroup((MergeTreeGroupMsg)op, seq, refSeq, clientId);

				case MergeTreeDeltaType.IntervalAdd:
					ApplyRemoteIntervalAdd((IntervalAddOpMsg)op);
					return Array.Empty<MergeTreeDelta>();

				case MergeTreeDeltaType.IntervalDelete:
					ApplyRemoteIntervalDelete((IntervalDeleteOpMsg)op);
					return Array.Empty<MergeTreeDelta>();

				case MergeTreeDeltaType.IntervalChange:
					ApplyRemoteIntervalChange((IntervalChangeOpMsg)op);
					return Array.Empty<MergeTreeDelta>();

				case MergeTreeDeltaType.IntervalPropertyChanged:
					ApplyRemoteIntervalPropertyChanged((IntervalPropertyChangedOpMsg)op);
					return Array.Empty<MergeTreeDelta>();

				default:
					throw new NotSupportedException($"Merge-tree op type '{op.Type}' is not supported by the POC Client.");
			}
		}

		/// <summary>
		/// Extracts text from the current local view.
		/// </summary>
		/// <param name="start">The inclusive start position.</param>
		/// <param name="end">The exclusive end position.</param>
		/// <returns>The requested text.</returns>
		public string GetText(int start, int end)
		{
			return MergeTree.GetText(start, end);
		}

		/// <summary>
		/// Gets the current local-view length.
		/// </summary>
		/// <returns>The current local-view length.</returns>
		public int GetLength()
		{
			return MergeTree.GetLength();
		}

		/// <summary>
		/// Finds the segment containing a position in the current local view.
		/// </summary>
		/// <param name="position">The current local-view position.</param>
		/// <returns>The containing segment and offset within that segment.</returns>
		public (ISegment segment, int offsetInSegment) GetContainingSegment(int position)
		{
			return MergeTree.GetContainingSegment(position, MergeTree.UnassignedSequenceNumber, ClientId);
		}

		public Marker? GetMarkerFromId(string id)
		{
			if (Markers.TryGet(id, out Marker? marker)
				&& marker is not null
				&& marker.RemoveStamps.Count == 0)
			{
				return marker;
			}

			return null;
		}

		public int? GetPositionOfMarker(Marker marker)
		{
			ArgumentNullException.ThrowIfNull(marker);

			string? markerId = marker.GetId();
			foreach ((ISegment segment, int segmentPosition) in WalkVisibleSegments())
			{
				if (segment is Marker candidate
					&& (ReferenceEquals(candidate, marker)
						|| (!string.IsNullOrEmpty(markerId)
							&& string.Equals(candidate.GetId(), markerId, StringComparison.Ordinal))))
				{
					return segmentPosition;
				}
			}

			return null;
		}

		public Marker? SearchForTileMarker(int startPos, bool forwards, string? tileLabel = null)
		{
			int length = GetLength();
			if (startPos < 0 || startPos >= length)
			{
				return null;
			}

			if (forwards)
			{
				foreach ((ISegment segment, int position) in WalkVisibleSegments())
				{
					if (position < startPos)
					{
						continue;
					}

					if (IsMatchingTileMarker(segment, tileLabel, out Marker? marker))
					{
						return marker;
					}
				}

				return null;
			}

			Marker? found = null;
			foreach ((ISegment segment, int position) in WalkVisibleSegments())
			{
				if (position > startPos)
				{
					break;
				}

				if (IsMatchingTileMarker(segment, tileLabel, out Marker? marker))
				{
					found = marker;
				}
			}

			return found;
		}

		internal IReadOnlyList<IMergeTreeOp> RebasePendingOps()
		{
			List<IMergeTreeOp> rebasedOps = new();
			if (_pendingOps.Count == 0)
			{
				return rebasedOps;
			}

			foreach (PendingOpEntry pending in _pendingOps)
			{
				pending.DropAfterRebase = false;
			}

			int index = 0;
			while (index < _pendingOps.Count)
			{
				long batchClientSeq = _pendingOps[index].ClientSeq;
				List<IMergeTreeOp> batchOps = new();
				do
				{
					batchOps.AddRange(RebasePendingOp(_pendingOps[index]));
					index++;
				}
				while (index < _pendingOps.Count && _pendingOps[index].ClientSeq == batchClientSeq);

				AddRebasedBatchOps(rebasedOps, batchOps);
			}

			for (int i = _pendingOps.Count - 1; i >= 0; i--)
			{
				if (_pendingOps[i].DropAfterRebase)
				{
					_pendingOps.RemoveAt(i);
				}
			}

			return rebasedOps;
		}

		private List<IMergeTreeOp> RebasePendingOp(PendingOpEntry pending)
		{
			if (pending.IsInterval)
			{
				return RebasePendingInterval(pending);
			}

			if (pending.IsInsert)
			{
				return RebasePendingInsert(pending);
			}

			if (pending.IsAnnotate)
			{
				return RebasePendingAnnotate(pending);
			}

			if (pending.IsObliterate)
			{
				return RebasePendingObliterate(pending);
			}

			if (pending.IsRemove)
			{
				return RebasePendingRemove(pending);
			}

			pending.DropAfterRebase = true;
			return new List<IMergeTreeOp>();
		}

		private List<IMergeTreeOp> RebasePendingInterval(PendingOpEntry pending)
		{
			if (pending.Op is null)
			{
				pending.DropAfterRebase = true;
				return new List<IMergeTreeOp>();
			}

			if (ShouldDropDetachedTransientPendingAdd(pending))
			{
				DropDetachedPendingInterval(pending);
				pending.DropAfterRebase = true;
				return new List<IMergeTreeOp>();
			}

			IMergeTreeOp rebased = CloneOp(pending.Op);
			rebased.ClientSeq = pending.LocalSeq;
			RebaseIntervalEndpointPositions(pending, rebased);
			pending.RefSeq = CollabWindowCurrentSeq;
			return new List<IMergeTreeOp>() { rebased };
		}

		private void DropDetachedPendingInterval(PendingOpEntry pending)
		{
			if (pending.IntervalCollectionName is null || pending.IntervalId is null)
			{
				return;
			}

			if (_intervalCollections.TryGetValue(pending.IntervalCollectionName, out IntervalCollection? collection))
			{
				collection.DropDetachedIntervalForRebase(pending.IntervalId);
			}
		}

		private static bool ShouldDropDetachedTransientPendingAdd(PendingOpEntry pending)
		{
			return pending.Op is IntervalAddOpMsg
				&& (IsDetachedTransientReference(pending.IntervalStartReference)
					|| IsDetachedTransientReference(pending.IntervalEndReference));
		}

		private static bool IsDetachedTransientReference(LocalReferencePosition? reference)
		{
			return reference is not null
				&& reference.IsDetached
				&& (reference.RefType & ReferenceType.Transient) == ReferenceType.Transient;
		}

		private void RebaseIntervalEndpointPositions(PendingOpEntry pending, IMergeTreeOp rebased)
		{
			long localSeq = pending.LocalSeq - 1;
			switch (rebased)
			{
				case IntervalAddOpMsg addOp:
					addOp.Start = RebaseIntervalEndpointPosition(
						pending.IntervalStartReference,
						addOp.Start,
						pending.RefSeq,
						localSeq);
					addOp.End = RebaseIntervalEndpointPosition(
						pending.IntervalEndReference,
						addOp.End,
						pending.RefSeq,
						localSeq);
					break;

				case IntervalChangeOpMsg changeOp:
					if (changeOp.Start.HasValue)
					{
						changeOp.Start = RebaseIntervalEndpointPosition(
							pending.IntervalStartReference,
							changeOp.Start.Value,
							pending.RefSeq,
							localSeq);
					}

					if (changeOp.End.HasValue)
					{
						changeOp.End = RebaseIntervalEndpointPosition(
							pending.IntervalEndReference,
							changeOp.End.Value,
							pending.RefSeq,
							localSeq);
					}

					break;
			}
		}

		private int RebaseIntervalEndpointPosition(
			LocalReferencePosition? reference,
			int stalePosition,
			long staleRefSeq,
			long localSeq)
		{
			if (reference is not null
				&& TryGetPositionOfReferenceForReconnect(reference, localSeq) is int referencePosition)
			{
				return referencePosition;
			}

			return RebaseStaleIntervalPosition(reference, stalePosition, staleRefSeq, localSeq);
		}

		private int? TryGetPositionOfReferenceForReconnect(LocalReferencePosition reference, long localSeq)
		{
			if (reference.Segment is null)
			{
				return null;
			}

			int? segmentPosition = MergeTree.GetPositionOfSegmentForReconnect(reference.Segment, ClientId, localSeq);
			if (segmentPosition is not int start)
			{
				return null;
			}

			int length = MergeTree.GetVisibleLengthForReconnect(reference.Segment, ClientId, localSeq);
			if (length == 0)
			{
				return null;
			}

			return Math.Min(start + reference.Offset, GetLengthForReconnect(localSeq));
		}

		private int RebaseStaleIntervalPosition(
			LocalReferencePosition? reference,
			int stalePosition,
			long staleRefSeq,
			long localSeq)
		{
			int staleLength = GetLengthForIntervalStalePerspective(staleRefSeq, localSeq);
			int clampedStalePosition = Math.Clamp(stalePosition, 0, staleLength);
			int currentLength = GetLengthForReconnect(localSeq);
			if (clampedStalePosition == staleLength)
			{
				return currentLength;
			}

			(ISegment segment, int offsetInSegment)? segmentInfo = TryGetContainingSegmentForIntervalStalePerspective(
				clampedStalePosition,
				staleRefSeq,
				localSeq);
			if (segmentInfo is null)
			{
				return reference is not null
						&& TryGetPositionOfReferenceForReconnect(reference, localSeq) is int referencePosition
					? referencePosition
					: Math.Min(clampedStalePosition, currentLength);
			}

			int? segmentPosition = MergeTree.GetPositionOfSegmentForReconnect(
				segmentInfo.Value.segment,
				ClientId,
				localSeq);
			if (segmentPosition is int start)
			{
				return Math.Min(start + segmentInfo.Value.offsetInSegment, currentLength);
			}

			return SlideIntervalPosition(segmentInfo.Value.segment, localSeq, reference?.SlidingPreference ?? SlidingPreference.Forward);
		}

		private (ISegment Segment, int OffsetInSegment)? TryGetContainingSegmentForIntervalStalePerspective(
			int position,
			long staleRefSeq,
			long localSeq)
		{
			if (position < 0)
			{
				return null;
			}

			int currentPosition = 0;
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				int length = GetVisibleLengthForIntervalStalePerspective(segment, staleRefSeq, localSeq);
				if (length == 0)
				{
					continue;
				}

				int nextPosition = currentPosition + length;
				if (position < nextPosition)
				{
					return (segment, position - currentPosition);
				}

				currentPosition = nextPosition;
			}

			return null;
		}

		private int GetLengthForIntervalStalePerspective(long staleRefSeq, long localSeq)
		{
			int length = 0;
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				length += GetVisibleLengthForIntervalStalePerspective(segment, staleRefSeq, localSeq);
			}

			return length;
		}

		private int GetVisibleLengthForIntervalStalePerspective(ISegment segment, long staleRefSeq, long localSeq)
		{
			int reconnectClientId = MergeTree.GetOrAddShortClientId(ClientId);
			bool inserted = IsAssignedAtOrBefore(segment.Seq, staleRefSeq)
				|| (segment.Seq == MergeTree.UnassignedSequenceNumber
					&& segment.LocalSeq is long insertLocalSeq
					&& segment.ClientId == reconnectClientId
					&& insertLocalSeq <= localSeq);
			if (!inserted)
			{
				return 0;
			}

			foreach (RemoveOperationStamp stamp in segment.RemoveStamps)
			{
				if (IsHiddenForIntervalStalePerspective(stamp, reconnectClientId, staleRefSeq, localSeq))
				{
					return 0;
				}
			}

			return segment.CachedLength;
		}

		private static bool IsHiddenForIntervalStalePerspective(
			RemoveOperationStamp stamp,
			int reconnectClientId,
			long staleRefSeq,
			long reconnectLocalSeq)
		{
			if (stamp.Seq != MergeTree.UnassignedSequenceNumber)
			{
				return stamp.Seq <= staleRefSeq;
			}

			return stamp.LocalSeq is long local
				&& stamp.ClientId == reconnectClientId
				&& local <= reconnectLocalSeq;
		}

		private static bool IsAssignedAtOrBefore(long seq, long refSeq)
		{
			return seq != MergeTree.UnassignedSequenceNumber && seq <= refSeq;
		}

		private int SlideIntervalPosition(ISegment removedSegment, long localSeq, SlidingPreference slidingPreference)
		{
			int position = 0;
			int? previousVisibleEnd = null;
			bool foundRemovedSegment = false;
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				int length = MergeTree.GetVisibleLengthForReconnect(segment, ClientId, localSeq);
				if (ReferenceEquals(segment, removedSegment))
				{
					foundRemovedSegment = true;
					if (length > 0)
					{
						return position;
					}
				}
				else if (length > 0)
				{
					if (foundRemovedSegment)
					{
						return slidingPreference == SlidingPreference.Backward ? previousVisibleEnd ?? 0 : position;
					}

					previousVisibleEnd = position + length;
				}

				position += length;
			}

			// If the endpoint reference is already detached, fall back to the legacy
			// interval-position rebase path.
			return foundRemovedSegment ? previousVisibleEnd ?? 0 : Math.Min(position, GetLengthForReconnect(localSeq));
		}

		private int GetLengthForReconnect(long localSeq)
		{
			int length = 0;
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				length += MergeTree.GetVisibleLengthForReconnect(segment, ClientId, localSeq);
			}

			return length;
		}

		private List<IMergeTreeOp> RebasePendingInsert(PendingOpEntry pending)
		{
			List<ISegment> insertedSegments = CollectSegmentsInTreeOrder(
				segment => segment.LocalSeq == pending.LocalSeq);
			if (insertedSegments.Count == 0)
			{
				insertedSegments.AddRange(pending.Segments);
			}

			ISegment? firstVisibleSegment = null;
			foreach (ISegment segment in insertedSegments)
			{
				if (MergeTree.GetPositionOfSegmentForReconnect(segment, ClientId, pending.LocalSeq) is not null)
				{
					firstVisibleSegment = segment;
					break;
				}
			}

			if (firstVisibleSegment is null)
			{
				CompleteSkippedInsert(pending, insertedSegments);
				return new List<IMergeTreeOp>();
			}

			int position = MergeTree.GetPositionOfSegmentForReconnect(firstVisibleSegment, ClientId, pending.LocalSeq)
				?? throw new InvalidOperationException("Pending insert segment must be visible in its reconnect perspective.");
			MergeTreeInsertMsg op = new()
			{
				Pos1 = position,
				Seg = CloneInsertSegmentSpec(pending),
				ClientSeq = pending.LocalSeq,
			};
			ReplacePendingSegments(pending, insertedSegments);
			pending.RefSeq = CollabWindowCurrentSeq;
			return new List<IMergeTreeOp>() { op };
		}

		private List<IMergeTreeOp> RebasePendingRemove(PendingOpEntry pending)
		{
			List<ISegment> removedSegments = CollectSegmentsInTreeOrder(
				segment => HasSetRemoveStamp(segment, stamp => stamp.LocalSeq == pending.LocalSeq));
			List<ISegment> activeRemovedSegments = new();
			List<ISegment> staleRemovedSegments = new();
			foreach (ISegment segment in removedSegments)
			{
				if (HasSetRemoveStamp(segment, stamp => stamp.LocalSeq == pending.LocalSeq && stamp.Seq == MergeTree.UnassignedSequenceNumber))
				{
					activeRemovedSegments.Add(segment);
				}
				else
				{
					staleRemovedSegments.Add(segment);
				}
			}

			ClearStaleRemoveLocalSeqs(staleRemovedSegments, pending.LocalSeq);
			List<(int Start, int End)> ranges = MakeRangesSequentialForRemove(
				BuildRebasedRanges(activeRemovedSegments, pending.LocalSeq - 1));
			if (ranges.Count == 0)
			{
				ClearStaleRemoveLocalSeqs(activeRemovedSegments, pending.LocalSeq);
				pending.DropAfterRebase = true;
				MergeTree.RebuildFromRoot();
				return new List<IMergeTreeOp>();
			}

			ReplacePendingSegments(pending, activeRemovedSegments);
			pending.RefSeq = CollabWindowCurrentSeq;
			List<IMergeTreeOp> ops = new();
			foreach ((int start, int end) in ranges)
			{
				ops.Add(new MergeTreeRemoveMsg()
				{
					Pos1 = start,
					Pos2 = end,
					ClientSeq = pending.LocalSeq,
				});
			}

			return ops;
		}

		private List<IMergeTreeOp> RebasePendingAnnotate(PendingOpEntry pending)
		{
			if (pending.AnnotateProps is null)
			{
				pending.DropAfterRebase = true;
				return new List<IMergeTreeOp>();
			}

			List<ISegment> annotateSegments = new();
			foreach (ISegment segment in CollectSegmentsInTreeOrder(
				segment => segment.PendingAnnotates is not null && segment.PendingAnnotates.ContainsKey(pending.LocalSeq)))
			{
				if (MergeTree.GetVisibleLengthForReconnect(segment, ClientId, pending.LocalSeq - 1) > 0)
				{
					annotateSegments.Add(segment);
				}
			}

			List<(int Start, int End)> ranges = BuildRebasedRanges(annotateSegments, pending.LocalSeq - 1);
			if (ranges.Count == 0)
			{
				AcknowledgeAnnotateSegments(pending.LocalSeq);
				pending.DropAfterRebase = true;
				return new List<IMergeTreeOp>();
			}

			ReplacePendingSegments(pending, annotateSegments);
			pending.RefSeq = CollabWindowCurrentSeq;
			List<IMergeTreeOp> ops = new();
			foreach ((int start, int end) in ranges)
			{
				ops.Add(new MergeTreeAnnotateMsg()
				{
					Pos1 = start,
					Pos2 = end,
					Props = CloneAnnotateProps(pending.AnnotateProps),
					ClientSeq = pending.LocalSeq,
				});
			}

			return ops;
		}

		private List<IMergeTreeOp> RebasePendingObliterate(PendingOpEntry pending)
		{
			List<ISegment> obliteratedSegments = CollectSegmentsInTreeOrder(
				segment => HasSliceRemoveStamp(segment, stamp => stamp.LocalSeq == pending.LocalSeq));
			List<ISegment> activeObliteratedSegments = new();
			List<ISegment> staleObliteratedSegments = new();
			bool removedExpandedStamp = false;
			foreach (ISegment segment in obliteratedSegments)
			{
				if (!WasPresentWhenPendingOpSubmitted(segment, pending))
				{
					removedExpandedStamp = RemoveSliceRemoveStamp(segment, pending.LocalSeq) || removedExpandedStamp;
					continue;
				}

				if (HasSliceRemoveStamp(segment, stamp => stamp.LocalSeq == pending.LocalSeq && stamp.Seq == MergeTree.UnassignedSequenceNumber))
				{
					activeObliteratedSegments.Add(segment);
				}
				else
				{
					staleObliteratedSegments.Add(segment);
				}
			}

			ClearStaleObliterateLocalSeqs(staleObliteratedSegments, pending.LocalSeq);
			List<RebasedSegmentRun> runs = BuildRebasedSegmentRuns(activeObliteratedSegments, pending.LocalSeq - 1);
			List<ISegment> representedSegments = CollectRunSegments(runs);
			HashSet<ISegment> representedSegmentSet = new(representedSegments);
			foreach (ISegment segment in activeObliteratedSegments)
			{
				if (!representedSegmentSet.Contains(segment))
				{
					removedExpandedStamp = RemoveSliceRemoveStamp(segment, pending.LocalSeq) || removedExpandedStamp;
				}
			}

			if (removedExpandedStamp)
			{
				MergeTree.RebuildFromRoot();
			}

			if (runs.Count == 0)
			{
				pending.DropAfterRebase = true;
				return new List<IMergeTreeOp>();
			}

			ReplacePendingSegments(pending, representedSegments);
			pending.RefSeq = CollabWindowCurrentSeq;
			UpdatePendingObliterateBoundarySides(pending, runs);
			List<(int Start, int End)> sequentialRanges = MakeRunsSequentialForRemove(runs);
			if (pending.Op is MergeTreeObliterateSidedMsg sidedObliterate)
			{
				List<IMergeTreeOp> sidedOps = new();
				for (int i = 0; i < sequentialRanges.Count; i++)
				{
					(int start, int end) = sequentialRanges[i];
					sidedOps.Add(new MergeTreeObliterateSidedMsg()
					{
						Pos1 = SequencePlace.At(start, i == 0 ? sidedObliterate.Pos1.Side : Side.Before),
						Pos2 = SequencePlace.At(end, i == sequentialRanges.Count - 1 ? sidedObliterate.Pos2.Side : Side.Before),
						ClientSeq = pending.LocalSeq,
					});
				}

				return sidedOps;
			}

			List<IMergeTreeOp> ops = new();
			foreach ((int start, int end) in sequentialRanges)
			{
				ops.Add(new MergeTreeObliterateMsg()
				{
					Pos1 = start,
					Pos2 = end,
					ClientSeq = pending.LocalSeq,
				});
			}

			return ops;
		}

		private bool WasPresentWhenPendingOpSubmitted(ISegment segment, PendingOpEntry pending)
		{
			int localClientId = MergeTree.GetOrAddShortClientId(ClientId);
			bool inserted = segment.Seq != MergeTree.UnassignedSequenceNumber
				? segment.Seq <= pending.RefSeq
				: segment.ClientId == localClientId
					&& segment.LocalSeq is long insertLocalSeq
					&& insertLocalSeq <= pending.LocalSeq;
			if (!inserted)
			{
				return false;
			}

			foreach (RemoveOperationStamp stamp in segment.RemoveStamps)
			{
				if (stamp.LocalSeq == pending.LocalSeq && stamp.ClientId == localClientId)
				{
					continue;
				}

				bool removed = stamp.Seq != MergeTree.UnassignedSequenceNumber
					? stamp.Seq <= pending.RefSeq
					: stamp.ClientId == localClientId
						&& stamp.LocalSeq is long removeLocalSeq
						&& removeLocalSeq < pending.LocalSeq;
				if (removed)
				{
					return false;
				}
			}

			return true;
		}

		private static List<ISegment> CollectRunSegments(IEnumerable<RebasedSegmentRun> runs)
		{
			List<ISegment> segments = new();
			foreach (RebasedSegmentRun run in runs)
			{
				foreach (ISegment segment in run.Segments)
				{
					segments.Add(segment);
				}
			}

			return segments;
		}

		private static void UpdatePendingObliterateBoundarySides(PendingOpEntry pending, IReadOnlyList<RebasedSegmentRun> runs)
		{
			Side startSide = pending.Op is MergeTreeObliterateSidedMsg sidedObliterate ? sidedObliterate.Pos1.Side : Side.Before;
			Side endSide = pending.Op is MergeTreeObliterateSidedMsg sidedObliterateEnd ? sidedObliterateEnd.Pos2.Side : Side.Before;
			foreach (RebasedSegmentRun run in runs)
			{
				foreach (ISegment segment in run.Segments)
				{
					UpdateSliceRemoveStamp(
						segment,
						pending.LocalSeq,
						stamp => stamp with
						{
							StartSide = null,
							EndSide = null,
						});
				}
			}

			for (int i = 0; i < runs.Count; i++)
			{
				RebasedSegmentRun run = runs[i];
				UpdateSliceRemoveStamp(
					run.Segments[0],
					pending.LocalSeq,
					stamp => stamp with { StartSide = i == 0 ? startSide : Side.Before });
				UpdateSliceRemoveStamp(
					run.Segments[run.Segments.Count - 1],
					pending.LocalSeq,
					stamp => stamp with { EndSide = i == runs.Count - 1 ? endSide : Side.Before });
			}
		}

		private static void AddRebasedBatchOps(List<IMergeTreeOp> rebasedOps, List<IMergeTreeOp> batchOps)
		{
			if (batchOps.Count == 0)
			{
				return;
			}

			if (batchOps.Count == 1)
			{
				rebasedOps.Add(batchOps[0]);
				return;
			}

			// Peel interval ops out — they go on the wire individually and cannot be
			// members of a MergeTreeGroupMsg (whose members are merge-tree delta ops only).
			MergeTreeGroupMsg groupOp = new();
			List<IMergeTreeOp> intervalOps = new();
			foreach (IMergeTreeOp op in batchOps)
			{
				if (op is IntervalOpMsg)
				{
					intervalOps.Add(op);
					continue;
				}

				if (op is not MergeTreeOp mergeTreeOp)
				{
					rebasedOps.Add(op);
					continue;
				}

				groupOp.Ops.Add(mergeTreeOp);
			}

			if (groupOp.Ops.Count == 1)
			{
				rebasedOps.Add(groupOp.Ops[0]);
			}
			else if (groupOp.Ops.Count > 1)
			{
				rebasedOps.Add(groupOp);
			}

			rebasedOps.AddRange(intervalOps);
		}

		private object? CloneInsertSegmentSpec(PendingOpEntry pending)
		{
			if (pending.Op is IMergeTreeInsertMsg insertMsg)
			{
				return CloneSegmentSpec(insertMsg.Seg);
			}

			return pending.Segments.Count == 0 ? null : CloneSegmentSpec(pending.Segments[0].ToJSONObject());
		}

		private List<(int Start, int End)> BuildRebasedRanges(IReadOnlyList<ISegment> segments, long localSeq)
		{
			List<(int Start, int End)> ranges = new();
			foreach (RebasedSegmentRun run in BuildRebasedSegmentRuns(segments, localSeq))
			{
				ranges.Add((run.Start, run.End));
			}

			return ranges;
		}

		private List<RebasedSegmentRun> BuildRebasedSegmentRuns(IReadOnlyList<ISegment> segments, long localSeq)
		{
			List<RebasedSegmentRun> runs = new();
			foreach (ISegment segment in segments)
			{
				int? position = MergeTree.GetPositionOfSegmentForReconnect(segment, ClientId, localSeq);
				if (position is not int start)
				{
					continue;
				}

				int length = MergeTree.GetVisibleLengthForReconnect(segment, ClientId, localSeq);
				if (length == 0)
				{
					continue;
				}

				int end = start + length;
				if (runs.Count > 0)
				{
					RebasedSegmentRun previous = runs[runs.Count - 1];
					if (previous.End == start)
					{
						previous.End = end;
						previous.Segments.Add(segment);
						continue;
					}
				}

				RebasedSegmentRun run = new()
				{
					Start = start,
					End = end,
				};
				run.Segments.Add(segment);
				runs.Add(run);
			}

			return runs;
		}

		private bool TryBuildEnclosingRange(IReadOnlyList<ISegment> segments, long localSeq, out int start, out int end)
		{
			start = 0;
			end = 0;
			bool found = false;
			foreach (ISegment segment in segments)
			{
				int? position = MergeTree.GetPositionOfSegmentForReconnect(segment, ClientId, localSeq);
				if (position is not int segmentStart)
				{
					continue;
				}

				int length = MergeTree.GetVisibleLengthForReconnect(segment, ClientId, localSeq);
				if (length == 0)
				{
					continue;
				}

				int segmentEnd = segmentStart + length;
				if (!found)
				{
					start = segmentStart;
					end = segmentEnd;
					found = true;
				}
				else
				{
					start = Math.Min(start, segmentStart);
					end = Math.Max(end, segmentEnd);
				}
			}

			return found && end > start;
		}

		private static List<(int Start, int End)> MakeRangesSequentialForRemove(IReadOnlyList<(int Start, int End)> ranges)
		{
			List<(int Start, int End)> sequentialRanges = new();
			int removedLength = 0;
			foreach ((int start, int end) in ranges)
			{
				int rebasedStart = start - removedLength;
				int rebasedEnd = end - removedLength;
				if (rebasedEnd <= rebasedStart)
				{
					continue;
				}

				sequentialRanges.Add((rebasedStart, rebasedEnd));
				removedLength += end - start;
			}

			return sequentialRanges;
		}

		private static List<(int Start, int End)> MakeRunsSequentialForRemove(IReadOnlyList<RebasedSegmentRun> runs)
		{
			List<(int Start, int End)> ranges = new();
			foreach (RebasedSegmentRun run in runs)
			{
				ranges.Add((run.Start, run.End));
			}

			return MakeRangesSequentialForRemove(ranges);
		}

		private List<ISegment> CollectVisibleSegmentsInRange(int start, int end)
		{
			List<ISegment> segments = new();
			if (end <= start)
			{
				return segments;
			}

			foreach ((ISegment segment, int _, int _) in MergeTree.GetSegments(start, end))
			{
				segments.Add(segment);
			}

			return segments;
		}

		private List<ISegment> CollectSegmentsInTreeOrder(Func<ISegment, bool> predicate)
		{
			List<ISegment> segments = new();
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				if (predicate(segment))
				{
					segments.Add(segment);
				}
			}

			return segments;
		}

		private static void ReplacePendingSegments(PendingOpEntry pending, IEnumerable<ISegment> segments)
		{
			pending.Segments.Clear();
			pending.Segments.AddRange(segments);
		}

		private void CompleteSkippedInsert(PendingOpEntry pending, IReadOnlyList<ISegment> insertedSegments)
		{
			foreach (ISegment segment in insertedSegments)
			{
				if (segment.LocalSeq == pending.LocalSeq)
				{
					segment.LocalSeq = null;
					if (segment.Seq == MergeTree.UnassignedSequenceNumber)
					{
						segment.Seq = MergeTree.UniversalSequenceNumber;
						segment.ClientId = Constants.NonCollabClient;
					}
				}
			}

			pending.DropAfterRebase = true;
			MergeTree.RebuildFromRoot();
		}

		private static void ClearStaleRemoveLocalSeqs(IEnumerable<ISegment> segments, long localSeq)
		{
			foreach (ISegment segment in segments)
			{
				for (int i = 0; i < segment.RemoveStamps.Count; i++)
				{
					if (segment.RemoveStamps[i] is SetRemoveOperationStamp stamp && stamp.LocalSeq == localSeq)
					{
						segment.RemoveStamps[i] = stamp with { LocalSeq = null };
					}
				}
			}
		}

		private static void ClearStaleObliterateLocalSeqs(IEnumerable<ISegment> segments, long localSeq)
		{
			foreach (ISegment segment in segments)
			{
				for (int i = 0; i < segment.RemoveStamps.Count; i++)
				{
					if (segment.RemoveStamps[i] is SliceRemoveOperationStamp stamp && stamp.LocalSeq == localSeq)
					{
						segment.RemoveStamps[i] = stamp with { LocalSeq = null };
					}
				}
			}
		}

		private static IMergeTreeOp CloneOp(IMergeTreeOp op)
		{
			switch (op)
			{
				case MergeTreeInsertMsg insertMsg:
					return new MergeTreeInsertMsg()
					{
						Pos1 = insertMsg.Pos1,
						RelativePos1 = insertMsg.RelativePos1,
						Pos2 = insertMsg.Pos2,
						RelativePos2 = insertMsg.RelativePos2,
						Seg = CloneSegmentSpec(insertMsg.Seg),
						ClientSeq = insertMsg.ClientSeq,
					};

				case MergeTreeRemoveMsg removeMsg:
					return new MergeTreeRemoveMsg()
					{
						Pos1 = removeMsg.Pos1,
						RelativePos1 = removeMsg.RelativePos1,
						Pos2 = removeMsg.Pos2,
						RelativePos2 = removeMsg.RelativePos2,
						ClientSeq = removeMsg.ClientSeq,
					};

				case MergeTreeObliterateMsg obliterateMsg:
					return new MergeTreeObliterateMsg()
					{
						Pos1 = obliterateMsg.Pos1,
						RelativePos1 = obliterateMsg.RelativePos1,
						Pos2 = obliterateMsg.Pos2,
						RelativePos2 = obliterateMsg.RelativePos2,
						ClientSeq = obliterateMsg.ClientSeq,
					};

				case MergeTreeObliterateSidedMsg sidedObliterateMsg:
					return new MergeTreeObliterateSidedMsg()
					{
						Pos1 = sidedObliterateMsg.Pos1.Clone(),
						RelativePos1 = sidedObliterateMsg.RelativePos1,
						Pos2 = sidedObliterateMsg.Pos2.Clone(),
						RelativePos2 = sidedObliterateMsg.RelativePos2,
						ClientSeq = sidedObliterateMsg.ClientSeq,
					};

				case MergeTreeAnnotateMsg annotateMsg:
					return new MergeTreeAnnotateMsg()
					{
						Pos1 = annotateMsg.Pos1,
						Pos2 = annotateMsg.Pos2,
						Props = CloneAnnotateProps(annotateMsg.Props),
						ClientSeq = annotateMsg.ClientSeq,
					};

				case MergeTreeGroupMsg groupMsg:
					MergeTreeGroupMsg groupClone = new()
					{
						ClientSeq = groupMsg.ClientSeq,
					};
					foreach (MergeTreeOp memberOp in groupMsg.Ops)
					{
						groupClone.Ops.Add((MergeTreeOp)CloneOp(memberOp));
					}

					return groupClone;

				case IntervalAddOpMsg addOp:
					return new IntervalAddOpMsg()
					{
						CollectionName = addOp.CollectionName,
						IntervalId = addOp.IntervalId,
						Start = addOp.Start,
						End = addOp.End,
						IntervalType = addOp.IntervalType,
						Stickiness = addOp.Stickiness,
						StartSide = addOp.StartSide,
						EndSide = addOp.EndSide,
						Props = PropertyMap.ClonePropertySet(addOp.Props),
						ClientSeq = addOp.ClientSeq,
					};

				case IntervalDeleteOpMsg deleteOp:
					return new IntervalDeleteOpMsg()
					{
						CollectionName = deleteOp.CollectionName,
						IntervalId = deleteOp.IntervalId,
						ClientSeq = deleteOp.ClientSeq,
					};

				case IntervalChangeOpMsg changeOp:
					return new IntervalChangeOpMsg()
					{
						CollectionName = changeOp.CollectionName,
						IntervalId = changeOp.IntervalId,
						Start = changeOp.Start,
						End = changeOp.End,
						Stickiness = changeOp.Stickiness,
						StartSide = changeOp.StartSide,
						EndSide = changeOp.EndSide,
						ClientSeq = changeOp.ClientSeq,
					};

				case IntervalPropertyChangedOpMsg propertyChangedOp:
					return new IntervalPropertyChangedOpMsg()
					{
						CollectionName = propertyChangedOp.CollectionName,
						IntervalId = propertyChangedOp.IntervalId,
						Props = PropertyMap.ClonePropertySet(propertyChangedOp.Props) ?? new PropertySet(),
						ClientSeq = propertyChangedOp.ClientSeq,
					};

				default:
					throw new OcsException(OcsGateErrorCode.UnknownOp, $"Cannot clone op runtime type: {op.GetType().FullName}");
			}
		}

		private static object? CloneSegmentSpec(object? spec)
		{
			switch (spec)
			{
				case null:
					return null;

				case string text:
					return text;

				case IJSONTextSegment textSpec:
					return new JSONTextSegment()
					{
						Text = textSpec.Text,
						Props = PropertyMap.ClonePropertySet(textSpec.Props),
					};

				case IJSONMarkerSegment markerSpec:
					return new JSONMarkerSegment()
					{
						Marker = new JSONMarkerDef()
						{
							RefType = markerSpec.Marker.RefType,
							Props = PropertyMap.ClonePropertySet(markerSpec.Marker.Props),
						},
						Props = PropertyMap.ClonePropertySet(markerSpec.Props),
					};

				case ISegment segment:
					return CloneSegmentSpec(segment.ToJSONObject());

				case JsonElement jsonElement:
					return jsonElement.Clone();

				default:
					return spec;
			}
		}

		private long NextClientSeq()
		{
			_clientSeq++;
			return _clientSeq;
		}

		private static IEnumerable<SetRemoveOperationStamp> FindSetRemoveStamps(ISegment segment)
		{
			foreach (RemoveOperationStamp stamp in segment.RemoveStamps)
			{
				if (stamp is SetRemoveOperationStamp setRemoveStamp)
				{
					yield return setRemoveStamp;
				}
			}
		}

		private static IEnumerable<SliceRemoveOperationStamp> FindSliceRemoveStamps(ISegment segment)
		{
			foreach (RemoveOperationStamp stamp in segment.RemoveStamps)
			{
				if (stamp is SliceRemoveOperationStamp sliceRemoveStamp)
				{
					yield return sliceRemoveStamp;
				}
			}
		}

		private static SliceRemoveOperationStamp? TryGetSliceRemoveStamp(ISegment segment, long localSeq)
		{
			foreach (SliceRemoveOperationStamp stamp in FindSliceRemoveStamps(segment))
			{
				if (stamp.LocalSeq == localSeq)
				{
					return stamp;
				}
			}

			return null;
		}

		private static bool UpdateSliceRemoveStamp(
			ISegment segment,
			long localSeq,
			Func<SliceRemoveOperationStamp, SliceRemoveOperationStamp> update)
		{
			for (int i = 0; i < segment.RemoveStamps.Count; i++)
			{
				if (segment.RemoveStamps[i] is SliceRemoveOperationStamp stamp && stamp.LocalSeq == localSeq)
				{
					segment.RemoveStamps[i] = update(stamp);
					return true;
				}
			}

			return false;
		}

		private static bool RemoveSliceRemoveStamp(ISegment segment, long localSeq)
		{
			bool removed = false;
			for (int i = segment.RemoveStamps.Count - 1; i >= 0; i--)
			{
				if (segment.RemoveStamps[i] is SliceRemoveOperationStamp stamp && stamp.LocalSeq == localSeq)
				{
					segment.RemoveStamps.RemoveAt(i);
					removed = true;
				}
			}

			return removed;
		}

		private static bool HasSetRemoveStamp(ISegment segment, Func<SetRemoveOperationStamp, bool> predicate)
		{
			foreach (SetRemoveOperationStamp stamp in FindSetRemoveStamps(segment))
			{
				if (predicate(stamp))
				{
					return true;
				}
			}

			return false;
		}

		private static bool HasSliceRemoveStamp(ISegment segment, Func<SliceRemoveOperationStamp, bool> predicate)
		{
			foreach (SliceRemoveOperationStamp stamp in FindSliceRemoveStamps(segment))
			{
				if (predicate(stamp))
				{
					return true;
				}
			}

			return false;
		}

		private void RecordSequence(long seq)
		{
			if (seq > CollabWindowCurrentSeq)
			{
				CollabWindowCurrentSeq = seq;
				MergeTree.CurrentSeq = seq;
			}
		}

		private bool SetMinimumSequenceNumber(long minSeq)
		{
			long effectiveMinSeq = GetMinimumSequenceNumberConstrainedByPendingOps(minSeq);
			if (effectiveMinSeq <= CollabWindowMinSeq)
			{
				return false;
			}

			CollabWindowMinSeq = effectiveMinSeq;
			return MergeTree.SetMinSeq(effectiveMinSeq);
		}

		private long GetMinimumSequenceNumberConstrainedByPendingOps(long minSeq)
		{
			long constrainedMinSeq = minSeq;
			foreach (PendingOpEntry pending in _pendingOps)
			{
				if (pending.IsInterval)
				{
					continue;
				}

				if (pending.RefSeq < constrainedMinSeq)
				{
					constrainedMinSeq = pending.RefSeq;
				}
			}

			return constrainedMinSeq;
		}

		private void AcknowledgeLocalOp(IMergeTreeOp op, long seq, long? clientSeq, bool clientSeqIsLocal)
		{
			List<int> pendingIndexes = FindPendingOpIndexes(op, clientSeq, clientSeqIsLocal);
			for (int i = 0; i < pendingIndexes.Count; i++)
			{
				AcknowledgePendingOp(_pendingOps[pendingIndexes[i]], seq);
			}

			for (int i = pendingIndexes.Count - 1; i >= 0; i--)
			{
				_pendingOps.RemoveAt(pendingIndexes[i]);
			}
		}

		private List<int> FindPendingOpIndexes(IMergeTreeOp op, long? clientSeq, bool clientSeqIsLocal)
		{
			List<int> indexes = new();
			if (clientSeq is long value)
			{
				for (int i = 0; i < _pendingOps.Count; i++)
				{
					if (clientSeqIsLocal ? _pendingOps[i].LocalSeq == value : _pendingOps[i].ClientSeq == value)
					{
						indexes.Add(i);
					}
				}

				return indexes;
			}

			int fallbackCount = op is MergeTreeGroupMsg groupOp ? groupOp.Ops.Count : 1;
			for (int i = 0; i < _pendingOps.Count && i < fallbackCount; i++)
			{
				indexes.Add(i);
			}

			return indexes;
		}

		private void AcknowledgePendingOp(PendingOpEntry pending, long seq)
		{
			if (pending.IsInterval)
			{
				// Interval ops already mutated the local interval collection when submitted.
			}
			else if (pending.IsAnnotate)
			{
				AcknowledgeAnnotateSegments(pending.LocalSeq);
			}
			else if (pending.IsInsert)
			{
				AcknowledgeInsertedSegments(pending.LocalSeq, pending.RefSeq, seq);
			}
			else if (pending.IsObliterate)
			{
				AcknowledgeObliteratedSegments(pending.LocalSeq, seq);
			}
			else if (pending.IsRemove)
			{
				AcknowledgeRemovedSegments(pending.LocalSeq, seq);
			}
		}

		private void AcknowledgeInsertedSegments(long localSeq, long refSeq, long seq)
		{
			List<ISegment> ackedSegments = new();
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				if (segment.LocalSeq == localSeq && segment.Seq == MergeTree.UnassignedSequenceNumber)
				{
					segment.Seq = seq;
					segment.LocalSeq = null;
					ackedSegments.Add(segment);
				}
			}

			if (ackedSegments.Count > 0)
			{
				int localClientId = MergeTree.GetOrAddShortClientId(ClientId);
				List<ISegment> newlyObliterated = MergeTree.ObliterateInsertedSegmentsIfCovered(
					ackedSegments,
					refSeq,
					seq,
					localClientId);
				RemoveMarkersFromIndex(newlyObliterated);
				MergeTree.UpdatePartialLengthsForSegments(
					ackedSegments,
					new OperationStamp()
					{
						Seq = seq,
						ClientId = localClientId,
					});
			}
		}

		private void AcknowledgeRemovedSegments(long localSeq, long seq)
		{
			List<ISegment> changedSegments = new();
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				for (int i = 0; i < segment.RemoveStamps.Count; i++)
				{
					if (segment.RemoveStamps[i] is not SetRemoveOperationStamp stamp || stamp.LocalSeq != localSeq)
					{
						continue;
					}

					segment.RemoveStamps[i] = stamp with
					{
						Seq = stamp.Seq == MergeTree.UnassignedSequenceNumber ? seq : stamp.Seq,
						LocalSeq = null,
					};
					Stamps.SortList(segment.RemoveStamps);
					changedSegments.Add(segment);
					break;
				}
			}

			if (changedSegments.Count > 0)
			{
				MergeTree.UpdatePartialLengthsForSegments(
					changedSegments,
					new OperationStamp()
					{
						Seq = seq,
						ClientId = MergeTree.GetOrAddShortClientId(ClientId),
					},
					newStructure: true);
			}
		}

		private void AcknowledgeObliteratedSegments(long localSeq, long seq)
		{
			List<ISegment> changedSegments = new();
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				for (int i = 0; i < segment.RemoveStamps.Count; i++)
				{
					if (segment.RemoveStamps[i] is not SliceRemoveOperationStamp stamp || stamp.LocalSeq != localSeq)
					{
						continue;
					}

					segment.RemoveStamps[i] = stamp with
					{
						Seq = stamp.Seq == MergeTree.UnassignedSequenceNumber ? seq : stamp.Seq,
						LocalSeq = null,
					};
					Stamps.SortList(segment.RemoveStamps);
					changedSegments.Add(segment);
					break;
				}
			}

			if (changedSegments.Count > 0)
			{
				MergeTree.UpdatePartialLengthsForSegments(
					changedSegments,
					new OperationStamp()
					{
						Seq = seq,
						ClientId = MergeTree.GetOrAddShortClientId(ClientId),
					},
					newStructure: true);
			}
		}

		private void AcknowledgeAnnotateSegments(long localSeq)
		{
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				segment.PendingAnnotates?.Remove(localSeq);
				if (segment.PendingAnnotates is { Count: 0 })
				{
					segment.PendingAnnotates = null;
				}
			}
		}

		private static void MarkPendingAnnotateSegments(IEnumerable<ISegment> segments, long localSeq, PropertySet props)
		{
			foreach (ISegment segment in segments)
			{
				segment.PendingAnnotates ??= new Dictionary<long, PropertySet>();
				segment.PendingAnnotates[localSeq] = CloneAnnotateProps(props);
			}
		}

		private IReadOnlyList<MergeTreeDelta> ApplyRemoteInsert(IMergeTreeInsertMsg op, long seq, long refSeq, string clientId, long? perspectiveSeq = null)
		{
			int position = ResolvePosition(op.Pos1, op.RelativePos1, nameof(op.Pos1));
			ISegment segment = SegmentFromSpec(op.Seg);
			List<ISegment> deltaSegments = MergeTree.InsertSegments(position, new ISegment[] { segment }, refSeq, seq, clientId, perspectiveSeq);
			if (segment is Marker marker && segment.RemoveStamps.Count == 0)
			{
				Markers.Add(marker);
			}

			return CreateRemoteDelta(MergeTreeDeltaType.Insert, op, deltaSegments, clientId);
		}

		private IReadOnlyList<MergeTreeDelta> ApplyRemoteRemove(IMergeTreeRemoveMsg op, long seq, long refSeq, string clientId, long? perspectiveSeq = null)
		{
			int start = ResolvePosition(op.Pos1, op.RelativePos1, nameof(op.Pos1));
			int end = ResolvePosition(op.Pos2, op.RelativePos2, nameof(op.Pos2));
			HashSet<ISegment> previouslyRemoved = CollectRemovedSegments();
			List<ISegment> deltaSegments = MergeTree.MarkRangeRemoved(start, end, refSeq, seq, clientId, perspectiveSeq);
			RemoveNewlyRemovedMarkers(previouslyRemoved);
			return CreateRemoteDelta(MergeTreeDeltaType.Remove, op, deltaSegments, clientId);
		}

		private IReadOnlyList<MergeTreeDelta> ApplyRemoteObliterate(IMergeTreeObliterateMsg op, long seq, long refSeq, string clientId, long? perspectiveSeq = null)
		{
			int start = RequirePosition(op.Pos1, nameof(op.Pos1));
			int end = RequirePosition(op.Pos2, nameof(op.Pos2));
			HashSet<ISegment> previouslyHidden = CollectHiddenSegments();
			List<ISegment> deltaSegments = MergeTree.ObliterateRange(start, end, refSeq, seq, clientId, perspectiveSeq);
			RemoveNewlyHiddenMarkers(previouslyHidden);
			return CreateRemoteDelta(MergeTreeDeltaType.Obliterate, op, deltaSegments, clientId);
		}

		private IReadOnlyList<MergeTreeDelta> ApplyRemoteObliterateSided(IMergeTreeObliterateSidedMsg op, long seq, long refSeq, string clientId, long? perspectiveSeq = null)
		{
			HashSet<ISegment> previouslyHidden = CollectHiddenSegments();
			List<ISegment> deltaSegments = MergeTree.ObliterateRangeSided(op.Pos1, op.Pos2, refSeq, seq, clientId, perspectiveSeq);
			RemoveNewlyHiddenMarkers(previouslyHidden);
			return CreateRemoteDelta(MergeTreeDeltaType.Obliterate, op, deltaSegments, clientId);
		}

		private IReadOnlyList<MergeTreeDelta> ApplyRemoteAnnotate(MergeTreeAnnotateMsg op, long seq, long refSeq, string clientId, long? perspectiveSeq = null)
		{
			int start = RequirePosition(op.Pos1, nameof(op.Pos1));
			int end = RequirePosition(op.Pos2, nameof(op.Pos2));
			PropertySet props = CloneAnnotateProps(op.Props);
			List<ISegment> deltaSegments = MergeTree.AnnotateRange(start, end, props, refSeq, seq, clientId, perspectiveSeq);
			IReadOnlyList<MergeTreeDelta> deltas = CreateRemoteDelta(MergeTreeDeltaType.Annotate, op, deltaSegments, clientId);
			ReapplyPendingAnnotates();
			return deltas;
		}

		private IReadOnlyList<MergeTreeDelta> ApplyRemoteGroup(MergeTreeGroupMsg op, long seq, long refSeq, string clientId)
		{
			List<MergeTreeDelta> deltas = new();
			long perspectiveSeq = seq == long.MaxValue ? seq : seq + 1;
			foreach (MergeTreeOp memberOp in op.Ops)
			{
				if (memberOp.Type == MergeTreeDeltaType.Group)
				{
					throw new NotSupportedException("Nested merge-tree group operations are not supported.");
				}

				deltas.AddRange(ApplyRemoteOp(memberOp, seq, refSeq, clientId, perspectiveSeq));
			}

			return deltas;
		}

		private IReadOnlyList<MergeTreeDelta> CreateRemoteDelta(
			MergeTreeDeltaType operation,
			IMergeTreeOp op,
			IReadOnlyList<ISegment> segments,
			string clientId)
		{
			List<MergeTreeDeltaRange> ranges = CreateDeltaRanges(segments);
			if (ranges.Count == 0)
			{
				return Array.Empty<MergeTreeDelta>();
			}

			return new List<MergeTreeDelta>()
			{
				new()
				{
					Operation = operation,
					Op = op,
					Segments = segments,
					Ranges = ranges,
					ClientId = clientId,
				},
			};
		}

		private List<MergeTreeDeltaRange> CreateDeltaRanges(IReadOnlyList<ISegment> segments)
		{
			List<MergeTreeDeltaRange> ranges = new();
			foreach (ISegment segment in segments)
			{
				int? position = MergeTree.GetPositionOfSegmentInCurrentView(segment);
				if (position is not int segmentPosition || segment.CachedLength == 0)
				{
					continue;
				}

				ranges.Add(new MergeTreeDeltaRange()
				{
					Segment = segment,
					Position = segmentPosition,
					Length = segment.CachedLength,
				});
			}

			return ranges;
		}

		private void ApplyRemoteIntervalAdd(IntervalAddOpMsg op)
		{
			GetOrCreateIntervalCollectionForOp(op).ApplyRemoteAdd(op);
		}

		private void ApplyRemoteIntervalDelete(IntervalDeleteOpMsg op)
		{
			GetOrCreateIntervalCollectionForOp(op).ApplyRemoteDelete(op);
		}

		private void ApplyRemoteIntervalChange(IntervalChangeOpMsg op)
		{
			GetOrCreateIntervalCollectionForOp(op).ApplyRemoteChange(op);
		}

		private void ApplyRemoteIntervalPropertyChanged(IntervalPropertyChangedOpMsg op)
		{
			GetOrCreateIntervalCollectionForOp(op).ApplyRemotePropertyChanged(op);
		}

		private void ApplyOwnIntervalAck(IMergeTreeOp op)
		{
			switch (op)
			{
				case IntervalAddOpMsg addOp:
					GetOrCreateIntervalCollectionForOp(addOp).ApplyOwnAck(addOp);
					break;

				case IntervalDeleteOpMsg deleteOp:
					GetOrCreateIntervalCollectionForOp(deleteOp).ApplyOwnAck(deleteOp);
					break;

				case IntervalChangeOpMsg changeOp:
					GetOrCreateIntervalCollectionForOp(changeOp).ApplyOwnAck(changeOp);
					break;

				case IntervalPropertyChangedOpMsg propertyChangedOp:
					GetOrCreateIntervalCollectionForOp(propertyChangedOp).ApplyOwnAck(propertyChangedOp);
					break;
			}
		}

		private IntervalCollection GetOrCreateIntervalCollectionForOp(IntervalOpMsg op)
		{
			return GetOrCreateIntervalCollection(
				op.CollectionName,
				op is IntervalAddOpMsg addOp ? addOp.IntervalType : IntervalType.SlideOnRemove);
		}

		private void EmitIntervalOp(IntervalOpMsg op)
		{
			long clientSeq = NextClientSeq();
			op.ClientSeq = clientSeq;
			SequenceInterval? interval = op is IntervalAddOpMsg or IntervalChangeOpMsg
				? GetOrCreateIntervalCollectionForOp(op).GetIntervalById(op.IntervalId)
				: null;
			_pendingOps.Add(new PendingOpEntry()
			{
				ClientSeq = clientSeq,
				LocalSeq = clientSeq,
				RefSeq = CollabWindowCurrentSeq,
				Op = CloneOp(op),
				IsInterval = true,
				IntervalOpKind = op.IntervalOpKind,
				IntervalCollectionName = op.CollectionName,
				IntervalId = op.IntervalId,
				IntervalStartReference = interval?.Start,
				IntervalEndReference = interval?.End,
			});
			_emittedOps.Enqueue(op);
		}

		private void ReapplyPendingAnnotates()
		{
			foreach (PendingOpEntry pending in _pendingOps)
			{
				if (!pending.IsAnnotate || pending.AnnotateProps is null)
				{
					continue;
				}

				List<ISegment> annotateSegments = CollectSegmentsInTreeOrder(
					segment => segment.PendingAnnotates is not null && segment.PendingAnnotates.ContainsKey(pending.LocalSeq));
				foreach ((int start, int end) in BuildRebasedRanges(annotateSegments, pending.LocalSeq - 1))
				{
					MergeTree.AnnotateRange(
						start,
						end,
						pending.AnnotateProps,
						MergeTree.UnassignedSequenceNumber,
						MergeTree.UnassignedSequenceNumber,
						ClientId);
				}
			}
		}

		private static PropertySet CloneAnnotateProps(IReadOnlyDictionary<string, object?>? props)
		{
			PropertySet clone = new();
			if (props is null)
			{
				return clone;
			}

			foreach (KeyValuePair<string, object?> property in props)
			{
				clone[property.Key] = property.Value;
			}

			return clone;
		}

		private HashSet<SetRemoveOperationStamp> CollectPendingRemovedStamps()
		{
			HashSet<SetRemoveOperationStamp> stamps = new(ReferenceEqualityComparer.Instance);
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				foreach (SetRemoveOperationStamp stamp in FindSetRemoveStamps(segment))
				{
					if (stamp.Seq == MergeTree.UnassignedSequenceNumber)
					{
						stamps.Add(stamp);
					}
				}
			}

			return stamps;
		}

		private HashSet<SliceRemoveOperationStamp> CollectPendingObliteratedStamps()
		{
			HashSet<SliceRemoveOperationStamp> stamps = new(ReferenceEqualityComparer.Instance);
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				foreach (SliceRemoveOperationStamp stamp in FindSliceRemoveStamps(segment))
				{
					if (stamp.Seq == MergeTree.UnassignedSequenceNumber)
					{
						stamps.Add(stamp);
					}
				}
			}

			return stamps;
		}

		private HashSet<ISegment> CollectRemovedSegments()
		{
			HashSet<ISegment> segments = new();
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				if (HasSetRemoveStamp(segment, _ => true))
				{
					segments.Add(segment);
				}
			}

			return segments;
		}

		private HashSet<ISegment> CollectHiddenSegments()
		{
			HashSet<ISegment> segments = new();
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				if (segment.RemoveStamps.Count > 0)
				{
					segments.Add(segment);
				}
			}

			return segments;
		}

		private void RemoveNewlyRemovedMarkers(HashSet<ISegment> previouslyRemoved)
		{
			List<ISegment> newlyRemoved = new();
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				if (HasSetRemoveStamp(segment, _ => true) && !previouslyRemoved.Contains(segment))
				{
					newlyRemoved.Add(segment);
				}
			}

			RemoveMarkersFromIndex(newlyRemoved);
		}

		private void RemoveNewlyHiddenMarkers(HashSet<ISegment> previouslyHidden)
		{
			List<ISegment> newlyHidden = new();
			foreach (ISegment segment in MergeTree.WalkAllSegments())
			{
				if (segment.RemoveStamps.Count > 0
					&& !previouslyHidden.Contains(segment))
				{
					newlyHidden.Add(segment);
				}
			}

			RemoveMarkersFromIndex(newlyHidden);
		}

		private void RemoveMarkersFromIndex(IEnumerable<ISegment> segments)
		{
			foreach (ISegment segment in segments)
			{
				if (segment is Marker marker)
				{
					Markers.Remove(marker.GetId());
				}
			}
		}

		private IEnumerable<(ISegment segment, int position)> WalkVisibleSegments()
		{
			int position = 0;
			foreach ((ISegment segment, int startOffset, int endOffset) in MergeTree.GetSegments(0, GetLength()))
			{
				yield return (segment, position);
				position += endOffset - startOffset;
			}
		}

		private static bool IsMatchingTileMarker(ISegment segment, string? tileLabel, out Marker? marker)
		{
			if (segment is Marker candidate && candidate.RefTypeIncludesFlag(ReferenceType.Tile))
			{
				if (tileLabel is null || candidate.HasTileLabel(tileLabel))
				{
					marker = candidate;
					return true;
				}
			}

			marker = null;
			return false;
		}

		private static int RequirePosition(int? position, string name)
		{
			if (position is not int value)
			{
				throw new ArgumentException("Operation position is required.", name);
			}

			return value;
		}

		private int ResolvePosition(int? position, object? relativePosition, string name)
		{
			if (position is int value)
			{
				return value;
			}

			if (relativePosition is not null)
			{
				return PositionFromRelativePosition(relativePosition, name);
			}

			return RequirePosition(position, name);
		}

		private int PositionFromRelativePosition(object relativePosition, string name)
		{
			if (!TryReadRelativePositionId(relativePosition, out string? id) || string.IsNullOrEmpty(id))
			{
				throw new ArgumentException("Relative position id is required.", name);
			}

			Marker? marker = GetMarkerFromId(id);
			if (marker is null)
			{
				throw new ArgumentException("Relative position marker could not be found.", name);
			}

			int? position = GetPositionOfMarker(marker);
			if (position is not int markerPosition)
			{
				throw new ArgumentException("Relative position marker is not in the current view.", name);
			}

			bool before = TryReadRelativePositionBefore(relativePosition, out bool beforeValue) && beforeValue;
			if (before)
			{
				return markerPosition - (TryReadRelativePositionOffset(relativePosition, out int beforeOffset) ? beforeOffset : 0);
			}

			return markerPosition
				+ marker.CachedLength
				+ (TryReadRelativePositionOffset(relativePosition, out int afterOffset) ? afterOffset : 0);
		}

		private static bool TryReadRelativePositionId(object relativePosition, out string? id)
		{
			if (TryReadRelativePositionMember(relativePosition, "id", out object? value)
				&& value is string stringValue)
			{
				id = stringValue;
				return true;
			}

			id = null;
			return false;
		}

		private static bool TryReadRelativePositionBefore(object relativePosition, out bool before)
		{
			if (TryReadRelativePositionMember(relativePosition, "before", out object? value))
			{
				switch (value)
				{
					case bool boolValue:
						before = boolValue;
						return true;

					case JsonElement jsonElement when jsonElement.ValueKind == JsonValueKind.True:
						before = true;
						return true;

					case JsonElement jsonElement when jsonElement.ValueKind == JsonValueKind.False:
						before = false;
						return true;
				}
			}

			before = false;
			return false;
		}

		private static bool TryReadRelativePositionOffset(object relativePosition, out int offset)
		{
			if (TryReadRelativePositionMember(relativePosition, "offset", out object? value))
			{
				switch (value)
				{
					case int intValue:
						offset = intValue;
						return true;

					case long longValue when longValue >= int.MinValue && longValue <= int.MaxValue:
						offset = (int)longValue;
						return true;

					case JsonElement jsonElement when jsonElement.ValueKind == JsonValueKind.Number
						&& jsonElement.TryGetInt32(out int jsonValue):
						offset = jsonValue;
						return true;
				}
			}

			offset = 0;
			return false;
		}

		private static bool TryReadRelativePositionMember(object relativePosition, string memberName, out object? value)
		{
			if (relativePosition is IReadOnlyDictionary<string, object?> dictionary
				&& dictionary.TryGetValue(memberName, out value))
			{
				return true;
			}

			if (relativePosition is JsonElement jsonElement
				&& jsonElement.ValueKind == JsonValueKind.Object
				&& jsonElement.TryGetProperty(memberName, out JsonElement property))
			{
				value = property;
				return true;
			}

			System.Reflection.PropertyInfo? propertyInfo = relativePosition.GetType().GetProperty(memberName)
				?? relativePosition.GetType().GetProperty(
					char.ToUpperInvariant(memberName[0]) + memberName.Substring(1));
			if (propertyInfo is not null)
			{
				value = propertyInfo.GetValue(relativePosition);
				return true;
			}

			value = null;
			return false;
		}

		private static int NormalizeSequencePlacePosition(SequencePlace place, int length)
		{
			if (place.IsStart)
			{
				return 0;
			}

			if (place.IsEnd)
			{
				return length;
			}

			return place.Position;
		}

		private static ISegment SegmentFromSpec(object? spec)
		{
			if (spec is ISegment segment)
			{
				return segment.Clone();
			}

			Marker? marker = Marker.FromJSONObject(spec);
			if (marker is not null)
			{
				return marker;
			}

			TextSegment? textSegment = TextSegment.FromJSONObject(spec);
			if (textSegment is not null)
			{
				return textSegment;
			}

			if (spec is JsonElement jsonElement)
			{
				Marker? markerFromJson = Marker.FromJSONObject(jsonElement);
				if (markerFromJson is not null)
				{
					return markerFromJson;
				}

				return TextSegmentFromJsonElement(jsonElement);
			}

			throw new NotSupportedException("Only TextSegment and Marker insert payloads are supported by the POC Client.");
		}

		private static TextSegment TextSegmentFromJsonElement(JsonElement jsonElement)
		{
			if (jsonElement.ValueKind == JsonValueKind.String)
			{
				return new TextSegment(jsonElement.GetString() ?? string.Empty);
			}

			if (jsonElement.ValueKind == JsonValueKind.Object
				&& jsonElement.TryGetProperty("text", out JsonElement textProperty)
				&& textProperty.ValueKind == JsonValueKind.String)
			{
				PropertySet? properties = null;
				if (jsonElement.TryGetProperty("props", out JsonElement propsProperty)
					&& propsProperty.ValueKind != JsonValueKind.Null
					&& propsProperty.ValueKind != JsonValueKind.Undefined)
				{
					properties = JsonElementToPropertySet(propsProperty);
				}

				return new TextSegment(textProperty.GetString() ?? string.Empty, properties);
			}

			throw new NotSupportedException("Only JSON text segment insert payloads are supported by the POC Client.");
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

		private sealed class ClientIntervalOpSender : IIntervalOpSender
		{
			private readonly Client _client;

			public ClientIntervalOpSender(Client client)
			{
				_client = client;
			}

			public void SendIntervalAdd(
				string collectionName,
				string intervalId,
				int start,
				int end,
				IntervalType intervalType,
				IntervalStickiness stickiness,
				Microsoft.Office.Web.Fluid.Intervals.Side startSide,
				Microsoft.Office.Web.Fluid.Intervals.Side endSide,
				PropertySet? props)
			{
				_client.EmitIntervalOp(new IntervalAddOpMsg()
				{
					CollectionName = collectionName,
					IntervalId = intervalId,
					Start = start,
					End = end,
					IntervalType = intervalType,
					Stickiness = stickiness,
					StartSide = startSide,
					EndSide = endSide,
					Props = PropertyMap.ClonePropertySet(props),
				});
			}

			public void SendIntervalDelete(string collectionName, string intervalId)
			{
				_client.EmitIntervalOp(new IntervalDeleteOpMsg()
				{
					CollectionName = collectionName,
					IntervalId = intervalId,
				});
			}

			public void SendIntervalChange(
				string collectionName,
				string intervalId,
				int? start,
				int? end,
				IntervalStickiness stickiness,
				Microsoft.Office.Web.Fluid.Intervals.Side startSide,
				Microsoft.Office.Web.Fluid.Intervals.Side endSide)
			{
				_client.EmitIntervalOp(new IntervalChangeOpMsg()
				{
					CollectionName = collectionName,
					IntervalId = intervalId,
					Start = start,
					End = end,
					Stickiness = stickiness,
					StartSide = startSide,
					EndSide = endSide,
				});
			}

			public void SendIntervalPropertyChanged(string collectionName, string intervalId, PropertySet props)
			{
				_client.EmitIntervalOp(new IntervalPropertyChangedOpMsg()
				{
					CollectionName = collectionName,
					IntervalId = intervalId,
					Props = PropertyMap.ClonePropertySet(props) ?? new PropertySet(),
				});
			}
		}
	}
}
