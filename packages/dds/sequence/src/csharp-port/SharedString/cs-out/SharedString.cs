// -----------------------------------------------------------------------------
// Ported from packages/dds/sequence/src/sharedString.ts + sharedSequence.ts +
// sequence.ts (subset for POC).
// Part of the SharedString C# feasibility port — Wave 6.
//
// POC APIs: GetLength / GetText / marker-aware text extraction / InsertText / RemoveText
// / DeleteText / ReplaceText / relative inserts / RunInBatch / ObliterateRange
// / SequenceDelta event / LoadFromSnapshot. Full Fluid runtime lifecycle parity remains deferred.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;
using System.Text.Json;

using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;

namespace Microsoft.Office.Web.Fluid
{
	public class SequenceDeltaEventArgs
	{
		public string OpType { get; set; } = string.Empty;

		public MergeTreeDeltaType? DeltaOperation { get; set; }

		public IMergeTreeOp? Op { get; set; }

		public int Position { get; set; }

		public int Length { get; set; }

		public string? Text { get; set; }

		public bool IsMarker { get; set; }

		public Marker? Marker { get; set; }

		public bool Local { get; set; }

		public bool IsLocal
		{
			get => Local;
			set => Local = value;
		}

		public PropertySet? AnnotatedProperties { get; set; }

		public IReadOnlyList<SequenceDeltaRange> Ranges { get; set; } = Array.Empty<SequenceDeltaRange>();

		public SequenceDeltaRange? First => Ranges.Count == 0 ? null : Ranges[0];

		public SequenceDeltaRange? Last => Ranges.Count == 0 ? null : Ranges[Ranges.Count - 1];

		public string? ClientId { get; set; }
	}

	public sealed class SequenceDeltaRange
	{
		public string OpType { get; set; } = string.Empty;

		public MergeTreeDeltaType Operation { get; set; }

		public int Position { get; set; }

		public int Length { get; set; }

		public ISegment Segment { get; set; } = null!;

		public PropertySet? PropertyDeltas { get; set; }
	}

	public delegate void SequenceDeltaEventHandler(object sender, SequenceDeltaEventArgs e);

	public sealed class RelativePosition
	{
		public string Id { get; set; } = string.Empty;

		public bool Before { get; set; }

		public int Offset { get; set; }
	}

	public sealed class SharedString : IFluidDataObject, IFluidDataObjectMessageHandler
	{
		private const string _insertOpType = "insert";
		private const string _removeOpType = "remove";
		private const string _obliterateOpType = "obliterate";
		private const string _annotateOpType = "annotate";
		private const string _intervalOpType = "interval";
		private const string _groupOpType = "group";

		private readonly string _id;
		private readonly object _lock = new object();
		private readonly IFluidDataObjectSender? _sender;
		private readonly IFluidDataObjectRegistry? _registry;
		private readonly Client _client;
		private readonly HashSet<IntervalCollection> _intervalCollectionsWithOutboundHandlers = new();
		private List<MergeTreeOp>? _batchOps;
		private int _localMutationDepth;

		public SharedString(string? id = null, IFluidDataObjectSender? sender = null, IFluidDataObjectRegistry? registry = null)
		{
			_id = id ?? FluidObjectId.CreateId();
			_sender = sender;
			_registry = registry;
			_client = new Client(clientId: _id);
		}

		public string Id => _id;

		public event SequenceDeltaEventHandler? OnSequenceDelta;

		public void RunInBatch(Action action)
		{
			ArgumentNullException.ThrowIfNull(action);

			lock (_lock)
			{
				if (_batchOps is not null)
				{
					throw new InvalidOperationException("Nested SharedString batches are not supported.");
				}

				_batchOps = new List<MergeTreeOp>();
			}

			try
			{
				action();
			}
			finally
			{
				List<MergeTreeOp> opsToFlush;
				lock (_lock)
				{
					opsToFlush = _batchOps ?? new List<MergeTreeOp>();
					_batchOps = null;
				}

				FlushBatchOps(opsToFlush);
			}
		}

		/// <summary>
		/// Returns (or creates) a named interval collection over this SharedString.
		/// Multiple named collections are supported (e.g., "comments", "highlights").
		/// </summary>
		public IntervalCollection GetIntervalCollection(string name, IntervalType endpointType = IntervalType.SlideOnRemove)
		{
			lock (_lock)
			{
				IntervalCollection collection = _client.GetOrCreateIntervalCollection(name, endpointType);
				EnsureIntervalCollectionOutboundHandlers(collection);
				return collection;
			}
		}

		public int GetLength()
		{
			lock (_lock)
			{
				return _client.GetLength();
			}
		}

		public int GetLength(long refSeq, string? clientId = null, long? localSeq = null)
		{
			lock (_lock)
			{
				return _client.MergeTree.GetLengthAt(refSeq, clientId, localSeq);
			}
		}

		public string GetText()
		{
			lock (_lock)
			{
				return _client.GetText(0, _client.GetLength());
			}
		}

		public string GetText(int start, int end)
		{
			lock (_lock)
			{
				return _client.GetText(start, end);
			}
		}

		public string GetTextWithPlaceholders(int? start = null, int? end = null)
		{
			lock (_lock)
			{
				return GetTextWithMarkerPlaceholder(start ?? 0, end ?? _client.GetLength(), " ", useMarkerString: false);
			}
		}

		public string GetTextRangeWithMarkers(int start, int end)
		{
			lock (_lock)
			{
				return GetTextWithMarkerPlaceholder(start, end, "*", useMarkerString: true);
			}
		}

		/// <summary>
		/// Returns the properties applied to the segment containing the given position.
		/// Returns null if the segment has no properties.
		/// </summary>
		/// <param name="position">The current local-view position.</param>
		/// <returns>The segment properties, or <see langword="null" /> if none are present or the position is outside content.</returns>
		public PropertySet? GetPropertiesAtPosition(int position)
		{
			lock (_lock)
			{
				// TS ref: packages/dds/merge-tree/src/client.ts getPropertiesAtPosition —
				// TS returns undefined for positions outside content (position < 0 or
				// position >= length). The port previously threw.
				if (position < 0 || position >= _client.GetLength())
				{
					return null;
				}

				var (segment, _) = _client.GetContainingSegment(position);
				return segment.Properties;
			}
		}

		// -----------------------------------------------------------------------------
		// Navigation APIs (Wave 13):
		// GetContainingSegment / GetPosition / GetRangeExtentsOfPosition /
		// LocalReferencePositionToPosition. All wrap existing MergeTree/Client
		// machinery — thin public surface for consumers.
		// -----------------------------------------------------------------------------

		/// <summary>
		/// Returns the segment containing the given position, plus the offset within that segment.
		/// Returns null if position is out of range.
		/// </summary>
		public (Microsoft.Office.Web.Fluid.MergeTree.ISegment segment, int offsetInSegment)? GetContainingSegment(int position)
		{
			lock (_lock)
			{
				return _client.MergeTree.TryGetContainingSegment(
					position,
					Microsoft.Office.Web.Fluid.MergeTree.MergeTree.UnassignedSequenceNumber,
					_client.ClientId);
			}
		}

		/// <summary>
		/// Returns the current absolute position of a segment in the string.
		/// Returns null if the segment is not currently in the tree (e.g., removed).
		/// </summary>
		public int? GetPosition(Microsoft.Office.Web.Fluid.MergeTree.ISegment segment)
		{
			lock (_lock)
			{
				return _client.MergeTree.GetPositionOfSegment(segment);
			}
		}

		public int? GetPosition(
			Microsoft.Office.Web.Fluid.MergeTree.ISegment segment,
			long refSeq,
			string? clientId = null,
			long? localSeq = null)
		{
			lock (_lock)
			{
				return _client.MergeTree.GetPositionOfSegmentAt(segment, refSeq, clientId, localSeq);
			}
		}

		/// <summary>
		/// Returns the [start, end) range of the segment containing the given position.
		/// Both bounds are in absolute positions.
		/// Returns null if position is out of range.
		/// </summary>
		public (int start, int end)? GetRangeExtentsOfPosition(int position)
		{
			lock (_lock)
			{
				var segInfo = _client.MergeTree.TryGetContainingSegment(
					position,
					Microsoft.Office.Web.Fluid.MergeTree.MergeTree.UnassignedSequenceNumber,
					_client.ClientId);
				if (segInfo is null)
				{
					return null;
				}

				int segStart = position - segInfo.Value.offsetInSegment;
				int segEnd = segStart + segInfo.Value.segment.CachedLength;
				return (segStart, segEnd);
			}
		}

		/// <summary>
		/// Returns the current position of the given LocalReferencePosition.
		/// Returns null if the reference is detached (its segment was removed).
		/// </summary>
		public int? LocalReferencePositionToPosition(Microsoft.Office.Web.Fluid.MergeTree.LocalReferencePosition reference)
		{
			lock (_lock)
			{
				return _client.MergeTree.GetPositionOfReference(reference);
			}
		}

		public void InsertText(int position, string text, PropertySet? props = null)
		{
			if (text is null)
			{
				throw new ArgumentNullException(nameof(text));
			}

			// TS ref: packages/dds/merge-tree/src/client.ts insertSegmentLocal —
			// empty segments early-return undefined, so TS emits neither a wire op
			// nor a sequence-delta event. Match the TS semantics here.
			if (text.Length == 0)
			{
				return;
			}

			using IDisposable mutation = EnterLocalMutation();
			IMergeTreeInsertMsg insertMsg;
			IReadOnlyList<SequenceDeltaRange> ranges;
			lock (_lock)
			{
				insertMsg = _client.InsertText(position, text, props);
				ranges = CreateEventRangesFromCurrentView(position, position + text.Length, MergeTreeDeltaType.Insert);
			}

			EmitOrBatchLocalOp(insertMsg, _insertOpType);

			RaiseSequenceDelta(new SequenceDeltaEventArgs()
			{
				OpType = _insertOpType,
				DeltaOperation = MergeTreeDeltaType.Insert,
				Op = insertMsg,
				Position = position,
				Length = text.Length,
				Text = text,
				Local = true,
				ClientId = _client.ClientId,
				Ranges = ranges,
			});
		}

		public void InsertMarker(int position, ReferenceType refType, PropertySet? props = null)
		{
			using IDisposable mutation = EnterLocalMutation();
			IMergeTreeInsertMsg insertMsg;
			Marker? marker;
			IReadOnlyList<SequenceDeltaRange> ranges;
			lock (_lock)
			{
				insertMsg = _client.InsertMarker(position, refType, props);
				marker = _client.GetContainingSegment(position).segment as Marker;
				ranges = CreateEventRangesFromCurrentView(position, position + (marker?.CachedLength ?? 1), MergeTreeDeltaType.Insert);
			}

			EmitOrBatchLocalOp(insertMsg, _insertOpType);

			RaiseSequenceDelta(new SequenceDeltaEventArgs()
			{
				OpType = _insertOpType,
				DeltaOperation = MergeTreeDeltaType.Insert,
				Op = insertMsg,
				Position = position,
				Length = marker?.CachedLength ?? 1,
				Text = null,
				IsMarker = true,
				Marker = marker,
				Local = true,
				ClientId = _client.ClientId,
				Ranges = ranges,
			});
		}

		public void InsertTextRelative(object relativePos1, string text, PropertySet? props = null)
		{
			InsertText(PositionFromRelativePosition(relativePos1, nameof(relativePos1)), text, props);
		}

		public void InsertMarkerRelative(object relativePos1, ReferenceType refType, PropertySet? props = null)
		{
			InsertMarker(PositionFromRelativePosition(relativePos1, nameof(relativePos1)), refType, props);
		}

		public void RemoveText(int start, int end)
		{
			DeleteText(start, end);
		}

		public void ReplaceText(int start, int end, string text, PropertySet? props = null)
		{
			if (text is null)
			{
				throw new ArgumentNullException(nameof(text));
			}

			InsertText(Math.Max(start, end), text, props);

			// TS ref: packages/dds/sequence/src/sequence.ts replaceRange — the
			// remove is guarded by the truthiness of the insert op. When TS's
			// insertSegmentLocal returns undefined (empty segment), TS skips the
			// remove. Empty replaceText is therefore a no-op, not a delete.
			if (text.Length > 0 && start < end)
			{
				RemoveText(start, end);
			}
		}

		public void DeleteText(int start, int end)
		{
			using IDisposable mutation = EnterLocalMutation();
			IMergeTreeRemoveMsg removeMsg;
			IReadOnlyList<SequenceDeltaRange> ranges;
			lock (_lock)
			{
				ranges = CreateEventRangesFromCurrentView(start, end, MergeTreeDeltaType.Remove);
				removeMsg = _client.RemoveText(start, end);
			}

			EmitOrBatchLocalOp(removeMsg, _removeOpType);

			RaiseSequenceDelta(new SequenceDeltaEventArgs()
			{
				OpType = _removeOpType,
				DeltaOperation = MergeTreeDeltaType.Remove,
				Op = removeMsg,
				Position = start,
				Length = end - start,
				Text = null,
				Local = true,
				ClientId = _client.ClientId,
				Ranges = ranges,
			});
		}

		public void ObliterateRange(int start, int end)
		{
			using IDisposable mutation = EnterLocalMutation();
			IMergeTreeObliterateMsg? obliterateMsg;
			IReadOnlyList<SequenceDeltaRange> ranges;
			lock (_lock)
			{
				ranges = CreateEventRangesFromCurrentView(start, end, MergeTreeDeltaType.Obliterate);
				obliterateMsg = _client.ObliterateRangeLocal(start, end);
			}

			if (obliterateMsg is null)
			{
				return;
			}

			EmitOrBatchLocalOp(obliterateMsg, _obliterateOpType);

			RaiseSequenceDelta(new SequenceDeltaEventArgs()
			{
				OpType = _obliterateOpType,
				DeltaOperation = MergeTreeDeltaType.Obliterate,
				Op = obliterateMsg,
				Position = start,
				Length = end - start,
				Text = null,
				Local = true,
				ClientId = _client.ClientId,
				Ranges = ranges,
			});
		}

		public void ObliterateRange(SequencePlace start, SequencePlace end)
		{
			if (start is null)
			{
				throw new ArgumentNullException(nameof(start));
			}

			if (end is null)
			{
				throw new ArgumentNullException(nameof(end));
			}

			using IDisposable mutation = EnterLocalMutation();
			IMergeTreeObliterateSidedMsg? obliterateMsg;
			int eventStart;
			int eventEnd;
			IReadOnlyList<SequenceDeltaRange> ranges;
			lock (_lock)
			{
				eventStart = start.IsStart ? 0 : start.Position;
				eventEnd = end.IsEnd ? _client.GetLength() : end.Position;
				ranges = CreateEventRangesFromCurrentView(
					eventStart,
					Math.Max(eventStart, eventEnd),
					MergeTreeDeltaType.Obliterate);
				obliterateMsg = _client.ObliterateRangeLocal(start, end);
			}

			if (obliterateMsg is null)
			{
				return;
			}

			EmitOrBatchLocalOp(obliterateMsg, _obliterateOpType);

			RaiseSequenceDelta(new SequenceDeltaEventArgs()
			{
				OpType = _obliterateOpType,
				DeltaOperation = MergeTreeDeltaType.ObliterateSided,
				Op = obliterateMsg,
				Position = eventStart,
				Length = Math.Max(0, eventEnd - eventStart),
				Text = null,
				Local = true,
				ClientId = _client.ClientId,
				Ranges = ranges,
			});
		}

		public Marker? GetMarkerFromId(string id)
		{
			lock (_lock)
			{
				return _client.GetMarkerFromId(id);
			}
		}

		public int? GetPositionOfMarker(Marker marker)
		{
			lock (_lock)
			{
				return _client.GetPositionOfMarker(marker);
			}
		}

		public Marker? SearchForMarker(int startPos, bool forwards = true, string? tileLabel = null)
		{
			if (tileLabel is null)
			{
				return null;
			}

			lock (_lock)
			{
				return _client.SearchForTileMarker(startPos, forwards, tileLabel);
			}
		}

		public Marker? SearchForMarker(int startPos, string markerLabel, bool forwards = true)
		{
			if (markerLabel is null)
			{
				throw new ArgumentNullException(nameof(markerLabel));
			}

			lock (_lock)
			{
				return _client.SearchForTileMarker(startPos, forwards, markerLabel);
			}
		}

		public void AnnotateRange(int start, int end, PropertySet props)
		{
			if (props is null)
			{
				throw new ArgumentNullException(nameof(props));
			}

			using IDisposable mutation = EnterLocalMutation();
			MergeTreeAnnotateMsg annotateMsg;
			IReadOnlyList<SequenceDeltaRange> ranges;
			lock (_lock)
			{
				ranges = CreateAnnotateEventRanges(start, end, props);
				annotateMsg = _client.Annotate(start, end, props);
			}

			EmitOrBatchLocalOp(annotateMsg, _annotateOpType);

			RaiseSequenceDelta(new SequenceDeltaEventArgs()
			{
				OpType = _annotateOpType,
				DeltaOperation = MergeTreeDeltaType.Annotate,
				Op = annotateMsg,
				Position = start,
				Length = end - start,
				Text = null,
				Local = true,
				ClientId = _client.ClientId,
				AnnotatedProperties = CloneAnnotateProps(props),
				Ranges = ranges,
			});
		}

		public void LoadFromSnapshot(SharedStringSnapshotDto snapshot)
		{
			if (snapshot is null)
			{
				throw new ArgumentNullException(nameof(snapshot));
			}

			lock (_lock)
			{
				SharedStringSnapshotLoader.PopulateFromSnapshot(_client, snapshot, _registry);
			}
		}

		public void LoadFromSnapshot(string snapshotJson, Func<string, string>? blobResolver = null)
		{
			LoadFromSnapshot(SharedStringSnapshotLoader.Load(snapshotJson, blobResolver, _registry));
		}

		// Not part of tranpiled output. Added later for testing.
		public string Z_GenerateSnapshot()
		{
			return "{todo:true}";
		}

		public void RegeneratePendingOps()
		{
			if (_sender is null)
			{
				return;
			}

			IReadOnlyList<IMergeTreeOp> rebasedOps;
			lock (_lock)
			{
				rebasedOps = _client.RebasePendingOps();
			}

			foreach (IMergeTreeOp op in rebasedOps)
			{
				SendRegeneratedPendingOp(op);
			}
		}

		public IMergeTreeOp ParseOp(string opJson)
		{
			if (opJson is null)
			{
				throw new ArgumentNullException(nameof(opJson));
			}

			IMergeTreeOp op = SharedStringOpSerializer.Deserialize(opJson, _registry);
			return op;
		}

		public void ProcessDataObjectOp(SequencedDocumentMessageDescriptor descriptor, string opJson)
		{
			IMergeTreeOp op = ParseOp(opJson);
			long? minimumSequenceNumber = TryGetMinimumSequenceNumber(descriptor);
			if (descriptor.Origin == OpOrigin.Local)
			{
				long? clientSeq = descriptor.Seq.HasClientSequenceNumber
					? descriptor.Seq.clientSequenceNumber
					: null;
				lock (_lock)
				{
					_client.ApplyOp(
						op,
						seq: descriptor.Seq.sequenceNumber,
						refSeq: descriptor.Seq.referenceSequenceNumber,
						clientId: _client.ClientId,
						clientSeq: clientSeq,
						minimumSequenceNumber: minimumSequenceNumber);
				}

				return;
			}

			IReadOnlyList<MergeTreeDelta> deltas;
			lock (_lock)
			{
				deltas = _client.ApplyOp(
					op,
					seq: descriptor.Seq.sequenceNumber,
					refSeq: descriptor.Seq.referenceSequenceNumber,
					clientId: descriptor.ClientId ?? "unknown",
					minimumSequenceNumber: minimumSequenceNumber);
			}

			List<SequenceDeltaEventArgs> args = CreateRemoteSequenceDeltaEventArgs(deltas);
			foreach (SequenceDeltaEventArgs arg in args)
			{
				RaiseSequenceDelta(arg);
			}
		}

		public void ProcessDataObjectAttach(SequencedDocumentMessageDescriptor descriptor, fluidDataStoreMessageAttach op)
		{
			// Server-side client does not act on attach; state comes from snapshot.
		}

		private List<SequenceDeltaEventArgs> CreateRemoteSequenceDeltaEventArgs(IReadOnlyList<MergeTreeDelta> deltas)
		{
			List<SequenceDeltaEventArgs> args = new();
			foreach (MergeTreeDelta delta in deltas)
			{
				List<SequenceDeltaRange> ranges = CreateSequenceDeltaRanges(delta);
				if (ranges.Count == 0)
				{
					continue;
				}

				string opType = GetSequenceDeltaOpType(delta.Operation);
				SequenceDeltaRange firstRange = ranges[0];
				SequenceDeltaEventArgs eventArgs = new()
				{
					OpType = opType,
					DeltaOperation = delta.Operation,
					Op = delta.Op,
					Position = firstRange.Position,
					Length = GetTotalRangeLength(ranges),
					Text = null,
					Local = false,
					Ranges = ranges,
					ClientId = delta.ClientId,
				};

				if (delta.Operation == MergeTreeDeltaType.Insert)
				{
					ISegment firstSegment = firstRange.Segment;
					Marker? marker = firstSegment as Marker;
					string text = marker is null ? ExtractInsertedText(firstSegment) : string.Empty;
					eventArgs.Length = marker?.CachedLength ?? text.Length;
					eventArgs.Text = marker is null ? text : null;
					eventArgs.IsMarker = marker is not null;
					eventArgs.Marker = marker;
				}
				else if (delta.Operation == MergeTreeDeltaType.Annotate && delta.Op is MergeTreeAnnotateMsg annotateMsg)
				{
					eventArgs.AnnotatedProperties = CloneAnnotateProps(annotateMsg.Props);
				}

				args.Add(eventArgs);
			}

			return args;
		}

		private List<SequenceDeltaRange> CreateSequenceDeltaRanges(MergeTreeDelta delta)
		{
			List<SequenceDeltaRange> ranges = new();
			string opType = GetSequenceDeltaOpType(delta.Operation);
			foreach (MergeTreeDeltaRange deltaRange in delta.Ranges)
			{
				if (deltaRange.Length == 0)
				{
					continue;
				}

				PropertySet? propertyDeltas = delta.Op is MergeTreeAnnotateMsg annotateMsg
					? CloneAnnotateProps(annotateMsg.Props)
					: null;
				if (ShouldMergeRange(delta.Operation, ranges, deltaRange.Position, propertyDeltas))
				{
					ranges[ranges.Count - 1].Length += deltaRange.Length;
					continue;
				}

				ranges.Add(new SequenceDeltaRange()
				{
					OpType = opType,
					Operation = delta.Operation,
					Position = deltaRange.Position,
					Length = deltaRange.Length,
					Segment = deltaRange.Segment,
					PropertyDeltas = propertyDeltas,
				});
			}

			return ranges;
		}

		private List<SequenceDeltaRange> CreateEventRangesFromCurrentView(
			int start,
			int end,
			MergeTreeDeltaType operation,
			PropertySet? propertyDeltas = null)
		{
			List<SequenceDeltaRange> ranges = new();
			if (end <= start)
			{
				return ranges;
			}

			string opType = GetSequenceDeltaOpType(operation);
			foreach ((ISegment segment, int startOffset, int endOffset) in _client.MergeTree.GetSegments(start, end))
			{
				int length = endOffset - startOffset;
				if (length == 0)
				{
					continue;
				}

				int rangePosition = start + GetTotalRangeLength(ranges);
				ranges.Add(new SequenceDeltaRange()
				{
					OpType = opType,
					Operation = operation,
					Position = rangePosition,
					Length = length,
					Segment = CloneSegmentSlice(segment, startOffset, endOffset),
					PropertyDeltas = PropertyMap.ClonePropertySet(propertyDeltas),
				});
			}

			return ranges;
		}

		private List<SequenceDeltaRange> CreateAnnotateEventRanges(int start, int end, PropertySet props)
		{
			List<SequenceDeltaRange> ranges = CreateEventRangesFromCurrentView(start, end, MergeTreeDeltaType.Annotate);
			foreach (SequenceDeltaRange range in ranges)
			{
				range.PropertyDeltas = CreatePropertyDeltas(range.Segment.Properties, props);
				range.Segment.Properties = PropertyMap.AddProperties(range.Segment.Properties, props);
			}

			return ranges;
		}

		private string GetTextWithMarkerPlaceholder(int start, int end, string markerPlaceholder, bool useMarkerString)
		{
			System.Text.StringBuilder builder = new();
			foreach ((ISegment segment, int startOffset, int endOffset) in _client.MergeTree.GetSegments(start, end))
			{
				if (segment is TextSegment textSegment)
				{
					builder.Append(textSegment.Text, startOffset, endOffset - startOffset);
					continue;
				}

				if (useMarkerString)
				{
					builder.Append('\n');
					builder.Append(segment);
				}
				else
				{
					for (int i = 0; i < endOffset - startOffset; i++)
					{
						builder.Append(markerPlaceholder);
					}
				}
			}

			return builder.ToString();
		}

		private static ISegment CloneSegmentSlice(ISegment segment, int startOffset, int endOffset)
		{
			if (segment is TextSegment textSegment)
			{
				return new TextSegment(
					textSegment.Text.Substring(startOffset, endOffset - startOffset),
					PropertyMap.ClonePropertySet(textSegment.Properties));
			}

			if (segment is Marker marker)
			{
				return marker.Clone();
			}

			ISegment clone = segment.Clone();
			clone.CachedLength = endOffset - startOffset;
			return clone;
		}

		private static PropertySet CreatePropertyDeltas(
			IReadOnlyDictionary<string, object?>? currentProperties,
			IReadOnlyDictionary<string, object?> props)
		{
			PropertySet propertyDeltas = new();
			foreach (KeyValuePair<string, object?> property in props)
			{
				propertyDeltas[property.Key] = currentProperties is not null
					&& currentProperties.TryGetValue(property.Key, out object? currentValue)
						? currentValue
						: null;
			}

			return propertyDeltas;
		}

		private static bool ShouldMergeRange(
			MergeTreeDeltaType operation,
			IReadOnlyList<SequenceDeltaRange> ranges,
			int position,
			PropertySet? propertyDeltas)
		{
			if (ranges.Count == 0)
			{
				return false;
			}

			SequenceDeltaRange previous = ranges[ranges.Count - 1];
			return operation switch
			{
				MergeTreeDeltaType.Remove or MergeTreeDeltaType.Obliterate => previous.Position == position,
				MergeTreeDeltaType.Annotate => previous.Position + previous.Length == position
					&& PropertyMap.MatchProperties(previous.PropertyDeltas, propertyDeltas),
				_ => false,
			};
		}

		private static int GetTotalRangeLength(IReadOnlyList<SequenceDeltaRange> ranges)
		{
			int length = 0;
			foreach (SequenceDeltaRange range in ranges)
			{
				length += range.Length;
			}

			return length;
		}

		private static string GetSequenceDeltaOpType(MergeTreeDeltaType operation)
		{
			return operation switch
			{
				MergeTreeDeltaType.Insert => _insertOpType,
				MergeTreeDeltaType.Remove => _removeOpType,
				MergeTreeDeltaType.Obliterate => _obliterateOpType,
				MergeTreeDeltaType.Annotate => _annotateOpType,
				_ => throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unhandled sequence delta op type: {operation}"),
			};
		}

		private static long? TryGetMinimumSequenceNumber(SequencedDocumentMessageDescriptor descriptor)
		{
			foreach (string memberName in new[] { "MinimumSequenceNumber", "minimumSequenceNumber", "MinSeq", "minSeq" })
			{
				if (ReadOptionalLongMember(descriptor, memberName) is long descriptorValue)
				{
					return descriptorValue;
				}

				if (ReadOptionalLongMember(descriptor.Seq, memberName) is long sequenceValue)
				{
					return sequenceValue;
				}
			}

			return null;
		}

		private static long? ReadOptionalLongMember(object target, string memberName)
		{
			Type targetType = target.GetType();
			System.Reflection.PropertyInfo? property = targetType.GetProperty(
				memberName,
				System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
			if (property is not null && TryConvertToLong(property.GetValue(target), out long propertyValue))
			{
				return propertyValue;
			}

			System.Reflection.FieldInfo? field = targetType.GetField(
				memberName,
				System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
			if (field is not null && TryConvertToLong(field.GetValue(target), out long fieldValue))
			{
				return fieldValue;
			}

			return null;
		}

		private static bool TryConvertToLong(object? value, out long result)
		{
			switch (value)
			{
				case long longValue:
					result = longValue;
					return true;

				case int intValue:
					result = intValue;
					return true;

				case JsonElement jsonElement when jsonElement.ValueKind == JsonValueKind.Number
					&& jsonElement.TryGetInt64(out long jsonValue):
					result = jsonValue;
					return true;

				default:
					result = 0;
					return false;
			}
		}

		private static string ExtractInsertedText(object? segment)
		{
			switch (segment)
			{
				case null:
					return string.Empty;

				case Marker:
				case IJSONMarkerSegment:
					return string.Empty;

				case string text:
					return text;

				case TextSegment textSegment:
					return textSegment.Text;

				case IJSONTextSegment jsonTextSegment:
					return jsonTextSegment.Text;

				case JsonElement jsonElement:
					return ExtractInsertedText(jsonElement);

				default:
					TextSegment? convertedTextSegment = TextSegment.FromJSONObject(segment);
					return convertedTextSegment?.Text ?? segment.ToString() ?? string.Empty;
			}
		}

		private static string ExtractInsertedText(JsonElement segment)
		{
			if (segment.ValueKind == JsonValueKind.String)
			{
				return segment.GetString() ?? string.Empty;
			}

			if (segment.ValueKind == JsonValueKind.Object
				&& segment.TryGetProperty("text", out JsonElement textProperty)
				&& textProperty.ValueKind == JsonValueKind.String)
			{
				return textProperty.GetString() ?? string.Empty;
			}

			if (segment.ValueKind == JsonValueKind.Object
				&& segment.TryGetProperty("marker", out _))
			{
				return string.Empty;
			}

			return segment.ToString();
		}

		private static Marker? ExtractInsertedMarker(object? segment)
		{
			return Marker.FromJSONObject(segment);
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

		private IDisposable EnterLocalMutation()
		{
			lock (_lock)
			{
				if (_localMutationDepth > 0)
				{
					throw new InvalidOperationException("Reentrancy detected in sequence local ops");
				}

				_localMutationDepth++;
			}

			return new LocalMutationScope(this);
		}

		private void ExitLocalMutation()
		{
			lock (_lock)
			{
				_localMutationDepth--;
			}
		}

		private int PositionFromRelativePosition(object? relativePosition, string name)
		{
			if (relativePosition is null)
			{
				throw new ArgumentNullException(name);
			}

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

		private sealed class LocalMutationScope : IDisposable
		{
			private SharedString? _owner;

			public LocalMutationScope(SharedString owner)
			{
				_owner = owner;
			}

			public void Dispose()
			{
				SharedString? owner = _owner;
				if (owner is not null)
				{
					_owner = null;
					owner.ExitLocalMutation();
				}
			}
		}

		private void EmitOrBatchLocalOp(IMergeTreeOp op, string opTypeName)
		{
			ArgumentNullException.ThrowIfNull(op);

			// Interval ops go on the wire individually and cannot be members of a group batch.
			if (op is IntervalOpMsg)
			{
				SendLocalOp(op, opTypeName);
				return;
			}

			lock (_lock)
			{
				if (_batchOps is not null)
				{
					if (op is not MergeTreeOp mergeTreeOp)
					{
						throw new OcsException(OcsGateErrorCode.UnknownOp, $"Cannot batch op runtime type: {op.GetType().FullName}");
					}

					_batchOps.Add(mergeTreeOp);
					return;
				}
			}

			SendLocalOp(op, opTypeName);
		}

		private void FlushBatchOps(IReadOnlyList<MergeTreeOp> ops)
		{
			if (ops.Count == 0)
			{
				return;
			}

			if (ops.Count == 1)
			{
				SendLocalOp(ops[0], GetLocalOpTypeName(ops[0]));
				return;
			}

			if (_sender is null)
			{
				return;
			}

			MergeTreeGroupMsg groupOp = new();
			groupOp.Ops.AddRange(ops);
			string opJson = SharedStringOpSerializer.Serialize(groupOp, _registry, currentSequenceNumber: _client.CollabWindowCurrentSeq);
			SequenceNumber sequenceNumber = _sender.QueueDataObjectMessage(_id, _groupOpType, opJson);
			AssociateSentClientSequence(ops, sequenceNumber);
		}

		private void SendLocalOp(IMergeTreeOp op, string opTypeName)
		{
			if (_sender is null)
			{
				return;
			}

			string opJson = SharedStringOpSerializer.Serialize(op, _registry, currentSequenceNumber: _client.CollabWindowCurrentSeq);
			SequenceNumber sequenceNumber = _sender.QueueDataObjectMessage(_id, opTypeName, opJson);
			if (op.ClientSeq is long localSeq)
			{
				AssociateSentClientSequence(new long[] { localSeq }, sequenceNumber);
			}
		}

		private void SendRegeneratedPendingOp(IMergeTreeOp op)
		{
			if (_sender is null)
			{
				return;
			}

			string opJson = SharedStringOpSerializer.Serialize(op, _registry, currentSequenceNumber: _client.CollabWindowCurrentSeq);
			SequenceNumber sequenceNumber = _sender.QueueDataObjectMessage(_id, GetLocalOpTypeName(op), opJson);
			if (op is MergeTreeGroupMsg groupOp)
			{
				AssociateSentClientSequence(groupOp.Ops, sequenceNumber);
			}
			else if (op.ClientSeq is long localSeq)
			{
				AssociateSentClientSequence(new long[] { localSeq }, sequenceNumber);
			}
		}

		private void AssociateSentClientSequence(IEnumerable<MergeTreeOp> ops, SequenceNumber sequenceNumber)
		{
			List<long> localSeqs = new();
			foreach (MergeTreeOp op in ops)
			{
				if (op.ClientSeq is long localSeq)
				{
					localSeqs.Add(localSeq);
				}
			}

			AssociateSentClientSequence(localSeqs, sequenceNumber);
		}

		private void AssociateSentClientSequence(IEnumerable<long> localSeqs, SequenceNumber sequenceNumber)
		{
			if (!sequenceNumber.HasClientSequenceNumber)
			{
				return;
			}

			lock (_lock)
			{
				_client.AssociatePendingOpsWithClientSequence(localSeqs, sequenceNumber.clientSequenceNumber);
			}
		}

		private static string GetLocalOpTypeName(IMergeTreeOp op)
		{
			return op.Type switch
			{
				MergeTreeDeltaType.Insert => _insertOpType,
				MergeTreeDeltaType.Remove => _removeOpType,
				MergeTreeDeltaType.Obliterate => _obliterateOpType,
				MergeTreeDeltaType.ObliterateSided => _obliterateOpType,
				MergeTreeDeltaType.Annotate => _annotateOpType,
				MergeTreeDeltaType.Group => _groupOpType,
				MergeTreeDeltaType.IntervalAdd => _intervalOpType,
				MergeTreeDeltaType.IntervalDelete => _intervalOpType,
				MergeTreeDeltaType.IntervalChange => _intervalOpType,
				MergeTreeDeltaType.IntervalPropertyChanged => _intervalOpType,
				_ => throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unhandled local op type: {op.Type}"),
			};
		}

		private void RaiseSequenceDelta(SequenceDeltaEventArgs args)
		{
			SequenceDeltaEventHandler? raiseEvent = OnSequenceDelta;
			if (raiseEvent != null)
			{
				raiseEvent(this, args);
			}
		}

		private void EnsureIntervalCollectionOutboundHandlers(IntervalCollection collection)
		{
			if (!_intervalCollectionsWithOutboundHandlers.Add(collection))
			{
				return;
			}

			collection.OnAddInterval += HandleIntervalAdded;
			collection.OnDeleteInterval += HandleIntervalDeleted;
			collection.OnChangeInterval += HandleIntervalChanged;
			collection.OnPropertyChanged += HandleIntervalPropertyChanged;
		}

		private void HandleIntervalAdded(object sender, IntervalAddedEventArgs e)
		{
			if (e.Local)
			{
				FlushPendingIntervalOps();
			}
		}

		private void HandleIntervalDeleted(object sender, IntervalDeletedEventArgs e)
		{
			if (e.Local)
			{
				FlushPendingIntervalOps();
			}
		}

		private void HandleIntervalChanged(object sender, IntervalChangedEventArgs e)
		{
			if (e.Local)
			{
				FlushPendingIntervalOps();
			}
		}

		private void HandleIntervalPropertyChanged(object sender, IntervalPropertyChangedEventArgs e)
		{
			if (e.Local)
			{
				FlushPendingIntervalOps();
			}
		}

		private void FlushPendingIntervalOps()
		{
			lock (_lock)
			{
				while (_client.TryDequeueEmittedOp(out IMergeTreeOp? intervalOp))
				{
					if (intervalOp != null)
					{
						EmitOrBatchLocalOp(intervalOp, _intervalOpType);
					}
				}
			}
		}
	}
}
