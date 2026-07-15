// -----------------------------------------------------------------------------
// Ported from packages/dds/sequence/src/intervalCollection.ts (subset for POC).
// Part of the SharedString C# feasibility port — Wave 12c.
//
// POC scope: Add/RemoveById/GetById/Change + 4 iterators + 4 events. Deferred:
// snapshot persistence, IntervalCollectionValueType registration, revertibles,
// attribution.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using MergeTree = Microsoft.Office.Web.Fluid.MergeTree;

namespace Microsoft.Office.Web.Fluid.Intervals
{
	/// <summary>
	/// Event fired when an interval is added.
	/// </summary>
	public sealed class IntervalAddedEventArgs
	{
		public SequenceInterval Interval { get; init; } = null!;

		public bool Local { get; init; }

		public MergeTree.IntervalOpMsg? Operation { get; init; }
	}

	/// <summary>Event args for interval deleted.</summary>
	public sealed class IntervalDeletedEventArgs
	{
		public SequenceInterval Interval { get; init; } = null!;

		public int? PreviousStart { get; init; }

		public int? PreviousEnd { get; init; }

		public bool Local { get; init; }

		public MergeTree.IntervalOpMsg? Operation { get; init; }
	}

	/// <summary>Event args for interval endpoint change.</summary>
	public sealed class IntervalChangedEventArgs
	{
		public SequenceInterval Interval { get; init; } = null!;

		public int? PreviousStart { get; init; }

		public int? PreviousEnd { get; init; }

		public bool Slide { get; init; }

		public bool Local { get; init; }

		public MergeTree.IntervalOpMsg? Operation { get; init; }
	}

	/// <summary>Event args for interval property change.</summary>
	public sealed class IntervalPropertyChangedEventArgs
	{
		public SequenceInterval Interval { get; init; } = null!;

		public MergeTree.PropertySet ChangedProps { get; init; } = new();

		public MergeTree.PropertySet PropertyDeltas { get; init; } = new();

		public bool Local { get; init; }

		public MergeTree.IntervalOpMsg? Operation { get; init; }
	}

	/// <summary>Event args for any interval endpoint or property change.</summary>
	public sealed class IntervalUpdatedEventArgs
	{
		public SequenceInterval Interval { get; init; } = null!;

		public MergeTree.PropertySet PropertyDeltas { get; init; } = new();

		public int? PreviousStart { get; init; }

		public int? PreviousEnd { get; init; }

		public bool Slide { get; init; }

		public bool Local { get; init; }

		public MergeTree.IntervalOpMsg? Operation { get; init; }
	}

	public delegate void IntervalAddedEventHandler(object sender, IntervalAddedEventArgs e);

	public delegate void IntervalDeletedEventHandler(object sender, IntervalDeletedEventArgs e);

	public delegate void IntervalChangedEventHandler(object sender, IntervalChangedEventArgs e);

	public delegate void IntervalPropertyChangedEventHandler(object sender, IntervalPropertyChangedEventArgs e);

	public delegate void IntervalUpdatedEventHandler(object sender, IntervalUpdatedEventArgs e);

	/// <summary>
	/// A named collection of SequenceIntervals over a SharedString.
	/// Owns all interval indexes for efficient queries.
	/// </summary>
	public sealed class IntervalCollection : IEnumerable<SequenceInterval>
	{
		private readonly object _lock = new();
		private readonly MergeTree.MergeTree _mergeTree;
		private readonly IIntervalOpSender _opSender;
		private readonly IdIntervalIndex _idIndex = new();
		private readonly OverlappingIntervalsIndex _overlappingIndex = new();
		private readonly StartpointInRangeIndex _startpointIndex = new();
		private readonly EndpointInRangeIndex _endpointInRangeIndex = new();
		private readonly EndpointIndex _endpointIndex = new();
		private readonly IIntervalIndex[] _indexes;

		internal IntervalCollection(
			string name,
			IntervalType endpointType,
			MergeTree.MergeTree mergeTree,
			IIntervalOpSender opSender)
		{
			if (string.IsNullOrEmpty(name))
			{
				throw new ArgumentException("Interval collection name must be provided.", nameof(name));
			}

			ArgumentNullException.ThrowIfNull(mergeTree);
			ArgumentNullException.ThrowIfNull(opSender);

			Name = name;
			EndpointType = endpointType;
			_mergeTree = mergeTree;
			_opSender = opSender;
			_indexes = new IIntervalIndex[]
			{
				_idIndex,
				_overlappingIndex,
				_startpointIndex,
				_endpointInRangeIndex,
				_endpointIndex,
			};
		}

		// Public event surface
		public event IntervalAddedEventHandler? OnAddInterval;

		public event IntervalDeletedEventHandler? OnDeleteInterval;

		public event IntervalChangedEventHandler? OnChangeInterval;

		public event IntervalPropertyChangedEventHandler? OnPropertyChanged;

		public event IntervalUpdatedEventHandler? OnChanged;

		// Public API
		public string Name { get; }

		public IntervalType EndpointType { get; }

		public SequenceInterval Add(
			int start,
			int end,
			MergeTree.PropertySet? properties = null,
			string? intervalId = null,
			IntervalStickiness stickiness = IntervalStickiness.End,
			IntervalType? intervalType = null)
		{
			(Side startSide, Side endSide) = IntervalUtils.SidesFromStickiness(stickiness);
			return Add(start, startSide, end, endSide, properties, intervalId, intervalType);
		}

		public SequenceInterval Add(
			int start,
			Side startSide,
			int end,
			Side endSide,
			MergeTree.PropertySet? properties = null,
			string? intervalId = null,
			IntervalType? intervalType = null)
		{
			SequenceInterval interval;
			lock (_lock)
			{
				interval = AddCore(start, end, properties, intervalId, intervalType ?? EndpointType, validateDuplicate: true, startSide, endSide);
				_opSender.SendIntervalAdd(
					Name,
					interval.Id!,
					start,
					end,
					interval.IntervalType,
					interval.Stickiness,
					interval.StartSide,
					interval.EndSide,
					CloneProperties(interval.Properties));
			}

			RaiseAdd(interval, local: true, operation: null);
			return interval;
		}

		public SequenceInterval? RemoveIntervalById(string id)
		{
			SequenceInterval? interval;
			int? previousStart;
			int? previousEnd;
			lock (_lock)
			{
				interval = RemoveIntervalByIdCore(id, out previousStart, out previousEnd);
				if (interval is null)
				{
					return null;
				}

				_opSender.SendIntervalDelete(Name, id);
			}

			RaiseDelete(interval, previousStart, previousEnd, local: true, operation: null);
			return interval;
		}

		public SequenceInterval? GetIntervalById(string id)
		{
			ArgumentException.ThrowIfNullOrEmpty(id);

			lock (_lock)
			{
				return GetActiveIntervalById(id);
			}
		}

		public SequenceInterval? Change(
			string id,
			int? newStart,
			int? newEnd,
			Side? newStartSide = null,
			Side? newEndSide = null,
			MergeTree.PropertySet? props = null)
		{
			ArgumentException.ThrowIfNullOrEmpty(id);

			SequenceInterval? interval;
			int? previousStart;
			int? previousEnd;
			MergeTree.PropertySet? changedProps = null;
			MergeTree.PropertySet? propertyDeltas = null;
			bool hasEndpointChange = newStart.HasValue || newEnd.HasValue || newStartSide.HasValue || newEndSide.HasValue;
			lock (_lock)
			{
				(interval, previousStart, previousEnd) = hasEndpointChange
					? ChangeCore(id, newStart, newEnd, newStartSide, newEndSide)
					: (GetActiveIntervalById(id), null, null);
				if (interval is null)
				{
					return null;
				}

				if (props is not null)
				{
					(changedProps, propertyDeltas) = ApplyPropertyChanges(interval, props);
				}

				if (hasEndpointChange)
				{
					_opSender.SendIntervalChange(
						Name,
						id,
						newStart,
						newEnd,
						interval.Stickiness,
						interval.StartSide,
						interval.EndSide);
				}

				if (changedProps is { Count: > 0 })
				{
					_opSender.SendIntervalPropertyChanged(Name, id, props!);
				}
			}

			if (changedProps is { Count: > 0 } && propertyDeltas is not null)
			{
				RaisePropertyChanged(interval, changedProps, propertyDeltas, local: true, operation: null);
				RaiseChanged(interval, propertyDeltas, previousStart: null, previousEnd: null, slide: false, local: true, operation: null);
			}

			if (hasEndpointChange)
			{
				RaiseChange(interval, previousStart, previousEnd, slide: false, local: true, operation: null);
				RaiseChanged(interval, new MergeTree.PropertySet(), previousStart, previousEnd, slide: false, local: true, operation: null);
			}

			return interval;
		}

		public SequenceInterval? ChangeProperties(string id, MergeTree.PropertySet props)
		{
			ArgumentNullException.ThrowIfNull(props);

			return Change(id, newStart: null, newEnd: null, props: props);
		}

		public IEnumerable<SequenceInterval> CreateForwardIteratorWithStartPosition(int startPos)
		{
			lock (_lock)
			{
				return SnapshotIntervals()
					.Where(interval => interval.StartPosition is int position && position >= startPos)
					.OrderBy(interval => interval.StartPosition)
					.ThenBy(interval => interval.EndPosition)
					.ToArray();
			}
		}

		public IEnumerable<SequenceInterval> CreateBackwardIteratorWithStartPosition(int startPos)
		{
			lock (_lock)
			{
				return SnapshotIntervals()
					.Where(interval => interval.StartPosition is int position && position < startPos)
					.OrderByDescending(interval => interval.StartPosition)
					.ThenByDescending(interval => interval.EndPosition)
					.ToArray();
			}
		}

		public IEnumerable<SequenceInterval> CreateForwardIteratorWithEndPosition(int endPos)
		{
			lock (_lock)
			{
				return SnapshotIntervals()
					.Where(interval => interval.EndPosition is int position && position >= endPos)
					.OrderBy(interval => interval.EndPosition)
					.ThenBy(interval => interval.StartPosition)
					.ToArray();
			}
		}

		public IEnumerable<SequenceInterval> CreateBackwardIteratorWithEndPosition(int endPos)
		{
			lock (_lock)
			{
				return SnapshotIntervals()
					.Where(interval => interval.EndPosition is int position && position < endPos)
					.OrderByDescending(interval => interval.EndPosition)
					.ThenByDescending(interval => interval.StartPosition)
					.ToArray();
			}
		}

		public IEnumerable<SequenceInterval> FindOverlappingIntervals(int startPosition, int endPosition)
		{
			if (endPosition < startPosition)
			{
				return Array.Empty<SequenceInterval>();
			}

			lock (_lock)
			{
				return SnapshotIntervals()
					.Where(interval => IntervalOverlapsInclusive(interval, startPosition, endPosition))
					.OrderBy(interval => interval.StartPosition)
					.ThenBy(interval => interval.EndPosition)
					.ThenBy(interval => interval.Id, StringComparer.Ordinal)
					.ToArray();
			}
		}

		public IEnumerable<SequenceInterval> FindIntervalsWithStartpointInRange(int start, int end)
		{
			if (start <= 0 || start > end)
			{
				return Array.Empty<SequenceInterval>();
			}

			lock (_lock)
			{
				return SnapshotIntervals()
					.Where(interval => interval.StartPosition is int position && position >= start && position <= end)
					.OrderBy(interval => interval.StartPosition)
					.ThenBy(interval => interval.EndPosition)
					.ThenBy(interval => interval.Id, StringComparer.Ordinal)
					.ToArray();
			}
		}

		public IEnumerable<SequenceInterval> FindIntervalsWithEndpointInRange(int start, int end)
		{
			if (start <= 0 || start > end)
			{
				return Array.Empty<SequenceInterval>();
			}

			lock (_lock)
			{
				return SnapshotIntervals()
					.Where(interval => interval.EndPosition is int position && position >= start && position <= end)
					.OrderBy(interval => interval.EndPosition)
					.ThenBy(interval => interval.StartPosition)
					.ThenBy(interval => interval.Id, StringComparer.Ordinal)
					.ToArray();
			}
		}

		public SequenceInterval? PreviousInterval(int position)
		{
			lock (_lock)
			{
				return SnapshotIntervals()
					.Where(interval => interval.EndPosition is int endPosition && endPosition <= position)
					.OrderByDescending(interval => interval.EndPosition)
					.ThenByDescending(interval => interval.StartPosition)
					.ThenByDescending(interval => interval.Id, StringComparer.Ordinal)
					.FirstOrDefault();
			}
		}

		public SequenceInterval? NextInterval(int position)
		{
			lock (_lock)
			{
				return SnapshotIntervals()
					.Where(interval => interval.EndPosition is int endPosition && endPosition >= position)
					.OrderBy(interval => interval.EndPosition)
					.ThenBy(interval => interval.StartPosition)
					.ThenBy(interval => interval.Id, StringComparer.Ordinal)
					.FirstOrDefault();
			}
		}

		public void Map(Action<SequenceInterval> action)
		{
			ArgumentNullException.ThrowIfNull(action);

			foreach (SequenceInterval interval in this)
			{
				action(interval);
			}
		}

		public void GatherIterationResults(
			ICollection<SequenceInterval> results,
			bool iteratesForward,
			int? start = null,
			int? end = null)
		{
			ArgumentNullException.ThrowIfNull(results);

			IEnumerable<SequenceInterval> intervals = start.HasValue
				? (iteratesForward
					? CreateForwardIteratorWithStartPosition(start.Value)
					: CreateBackwardIteratorWithStartPosition(start.Value))
				: end.HasValue
					? (iteratesForward
						? CreateForwardIteratorWithEndPosition(end.Value)
						: CreateBackwardIteratorWithEndPosition(end.Value))
					: this;

			foreach (SequenceInterval interval in intervals)
			{
				results.Add(interval);
			}
		}

		public IEnumerator<SequenceInterval> GetEnumerator()
		{
			lock (_lock)
			{
				return ((IEnumerable<SequenceInterval>)SnapshotIntervals()).GetEnumerator();
			}
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			return GetEnumerator();
		}

		public MergeTree.ISegment? GetSegment(MergeTree.LocalReferencePosition endpoint)
		{
			ArgumentNullException.ThrowIfNull(endpoint);

			lock (_lock)
			{
				return endpoint.Segment;
			}
		}

		internal void ApplyRemoteAdd(MergeTree.IntervalAddOpMsg op)
		{
			ArgumentNullException.ThrowIfNull(op);

			SequenceInterval interval;
			lock (_lock)
			{
				interval = AddCore(op.Start, op.End, op.Props, op.IntervalId, op.IntervalType, validateDuplicate: true, op.StartSide, op.EndSide);
			}

			RaiseAdd(interval, local: false, operation: op);
		}

		internal void ApplyRemoteDelete(MergeTree.IntervalDeleteOpMsg op)
		{
			ArgumentNullException.ThrowIfNull(op);

			SequenceInterval? interval;
			int? previousStart;
			int? previousEnd;
			lock (_lock)
			{
				interval = RemoveIntervalByIdCore(op.IntervalId, out previousStart, out previousEnd);
			}

			if (interval is not null)
			{
				RaiseDelete(interval, previousStart, previousEnd, local: false, operation: op);
			}
		}

		internal void ApplyRemoteChange(MergeTree.IntervalChangeOpMsg op)
		{
			ArgumentNullException.ThrowIfNull(op);

			SequenceInterval? interval;
			int? previousStart;
			int? previousEnd;
			lock (_lock)
			{
				(interval, previousStart, previousEnd) = ChangeCore(op.IntervalId, op.Start, op.End, op.StartSide, op.EndSide);
			}

			if (interval is not null)
			{
				RaiseChange(interval, previousStart, previousEnd, slide: false, local: false, operation: op);
				RaiseChanged(interval, new MergeTree.PropertySet(), previousStart, previousEnd, slide: false, local: false, operation: op);
			}
		}

		internal void ApplyRemotePropertyChanged(MergeTree.IntervalPropertyChangedOpMsg op)
		{
			ArgumentNullException.ThrowIfNull(op);

			SequenceInterval? interval;
			MergeTree.PropertySet changedProps;
			MergeTree.PropertySet propertyDeltas;
			lock (_lock)
			{
				interval = GetActiveIntervalById(op.IntervalId);
				if (interval is null)
				{
					return;
				}

				(changedProps, propertyDeltas) = ApplyPropertyChanges(interval, op.Props);
			}

			if (changedProps.Count > 0)
			{
				RaisePropertyChanged(interval, changedProps, propertyDeltas, local: false, operation: op);
				RaiseChanged(interval, propertyDeltas, previousStart: null, previousEnd: null, slide: false, local: false, operation: op);
			}
		}

		internal void ApplyOwnAck(MergeTree.IntervalAddOpMsg op)
		{
			ArgumentNullException.ThrowIfNull(op);
		}

		internal void ApplyOwnAck(MergeTree.IntervalDeleteOpMsg op)
		{
			ArgumentNullException.ThrowIfNull(op);
		}

		internal void ApplyOwnAck(MergeTree.IntervalChangeOpMsg op)
		{
			ArgumentNullException.ThrowIfNull(op);
		}

		internal void ApplyOwnAck(MergeTree.IntervalPropertyChangedOpMsg op)
		{
			ArgumentNullException.ThrowIfNull(op);
		}

		internal void DropDetachedIntervalForRebase(string id)
		{
			ArgumentException.ThrowIfNullOrEmpty(id);

			SequenceInterval? interval;
			lock (_lock)
			{
				interval = _idIndex.GetStoredIntervalById(id);
				if (interval is null || !interval.HasDetachedEndpoint)
				{
					return;
				}

				RemoveFromIndexes(interval);
				_mergeTree.RemoveReferencePosition(interval.Start);
				_mergeTree.RemoveReferencePosition(interval.End);
			}

			RaiseDelete(interval, previousStart: null, previousEnd: null, local: true, operation: null);
		}

		private SequenceInterval AddCore(
			int start,
			int end,
			MergeTree.PropertySet? properties,
			string? intervalId,
			IntervalType intervalType,
			bool validateDuplicate,
			Side startSide,
			Side endSide)
		{
			ValidateEndpointOrder(start, end);

			string id = string.IsNullOrEmpty(intervalId) ? Guid.NewGuid().ToString() : intervalId;
			if (validateDuplicate && GetActiveIntervalById(id) is not null)
			{
				throw new InvalidOperationException($"Interval '{id}' already exists in collection '{Name}'.");
			}

			MergeTree.LocalReferencePosition startReference = CreateEndpointReference(start, intervalType, isStartEndpoint: true, startSide, endSide);
			MergeTree.LocalReferencePosition endReference;
			try
			{
				endReference = CreateEndpointReference(end, intervalType, isStartEndpoint: false, startSide, endSide);
			}
			catch
			{
				_mergeTree.RemoveReferencePosition(startReference);
				throw;
			}

			SequenceInterval interval = new(
				_mergeTree,
				startReference,
				endReference,
				intervalType,
				startSide,
				endSide,
				id,
				properties);

			AddToIndexes(interval);
			return interval;
		}

		private SequenceInterval? RemoveIntervalByIdCore(string id, out int? previousStart, out int? previousEnd)
		{
			ArgumentException.ThrowIfNullOrEmpty(id);
			previousStart = null;
			previousEnd = null;

			SequenceInterval? interval = GetActiveIntervalById(id);
			if (interval is null)
			{
				return null;
			}

			previousStart = interval.StartPosition;
			previousEnd = interval.EndPosition;
			RemoveFromIndexes(interval);
			_mergeTree.RemoveReferencePosition(interval.Start);
			_mergeTree.RemoveReferencePosition(interval.End);
			return interval;
		}

		private (SequenceInterval? Interval, int? PreviousStart, int? PreviousEnd) ChangeCore(
			string id,
			int? newStart,
			int? newEnd,
			Side? newStartSide = null,
			Side? newEndSide = null)
		{
			ArgumentException.ThrowIfNullOrEmpty(id);

			SequenceInterval? interval = GetActiveIntervalById(id);
			if (interval is null)
			{
				return (null, null, null);
			}

			int? previousStart = interval.StartPosition;
			int? previousEnd = interval.EndPosition;
			int? resultingStart = newStart ?? previousStart;
			int? resultingEnd = newEnd ?? previousEnd;
			Side resultingStartSide = newStartSide ?? (newStart.HasValue ? IntervalUtils.DefaultSide : interval.StartSide);
			Side resultingEndSide = newEndSide ?? (newEnd.HasValue ? IntervalUtils.DefaultSide : interval.EndSide);
			if (resultingStart.HasValue && resultingEnd.HasValue)
			{
				ValidateEndpointOrder(resultingStart.Value, resultingEnd.Value);
			}

			MergeTree.LocalReferencePosition? replacementStart = null;
			MergeTree.LocalReferencePosition? replacementEnd = null;
			try
			{
				if (newStart.HasValue || (newStartSide.HasValue && previousStart.HasValue))
				{
					replacementStart = CreateEndpointReference(
						newStart ?? previousStart!.Value,
						interval.IntervalType,
						isStartEndpoint: true,
						resultingStartSide,
						resultingEndSide);
				}

				if (newEnd.HasValue || (newEndSide.HasValue && previousEnd.HasValue))
				{
					replacementEnd = CreateEndpointReference(
						newEnd ?? previousEnd!.Value,
						interval.IntervalType,
						isStartEndpoint: false,
						resultingStartSide,
						resultingEndSide);
				}
			}
			catch
			{
				if (replacementStart is not null)
				{
					_mergeTree.RemoveReferencePosition(replacementStart);
				}

				throw;
			}

			RemoveFromIndexes(interval);
			if (replacementStart is not null)
			{
				_mergeTree.RemoveReferencePosition(interval.Start);
				interval.SetStart(replacementStart, resultingStartSide);
			}
			else
			{
				interval.StartSide = resultingStartSide;
			}

			if (replacementEnd is not null)
			{
				_mergeTree.RemoveReferencePosition(interval.End);
				interval.SetEnd(replacementEnd, resultingEndSide);
			}
			else
			{
				interval.EndSide = resultingEndSide;
			}

			AddToIndexes(interval);
			return (interval, previousStart, previousEnd);
		}

		private MergeTree.LocalReferencePosition CreateEndpointReference(
			int position,
			IntervalType intervalType,
			bool isStartEndpoint,
			Side startSide,
			Side endSide)
		{
			MergeTree.SlidingPreference slidingPreference = isStartEndpoint
				? IntervalUtils.StartReferenceSlidingPreference(startSide, endSide)
				: IntervalUtils.EndReferenceSlidingPreference(startSide, endSide);

			return _mergeTree.CreateReferencePosition(
				position,
				EndpointTypeToReferenceType(intervalType, isStartEndpoint),
				slidingPreference);
		}

		private void AddToIndexes(SequenceInterval interval)
		{
			foreach (IIntervalIndex index in _indexes)
			{
				index.Add(interval);
			}
		}

		private void RemoveFromIndexes(SequenceInterval interval)
		{
			foreach (IIntervalIndex index in _indexes)
			{
				index.Remove(interval);
			}
		}

		private SequenceInterval[] SnapshotIntervals()
		{
			return _idIndex.All.ToArray();
		}

		private SequenceInterval? GetActiveIntervalById(string id)
		{
			return _idIndex.GetIntervalById(id);
		}

		private static void ValidateEndpointOrder(int start, int end)
		{
			if (start > end)
			{
				throw new ArgumentOutOfRangeException(nameof(end), "Interval start must be less than or equal to end.");
			}
		}

		private static MergeTree.ReferenceType EndpointTypeToReferenceType(IntervalType endpointType, bool isStartEndpoint)
		{
			if ((endpointType & IntervalType.Transient) != 0)
			{
				return MergeTree.ReferenceType.Transient;
			}

			MergeTree.ReferenceType rangeType = isStartEndpoint
				? MergeTree.ReferenceType.RangeBegin
				: MergeTree.ReferenceType.RangeEnd;
			return rangeType | MergeTree.ReferenceType.SlideOnRemove;
		}

		private static MergeTree.PropertySet? CloneProperties(IReadOnlyDictionary<string, object?>? properties)
		{
			return MergeTree.PropertyMap.ClonePropertySet(properties);
		}

		private static (MergeTree.PropertySet ChangedProps, MergeTree.PropertySet PropertyDeltas) ApplyPropertyChanges(
			SequenceInterval interval,
			IReadOnlyDictionary<string, object?> props)
		{
			MergeTree.PropertySet changedProps = new();
			MergeTree.PropertySet propertyDeltas = new();
			if (props.Count == 0)
			{
				return (changedProps, propertyDeltas);
			}

			MergeTree.PropertySet? current = interval.Properties;
			foreach (KeyValuePair<string, object?> property in props)
			{
				object? previousValue = null;
				bool hadPreviousValue = current is not null && current.TryGetValue(property.Key, out previousValue);
				if (property.Value is null)
				{
					if (!hadPreviousValue)
					{
						continue;
					}

					propertyDeltas[property.Key] = previousValue;
					changedProps[property.Key] = null;
					current!.Remove(property.Key);
					continue;
				}

				if (hadPreviousValue && object.Equals(previousValue, property.Value))
				{
					continue;
				}

				propertyDeltas[property.Key] = hadPreviousValue ? previousValue : null;
				changedProps[property.Key] = property.Value;
				current ??= new MergeTree.PropertySet();
				current[property.Key] = property.Value;
			}

			interval.Properties = current is { Count: > 0 } ? current : null;
			return (changedProps, propertyDeltas);
		}

		private static bool IntervalOverlapsInclusive(SequenceInterval interval, int startPosition, int endPosition)
		{
			if (interval.StartPosition is not int intervalStartPosition || interval.EndPosition is not int intervalEndPosition)
			{
				return false;
			}

			int intervalStart = Math.Min(intervalStartPosition, intervalEndPosition);
			int intervalEnd = Math.Max(intervalStartPosition, intervalEndPosition);
			return intervalStart <= endPosition && startPosition <= intervalEnd;
		}

		private void RaiseAdd(SequenceInterval interval, bool local, MergeTree.IntervalOpMsg? operation)
		{
			OnAddInterval?.Invoke(
				this,
				new IntervalAddedEventArgs()
				{
					Interval = interval,
					Local = local,
					Operation = operation,
				});
		}

		private void RaiseDelete(
			SequenceInterval interval,
			int? previousStart,
			int? previousEnd,
			bool local,
			MergeTree.IntervalOpMsg? operation)
		{
			OnDeleteInterval?.Invoke(
				this,
				new IntervalDeletedEventArgs()
				{
					Interval = interval,
					PreviousStart = previousStart,
					PreviousEnd = previousEnd,
					Local = local,
					Operation = operation,
				});
		}

		private void RaiseChange(
			SequenceInterval interval,
			int? previousStart,
			int? previousEnd,
			bool slide,
			bool local,
			MergeTree.IntervalOpMsg? operation)
		{
			OnChangeInterval?.Invoke(
				this,
				new IntervalChangedEventArgs()
				{
					Interval = interval,
					PreviousStart = previousStart,
					PreviousEnd = previousEnd,
					Slide = slide,
					Local = local,
					Operation = operation,
				});
		}

		private void RaisePropertyChanged(
			SequenceInterval interval,
			MergeTree.PropertySet changedProps,
			MergeTree.PropertySet propertyDeltas,
			bool local,
			MergeTree.IntervalOpMsg? operation)
		{
			OnPropertyChanged?.Invoke(
				this,
				new IntervalPropertyChangedEventArgs()
				{
					Interval = interval,
					ChangedProps = changedProps,
					PropertyDeltas = propertyDeltas,
					Local = local,
					Operation = operation,
				});
		}

		private void RaiseChanged(
			SequenceInterval interval,
			MergeTree.PropertySet propertyDeltas,
			int? previousStart,
			int? previousEnd,
			bool slide,
			bool local,
			MergeTree.IntervalOpMsg? operation)
		{
			OnChanged?.Invoke(
				this,
				new IntervalUpdatedEventArgs()
				{
					Interval = interval,
					PropertyDeltas = propertyDeltas,
					PreviousStart = previousStart,
					PreviousEnd = previousEnd,
					Slide = slide,
					Local = local,
					Operation = operation,
				});
		}
	}

	/// <summary>
	/// Abstraction over how IntervalCollection sends ops. In production, this is Client.
	/// In tests, a fake or Client-with-FakeSender.
	/// </summary>
	public interface IIntervalOpSender
	{
		void SendIntervalAdd(
			string collectionName,
			string intervalId,
			int start,
			int end,
			IntervalType intervalType,
			IntervalStickiness stickiness,
			Side startSide,
			Side endSide,
			MergeTree.PropertySet? props);

		void SendIntervalDelete(string collectionName, string intervalId);

		void SendIntervalChange(
			string collectionName,
			string intervalId,
			int? start,
			int? end,
			IntervalStickiness stickiness,
			Side startSide,
			Side endSide);

		void SendIntervalPropertyChanged(string collectionName, string intervalId, MergeTree.PropertySet props);
	}
}
