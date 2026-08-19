// -----------------------------------------------------------------------------
// Ported from packages/dds/sequence/src/intervalIndex/idIntervalIndex.ts (subset for POC).
// Part of the SharedString C# feasibility port — Wave 12b.
// POC uses List/SortedDictionary-based implementations instead of the TS
// RedBlackTree-backed indexes.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;

namespace Microsoft.Office.Web.Fluid.Intervals
{
    public sealed class IdIntervalIndex : IIntervalIndex
    {
        private readonly Dictionary<string, SequenceInterval> _byId = new();

        public void Add(SequenceInterval interval)
        {
            ArgumentNullException.ThrowIfNull(interval);
            if (interval.Id is null)
            {
                throw new ArgumentException("Interval ID must exist before adding interval to the ID index.", nameof(interval));
            }

            _byId[interval.Id] = interval;
        }

        public void Remove(SequenceInterval interval)
        {
            ArgumentNullException.ThrowIfNull(interval);
            if (interval.Id is null)
            {
                throw new ArgumentException("Interval ID must exist before removing interval from the ID index.", nameof(interval));
            }

            _byId.Remove(interval.Id);
        }

        // TS ref: packages/dds/sequence/src/intervalIndex/idIntervalIndex.ts —
        // TS keeps detached intervals addressable via getIntervalById and
        // included in iteration. Detached intervals continue to report -1 for
        // their endpoints; TS callers rely on that shape for id-based lookup
        // and enumeration across the lifecycle.
        //
        // Transient intervals (a port-specific interval type) auto-remove on
        // detachment — they represent short-lived reference pairs and their
        // consumers expect them to disappear once anchoring is lost. The
        // filter below preserves that lifecycle while restoring TS parity for
        // regular (SlideOnRemove) intervals.
        public SequenceInterval? GetIntervalById(string id) =>
            _byId.TryGetValue(id, out SequenceInterval? interval) && !IsDetachedTransient(interval) ? interval : null;

        internal SequenceInterval? GetStoredIntervalById(string id) =>
            _byId.TryGetValue(id, out SequenceInterval? interval) ? interval : null;

        public IEnumerable<SequenceInterval> All => _byId.Values.Where(interval => !IsDetachedTransient(interval));

        private static bool IsDetachedTransient(SequenceInterval interval) =>
            (interval.IntervalType & IntervalType.Transient) != 0 && interval.HasDetachedEndpoint;
    }
}
