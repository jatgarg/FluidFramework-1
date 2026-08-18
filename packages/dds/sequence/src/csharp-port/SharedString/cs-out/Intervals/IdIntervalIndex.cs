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

        public SequenceInterval? GetIntervalById(string id) =>
            _byId.TryGetValue(id, out SequenceInterval? interval) && !interval.HasDetachedEndpoint ? interval : null;

        internal SequenceInterval? GetStoredIntervalById(string id) =>
            _byId.TryGetValue(id, out SequenceInterval? interval) ? interval : null;

        public IEnumerable<SequenceInterval> All => _byId.Values.Where(interval => !interval.HasDetachedEndpoint);
    }
}
