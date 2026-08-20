// -----------------------------------------------------------------------------
// Ported from packages/dds/merge-tree/src/stamps.ts
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid.MergeTree
{
    /// <summary>
    /// A stamp that identifies provenance of an operation performed on the MergeTree.
    /// </summary>
    /// <remarks>
    /// Stamps identify a point in time (Seq/LocalSeq) as well as the source (ClientId) for the operation.
    /// Treat stamps as immutable: new operations applied to a merge-tree should create new stamps rather than modify existing ones.
    /// </remarks>
    public record OperationStamp
    {
        /// <summary>
        /// Gets the sequence number at which this operation was applied.
        /// </summary>
        public long Seq { get; init; }

        /// <summary>
        /// Gets the short client id for the client that performed this operation.
        /// </summary>
        public int ClientId { get; init; }

        /// <summary>
        /// Gets the local sequence at which this operation was applied, when pending an ack.
        /// </summary>
        public long? LocalSeq { get; init; }
    }

    /// <summary>
    /// <see cref="OperationStamp" /> for an insert operation.
    /// </summary>
    public sealed record InsertOperationStamp : OperationStamp
    {
        /// <summary>
        /// Gets the stamp type discriminator.
        /// </summary>
        public string Type => "insert";
    }

    /// <summary>
    /// <see cref="OperationStamp" /> for an operation that removes a segment from some perspective.
    /// </summary>
    public abstract record RemoveOperationStamp : OperationStamp
    {
        /// <summary>
        /// Gets the stamp type discriminator.
        /// </summary>
        public abstract string Type { get; }
    }

    /// <summary>
    /// <see cref="OperationStamp" /> for a set remove operation.
    /// </summary>
    public sealed record SetRemoveOperationStamp : RemoveOperationStamp
    {
        /// <summary>
        /// Gets the stamp type discriminator.
        /// </summary>
        public override string Type => "setRemove";
    }

    /// <summary>
    /// <see cref="OperationStamp" /> for a slice remove operation, the stamp kind used by obliterate.
    /// </summary>
    public sealed record SliceRemoveOperationStamp : RemoveOperationStamp
    {
        /// <summary>
        /// Gets the stamp type discriminator.
        /// </summary>
        public override string Type => "sliceRemove";

        /// <summary>
        /// Gets the sidedness of the start boundary when this stamp sits on a boundary segment.
        /// </summary>
        public Side? StartSide { get; init; }

        /// <summary>
        /// Gets the sidedness of the end boundary when this stamp sits on a boundary segment.
        /// </summary>
        public Side? EndSide { get; init; }
    }

    /// <summary>
    /// Utility methods for comparing and sorting operation stamps.
    /// </summary>
    public static class Stamps
    {
        /// <summary>
        /// Returns whether <paramref name="a" /> occurs before <paramref name="b" />.
        /// </summary>
        /// <param name="a">The first stamp.</param>
        /// <param name="b">The second stamp.</param>
        /// <returns><see langword="true" /> when <paramref name="a" /> occurs before <paramref name="b" />.</returns>
        public static bool LessThan(OperationStamp a, OperationStamp b)
        {
            if (a.Seq == Constants.UnassignedSequenceNumber)
            {
                return b.Seq == Constants.UnassignedSequenceNumber && LocalSeqValue(a) < LocalSeqValue(b);
            }

            if (b.Seq == Constants.UnassignedSequenceNumber)
            {
                return true;
            }

            return a.Seq < b.Seq;
        }

        /// <summary>
        /// Returns whether <paramref name="a" /> occurs at or after <paramref name="b" />.
        /// </summary>
        /// <param name="a">The first stamp.</param>
        /// <param name="b">The second stamp.</param>
        /// <returns><see langword="true" /> when <paramref name="a" /> occurs at or after <paramref name="b" />.</returns>
        public static bool Gte(OperationStamp a, OperationStamp b)
        {
            return !LessThan(a, b);
        }

        /// <summary>
        /// Returns whether <paramref name="a" /> occurs after <paramref name="b" />.
        /// </summary>
        /// <param name="a">The first stamp.</param>
        /// <param name="b">The second stamp.</param>
        /// <returns><see langword="true" /> when <paramref name="a" /> occurs after <paramref name="b" />.</returns>
        public static bool GreaterThan(OperationStamp a, OperationStamp b)
        {
            if (a.Seq == Constants.UnassignedSequenceNumber)
            {
                return b.Seq != Constants.UnassignedSequenceNumber || LocalSeqValue(a) > LocalSeqValue(b);
            }

            if (b.Seq == Constants.UnassignedSequenceNumber)
            {
                return false;
            }

            return a.Seq > b.Seq;
        }

        /// <summary>
        /// Returns whether <paramref name="a" /> occurs at or before <paramref name="b" />.
        /// </summary>
        /// <param name="a">The first stamp.</param>
        /// <param name="b">The second stamp.</param>
        /// <returns><see langword="true" /> when <paramref name="a" /> occurs at or before <paramref name="b" />.</returns>
        public static bool Lte(OperationStamp a, OperationStamp b)
        {
            return !GreaterThan(a, b);
        }

        /// <summary>
        /// Returns whether two operation stamps identify the same operation provenance.
        /// </summary>
        /// <param name="a">The first stamp.</param>
        /// <param name="b">The second stamp.</param>
        /// <returns><see langword="true" /> when the stamps are equal.</returns>
        public static bool Equal(OperationStamp a, OperationStamp b)
        {
            return a.Seq == b.Seq && a.ClientId == b.ClientId && a.LocalSeq == b.LocalSeq;
        }

        /// <summary>
        /// Returns whether a stamp is local and pending ack.
        /// </summary>
        /// <param name="a">The stamp to test.</param>
        /// <returns><see langword="true" /> when the stamp is local.</returns>
        public static bool IsLocal(OperationStamp a)
        {
            return a.Seq == Constants.UnassignedSequenceNumber;
        }

        /// <summary>
        /// Returns whether a stamp represents a squashed operation.
        /// </summary>
        /// <param name="a">The stamp to test.</param>
        /// <returns><see langword="true" /> when the stamp is squashed.</returns>
        public static bool IsSquashedOp(OperationStamp a)
        {
            return a.ClientId == Constants.SquashClient && a.Seq == Constants.UniversalSequenceNumber;
        }

        /// <summary>
        /// Returns whether a stamp has been acked.
        /// </summary>
        /// <param name="a">The stamp to test.</param>
        /// <returns><see langword="true" /> when the stamp has been acked.</returns>
        public static bool IsAcked(OperationStamp a)
        {
            return a.Seq != Constants.UnassignedSequenceNumber;
        }

        /// <summary>
        /// Inserts a stamp into a sorted list of stamps in the correct sorted position.
        /// </summary>
        /// <param name="list">The sorted stamp list to update.</param>
        /// <param name="stamp">The stamp to insert.</param>
        public static void SpliceIntoList<TStamp>(IList<TStamp> list, TStamp stamp)
            where TStamp : OperationStamp
        {
            if (IsLocal(stamp) || list.Count == 0)
            {
                list.Add(stamp);
            }
            else
            {
                for (int i = list.Count - 1; i >= 0; i--)
                {
                    if (GreaterThan(stamp, list[i]))
                    {
                        list.Insert(i + 1, stamp);
                        return;
                    }
                }

                list.Insert(0, stamp);
            }
        }

        /// <summary>
        /// Sorts a stamp list using the same ordering as <see cref="SpliceIntoList{TStamp}" />.
        /// </summary>
        /// <typeparam name="TStamp">The operation stamp type.</typeparam>
        /// <param name="list">The stamp list to sort.</param>
        public static void SortList<TStamp>(List<TStamp> list)
            where TStamp : OperationStamp
        {
            list.Sort(Compare);
        }

        /// <summary>
        /// Returns whether a stamp list contains any acked operation.
        /// </summary>
        /// <param name="list">The stamp list to scan.</param>
        /// <returns><see langword="true" /> when any stamp is acked.</returns>
        public static bool HasAnyAckedOperation(IEnumerable<OperationStamp> list)
        {
            foreach (OperationStamp stamp in list)
            {
                if (IsAcked(stamp))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Compares two operation stamps for sorting.
        /// </summary>
        /// <param name="a">The first stamp.</param>
        /// <param name="b">The second stamp.</param>
        /// <returns>1 when <paramref name="a" /> is greater, -1 when it is less, and 0 when equal.</returns>
        public static int Compare(OperationStamp a, OperationStamp b)
        {
            if (GreaterThan(a, b))
            {
                return 1;
            }
            else if (LessThan(a, b))
            {
                return -1;
            }
            else
            {
                return 0;
            }
        }

        private static long LocalSeqValue(OperationStamp stamp)
        {
            if (stamp.LocalSeq is long localSeq)
            {
                return localSeq;
            }

            throw new System.InvalidOperationException("Local sequence number is required for unassigned operation stamps.");
        }
    }
}
