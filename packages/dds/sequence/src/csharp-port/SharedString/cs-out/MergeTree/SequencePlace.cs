// -----------------------------------------------------------------------------
// Ported from packages/dds/merge-tree/src/sequencePlace.ts.
// -----------------------------------------------------------------------------

#nullable enable

namespace Microsoft.Office.Web.Fluid.MergeTree
{
    public enum Side
    {
        Before = 0,
        After = 1,
    }

    public sealed class SequencePlace
    {
        public int Position { get; init; }

        public Side Side { get; init; }

        public bool IsStart { get; init; }

        public bool IsEnd { get; init; }

        public static SequencePlace Start => new()
        {
            Position = -1,
            Side = Side.After,
            IsStart = true,
        };

        public static SequencePlace End => new()
        {
            Position = -1,
            Side = Side.Before,
            IsEnd = true,
        };

        public static SequencePlace At(int position, Side side)
        {
            return new SequencePlace()
            {
                Position = position,
                Side = side,
            };
        }

        internal SequencePlace Clone()
        {
            return new SequencePlace()
            {
                Position = Position,
                Side = Side,
                IsStart = IsStart,
                IsEnd = IsEnd,
            };
        }
    }
}
