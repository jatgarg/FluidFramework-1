// -----------------------------------------------------------------------------
// Ported from packages/dds/merge-tree/src/mergeTreeNodes.ts
// Scope: base segment/block types, TextSegment-only. Obliterate,
// attribution, local refs, and Marker deferred to later waves.
// -----------------------------------------------------------------------------

#nullable enable

namespace Microsoft.Office.Web.Fluid.MergeTree
{
    /// <summary>
    /// Interior merge-tree node containing segments or child blocks.
    /// </summary>
    public sealed class MergeBlock : IHierBlock
    {
        /// <summary>
        /// Maximum number of child slots in a block.
        /// </summary>
        public const int MaxChildren = 8;

        /// <summary>
        /// Minimum preferred child count before rebalancing.
        /// </summary>
        public const int MinChildren = MaxChildren / 2;

        private int childCount;

        /// <summary>
        /// Initializes a new instance of the <see cref="MergeBlock" /> class.
        /// </summary>
        /// <param name="childCount">The initial number of populated child slots.</param>
        public MergeBlock(int childCount = 0)
        {
            Children = new IMergeNode?[MaxChildren];
            ChildCount = childCount;
        }

        /// <inheritdoc />
        public MergeBlock? Parent { get; set; }

        /// <inheritdoc />
        public int Index { get; set; }

        /// <inheritdoc />
        public string Ordinal { get; set; } = string.Empty;

        /// <inheritdoc />
        public int CachedLength { get; set; }

        /// <inheritdoc />
        public int ChildCount
        {
            get
            {
                return childCount;
            }

            set
            {
                if (value < 0 || value > MaxChildren)
                {
                    throw new System.ArgumentOutOfRangeException(nameof(value), "Child count must be within the block capacity.");
                }

                childCount = value;
            }
        }

        /// <inheritdoc />
        public int Capacity
        {
            get
            {
                return MaxChildren;
            }
        }

        /// <inheritdoc />
        public IMergeNode?[] Children { get; }

        internal PartialLengths PartialLengths { get; set; } = new();

        /// <inheritdoc />
        public bool IsLeaf()
        {
            return false;
        }

        internal IMergeNode? GetChild(int index)
        {
            ValidateChildIndex(index);
            return Children[index];
        }

        internal void SetChild(int index, IMergeNode child, bool updateOrdinal = true)
        {
            ValidateChildIndex(index);

            child.Parent = this;
            child.Index = index;
            Children[index] = child;
            if (index >= ChildCount)
            {
                ChildCount = index + 1;
            }

            if (updateOrdinal)
            {
                SetOrdinal(child, index);
            }
        }

        internal void AppendChild(IMergeNode child, bool updateOrdinal = true)
        {
            SetChild(ChildCount, child, updateOrdinal);
        }

        internal void MoveChildrenTo(MergeBlock target, bool updateOrdinal = true)
        {
            MoveChildrenTo(target, ChildCount, updateOrdinal);
        }

        internal void MoveChildrenTo(MergeBlock target, int count, bool updateOrdinal = true)
        {
            if (target is null)
            {
                throw new System.ArgumentNullException(nameof(target));
            }

            if (count < 0 || count > ChildCount)
            {
                throw new System.ArgumentOutOfRangeException(nameof(count), "Moved child count must be within the populated child range.");
            }

            if (target.ChildCount + count > MaxChildren)
            {
                throw new LoggingError("Moved children must fit in the target block.");
            }

            for (int i = 0; i < count; i++)
            {
                IMergeNode child = Children[0]
                    ?? throw new LoggingError("Cannot move a missing child.");
                RemoveChildAt(0);
                target.AppendChild(child, updateOrdinal);
            }
        }

        internal IMergeNode RemoveChildAt(int index)
        {
            if (index < 0 || index >= ChildCount)
            {
                throw new System.ArgumentOutOfRangeException(nameof(index), "Child index must be within the populated child range.");
            }

            IMergeNode removedChild = Children[index]
                ?? throw new LoggingError("Cannot remove a missing child.");
            for (int i = index; i < ChildCount - 1; i++)
            {
                IMergeNode? movedChild = Children[i + 1];
                Children[i] = movedChild;
                if (movedChild is not null)
                {
                    movedChild.Parent = this;
                    movedChild.Index = i;
                }
            }

            Children[ChildCount - 1] = null;
            ChildCount--;
            removedChild.Parent = null;
            removedChild.Index = 0;
            removedChild.Ordinal = string.Empty;
            return removedChild;
        }

        internal bool DropEmptyBlockAt(int index)
        {
            IMergeNode? child = GetChild(index);
            if (child is not MergeBlock childBlock || childBlock.ChildCount != 0)
            {
                return false;
            }

            RemoveChildAt(index);
            return true;
        }

        internal void SetOrdinal(IMergeNode child, int index)
        {
            ValidateChildIndex(index);
            if (ChildCount < 1 || ChildCount > MaxChildren)
            {
                throw new LoggingError("Child count must be within [1,8] before assigning ordinals.");
            }

            string? previousOrdinal = index == 0 ? null : Children[index - 1]?.Ordinal;
            child.Ordinal = ComputeHierarchicalOrdinal(MaxChildren, ChildCount, Ordinal, previousOrdinal);
        }

        private static string ComputeHierarchicalOrdinal(
            int maxCount,
            int actualCount,
            string parentOrdinal,
            string? previousOrdinal)
        {
            FluidAssert.That(maxCount <= 16 && actualCount <= maxCount, 0x3f0 /* count must be less than max, and max must be 16 or less */);
            int ordinalWidth = 1 << (maxCount - actualCount);
            if (previousOrdinal is null)
            {
                return string.Concat(parentOrdinal, (char)(ordinalWidth - 1));
            }

            char previousLocalOrdinal = previousOrdinal[previousOrdinal.Length - 1];
            return string.Concat(parentOrdinal, (char)(previousLocalOrdinal + ordinalWidth));
        }

        private static void ValidateChildIndex(int index)
        {
            if (index < 0 || index >= MaxChildren)
            {
                throw new System.ArgumentOutOfRangeException(nameof(index), "Child index must be within the block capacity.");
            }
        }
    }
}
