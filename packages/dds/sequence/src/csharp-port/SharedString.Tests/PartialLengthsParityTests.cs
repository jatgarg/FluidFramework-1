// -----------------------------------------------------------------------------
// P-series partial-length parity regressions for SharedString C# port.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeBlockModel = Microsoft.Office.Web.Fluid.MergeTree.MergeBlock;
using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;
using TextSegmentModel = Microsoft.Office.Web.Fluid.MergeTree.TextSegment;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class PartialLengthsParityTests
	{
		[Fact]
		public void P1_LengthQueriesMatchNaiveWalkForMixedOperations()
		{
			MergeTreeModel tree = CreateSingleCharacterSegmentTree(10);
			tree.InsertSegments(5, SegmentArray("I"), refSeq: 10, seq: 11, clientId: "client-a");
			tree.MarkRangeRemoved(2, 7, refSeq: 11, seq: 12, clientId: "client-b");
			tree.ObliterateRange(4, 9, refSeq: 11, seq: 14, clientId: "client-c");

			foreach (long refSeq in new long[] { 10, 11, 12, 13, 14 })
			{
				Assert.Equal(NaiveLengthAt(tree, refSeq), tree.GetLength(refSeq));
			}
		}

		[Fact]
		public void P1_BlockPartialLengthInvariantsHoldAfterZamboni()
		{
			MergeTreeModel tree = CreateSingleCharacterSegmentTree(24);
			tree.MarkRangeRemoved(4, 20, refSeq: 24, seq: 30, clientId: "remover");

			AdvanceMinSeq(tree, 30);
			tree.ZamboniSegments();

			AssertBlockInvariants(tree.Root);
			Assert.Equal(8, tree.GetLength());
		}

		[Fact]
		public void P2_LocalInsertVisibleOnlyToOwningClientPerspective()
		{
			MergeTreeModel tree = CreateAckedTree("abc");

			tree.InsertSegments(1, SegmentArray("X"), refSeq: 1, seq: MergeTreeModel.UnassignedSequenceNumber, clientId: "client-a");

			Assert.Equal(4, tree.GetLength());
			Assert.Equal(3, tree.GetLength(1));
			Assert.Equal(4, tree.GetLength(1, "client-a"));
			Assert.Equal(3, tree.GetLength(1, "client-b"));
		}

		[Fact]
		public void P2_LocalRemoveVisibleOnlyToOwningClientPerspective()
		{
			MergeTreeModel tree = CreateAckedTree("abc");

			tree.MarkRangeRemoved(0, 3, refSeq: 1, seq: MergeTreeModel.UnassignedSequenceNumber, clientId: "client-a");

			Assert.Equal(0, tree.GetLength());
			Assert.Equal(3, tree.GetLength(1));
			Assert.Equal(0, tree.GetLength(1, "client-a"));
			Assert.Equal(3, tree.GetLength(1, "client-b"));
		}

		[Fact]
		public void P2_SameClientSequencedInsertIsVisibleBeforeRefSeq()
		{
			MergeTreeModel tree = CreateAckedTree("abc");

			tree.InsertSegments(1, SegmentArray("X"), refSeq: 1, seq: 10, clientId: "client-a");

			Assert.Equal(3, tree.GetLength(1));
			Assert.Equal(4, tree.GetLength(1, "client-a"));
			Assert.Equal(3, tree.GetLength(1, "client-b"));
		}

		[Fact]
		public void P3_HistoricalLengthUnaffectedByLaterRemove()
		{
			MergeTreeModel tree = CreateAckedTree("abcde");

			tree.MarkRangeRemoved(1, 4, refSeq: 1, seq: 10, clientId: "remover");

			Assert.Equal(5, tree.GetLength(9));
			Assert.Equal(2, tree.GetLength(10));
			Assert.Equal(2, tree.GetLength());
		}

		[Fact]
		public void P3_HistoricalLengthTracksLaterInsertAndRemoveSeparately()
		{
			MergeTreeModel tree = CreateAckedTree("abc");
			tree.InsertSegments(1, SegmentArray("X"), refSeq: 1, seq: 5, clientId: "client-a");
			tree.MarkRangeRemoved(1, 3, refSeq: 5, seq: 10, clientId: "client-b");

			Assert.Equal(3, tree.GetLength(1));
			Assert.Equal(4, tree.GetLength(5));
			Assert.Equal(2, tree.GetLength(10));
		}

		[Fact]
		public void P3_SameClientSequencedRemoveAppliesBeforeRefSeqForThatClient()
		{
			MergeTreeModel tree = CreateAckedTree("abcd");

			tree.MarkRangeRemoved(1, 3, refSeq: 1, seq: 10, clientId: "client-a");

			Assert.Equal(4, tree.GetLength(1));
			Assert.Equal(2, tree.GetLength(1, "client-a"));
			Assert.Equal(4, tree.GetLength(1, "client-b"));
		}

		[Fact]
		public void P5_LocalUnackedInsertSplitByAnnotationKeepsLengthAdjustment()
		{
			MergeTreeModel tree = new();
			tree.InsertSegments(
				0,
				SegmentArray("abcd"),
				refSeq: MergeTreeModel.UnassignedSequenceNumber,
				seq: MergeTreeModel.UnassignedSequenceNumber,
				clientId: "client-a");

			tree.AnnotateRange(
				1,
				3,
				new PropertySet() { ["style"] = "middle" },
				refSeq: MergeTreeModel.UnassignedSequenceNumber,
				seq: MergeTreeModel.UnassignedSequenceNumber,
				clientId: "client-a");

			Assert.Equal(3, tree.WalkAllSegments().Count());
			Assert.Equal(4, tree.GetLength(MergeTreeModel.UnassignedSequenceNumber, "client-a"));
			Assert.Equal(0, tree.GetLength(1, "client-b"));
			Assert.Equal(4, tree.GetLength(1, "client-a"));
		}

		[Fact]
		public void P5_LocalUnackedRemoveSplitByRemoteAnnotationKeepsLengthAdjustment()
		{
			MergeTreeModel tree = CreateAckedTree("abcd");
			tree.MarkRangeRemoved(
				0,
				4,
				refSeq: 1,
				seq: MergeTreeModel.UnassignedSequenceNumber,
				clientId: "client-a");

			tree.AnnotateRange(1, 3, new PropertySet() { ["remote"] = true }, refSeq: 1, seq: 10, clientId: "client-b");

			Assert.Equal(3, tree.WalkAllSegments().Count());
			Assert.Equal(0, tree.GetLength(1, "client-a"));
			Assert.Equal(4, tree.GetLength(1, "client-b"));
		}

		[Fact]
		public void P6_InvariantsHoldAfterBlockSplitAndOverlappingRemove()
		{
			MergeTreeModel tree = CreateSingleCharacterSegmentTree(20);

			tree.MarkRangeRemoved(3, 14, refSeq: 20, seq: 30, clientId: "client-a");
			tree.MarkRangeRemoved(8, 18, refSeq: 20, seq: 31, clientId: "client-b");

			AssertBlockInvariants(tree.Root);
			Assert.Equal(5, tree.GetLength(31));
		}

		[Fact]
		public void P6_InvariantsHoldAfterOverlappingObliterateAndInsertOnInsert()
		{
			MergeTreeModel tree = CreateSingleCharacterSegmentTree(20);
			tree.ObliterateRange(5, 15, refSeq: 20, seq: 30, clientId: "client-a");

			tree.InsertSegments(10, SegmentArray("X"), refSeq: 20, seq: 31, clientId: "client-b");

			AssertBlockInvariants(tree.Root);
			Assert.Equal(10, tree.GetLength());
			Assert.Equal(20, tree.GetLength(20));
		}

		[Fact]
		public void P6_InvariantsHoldAfterStructureRebuildFromSegments()
		{
			MergeTreeModel tree = CreateSingleCharacterSegmentTree(18);
			tree.ObliterateRange(2, 8, refSeq: 18, seq: 30, clientId: "client-a");
			tree.MarkRangeRemoved(10, 15, refSeq: 18, seq: 31, clientId: "client-b");
			List<ISegment> segments = tree.WalkAllSegments().ToList();

			tree.RebuildFromSegments(segments);

			AssertBlockInvariants(tree.Root);
			Assert.Equal(7, tree.GetLength());
		}

		private static MergeTreeModel CreateAckedTree(string text)
		{
			MergeTreeModel tree = new();
			tree.InsertSegments(0, SegmentArray(text), MergeTreeModel.UnassignedSequenceNumber, seq: 1, clientId: "seed");
			return tree;
		}

		private static MergeTreeModel CreateSingleCharacterSegmentTree(int count)
		{
			MergeTreeModel tree = new();
			for (int i = 0; i < count; i++)
			{
				tree.InsertSegments(i, SegmentArray("x"), refSeq: i, seq: i + 1, clientId: "seed");
			}

			return tree;
		}

		private static ISegment[] SegmentArray(string text)
		{
			return new ISegment[] { TextSegmentModel.Make(text) };
		}

		private static int NaiveLengthAt(MergeTreeModel tree, long refSeq)
		{
			int length = 0;
			foreach (ISegment segment in tree.WalkAllSegments())
			{
				if (segment.Seq != MergeTreeModel.UnassignedSequenceNumber && segment.Seq <= refSeq
					&& !segment.RemoveStamps.Any(stamp => stamp.Seq != MergeTreeModel.UnassignedSequenceNumber && stamp.Seq <= refSeq))
				{
					length += segment.CachedLength;
				}
			}

			return length;
		}

		private static void AssertBlockInvariants(MergeBlockModel block)
		{
			PartialLengths.ValidateBlockPartialLengthInvariants(block);
			for (int i = 0; i < block.ChildCount; i++)
			{
				if (block.Children[i] is MergeBlockModel childBlock)
				{
					AssertBlockInvariants(childBlock);
				}
			}
		}

		private static void AdvanceMinSeq(MergeTreeModel tree, long minSeq)
		{
			if (tree.CurrentSeq < minSeq)
			{
				tree.CurrentSeq = minSeq;
			}

			tree.SetMinSeq(minSeq);
		}
	}
}
