#nullable enable

using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;
using TextSegmentModel = Microsoft.Office.Web.Fluid.MergeTree.TextSegment;
using static Microsoft.Office.Web.Fluid.Tests.PortedTestUtilities;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class PartialLengthTestsFromTS
	{
		// Ported from packages/dds/merge-tree/src/test/partialLength.spec.ts — "passes with no additional ops"
		[Fact]
		public void NoAdditionalOps_LengthMatchesInitialText()
		{
			MergeTreeModel tree = CreateAckedTree("hello world!");

			Assert.Equal(12, tree.GetLength(1));
			Assert.Equal(12, tree.GetLength(1, "local"));
			Assert.Equal(12, tree.GetLength(1, "remote"));
		}

		// Ported from packages/dds/merge-tree/src/test/partialLength.spec.ts — "includes length of local insert for local view"
		[Fact]
		public void SingleInsertedElement_LocalInsertVisibleForLocalView()
		{
			MergeTreeModel tree = CreateAckedTree("hello world!");

			tree.InsertSegments(0, SegmentArray("more "), refSeq: 1, seq: 2, clientId: "local");

			Assert.Equal(17, tree.GetLength(2, "local"));
			Assert.Equal(17, tree.GetLength(2));
		}

		// Ported from packages/dds/merge-tree/src/test/partialLength.spec.ts — "includes length of remote insert for local view"
		[Fact]
		public void SingleInsertedElement_RemoteInsertVisibleForLocalView()
		{
			MergeTreeModel tree = CreateAckedTree("hello world!");

			tree.InsertSegments(0, SegmentArray("more "), refSeq: 1, seq: 2, clientId: "remote");

			Assert.Equal(17, tree.GetLength(2, "local"));
			Assert.Equal(12, tree.GetLength(1, "local"));
		}

		// Ported from packages/dds/merge-tree/src/test/partialLength.spec.ts — "includes result of local delete for local view"
		[Fact]
		public void SingleRemovedSegment_LocalDeleteVisibleForLocalView()
		{
			MergeTreeModel tree = CreateAckedTree("hello world!");

			tree.MarkRangeRemoved(0, 12, refSeq: 1, seq: 2, clientId: "local");

			Assert.Equal(0, tree.GetLength(2, "local"));
			Assert.Equal(0, tree.GetLength(1, "local"));
		}

		// Ported from packages/dds/merge-tree/src/test/partialLength.spec.ts — "includes result of remote delete for local view"
		[Fact]
		public void SingleRemovedSegment_RemoteDeleteVisibleForLocalView()
		{
			MergeTreeModel tree = CreateAckedTree("hello world!");

			tree.MarkRangeRemoved(0, 12, refSeq: 1, seq: 2, clientId: "remote");

			Assert.Equal(0, tree.GetLength(2, "local"));
			Assert.Equal(12, tree.GetLength(1, "local"));
		}

		// Ported from packages/dds/merge-tree/src/test/partialLength.spec.ts — "includes lengths from multiple permutations in single tree"
		[Fact]
		public void Aggregation_MultipleInsertPermutationsInSingleTree()
		{
			MergeTreeModel tree = CreateAckedTree("hello world!");

			tree.InsertSegments(0, SegmentArray("1"), refSeq: 1, seq: 2, clientId: "local");
			tree.InsertSegments(0, SegmentArray("2"), refSeq: 2, seq: 3, clientId: "remote");
			tree.InsertSegments(0, SegmentArray("3"), refSeq: 3, seq: 4, clientId: "local");
			tree.InsertSegments(0, SegmentArray("4"), refSeq: 4, seq: 5, clientId: "remote");

			Assert.Equal(16, tree.GetLength(5, "local"));
			Assert.Equal(16, tree.GetLength(5, "remote"));
			Assert.Equal("4321hello world!", tree.GetText());
		}

		// Ported from packages/dds/merge-tree/src/test/partialLength.spec.ts — "is correct for different heights"
		[Fact]
		public void Aggregation_DifferentTreeHeights()
		{
			MergeTreeModel tree = CreateAckedTree("hello world!");
			for (int i = 0; i < 100; i++)
			{
				tree.InsertSegments(0, new ISegment[] { TextSegmentModel.Make("a") }, refSeq: i + 1, seq: i + 2, clientId: "local");
			}

			Assert.Equal(112, tree.GetLength(101, "local"));
			Assert.Equal(112, tree.GetLength(101, "remote"));
			PartialLengths.ValidateBlockPartialLengthInvariants(tree.Root);
		}

		// Ported from packages/dds/merge-tree/src/test/partialLength.spec.ts — "concurrent remote changes are visible to local"
		[Fact]
		public void ConcurrentOverlappingDeletes_RemoteChangesVisibleToLocal()
		{
			MergeTreeModel tree = CreateAckedTree("hello world!");

			tree.MarkRangeRemoved(0, 10, refSeq: 1, seq: 2, clientId: "remote-1");
			tree.MarkRangeRemoved(0, 10, refSeq: 1, seq: 3, clientId: "remote-2");

			Assert.Equal(2, tree.GetLength(2, "local"));
			Assert.Equal(2, tree.GetLength(3, "local"));
			Assert.Equal("d!", tree.GetText());
		}

		// Ported from packages/dds/merge-tree/src/test/partialLength.spec.ts — "concurrent local and remote changes are visible"
		[Fact]
		public void ConcurrentOverlappingDeletes_LocalAndRemoteChangesVisible()
		{
			MergeTreeModel tree = CreateAckedTree("hello world!");

			tree.MarkRangeRemoved(0, 10, refSeq: 1, seq: 2, clientId: "local");
			tree.MarkRangeRemoved(0, 10, refSeq: 1, seq: 3, clientId: "remote");

			Assert.Equal(2, tree.GetLength(2, "local"));
			Assert.Equal(2, tree.GetLength(3, "remote"));
			Assert.Equal("d!", tree.GetText());
		}

	}
}
