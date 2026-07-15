#nullable enable

using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;
using static Microsoft.Office.Web.Fluid.Tests.PortedTestUtilities;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class ObliterateConcurrentTestsFromTS
	{
		// Ported from packages/dds/merge-tree/src/test/obliterate.spec.ts:38 — "removes text"
		[Fact]
		public void Obliterate_RemovesText()
		{
			SharedString sharedString = CreateSharedStringWithText("hello world");

			sharedString.ObliterateRange(0, sharedString.GetLength());

			Assert.Equal(string.Empty, sharedString.GetText());
		}

		// Ported from packages/dds/merge-tree/src/test/obliterate.spec.ts:44 — "removes text for obliterate then insert"
		[Fact]
		public void ConcurrentObliterateThenInsert_InsertedTextInRangeIsEaten()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.ObliterateRange(0, tree.GetLength(), refSeq: 1, seq: 2, clientId: "obliterator");
			tree.InsertSegments(1, SegmentArray("more "), refSeq: 1, seq: 3, clientId: "inserter");

			Assert.Equal(string.Empty, tree.GetText());
			Assert.Contains(SegmentWithText(tree, "more ").RemoveStamps, stamp => stamp.Seq == 2);
		}

		// Ported from packages/dds/merge-tree/src/test/obliterate.spec.ts:98 — "does not expand to include text inserted at start"
		[Fact]
		public void EndpointBehavior_InsertAtStartBoundaryNotIncluded()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.ObliterateRange(5, tree.GetLength(), refSeq: 1, seq: 2, clientId: "obliterator");
			tree.InsertSegments(5, SegmentArray("XXX"), refSeq: 1, seq: 3, clientId: "inserter");

			Assert.Equal("helloXXX", tree.GetText());
		}

		// Ported from packages/dds/merge-tree/src/test/obliterate.partialLength.spec.ts:72 — "correctly applies local remove after local obliterate"
		[Fact]
		public void PartialLength_LocalRemoveAfterLocalObliterate_HasZeroFinalLength()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.ObliterateRange(0, "hello ".Length, refSeq: 1, seq: MergeTreeModel.UnassignedSequenceNumber, clientId: "local");
			tree.MarkRangeRemoved(0, "world".Length, refSeq: MergeTreeModel.UnassignedSequenceNumber, seq: MergeTreeModel.UnassignedSequenceNumber, clientId: "local");

			Assert.Equal(0, tree.GetLength(1, "local"));
			Assert.Equal(string.Empty, tree.GetText());
		}

		// Ported from packages/dds/merge-tree/src/test/obliterate.partialLength.spec.ts:133 — "passes for local remove and remote obliterate"
		[Fact]
		public void PartialLength_OverlappingLocalRemoveAndRemoteObliterate_RemainsRemovedOnce()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.MarkRangeRemoved(0, "hello ".Length, refSeq: MergeTreeModel.UnassignedSequenceNumber, seq: MergeTreeModel.UnassignedSequenceNumber, clientId: "local");
			tree.ObliterateRange(0, "hello ".Length, refSeq: 1, seq: 2, clientId: "remote");

			Assert.Equal("world", tree.GetText());
			Assert.Equal(5, tree.GetLength(2, "local"));
		}

		// Ported from packages/dds/merge-tree/src/test/obliterate.partialLength.spec.ts:173 — "passes for remote remove and local obliterate"
		[Fact]
		public void PartialLength_OverlappingRemoteRemoveAndLocalObliterate_RemainsRemovedOnce()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.ObliterateRange(0, "hello ".Length, refSeq: MergeTreeModel.UnassignedSequenceNumber, seq: MergeTreeModel.UnassignedSequenceNumber, clientId: "local");
			tree.MarkRangeRemoved(0, "hello ".Length, refSeq: 1, seq: 2, clientId: "remote");

			Assert.Equal("world", tree.GetText());
			Assert.Equal(5, tree.GetLength(2, "remote"));
		}

		// Ported from packages/dds/merge-tree/src/test/obliterate.partialLength.spec.ts:249 — "passes for local obliterate and remote obliterate"
		[Fact]
		public void PartialLength_OverlappingLocalAndRemoteObliterate_RemainsRemovedOnce()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.ObliterateRange(0, "hello ".Length, refSeq: MergeTreeModel.UnassignedSequenceNumber, seq: MergeTreeModel.UnassignedSequenceNumber, clientId: "local");
			tree.ObliterateRange(0, "hello ".Length, refSeq: 1, seq: 2, clientId: "remote");

			Assert.Equal("world", tree.GetText());
			Assert.Equal(5, tree.GetLength(2, "local"));
		}

		// Ported from packages/dds/merge-tree/src/test/obliterate.partialLength.spec.ts:291 — "obliterates when concurrent insert in middle of string"
		[Fact]
		public void PartialLength_ConcurrentInsertInMiddleOfObliterate_IsObliterated()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.ObliterateRange(0, tree.GetLength(), refSeq: MergeTreeModel.UnassignedSequenceNumber, seq: MergeTreeModel.UnassignedSequenceNumber, clientId: "local");
			tree.InsertSegments("hello".Length, SegmentArray("more "), refSeq: 1, seq: 2, clientId: "remote");

			Assert.Equal(string.Empty, tree.GetText());
			Assert.Equal(0, tree.GetLength(2, "local"));
		}

		// Ported from packages/dds/merge-tree/src/test/obliterate.partialLength.spec.ts:334 — "obliterate does not affect concurrent insert at start of string"
		[Fact]
		public void PartialLength_ConcurrentInsertAtStartOfObliterate_RemainsVisible()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.ObliterateRange(0, tree.GetLength(), refSeq: MergeTreeModel.UnassignedSequenceNumber, seq: MergeTreeModel.UnassignedSequenceNumber, clientId: "local");
			tree.InsertSegments(0, SegmentArray("more "), refSeq: 1, seq: 2, clientId: "remote");

			Assert.Equal("more ", tree.GetText());
			Assert.Equal(5, tree.GetLength(2, "local"));
		}

	}
}
