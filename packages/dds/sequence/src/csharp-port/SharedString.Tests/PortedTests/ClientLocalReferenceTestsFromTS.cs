#nullable enable

using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;

namespace Microsoft.Office.Web.Fluid.Tests
{
	// Ported from packages/dds/merge-tree/src/test/client.localReference.spec.ts
	//
	// Focused subset covering slide-on-remove semantics — the space where
	// (slide-target all-acked perspective) applies. The port already
	// has LocalReferenceParityTests and LocalReferenceTests; this file adds
	// the specific behavioral scenarios TS's spec calls out.
	public sealed class ClientLocalReferenceTestsFromTS
	{
		// Ported from client.localReference.spec.ts —
		// "Remove segment of non-sliding local reference"
		[Fact]
		public void RemoveSegmentOfNonSlidingReference_DetachesReference()
		{
			MergeTreeModel tree = BuildAckedTree("01234");
			ISegment segAt2 = tree.GetContainingSegment(2, MergeTreeModel.UnassignedSequenceNumber).segment;
			LocalReferencePosition reference = tree.CreateReferencePosition(
				segAt2,
				offset: 0,
				refType: ReferenceType.Simple);

			Assert.Equal(2, tree.GetPositionOfReference(reference));

			tree.MarkRangeRemoved(2, 3, refSeq: 5, seq: 6, clientId: "peer");

			// Non-sliding reference detaches on segment removal.
			Assert.Equal(MergeTreeModel.DetachedReferencePosition, tree.GetPositionOfReference(reference));
		}

		// Ported from client.localReference.spec.ts —
		// "Remove segment of sliding local reference"
		[Fact]
		public void RemoveSegmentOfSlidingReference_KeepsPositionByForwardSlide()
		{
			MergeTreeModel tree = BuildAckedTree("01234");
			ISegment segAt2 = tree.GetContainingSegment(2, MergeTreeModel.UnassignedSequenceNumber).segment;
			LocalReferencePosition reference = tree.CreateReferencePosition(
				segAt2,
				offset: 0,
				refType: ReferenceType.SlideOnRemove,
				slidingPreference: SlidingPreference.Forward);

			Assert.Equal(2, tree.GetPositionOfReference(reference));

			// Remove the segment holding the reference. Under forward-slide,
			// the reference lands on the next visible segment — position 2
			// stays valid because "3" (originally at pos 3) has now slid to
			// pos 2.
			tree.MarkRangeRemoved(2, 3, refSeq: 5, seq: 6, clientId: "peer");

			Assert.Equal(2, tree.GetPositionOfReference(reference));
		}

		// Ported from client.localReference.spec.ts —
		// "Remove segments to end with sliding local reference"
		[Fact]
		public void RemoveSegmentToEnd_SlidingReferenceLandsAtLastVisible()
		{
			MergeTreeModel tree = BuildAckedTree("01234");
			ISegment segAt3 = tree.GetContainingSegment(3, MergeTreeModel.UnassignedSequenceNumber).segment;
			LocalReferencePosition reference = tree.CreateReferencePosition(
				segAt3,
				offset: 0,
				refType: ReferenceType.SlideOnRemove,
				slidingPreference: SlidingPreference.Forward);

			// Remove from position 3 through end. Forward-slide has no
			// successor; falls back to previous visible position.
			tree.MarkRangeRemoved(3, 5, refSeq: 5, seq: 6, clientId: "peer");

			int? pos = tree.GetPositionOfReference(reference);
			Assert.NotNull(pos);
			Assert.NotEqual(MergeTreeModel.DetachedReferencePosition, pos!.Value);
			// Reference lands within the remaining visible content [0..3).
			Assert.InRange(pos.Value, 0, 3);
		}

		// Ported from client.localReference.spec.ts —
		// "References can have offsets on removed segment"
		[Fact]
		public void ReferenceOffset_PreservedAcrossSlideOnRemove()
		{
			MergeTreeModel tree = BuildAckedTree("01234");
			ISegment segAt2 = tree.GetContainingSegment(2, MergeTreeModel.UnassignedSequenceNumber).segment;
			// Reference at offset 0 within the "2" segment == position 2.
			LocalReferencePosition reference = tree.CreateReferencePosition(
				segAt2,
				offset: 0,
				refType: ReferenceType.SlideOnRemove,
				slidingPreference: SlidingPreference.Forward);

			// Remove segment "2"; reference slides forward to the "3" segment.
			tree.MarkRangeRemoved(2, 3, refSeq: 5, seq: 6, clientId: "peer");

			int? pos = tree.GetPositionOfReference(reference);
			Assert.Equal(2, pos);
		}

		// Ported from client.localReference.spec.ts —
		// "Remove all segments with sliding local reference"
		[Fact]
		public void RemoveAllSegments_SlidingReferenceDetaches()
		{
			MergeTreeModel tree = BuildAckedTree("01234");
			ISegment segAt2 = tree.GetContainingSegment(2, MergeTreeModel.UnassignedSequenceNumber).segment;
			LocalReferencePosition reference = tree.CreateReferencePosition(
				segAt2,
				offset: 0,
				refType: ReferenceType.SlideOnRemove,
				slidingPreference: SlidingPreference.Forward);

			// Remove EVERYTHING — no slide target available in any direction.
			tree.MarkRangeRemoved(0, 5, refSeq: 5, seq: 6, clientId: "peer");

			Assert.Equal(MergeTreeModel.DetachedReferencePosition, tree.GetPositionOfReference(reference));
		}

		private static MergeTreeModel BuildAckedTree(string text)
		{
			MergeTreeModel tree = new();
			long seq = 1;
			for (int i = 0; i < text.Length; i++)
			{
				tree.InsertSegments(
					tree.GetLength(),
					new ISegment[] { TextSegment.Make(text[i].ToString()) },
					refSeq: seq - 1,
					seq: seq,
					clientId: "seed");
				seq++;
			}

			return tree;
		}
	}
}
