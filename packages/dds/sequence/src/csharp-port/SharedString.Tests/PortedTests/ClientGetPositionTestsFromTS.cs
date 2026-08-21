#nullable enable

using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;

namespace Microsoft.Office.Web.Fluid.Tests
{
	// Ported from packages/dds/merge-tree/src/test/client.getPosition.spec.ts
	//
	// TS spec covers four getPosition cases: Existing, Deleted (tombstone),
	// Detached (zamboni'd), and Removed (preceding content removed). Two of
	// them already live in PortedTests/PartialLengthPerspectiveTestsFromTS.cs
	// (DetachedSegmentReturnsMinusOne, DeletedSegmentReturnsTombstonePosition).
	// The remaining two — Existing and Removed — round out the coverage of
	// SS-A15 (tombstones report the collapsed position, not -1).
	public sealed class ClientGetPositionTestsFromTS
	{
		// Segment "o" in "hello world" sits at position 4. The whole file
		// uses this segPos + "hello world" fixture.
		private const int _segPos = 4;

		// Ported from client.getPosition.spec.ts — "Existing Segment"
		[Fact]
		public void GetPosition_ExistingSegment_ReturnsCurrentPosition()
		{
			SharedString sharedString = CreateFixture(out ISegment segment);

			int? pos = sharedString.GetPosition(segment);

			Assert.Equal(_segPos, pos);
		}

		// Ported from client.getPosition.spec.ts — "Deleted Segment"
		[Fact]
		public void GetPosition_DeletedSegment_ReturnsTombstonePosition()
		{
			// SS-A15: TS reports the collapsed position — where the segment
			// used to sit before removal — not -1 or null.
			SharedString sharedString = CreateFixture(out ISegment segment);

			sharedString.DeleteText(_segPos, _segPos + 1);

			int? pos = sharedString.GetPosition(segment);
			Assert.Equal(_segPos, pos);
		}

		// Ported from client.getPosition.spec.ts — "Removed Segment"
		[Fact]
		public void GetPosition_PrecedingSegmentRemoved_ReturnsCollapsedPosition()
		{
			// Removing content before the target shifts its collapsed
			// position leftward by the removed length.
			SharedString sharedString = CreateFixture(out ISegment segment);

			sharedString.DeleteText(_segPos - 1, _segPos);

			int? pos = sharedString.GetPosition(segment);
			Assert.Equal(_segPos - 1, pos);
		}

		// Ported from client.getPosition.spec.ts — "Detached Segment"
		// (also covered in PartialLengthPerspectiveTestsFromTS.
		// GetPosition_DetachedSegmentReturnsMinusOne — kept here as a
		// contrast with the tombstone case above so the file is self-
		// contained.)
		[Fact]
		public void GetPosition_DetachedSegment_ReturnsMinusOne()
		{
			MergeTreeModel tree = new();
			ISegment detached = TextSegment.Make("x");

			int? pos = tree.GetPositionOfSegment(detached);
			Assert.Equal(MergeTreeModel.DetachedReferencePosition, pos);
		}

		private static SharedString CreateFixture(out ISegment segment)
		{
			SharedString sharedString = new();
			// TS test inserts one char at a time to create 11 single-char
			// segments; port matches so `getContainingSegment(4)` isolates
			// the "o" segment.
			for (int i = 0; i < "hello world".Length; i++)
			{
				sharedString.InsertText(sharedString.GetLength(), "hello world"[i].ToString());
			}

			segment = sharedString.GetContainingSegment(_segPos)!.Value.segment;
			return sharedString;
		}
	}
}
