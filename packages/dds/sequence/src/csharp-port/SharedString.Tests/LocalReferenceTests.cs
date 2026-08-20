// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------

using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class LocalReferenceTests
	{
		[Fact]
		public void CreateReferencePosition_AtValidPosition_ReturnsReference()
		{
			MergeTreeModel tree = CreateTree("hello");
			PropertySet properties = new()
			{
				["kind"] = "test",
			};

			LocalReferencePosition reference = tree.CreateReferencePosition(
				2,
				ReferenceType.Simple,
				SlidingPreference.Forward,
				properties);

			Assert.False(reference.IsDetached);
			Assert.Equal(ReferenceType.Simple, reference.RefType);
			Assert.Equal(SlidingPreference.Forward, reference.SlidingPreference);
			Assert.Same(properties, reference.Properties);
		}

		[Fact]
		public void GetPositionOfReference_ReturnsCreatedPosition()
		{
			MergeTreeModel tree = CreateTree("hello");

			LocalReferencePosition reference = tree.CreateReferencePosition(3, ReferenceType.Simple);

			Assert.Equal(3, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void RemoveReferencePosition_ReturnsTrue_RemovesFromSegment()
		{
			MergeTreeModel tree = CreateTree("hello");
			LocalReferencePosition reference = tree.CreateReferencePosition(3, ReferenceType.Simple);

			bool removed = tree.RemoveReferencePosition(reference);

			Assert.True(removed);
			Assert.True(reference.IsDetached);
			Assert.Equal(MergeTreeModel.DetachedReferencePosition, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void RemoveReferencePosition_OnAlreadyDetached_ReturnsFalse()
		{
			MergeTreeModel tree = CreateTree("hello");
			LocalReferencePosition reference = tree.CreateReferencePosition(3, ReferenceType.Simple);
			tree.RemoveReferencePosition(reference);

			bool removedAgain = tree.RemoveReferencePosition(reference);

			Assert.False(removedAgain);
		}

		[Fact]
		public void Insert_BeforeReference_ShiftsPositionRight()
		{
			MergeTreeModel tree = CreateTree("hello");
			LocalReferencePosition reference = tree.CreateReferencePosition(3, ReferenceType.Simple);

			InsertText(tree, 0, "XX");

			Assert.Equal(5, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void Insert_AfterReference_LeavesPositionUnchanged()
		{
			MergeTreeModel tree = CreateTree("hello");
			LocalReferencePosition reference = tree.CreateReferencePosition(3, ReferenceType.Simple);

			InsertText(tree, 5, "XX");

			Assert.Equal(3, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void Insert_AtReferencePosition_ReferenceStaysAtOldContent()
		{
			MergeTreeModel tree = CreateTree("hello");
			LocalReferencePosition reference = tree.CreateReferencePosition(
				3,
				ReferenceType.Simple,
				SlidingPreference.Forward);

			InsertText(tree, 3, "XX");

			Assert.Equal(5, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void Insert_IntoContainingSegment_SplitsCorrectly()
		{
			MergeTreeModel tree = CreateTree("hello world");
			LocalReferencePosition reference = tree.CreateReferencePosition(5, ReferenceType.Simple);

			InsertText(tree, 5, "X");

			Assert.Equal(6, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void MarkRemoved_ContainingReference_WithForwardSlidingPref_SlidesToNext()
		{
			MergeTreeModel tree = CreateTree("hello world");
			LocalReferencePosition reference = tree.CreateReferencePosition(
				3,
				ReferenceType.SlideOnRemove,
				SlidingPreference.Forward);

			RemoveRange(tree, 0, 5);

			Assert.False(reference.IsDetached);
			Assert.Equal(0, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void MarkRemoved_ContainingReference_WithStayOnRemovePref_StaysOnRemovedSegment()
		{
			MergeTreeModel tree = CreateTree("hello world");
			LocalReferencePosition reference = tree.CreateReferencePosition(
				3,
				ReferenceType.StayOnRemove,
				SlidingPreference.Forward);

			RemoveRange(tree, 0, 5);

			// TS-parity: StayOnRemove references remain attached to removed segments (Finding L2).
			Assert.False(reference.IsDetached);
			Assert.Equal(0, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void MarkRemoved_AdjacentSegment_LeavesReferenceUnchanged()
		{
			MergeTreeModel tree = new();
			InsertText(tree, 0, "hello");
			InsertText(tree, 5, " world");
			LocalReferencePosition reference = tree.CreateReferencePosition(3, ReferenceType.Simple);

			RemoveRange(tree, 7, 11);

			Assert.Equal(3, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void MarkRemoved_BeforeReference_ShiftsPositionLeft()
		{
			MergeTreeModel tree = CreateTree("hello world");
			LocalReferencePosition reference = tree.CreateReferencePosition(8, ReferenceType.Simple);

			RemoveRange(tree, 0, 6);

			Assert.Equal(2, tree.GetPositionOfReference(reference));
		}

		private static MergeTreeModel CreateTree(string text)
		{
			MergeTreeModel tree = new();
			InsertText(tree, 0, text);
			return tree;
		}

		private static void InsertText(MergeTreeModel tree, int position, string text)
		{
			tree.InsertSegments(
				position,
				new ISegment[] { new TextSegment(text) },
				MergeTreeModel.UnassignedSequenceNumber,
				MergeTreeModel.UnassignedSequenceNumber,
				"client");
		}

		private static void RemoveRange(MergeTreeModel tree, int start, int end)
		{
			tree.MarkRangeRemoved(
				start,
				end,
				MergeTreeModel.UnassignedSequenceNumber,
				MergeTreeModel.UnassignedSequenceNumber,
				"client");
		}
	}
}
