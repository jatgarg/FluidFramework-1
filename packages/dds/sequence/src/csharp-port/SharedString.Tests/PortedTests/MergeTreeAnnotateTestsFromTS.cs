#nullable enable

using System.Linq;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;
using TextSegmentModel = Microsoft.Office.Web.Fluid.MergeTree.TextSegment;
using static Microsoft.Office.Web.Fluid.Tests.PortedTestUtilities;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class MergeTreeAnnotateTestsFromTS
	{
		// Ported from packages/dds/merge-tree/src/test/mergeTree.annotate.spec.ts — "remote"
		[Fact]
		public void RemoteAnnotate_AppliesPropertyToRequestedRange()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.AnnotateRange(2, 7, new PropertySet() { ["propertySource"] = "remote" }, refSeq: 1, seq: 2, clientId: "remote");

			ISegment segment = tree.GetContainingSegment(2, MergeTreeModel.UnassignedSequenceNumber).segment;
			Assert.Equal("remote", Assert.IsType<string>(segment.Properties!["propertySource"]));
			Assert.Equal("llo w", Assert.IsType<TextSegmentModel>(segment).Text);
		}

		// Ported from packages/dds/merge-tree/src/test/mergeTree.annotate.spec.ts — "local"
		[Fact]
		public void LocalAnnotate_AppliesUnsequencedPropertyToRequestedRange()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.AnnotateRange(
				2,
				7,
				new PropertySet() { ["propertySource"] = "local" },
				refSeq: MergeTreeModel.UnassignedSequenceNumber,
				seq: MergeTreeModel.UnassignedSequenceNumber,
				clientId: "local");

			ISegment segment = tree.GetContainingSegment(2, MergeTreeModel.UnassignedSequenceNumber, "local").segment;
			Assert.Equal("local", Assert.IsType<string>(segment.Properties!["propertySource"]));
		}

		// Ported from packages/dds/merge-tree/src/test/mergeTree.annotate.spec.ts — "unsequenced local after unsequenced local"
		[Fact]
		public void UnsequencedLocalAfterUnsequencedLocal_MergesProperties()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.AnnotateRange(2, 7, new PropertySet() { ["propertySource"] = "local" }, MergeTreeModel.UnassignedSequenceNumber, MergeTreeModel.UnassignedSequenceNumber, "local");
			tree.AnnotateRange(2, 7, new PropertySet() { ["secondProperty"] = "local" }, MergeTreeModel.UnassignedSequenceNumber, MergeTreeModel.UnassignedSequenceNumber, "local");

			PropertySet props = tree.GetContainingSegment(2, MergeTreeModel.UnassignedSequenceNumber, "local").segment.Properties!;
			Assert.Equal("local", Assert.IsType<string>(props["propertySource"]));
			Assert.Equal("local", Assert.IsType<string>(props["secondProperty"]));
		}

		// Ported from packages/dds/merge-tree/src/test/mergeTree.annotate.spec.ts — "unsequenced local split"
		[Fact]
		public void UnsequencedLocalSplit_CopiesPropertiesToSplitSegment()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");
			tree.AnnotateRange(2, 7, new PropertySet() { ["propertySource"] = "local" }, MergeTreeModel.UnassignedSequenceNumber, MergeTreeModel.UnassignedSequenceNumber, "local");

			tree.InsertSegments(4, new ISegment[] { TextSegmentModel.Make("X") }, refSeq: MergeTreeModel.UnassignedSequenceNumber, seq: MergeTreeModel.UnassignedSequenceNumber, clientId: "local");

			ISegment left = tree.GetContainingSegment(2, MergeTreeModel.UnassignedSequenceNumber, "local").segment;
			ISegment right = tree.GetContainingSegment(5, MergeTreeModel.UnassignedSequenceNumber, "local").segment;
			Assert.Equal("local", Assert.IsType<string>(left.Properties!["propertySource"]));
			Assert.Equal("local", Assert.IsType<string>(right.Properties!["propertySource"]));
		}

		// Ported from packages/dds/merge-tree/src/test/mergeTree.annotate.spec.ts — "unsequenced local before remote"
		[Fact]
		public void UnsequencedLocalBeforeRemote_RemoteAddsIndependentPropertyWithoutOverwritingLocal()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");
			tree.AnnotateRange(2, 7, new PropertySet() { ["propertySource"] = "local" }, MergeTreeModel.UnassignedSequenceNumber, MergeTreeModel.UnassignedSequenceNumber, "local");

			tree.AnnotateRange(2, 7, new PropertySet() { ["remoteProperty"] = 1 }, refSeq: 1, seq: 2, clientId: "remote");

			PropertySet props = tree.GetContainingSegment(2, MergeTreeModel.UnassignedSequenceNumber, "local").segment.Properties!;
			Assert.Equal("local", Assert.IsType<string>(props["propertySource"]));
			Assert.Equal(1, Assert.IsType<int>(props["remoteProperty"]));
		}

		// Ported from packages/dds/merge-tree/src/test/mergeTree.annotate.spec.ts — "sequenced local before remote"
		[Fact]
		public void SequencedLocalBeforeRemote_RemoteOverwritesSharedProperty()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");
			tree.AnnotateRange(2, 7, new PropertySet() { ["propertySource"] = "local" }, refSeq: 1, seq: 2, clientId: "local");

			tree.AnnotateRange(2, 7, new PropertySet() { ["propertySource"] = "remote", ["remoteProperty"] = 1 }, refSeq: 2, seq: 3, clientId: "remote");

			PropertySet props = tree.GetContainingSegment(2, MergeTreeModel.UnassignedSequenceNumber).segment.Properties!;
			Assert.Equal("remote", Assert.IsType<string>(props["propertySource"]));
			Assert.Equal(1, Assert.IsType<int>(props["remoteProperty"]));
		}
	}
}
