#nullable enable

using System.Linq;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;
using static Microsoft.Office.Web.Fluid.Tests.PortedTestUtilities;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class MergeTreeZamboniTestsFromTS
	{
		// Ported from packages/dds/merge-tree/src/test/mergeTree.zamboni.spec.ts:45 — "zamboni with no segments to scour"
		[Fact]
		public void ZamboniWithNoSegmentsToScour_LeavesTreeUnchanged()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.SetMinSeq(10);
			tree.ZamboniSegments();

			Assert.Equal("hello world", tree.GetText());
			Assert.Single(tree.WalkAllSegments());
		}

		// Ported from packages/dds/merge-tree/src/test/mergeTree.zamboni.spec.ts:54 — "zamboni with one segment to scour"
		[Fact]
		public void ZamboniWithOneSegmentToScour_RemovesStableTombstone()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");
			tree.MarkRangeRemoved(0, 5, refSeq: 1, seq: 2, clientId: "remote");

			tree.SetMinSeq(3);
			tree.ZamboniSegments();

			Assert.Equal(" world", tree.GetText());
			Assert.DoesNotContain(tree.WalkAllSegments(), segment => segment is TextSegment textSegment && textSegment.Text == "hello");
		}

		// Ported from packages/dds/merge-tree/src/test/mergeTree.zamboni.spec.ts:66 — "zamboni with many segments to scour"
		[Fact]
		public void ZamboniWithManySegmentsToScour_RemovesAllStableTombstones()
		{
			MergeTreeModel tree = CreateSingleCharacterTree("abcdefghijklmnopqrst");
			tree.MarkRangeRemoved(0, 10, refSeq: 20, seq: 30, clientId: "remote");

			tree.SetMinSeq(31);
			tree.ZamboniSegments();

			Assert.Equal("klmnopqrst", tree.GetText());
			Assert.All(tree.WalkAllSegments(), segment => Assert.Empty(segment.RemoveStamps));
		}
	}
}
