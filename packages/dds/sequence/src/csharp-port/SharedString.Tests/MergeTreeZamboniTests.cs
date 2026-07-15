// -----------------------------------------------------------------------------
// Zamboni merge-tree cleanup tests for SharedString C# port.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeBlockModel = Microsoft.Office.Web.Fluid.MergeTree.MergeBlock;
using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;
using TextSegmentModel = Microsoft.Office.Web.Fluid.MergeTree.TextSegment;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class MergeTreeZamboniTests
	{
		[Fact]
		public void Zamboni_DropsTombstones_AfterMinSeqAdvance()
		{
			MergeTreeModel tree = CreateTreeWithSegments(
				("left", null),
				("drop", null),
				("right", null));
			tree.MarkRangeRemoved(4, 8, tree.CurrentSeq, seq: 10, clientId: "remover");
			int countBefore = CountSegments(tree);

			AdvanceMinSeq(tree, 15);
			tree.ZamboniSegments();

			Assert.Equal(3, countBefore);
			Assert.Equal(1, CountSegments(tree));
			Assert.Equal("leftright", tree.GetText());
			Assert.All(tree.WalkAllSegments(), segment => Assert.Null(segment.RemovedSeq));
		}

		[Fact]
		public void Zamboni_KeepsTombstones_BeforeMinSeqAdvance()
		{
			MergeTreeModel tree = CreateTreeWithSegments(
				("left", Property("side", "left")),
				("drop", null),
				("right", Property("side", "right")));
			tree.MarkRangeRemoved(4, 8, tree.CurrentSeq, seq: 10, clientId: "remover");

			AdvanceMinSeq(tree, 5);
			tree.ZamboniSegments();

			Assert.Equal(3, CountSegments(tree));
			Assert.Contains(tree.WalkAllSegments(), segment => segment.RemovedSeq == 10);
			Assert.Equal("leftright", tree.GetText());
		}

		[Fact]
		public void Zamboni_CoalescesAdjacentSameProps()
		{
			MergeTreeModel tree = CreateTreeWithSegments(("hello", null), (" world", null));

			AdvanceMinSeq(tree, 10);
			tree.ZamboniSegments();

			TextSegmentModel segment = Assert.Single(tree.WalkAllSegments().Cast<TextSegmentModel>());
			Assert.Equal("hello world", segment.Text);
			Assert.Equal("hello world", tree.GetText());
		}

		[Fact]
		public void Zamboni_DoesNotCoalesceDifferentProps()
		{
			MergeTreeModel tree = CreateTreeWithSegments(
				("bold", Property("bold", true)),
				(" normal", null));

			AdvanceMinSeq(tree, 10);
			tree.ZamboniSegments();

			Assert.Equal(2, CountSegments(tree));
			Assert.Equal("bold normal", tree.GetText());
		}

		[Fact]
		public void Zamboni_SlidesLocalReference()
		{
			MergeTreeModel tree = CreateTreeWithSegments(("keep", null), ("drop", null), ("tail", null));
			ISegment droppedSegment = GetRequiredSegment(tree, 5);
			tree.MarkRangeRemoved(4, 8, tree.CurrentSeq, seq: 10, clientId: "remover");
			LocalReferencePosition reference = AttachReference(
				droppedSegment,
				offset: 1,
				SlidingPreference.Forward);

			AdvanceMinSeq(tree, 15);
			tree.ZamboniSegments();

			Assert.False(reference.IsDetached);
			Assert.Equal(4, tree.GetPositionOfReference(reference));
			Assert.DoesNotContain(tree.WalkAllSegments(), segment => ReferenceEquals(segment, droppedSegment));
		}

		[Fact]
		public void Zamboni_SlidesIntervalEndpoint()
		{
			Client client = CreateClientWithAckedSegments("keep", "drop", "tail");
			ISegment droppedSegment = GetRequiredSegment(client.MergeTree, 5);
			IntervalCollection collection = client.GetOrCreateIntervalCollection("zamboni", IntervalType.SlideOnRemove);
			SequenceInterval interval = collection.Add(5, 5, intervalId: "interval-1");
			IMergeTreeRemoveMsg remove = client.RemoveText(4, 8);
			client.ApplyOp(remove, seq: 10, refSeq: 3, clientId: "doc");
			LocalReferencePosition tombstoneEndpoint = AttachReference(
				droppedSegment,
				offset: 1,
				SlidingPreference.Forward);
			interval.SetStart(tombstoneEndpoint);

			client.SetCollaborationWindow(minSeq: 15, currentSeq: 15);

			SequenceInterval? retrieved = collection.GetIntervalById("interval-1");
			Assert.Same(interval, retrieved);
			Assert.False(interval.Start.IsDetached);
			Assert.Equal(4, interval.StartPosition);
		}

		[Fact]
		public void Zamboni_ObliteratedSegments_Dropped()
		{
			MergeTreeModel tree = CreateTreeWithSegments(
				("left", Property("side", "left")),
				("drop", null),
				("right", Property("side", "right")));
			tree.ObliterateRange(4, 8, tree.CurrentSeq, seq: 10, clientId: "obliterator");

			AdvanceMinSeq(tree, 15);
			tree.ZamboniSegments();

			Assert.Equal(2, CountSegments(tree));
			Assert.Equal("leftright", tree.GetText());
			Assert.All(tree.WalkAllSegments(), segment => Assert.Null(segment.ObliteratedSeq));
		}

		[Fact]
		public void Zamboni_ConcurrentOpBeforeZamboni_NoCrash()
		{
			Client client = CreateClientWithAckedSegments("hello world");
			client.ApplyOp(
				new MergeTreeRemoveMsg() { Pos1 = 0, Pos2 = 5 },
				seq: 10,
				refSeq: 1,
				clientId: "remote-a");
			client.ApplyOp(
				new MergeTreeInsertMsg() { Pos1 = 11, Seg = "!" },
				seq: 11,
				refSeq: 8,
				clientId: "remote-b",
				minimumSequenceNumber: 15);

			Assert.Equal(" world!", client.GetText(0, client.GetLength()));
		}

		[Fact]
		public void Zamboni_MaxCountLimit_SpreadsWork()
		{
			MergeTreeModel tree = CreateSingleCharacterSegmentTree(100);
			tree.MarkRangeRemoved(0, 100, tree.CurrentSeq, seq: 150, clientId: "remover");
			AdvanceMinSeq(tree, 200);

			tree.ZamboniSegments(maxCount: 20);

			Assert.Equal(80, CountSegments(tree));
			tree.ZamboniSegments(maxCount: 20);
			Assert.Equal(60, CountSegments(tree));
		}

		[Fact]
		public void Zamboni_PartialLengthsCache_CorrectAfterCoalesce()
		{
			MergeTreeModel tree = CreateTreeWithSegments(("a", null), ("b", null), ("c", null));

			AdvanceMinSeq(tree, 10);
			tree.ZamboniSegments();
			tree.InsertSegments(
				2,
				new ISegment[] { TextSegmentModel.Make("!") },
				MergeTreeModel.UnassignedSequenceNumber,
				MergeTreeModel.UnassignedSequenceNumber,
				"local");

			Assert.Equal("ab!c", tree.GetText());
			Assert.Equal(4, tree.GetLength());
			Assert.Equal(1, tree.GetContainingSegment(1, MergeTreeModel.UnassignedSequenceNumber, "local").offsetInSegment);
			Assert.Equal(0, tree.GetContainingSegment(2, MergeTreeModel.UnassignedSequenceNumber, "local").offsetInSegment);
			Assert.Equal(0, tree.GetContainingSegment(3, MergeTreeModel.UnassignedSequenceNumber, "local").offsetInSegment);
		}

		[Fact]
		public void Zamboni_CrossBlockCoalesce_MergesAdjacent()
		{
			MergeTreeModel tree = CreateSingleCharacterSegmentTree(8);
			ISegment leftBoundary = tree.GetContainingSegment(3, MergeTreeModel.UnassignedSequenceNumber, "client").segment;
			ISegment rightBoundary = tree.GetContainingSegment(4, MergeTreeModel.UnassignedSequenceNumber, "client").segment;
			Assert.NotSame(leftBoundary.Parent, rightBoundary.Parent);

			AdvanceMinSeq(tree, 20);
			tree.ZamboniSegments();

			TextSegmentModel segment = Assert.Single(tree.WalkAllSegments().Cast<TextSegmentModel>());
			Assert.Equal("xxxxxxxx", segment.Text);
			Assert.Equal("xxxxxxxx", tree.GetText());
		}

		[Fact]
		public void Zamboni_SparseBlock_MergedIntoSibling()
		{
			MergeTreeModel tree = CreateSingleCharacterSegmentTree(12, i => Property("segment", i));
			Assert.Equal(3, tree.Root.ChildCount);

			tree.MarkRangeRemoved(4, 7, tree.CurrentSeq, seq: 50, clientId: "remover");
			AdvanceMinSeq(tree, 100);
			tree.ZamboniSegments();

			Assert.Equal(9, CountSegments(tree));
			Assert.Equal(2, tree.Root.ChildCount);
			Assert.All(
				tree.Root.Children.Take(tree.Root.ChildCount).Cast<MergeBlockModel>(),
				block => Assert.True(block.ChildCount >= MergeBlockModel.MinChildren));
			Assert.Equal(new string('x', 9), tree.GetText());
		}

		[Fact]
		public void Zamboni_BlockPack_ReducesTreeDepth()
		{
			MergeTreeModel tree = CreateSingleCharacterSegmentTree(200, i => Property("segment", i));
			int depthBefore = GetDepth(tree.Root);

			tree.MarkRangeRemoved(5, 200, tree.CurrentSeq, seq: 300, clientId: "remover");
			AdvanceMinSeq(tree, 400);
			tree.ZamboniSegments();

			Assert.True(depthBefore >= 3);
			Assert.True(GetDepth(tree.Root) < depthBefore);
			Assert.Equal(5, CountSegments(tree));
			Assert.Equal(new string('x', 5), tree.GetText());
		}

		[Fact]
		public void Zamboni_CrossBlock_ReferenceStillValid()
		{
			MergeTreeModel tree = CreateSingleCharacterSegmentTree(8);
			LocalReferencePosition reference = tree.CreateReferencePosition(
				4,
				ReferenceType.Simple,
				SlidingPreference.Forward);

			AdvanceMinSeq(tree, 20);
			tree.ZamboniSegments();

			Assert.False(reference.IsDetached);
			Assert.Equal(4, tree.GetPositionOfReference(reference));
			Assert.Same(Assert.Single(tree.WalkAllSegments()), reference.Segment);
		}

		[Fact]
		public void Zamboni_CrossBlock_IntervalStillValid()
		{
			Client client = CreateClientWithAckedSegments(Enumerable.Repeat("x", 8).ToArray());
			IntervalCollection collection = client.GetOrCreateIntervalCollection("cross-block", IntervalType.SlideOnRemove);
			SequenceInterval interval = collection.Add(4, 4, intervalId: "interval-1");

			client.SetCollaborationWindow(minSeq: 20, currentSeq: 20);

			Assert.Same(interval, collection.GetIntervalById("interval-1"));
			Assert.False(interval.Start.IsDetached);
			Assert.False(interval.End.IsDetached);
			Assert.Equal(4, interval.StartPosition);
			Assert.Equal(4, interval.EndPosition);
		}

		[Fact]
		public void Zamboni_CrossBlock_MatchesFullRebuild()
		{
			MergeTreeModel tree = CreateSingleCharacterSegmentTree(
				20,
				i => i < 8 ? Property("coalesce", true) : Property("segment", i));

			AdvanceMinSeq(tree, 100);
			tree.ZamboniSegments();

			AssertPartialLengthsCacheMatchesFullRebuild(tree);
		}

		private static MergeTreeModel CreateTreeWithSegments(params (string Text, PropertySet? Props)[] segmentSpecs)
		{
			MergeTreeModel tree = new();
			int position = 0;
			long seq = 1;
			foreach ((string text, PropertySet? props) in segmentSpecs)
			{
				tree.InsertSegments(
					position,
					new ISegment[] { TextSegmentModel.Make(text, PropertyMap.ClonePropertySet(props)) },
					refSeq: seq - 1,
					seq: seq,
					clientId: "client");
				position += text.Length;
				seq++;
			}

			return tree;
		}

		private static Client CreateClientWithAckedSegments(params string[] texts)
		{
			Client client = new("doc");
			int position = 0;
			long refSeq = 0;
			long seq = 1;
			foreach (string text in texts)
			{
				IMergeTreeInsertMsg insert = client.InsertText(position, text);
				client.ApplyOp(insert, seq, refSeq, "doc");
				position += text.Length;
				refSeq = seq;
				seq++;
			}

			return client;
		}

		private static MergeTreeModel CreateSingleCharacterSegmentTree(int count)
		{
			return CreateSingleCharacterSegmentTree(count, _ => null);
		}

		private static MergeTreeModel CreateSingleCharacterSegmentTree(
			int count,
			Func<int, PropertySet?> propertyFactory)
		{
			MergeTreeModel tree = new();
			for (int i = 0; i < count; i++)
			{
				tree.InsertSegments(
					i,
					new ISegment[] { TextSegmentModel.Make("x", propertyFactory(i)) },
					refSeq: i,
					seq: i + 1,
					clientId: "client");
			}

			return tree;
		}

		private static void AdvanceMinSeq(MergeTreeModel tree, long minSeq)
		{
			if (tree.CurrentSeq < minSeq)
			{
				tree.CurrentSeq = minSeq;
			}

			tree.SetMinSeq(minSeq);
		}

		private static LocalReferencePosition AttachReference(
			ISegment segment,
			int offset,
			SlidingPreference slidingPreference)
		{
			LocalReferencePosition reference = new(
				segment,
				offset,
				ReferenceType.SlideOnRemove,
				slidingPreference);
			segment.LocalRefs.Add(reference);
			return reference;
		}

		private static ISegment GetRequiredSegment(MergeTreeModel tree, int position)
		{
			return tree.GetContainingSegment(position, MergeTreeModel.UnassignedSequenceNumber, "client").segment;
		}

		private static int CountSegments(MergeTreeModel tree)
		{
			return tree.WalkAllSegments().Count();
		}

		private static PropertySet Property(string key, object value)
		{
			return new PropertySet()
			{
				[key] = value,
			};
		}

		private static int GetDepth(MergeBlockModel block)
		{
			int maxChildDepth = 0;
			for (int i = 0; i < block.ChildCount; i++)
			{
				if (block.Children[i] is MergeBlockModel childBlock)
				{
					maxChildDepth = Math.Max(maxChildDepth, GetDepth(childBlock));
				}
			}

			return maxChildDepth + 1;
		}

		private static void AssertPartialLengthsCacheMatchesFullRebuild(MergeTreeModel tree)
		{
			Dictionary<MergeBlockModel, BlockLengthSnapshot> incremental = CaptureBlockLengthSnapshots(tree);
			InvokeRebuildFromRoot(tree);
			Dictionary<MergeBlockModel, BlockLengthSnapshot> rebuilt = CaptureBlockLengthSnapshots(tree);

			Assert.Equal(incremental.Count, rebuilt.Count);
			foreach (KeyValuePair<MergeBlockModel, BlockLengthSnapshot> entry in incremental)
			{
				Assert.True(rebuilt.TryGetValue(entry.Key, out BlockLengthSnapshot? rebuiltSnapshot));
				Assert.Equal(entry.Value, rebuiltSnapshot);
			}
		}

		private static Dictionary<MergeBlockModel, BlockLengthSnapshot> CaptureBlockLengthSnapshots(MergeTreeModel tree)
		{
			Dictionary<MergeBlockModel, BlockLengthSnapshot> snapshots = new();
			foreach (MergeBlockModel block in WalkBlocks(tree.Root))
			{
				snapshots.Add(block, CaptureBlockLengthSnapshot(block));
			}

			return snapshots;
		}

		private static IEnumerable<MergeBlockModel> WalkBlocks(MergeBlockModel block)
		{
			yield return block;
			for (int i = 0; i < block.ChildCount; i++)
			{
				if (block.Children[i] is MergeBlockModel childBlock)
				{
					foreach (MergeBlockModel descendant in WalkBlocks(childBlock))
					{
						yield return descendant;
					}
				}
			}
		}

		private static BlockLengthSnapshot CaptureBlockLengthSnapshot(MergeBlockModel block)
		{
			PartialLengths partialLengths = (PartialLengths)PartialLengthsProperty.GetValue(block)!;
			return new BlockLengthSnapshot(
				block.CachedLength,
				partialLengths.TotalCurrentLength,
				FormatDeltas((SortedDictionary<long, int>)DeltasBySeqField.GetValue(partialLengths)!),
				FormatClientDeltas((Dictionary<int, SortedDictionary<long, int>>)ClientDeltasBySeqField.GetValue(partialLengths)!),
				FormatLocalUnackedSegments((Dictionary<int, HashSet<ISegment>>)LocalUnackedSegmentsByClientField.GetValue(partialLengths)!));
		}

		private static string FormatDeltas(SortedDictionary<long, int> deltas)
		{
			return string.Join(",", deltas.Select(delta => $"{delta.Key}:{delta.Value}"));
		}

		private static string FormatClientDeltas(Dictionary<int, SortedDictionary<long, int>> clientDeltas)
		{
			return string.Join(
				"|",
				clientDeltas
					.OrderBy(entry => entry.Key)
					.Select(entry => $"{entry.Key}=[{FormatDeltas(entry.Value)}]"));
		}

		private static string FormatLocalUnackedSegments(Dictionary<int, HashSet<ISegment>> localSegments)
		{
			return string.Join(
				"|",
				localSegments
					.OrderBy(entry => entry.Key)
					.Select(entry => $"{entry.Key}=[{string.Join(",", entry.Value.Select(segment => segment.Ordinal).OrderBy(ordinal => ordinal, StringComparer.Ordinal))}]"));
		}

		private static void InvokeRebuildFromRoot(MergeTreeModel tree)
		{
			RebuildFromRootMethod.Invoke(tree, Array.Empty<object>());
		}

		private static PropertyInfo PartialLengthsProperty { get; } =
			typeof(MergeBlockModel).GetProperty("PartialLengths", BindingFlags.Instance | BindingFlags.NonPublic)
			?? throw new InvalidOperationException("MergeBlock.PartialLengths was not found.");

		private static MethodInfo RebuildFromRootMethod { get; } =
			typeof(MergeTreeModel).GetMethod(
				"RebuildFromRoot",
				BindingFlags.Instance | BindingFlags.NonPublic,
				binder: null,
				Type.EmptyTypes,
				modifiers: null)
			?? throw new InvalidOperationException("MergeTree.RebuildFromRoot was not found.");

		private static FieldInfo DeltasBySeqField { get; } =
			typeof(PartialLengths).GetField("_deltasBySeq", BindingFlags.Instance | BindingFlags.NonPublic)
			?? throw new InvalidOperationException("PartialLengths._deltasBySeq was not found.");

		private static FieldInfo ClientDeltasBySeqField { get; } =
			typeof(PartialLengths).GetField("_clientDeltasBySeq", BindingFlags.Instance | BindingFlags.NonPublic)
			?? throw new InvalidOperationException("PartialLengths._clientDeltasBySeq was not found.");

		private static FieldInfo LocalUnackedSegmentsByClientField { get; } =
			typeof(PartialLengths).GetField("_localUnackedSegmentsByClient", BindingFlags.Instance | BindingFlags.NonPublic)
			?? throw new InvalidOperationException("PartialLengths._localUnackedSegmentsByClient was not found.");

		private sealed record BlockLengthSnapshot(
			int CachedLength,
			int TotalCurrentLength,
			string DeltasBySeq,
			string ClientDeltasBySeq,
			string LocalUnackedSegmentsByClient);
	}
}
