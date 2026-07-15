// -----------------------------------------------------------------------------
// Incremental merge-tree split/rebalance tests for SharedString C# port.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using MergeBlockModel = Microsoft.Office.Web.Fluid.MergeTree.MergeBlock;
using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;
using TextSegmentModel = Microsoft.Office.Web.Fluid.MergeTree.TextSegment;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class MergeTreeSplitTests
	{
		[Fact]
		public void Insert_ManySegments_DoesNotRebuildTree()
		{
			MergeTreeModel tree = CreateTreeWithIndividualSingleCharacterSegments(128);
			ISegment firstSegment = tree.WalkAllSegments().First();
			MergeBlockModel? originalFirstParent = firstSegment.Parent;

			tree.InsertSegments(
				tree.GetLength(),
				new ISegment[] { TextSegmentModel.Make("z") },
				MergeTreeModel.UnassignedSequenceNumber,
				MergeTreeModel.UnassignedSequenceNumber,
				"local");

			Assert.NotNull(originalFirstParent);
			Assert.Same(originalFirstParent, firstSegment.Parent);
			Assert.True(tree.Root.ChildCount > 1);
			Assert.True(GetDepth(tree.Root) > 1);
			Assert.Equal(new string('x', 128) + "z", tree.GetText());
		}

		[Fact]
		public void Insert_TriggersSplitAt_MaxNodesInBlock()
		{
			MergeTreeModel tree = CreateTreeWithSingleCharacterSegments(MergeBlockModel.MaxChildren);

			Assert.Equal(2, tree.Root.ChildCount);
			MergeBlockModel left = Assert.IsType<MergeBlockModel>(tree.Root.Children[0]);
			MergeBlockModel right = Assert.IsType<MergeBlockModel>(tree.Root.Children[1]);
			Assert.Equal(MergeBlockModel.MaxChildren / 2, left.ChildCount);
			Assert.Equal(MergeBlockModel.MaxChildren / 2, right.ChildCount);
			Assert.Equal(new string('x', MergeBlockModel.MaxChildren), tree.GetText());
		}

		[Fact]
		public void Insert_CascadingSplit_ReachesRoot()
		{
			MergeTreeModel tree = CreateTreeWithSingleCharacterSegments(200);

			Assert.Null(tree.Root.Parent);
			Assert.Equal(0, tree.Root.Index);
			Assert.True(tree.Root.ChildCount > 1);
			Assert.True(GetDepth(tree.Root) >= 3);
			Assert.All(tree.Root.Children.Take(tree.Root.ChildCount), child => Assert.IsType<MergeBlockModel>(child));
			Assert.Equal(200, tree.GetLength());
		}

		[Fact]
		public void Split_PreservesOrdinalOrdering()
		{
			MergeTreeModel tree = CreateTreeWithSingleCharacterSegments(200);

			string? previousOrdinal = null;
			foreach (ISegment segment in tree.WalkAllSegments())
			{
				if (previousOrdinal is not null)
				{
					Assert.True(
						string.CompareOrdinal(previousOrdinal, segment.Ordinal) < 0,
						$"Expected '{FormatOrdinal(previousOrdinal)}' < '{FormatOrdinal(segment.Ordinal)}'.");
				}

				previousOrdinal = segment.Ordinal;
			}
		}

		[Fact]
		public void Insert_MidSegment_Splits_TextSegment()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");

			sharedString.InsertText(5, " beautiful");

			Assert.Equal("hello beautiful world", sharedString.GetText());
			(ISegment segment, int offsetInSegment)? left = sharedString.GetContainingSegment(0);
			(ISegment segment, int offsetInSegment)? inserted = sharedString.GetContainingSegment(6);
			(ISegment segment, int offsetInSegment)? right = sharedString.GetContainingSegment(16);
			Assert.NotNull(left);
			Assert.NotNull(inserted);
			Assert.NotNull(right);
			AssertText(left.Value.segment, "hello");
			AssertText(inserted.Value.segment, " beautiful");
			AssertText(right.Value.segment, " world");
			Assert.Equal(0, left.Value.offsetInSegment);
			Assert.Equal(1, inserted.Value.offsetInSegment);
			Assert.Equal(1, right.Value.offsetInSegment);
		}

		[Fact]
		public void LargeTree_PositionQueries_Correct()
		{
			var sharedString = new SharedString();
			string expectedText = BuildSingleCharacterSharedString(sharedString, 1000);

			Assert.Equal(expectedText, sharedString.GetText());
			foreach (int position in new[] { 0, 1, 7, 123, 500, 999 })
			{
				(ISegment segment, int offsetInSegment)? result = sharedString.GetContainingSegment(position);
				Assert.NotNull(result);
				Assert.Equal(0, result.Value.offsetInSegment);
				AssertText(result.Value.segment, expectedText[position].ToString());
				Assert.Equal(position, sharedString.GetPosition(result.Value.segment));
				AssertRange(sharedString.GetRangeExtentsOfPosition(position), position, position + 1);
			}
		}

		[Fact]
		public void LargeTree_IntervalReferences_SlideCorrectly()
		{
			var sharedString = new SharedString();
			BuildSingleCharacterSharedString(sharedString, 500);
			IntervalCollection collection = sharedString.GetIntervalCollection("large-tree");
			SequenceInterval first = collection.Add(100, 120, intervalId: "first");
			SequenceInterval second = collection.Add(300, 350, intervalId: "second");

			sharedString.InsertText(50, "abcdefghij");
			sharedString.DeleteText(0, 20);

			AssertIntervalPositions(sharedString, first, 90, 110);
			AssertIntervalPositions(sharedString, second, 290, 340);
		}

		[Fact]
		public void LargeTree_ObliterateEating_StillWorks()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			long seq = InsertAckedCharacters(sharedString, sender, 200);
			sharedString.ObliterateRange(50, 150);
			ProcessLocalAck(sharedString, sender.Sent[^1], refSeq: seq, seq: seq + 1);

			ProcessRemoteInsert(sharedString, position: 75, text: "Z", refSeq: 100, seq: seq + 2);

			Assert.Equal(100, sharedString.GetLength());
			Assert.DoesNotContain("Z", sharedString.GetText(), StringComparison.Ordinal);
		}

		[Fact]
		public void Perf_LargeInsert_DoesNotFullRebuild()
		{
			Stopwatch stopwatch = Stopwatch.StartNew();
			MergeTreeModel tree = CreateTreeWithSingleCharacterSegments(5000);
			ISegment firstSegment = tree.WalkAllSegments().First();
			MergeBlockModel? originalFirstParent = firstSegment.Parent;

			tree.InsertSegments(
				tree.GetLength(),
				new ISegment[] { TextSegmentModel.Make("z") },
				MergeTreeModel.UnassignedSequenceNumber,
				MergeTreeModel.UnassignedSequenceNumber,
				"local");
			stopwatch.Stop();

			Assert.NotNull(originalFirstParent);
			Assert.Same(originalFirstParent, firstSegment.Parent);
			Assert.True(stopwatch.ElapsedMilliseconds < 2000, $"Large-tree insert took {stopwatch.ElapsedMilliseconds}ms.");
		}

		[Fact]
		public void IncrementalLength_AfterInsert_MatchesFullRebuild()
		{
			MergeTreeModel tree = CreateTreeWithSingleCharacterSegments(96);

			tree.InsertSegments(
				17,
				new ISegment[] { TextSegmentModel.Make("abc"), TextSegmentModel.Make("de") },
				MergeTreeModel.UnassignedSequenceNumber,
				101,
				"remote-client");

			Assert.Equal(101, tree.GetLength());
			AssertIncrementalMatchesFullRebuild(tree);
		}

		[Fact]
		public void IncrementalLength_AfterRemove_MatchesFullRebuild()
		{
			MergeTreeModel tree = CreateTreeWithSingleCharacterSegments(96);

			tree.MarkRangeRemoved(
				10,
				45,
				MergeTreeModel.UnassignedSequenceNumber,
				102,
				"remote-client");

			Assert.Equal(61, tree.GetLength());
			AssertIncrementalMatchesFullRebuild(tree);
		}

		[Fact]
		public void IncrementalLength_AfterObliterate_MatchesFullRebuild()
		{
			MergeTreeModel tree = CreateTreeWithSingleCharacterSegments(96);

			tree.ObliterateRange(
				12,
				55,
				MergeTreeModel.UnassignedSequenceNumber,
				103,
				"remote-client");

			Assert.Equal(53, tree.GetLength());
			AssertIncrementalMatchesFullRebuild(tree);
		}

		[Fact]
		public void IncrementalLength_AfterCascadingSplit_MatchesFullRebuild()
		{
			MergeTreeModel tree = new();
			for (int i = 0; i < 256; i++)
			{
				tree.InsertSegments(
					tree.GetLength(),
					new ISegment[] { TextSegmentModel.Make("x") },
					MergeTreeModel.UnassignedSequenceNumber,
					i + 1,
					"remote-client");
			}

			Assert.True(GetDepth(tree.Root) >= 3);
			AssertIncrementalMatchesFullRebuild(tree);
		}

		private static MergeTreeModel CreateTreeWithSingleCharacterSegments(int count)
		{
			MergeTreeModel tree = new();
			ISegment[] segments = new ISegment[count];
			for (int i = 0; i < count; i++)
			{
				segments[i] = TextSegmentModel.Make("x");
			}

			tree.InsertSegments(
				0,
				segments,
				MergeTreeModel.UnassignedSequenceNumber,
				MergeTreeModel.UnassignedSequenceNumber,
				"local");
			return tree;
		}

		private static MergeTreeModel CreateTreeWithIndividualSingleCharacterSegments(int count)
		{
			MergeTreeModel tree = new();
			for (int i = 0; i < count; i++)
			{
				tree.InsertSegments(
					tree.GetLength(),
					new ISegment[] { TextSegmentModel.Make("x") },
					MergeTreeModel.UnassignedSequenceNumber,
					MergeTreeModel.UnassignedSequenceNumber,
					"local");
			}

			return tree;
		}

		private static string BuildSingleCharacterSharedString(SharedString sharedString, int count)
		{
			char[] expected = new char[count];
			for (int i = 0; i < count; i++)
			{
				char value = (char)('a' + (i % 26));
				expected[i] = value;
				sharedString.InsertText(sharedString.GetLength(), value.ToString());
			}

			return new string(expected);
		}

		private static long InsertAckedCharacters(
			SharedString sharedString,
			FakeFluidDataObjectSender sender,
			int count)
		{
			long seq = 0;
			for (int i = 0; i < count; i++)
			{
				sharedString.InsertText(sharedString.GetLength(), "x");
				ProcessLocalAck(sharedString, sender.Sent[^1], refSeq: seq, seq: seq + 1);
				seq++;
			}

			return seq;
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

		private static void AssertIncrementalMatchesFullRebuild(MergeTreeModel tree)
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

		private static void ProcessRemoteInsert(
			SharedString sharedString,
			int position,
			string text,
			long refSeq,
			long seq,
			string clientId = "remote-client")
		{
			string opJson = SharedStringOpSerializer.Serialize(new MergeTreeInsertMsg()
			{
				Pos1 = position,
				Seg = text,
			});
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq, seq, clientId), opJson);
		}

		private static void ProcessLocalAck(
			SharedString sharedString,
			(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
			long refSeq,
			long seq)
		{
			sharedString.ProcessDataObjectOp(LocalAck(sent.ClientSeq, refSeq, seq), sent.OpJson);
		}

		private static SequencedDocumentMessageDescriptor RemoteMessage(
			long refSeq,
			long seq,
			string clientId)
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: 0, refSeq: refSeq, seq: seq),
				OpOrigin.Remote,
				clientId);
		}

		private static SequencedDocumentMessageDescriptor LocalAck(long clientSeq, long refSeq, long seq)
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: clientSeq, refSeq: refSeq, seq: seq),
				OpOrigin.Local);
		}

		private static void AssertText(ISegment segment, string expectedText)
		{
			TextSegmentModel textSegment = Assert.IsType<TextSegmentModel>(segment);
			Assert.Equal(expectedText, textSegment.Text);
		}

		private static void AssertRange((int start, int end)? actual, int expectedStart, int expectedEnd)
		{
			Assert.NotNull(actual);
			Assert.Equal(expectedStart, actual.Value.start);
			Assert.Equal(expectedEnd, actual.Value.end);
		}

		private static void AssertIntervalPositions(
			SharedString sharedString,
			SequenceInterval interval,
			int expectedStart,
			int expectedEnd)
		{
			Assert.Equal(expectedStart, sharedString.LocalReferencePositionToPosition(interval.Start));
			Assert.Equal(expectedEnd, sharedString.LocalReferencePositionToPosition(interval.End));
		}

		private static string FormatOrdinal(string ordinal)
		{
			return string.Join(",", ordinal.Select(c => ((int)c).ToString()));
		}
	}
}
