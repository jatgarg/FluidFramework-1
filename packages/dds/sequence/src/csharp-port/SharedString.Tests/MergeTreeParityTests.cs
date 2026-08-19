// -----------------------------------------------------------------------------
// M-series merge-tree parity regressions for SharedString C# port.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using IntervalSide = Microsoft.Office.Web.Fluid.Intervals.Side;
using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;
using TextSegmentModel = Microsoft.Office.Web.Fluid.MergeTree.TextSegment;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class MergeTreeParityTests
	{
		[Fact]
		public void M2_InsertInsideObliterate_IsStampedAsRemovedOnInsert()
		{
			MergeTreeModel tree = CreateAckedTree("0123456789AB");
			tree.ObliterateRange(5, 10, refSeq: 1, seq: 10, clientId: "obliterator");

			tree.InsertSegments(7, SegmentArray("X"), refSeq: 5, seq: 11, clientId: "inserter");

			ISegment inserted = SegmentWithText(tree, "X");
			Assert.Equal("01234AB", tree.GetText());
			Assert.Contains(inserted.RemoveStamps, stamp => stamp is SliceRemoveOperationStamp && stamp.Seq == 10);
		}

		[Fact]
		public void M2_NewerSameClientObliterate_AllowsInsertIntoOlderObliterate()
		{
			MergeTreeModel tree = CreateAckedTree("0123456789AB");
			tree.ObliterateRange(5, 10, refSeq: 1, seq: 10, clientId: "client-a");
			tree.ObliterateRange(5, 10, refSeq: 1, seq: 12, clientId: "client-b");

			tree.InsertSegments(7, SegmentArray("X"), refSeq: 5, seq: 13, clientId: "client-b");

			ISegment inserted = SegmentWithText(tree, "X");
			Assert.Empty(inserted.RemoveStamps);
			Assert.Contains("X", tree.GetText(), StringComparison.Ordinal);
		}

		[Fact]
		public void M2_MultipleOlderOverlappingObliterates_ArePreservedOnInsertedSegment()
		{
			MergeTreeModel tree = CreateAckedTree("0123456789AB");
			tree.ObliterateRange(5, 10, refSeq: 1, seq: 10, clientId: "client-a");
			tree.ObliterateRange(5, 10, refSeq: 1, seq: 12, clientId: "client-c");

			tree.InsertSegments(7, SegmentArray("X"), refSeq: 5, seq: 13, clientId: "client-b");

			ISegment inserted = SegmentWithText(tree, "X");
			Assert.Equal(new long[] { 10, 12 }, inserted.RemoveStamps.Select(stamp => stamp.Seq).ToArray());
			Assert.Equal("01234AB", tree.GetText());
		}

		[Fact]
		public void M3_NumericObliterate_EatsConcurrentInsertAtEndBoundary()
		{
			MergeTreeModel tree = CreateAckedTree("0123456789AB");
			tree.ObliterateRange(5, 10, refSeq: 1, seq: 10, clientId: "obliterator");

			tree.InsertSegments(10, SegmentArray("E"), refSeq: 5, seq: 11, clientId: "inserter");

			Assert.Equal("01234AB", tree.GetText());
			Assert.Contains(SegmentWithText(tree, "E").RemoveStamps, stamp => stamp.Seq == 10);
		}

		[Fact]
		public void M3_SidedObliterateEndBefore_SparesConcurrentInsertAtEndBoundary()
		{
			MergeTreeModel tree = CreateAckedTree("0123456789AB");
			tree.ObliterateRangeSided(
				SequencePlace.At(5, Microsoft.Office.Web.Fluid.MergeTree.Side.Before),
				SequencePlace.At(10, Microsoft.Office.Web.Fluid.MergeTree.Side.Before),
				refSeq: 1,
				seq: 10,
				clientId: "obliterator");

			tree.InsertSegments(10, SegmentArray("E"), refSeq: 5, seq: 11, clientId: "inserter");

			Assert.Equal("01234EAB", tree.GetText());
			Assert.Empty(SegmentWithText(tree, "E").RemoveStamps);
		}

		[Fact]
		public void M3_SidedObliterateStartAfter_EatsConcurrentInsertAtStartBoundary()
		{
			MergeTreeModel tree = CreateAckedTree("0123456789AB");
			tree.ObliterateRangeSided(
				SequencePlace.At(5, Microsoft.Office.Web.Fluid.MergeTree.Side.After),
				SequencePlace.At(10, Microsoft.Office.Web.Fluid.MergeTree.Side.Before),
				refSeq: 1,
				seq: 10,
				clientId: "obliterator");

			tree.InsertSegments(5, SegmentArray("S"), refSeq: 5, seq: 11, clientId: "inserter");

			Assert.Equal("01234AB", tree.GetText());
			Assert.Contains(SegmentWithText(tree, "S").RemoveStamps, stamp => stamp.Seq == 10);
		}

		[Fact]
		public void M3_NumericObliterate_RecordsAfterEndSide()
		{
			MergeTreeModel tree = CreateAckedTree("0123456789AB");

			tree.ObliterateRange(5, 10, refSeq: 1, seq: 10, clientId: "obliterator");

			SliceRemoveOperationStamp stamp = Assert.IsType<SliceRemoveOperationStamp>(
				SegmentWithText(tree, "56789").RemoveStamps.Single());
			Assert.Equal(Microsoft.Office.Web.Fluid.MergeTree.Side.Before, stamp.StartSide);
			Assert.Equal(Microsoft.Office.Web.Fluid.MergeTree.Side.After, stamp.EndSide);
		}

		[Fact]
		public void M3_ZeroLengthObliterate_DoesNotStampSegments()
		{
			MergeTreeModel tree = CreateAckedTree("0123456789AB");

			List<ISegment> delta = tree.ObliterateRange(5, 5, refSeq: 1, seq: 10, clientId: "obliterator");

			Assert.Empty(delta);
			Assert.Equal("0123456789AB", tree.GetText());
			Assert.All(tree.WalkAllSegments(), segment => Assert.Empty(segment.RemoveStamps));
		}

		[Fact]
		public void ClientIdAllocator_NumericIdReservesShortId_NonNumericCannotCollide()
		{
			// TS ref: merge-tree/src/client.ts client-id allocator — TS reserves
			// numeric-parseable ids so a later non-numeric client never receives
			// the same short id. Without the reservation, a numeric id like "5"
			// is used verbatim, and the allocator later hands out short id 5 to
			// a non-numeric client, causing the two clients' stamps to be
			// indistinguishable.
			MergeTreeModel tree = new();
			tree.InsertSegments(0, SegmentArray("a"), refSeq: 0, seq: 1, clientId: "5");
			// Any allocator counter that hadn't reserved 5 would soon reach it.
			tree.InsertSegments(1, SegmentArray("b"), refSeq: 1, seq: 2, clientId: "alice");
			tree.InsertSegments(2, SegmentArray("c"), refSeq: 2, seq: 3, clientId: "bob");
			tree.InsertSegments(3, SegmentArray("d"), refSeq: 3, seq: 4, clientId: "charlie");
			tree.InsertSegments(4, SegmentArray("e"), refSeq: 4, seq: 5, clientId: "dan");
			tree.InsertSegments(5, SegmentArray("f"), refSeq: 5, seq: 6, clientId: "eve");

			int? numericClientShortId = SegmentWithText(tree, "a").InsertionStamp?.ClientId;
			HashSet<int> nonNumericShortIds = new();
			foreach (string letter in new[] { "b", "c", "d", "e", "f" })
			{
				int? shortId = SegmentWithText(tree, letter).InsertionStamp?.ClientId;
				Assert.NotNull(shortId);
				Assert.NotEqual(numericClientShortId, shortId);
				Assert.True(nonNumericShortIds.Add(shortId!.Value));
			}
		}

		[Fact]
		public void M4_IntervalEndpointSurvivesAnnotateBoundarySplit()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdefghijklmnop");
			IntervalCollection collection = sharedString.GetIntervalCollection("m4-annotate");
			SequenceInterval interval = collection.Add(5, IntervalSide.Before, 10, IntervalSide.Before, intervalId: "i1");

			sharedString.AnnotateRange(0, 6, new PropertySet() { ["style"] = "left" });

			Assert.Equal(5, interval.StartPosition);
			Assert.Equal(10, interval.EndPosition);
		}

		[Fact]
		public void M4_IntervalEndpointSurvivesAdjacentRemoveAndInsert()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdefghijklmnop");
			IntervalCollection collection = sharedString.GetIntervalCollection("m4-edit");
			SequenceInterval interval = collection.Add(5, IntervalSide.Before, 10, IntervalSide.Before, intervalId: "i1");

			sharedString.DeleteText(0, 2);
			sharedString.InsertText(3, "XX");

			Assert.Equal(5, interval.StartPosition);
			Assert.Equal(10, interval.EndPosition);
		}

		[Fact]
		public void M5_AnnotatingMarkerWithSameId_IsAllowed()
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, "hi");
			sharedString.InsertMarker(2, ReferenceType.Tile, MarkerProps("m1", "tile"));

			sharedString.AnnotateRange(2, 3, new PropertySet()
			{
				[Marker.ReservedMarkerIdKey] = "m1",
				["label"] = "updated",
			});

			Marker? marker = sharedString.GetMarkerFromId("m1");
			Assert.NotNull(marker);
			Assert.Equal("updated", Assert.IsType<string>(marker!.Properties!["label"]));
		}

		[Fact]
		public void M5_AnnotatingMarkerWithDifferentId_Throws()
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, "hi");
			sharedString.InsertMarker(2, ReferenceType.Tile, MarkerProps("m1", "tile"));

			Assert.Throws<InvalidOperationException>(() => sharedString.AnnotateRange(2, 3, new PropertySet()
			{
				[Marker.ReservedMarkerIdKey] = "m2",
			}));
			Assert.NotNull(sharedString.GetMarkerFromId("m1"));
			Assert.Null(sharedString.GetMarkerFromId("m2"));
		}

		[Fact]
		public void M5_AnnotatingMarkerWithNullId_Throws()
		{
			SharedString sharedString = new();
			sharedString.InsertMarker(0, ReferenceType.Tile, MarkerProps("m1", "tile"));

			Assert.Throws<InvalidOperationException>(() => sharedString.AnnotateRange(0, 1, new PropertySet()
			{
				[Marker.ReservedMarkerIdKey] = null,
			}));
		}

		[Fact]
		public void M5_RemoteMarkerIdChange_ThrowsBeforeMutatingIndex()
		{
			SharedString sharedString = CreateSharedStringWithMarker("m1");
			string opJson = SharedStringOpSerializer.Serialize(new MergeTreeAnnotateMsg()
			{
				Pos1 = 0,
				Pos2 = 1,
				Props = new PropertySet() { [Marker.ReservedMarkerIdKey] = "m2" },
			});

			Assert.Throws<InvalidOperationException>(() => sharedString.ProcessDataObjectOp(RemoteMessage(1, 2), opJson));
			Assert.NotNull(sharedString.GetMarkerFromId("m1"));
			Assert.Null(sharedString.GetMarkerFromId("m2"));
		}

		[Fact]
		public void M6_ZamboniKeepsUnackedRemovedSegment()
		{
			MergeTreeModel tree = CreateAckedTree("drop");
			tree.MarkRangeRemoved(
				0,
				4,
				MergeTreeModel.UnassignedSequenceNumber,
				MergeTreeModel.UnassignedSequenceNumber,
				"client-a");

			AdvanceMinSeq(tree, 100);
			tree.ZamboniSegments();

			ISegment segment = Assert.Single(tree.WalkAllSegments());
			Assert.Equal(MergeTreeModel.UnassignedSequenceNumber, segment.RemoveStamps.Single().Seq);
		}

		[Fact]
		public void M6_ZamboniKeepsSegmentUntilAllRemoveStampsAreStable()
		{
			MergeTreeModel tree = CreateAckedTree("drop");
			tree.MarkRangeRemoved(0, 4, refSeq: 1, seq: 10, clientId: "client-a");
			tree.MarkRangeRemoved(0, 4, refSeq: 1, seq: 12, clientId: "client-b");

			AdvanceMinSeq(tree, 11);
			tree.ZamboniSegments();

			Assert.Single(tree.WalkAllSegments());
			AdvanceMinSeq(tree, 12);
			tree.ZamboniSegments();
			Assert.Single(tree.WalkAllSegments());
			AdvanceMinSeq(tree, 13);
			tree.ZamboniSegments();
			Assert.Empty(tree.WalkAllSegments());
		}

		[Fact]
		public void M6_ZamboniCoalescePreservesReferenceIdentity()
		{
			MergeTreeModel tree = new();
			tree.InsertSegments(0, SegmentArray("a"), refSeq: 0, seq: 1, clientId: "seed");
			tree.InsertSegments(1, SegmentArray("b"), refSeq: 1, seq: 2, clientId: "seed");
			LocalReferencePosition reference = tree.CreateReferencePosition(1, ReferenceType.Simple);

			AdvanceMinSeq(tree, 3);
			tree.ZamboniSegments();

			Assert.Equal(1, tree.GetPositionOfReference(reference));
			Assert.Equal("ab", Assert.IsType<TextSegmentModel>(Assert.Single(tree.WalkAllSegments())).Text);
		}

		[Fact]
		public void M7_ZamboniKeepsTombstoneAtMinimumSequenceBoundary()
		{
			MergeTreeModel tree = CreateAckedTree("drop");
			tree.MarkRangeRemoved(0, 4, refSeq: 1, seq: 10, clientId: "remover");

			AdvanceMinSeq(tree, 10);
			tree.ZamboniSegments();

			Assert.Single(tree.WalkAllSegments());
		}

		[Fact]
		public void M7_ZamboniCoalescesLiveSegmentsAfterMinimumSequenceBoundary()
		{
			MergeTreeModel tree = new();
			tree.InsertSegments(0, SegmentArray("a"), refSeq: 0, seq: 1, clientId: "seed");
			tree.InsertSegments(1, SegmentArray("b"), refSeq: 1, seq: 2, clientId: "seed");

			AdvanceMinSeq(tree, 3);
			tree.ZamboniSegments();

			TextSegmentModel segment = Assert.IsType<TextSegmentModel>(Assert.Single(tree.WalkAllSegments()));
			Assert.Equal("ab", segment.Text);
		}

		[Fact]
		public void M8_SearchForMarkerForwardAcrossMultipleBlocks()
		{
			SharedString sharedString = CreateSharedStringWithManyMarkers();

			Marker? marker = sharedString.SearchForMarker(0, forwards: true, tileLabel: "target");
			Assert.NotNull(marker);

			Assert.Equal("target-8", marker!.GetId());
		}

		[Fact]
		public void M8_SearchForMarkerBackwardAcrossMultipleBlocks()
		{
			SharedString sharedString = CreateSharedStringWithManyMarkers();

			Marker? marker = sharedString.SearchForMarker(sharedString.GetLength() - 1, forwards: false, tileLabel: "target");
			Assert.NotNull(marker);

			Assert.Equal("target-8", marker!.GetId());
		}

		[Fact]
		public void M8_SearchForMarkerAtMarkerPositionFindsItBothDirections()
		{
			SharedString sharedString = CreateSharedStringWithManyMarkers();
			Marker? marker = sharedString.GetMarkerFromId("target-8");
			Assert.NotNull(marker);
			int? position = sharedString.GetPositionOfMarker(marker!);
			Assert.NotNull(position);

			Assert.Same(marker, sharedString.SearchForMarker(position.Value, forwards: true, tileLabel: "target"));
			Assert.Same(marker, sharedString.SearchForMarker(position.Value, forwards: false, tileLabel: "target"));
		}

		[Fact]
		public void M9_GetPositionOfRemovedSegmentReturnsMinusOne()
		{
			MergeTreeModel tree = CreateAckedTree("hello");
			ISegment segment = Assert.Single(tree.WalkAllSegments());

			tree.MarkRangeRemoved(0, 5, refSeq: 1, seq: 10, clientId: "remover");

			Assert.Equal(MergeTreeModel.DetachedReferencePosition, tree.GetPositionOfSegment(segment));
		}

		[Fact]
		public void M9_GetPositionOfDetachedSegmentReturnsMinusOne()
		{
			MergeTreeModel tree = CreateAckedTree("hello");
			ISegment detached = TextSegmentModel.Make("x");

			Assert.Equal(MergeTreeModel.DetachedReferencePosition, tree.GetPositionOfSegment(detached));
		}

		[Fact]
		public void M9_GetPositionOfDetachedReferenceReturnsMinusOne()
		{
			MergeTreeModel tree = CreateAckedTree("hello");
			LocalReferencePosition reference = tree.CreateReferencePosition(2, ReferenceType.Simple);

			tree.RemoveReferencePosition(reference);

			Assert.Equal(MergeTreeModel.DetachedReferencePosition, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void M13_HistoricalContainingSegmentDescendsIntoZeroCurrentBlock()
		{
			MergeTreeModel tree = CreateSingleCharacterSegmentTree(12);
			tree.MarkRangeRemoved(0, 8, refSeq: 12, seq: 20, clientId: "remover");

			(ISegment segment, int offset) = tree.GetContainingSegment(5, refSeq: 12, clientId: "reader");

			Assert.Equal(0, offset);
			Assert.Equal(4, tree.GetLength());
			Assert.Equal(12, tree.GetLength(12));
			Assert.Contains(segment, tree.WalkAllSegments());
		}

		[Fact]
		public void M13_InsertIntoObliteratedBlockBoundaryIsNotPruned()
		{
			MergeTreeModel tree = CreateSingleCharacterSegmentTree(12);
			tree.ObliterateRange(3, 9, refSeq: 12, seq: 20, clientId: "obliterator");

			tree.InsertSegments(5, SegmentArray("X"), refSeq: 12, seq: 21, clientId: "inserter");

			Assert.DoesNotContain("X", tree.GetText());
			Assert.Contains(SegmentWithText(tree, "X").RemoveStamps, stamp => stamp.Seq == 20);
		}

		[Fact]
		public void M13_OverlappingRemoveAndObliterateTraversalKeepsHistoricalLength()
		{
			MergeTreeModel tree = CreateSingleCharacterSegmentTree(16);
			tree.MarkRangeRemoved(2, 12, refSeq: 16, seq: 20, clientId: "remover");
			tree.ObliterateRange(6, 14, refSeq: 16, seq: 22, clientId: "obliterator");

			Assert.Equal(16, tree.GetLength(16));
			Assert.Equal(6, tree.GetLength(20));
			Assert.Equal(4, tree.GetLength(22));
		}

		[Fact]
		public void M14_MatchPropertiesDistinguishesNullFromMissing()
		{
			Assert.False(PropertyMap.MatchProperties(
				new PropertySet() { ["value"] = null },
				new PropertySet()));
			Assert.True(PropertyMap.MatchProperties(null, new PropertySet()));
		}

		[Fact]
		public void M14_MatchPropertiesDeepMatchesNestedDictionaries()
		{
			Assert.True(PropertyMap.MatchProperties(
				new PropertySet() { ["config"] = new PropertySet() { ["color"] = "red" } },
				new PropertySet() { ["config"] = new PropertySet() { ["color"] = "red" } }));
		}

		[Fact]
		public void M14_ZamboniCoalescesSegmentsWithArrayEquivalentProperties()
		{
			MergeTreeModel tree = new();
			PropertySet arrayProps = new()
			{
				["labels"] = new object?[] { "a", 1, true },
			};
			PropertySet listProps = new()
			{
				["labels"] = new List<object?>() { "a", 1L, true },
			};
			tree.InsertSegments(0, new ISegment[] { TextSegmentModel.Make("a", arrayProps) }, refSeq: 0, seq: 1, clientId: "seed");
			tree.InsertSegments(1, new ISegment[] { TextSegmentModel.Make("b", listProps) }, refSeq: 1, seq: 2, clientId: "seed");

			AdvanceMinSeq(tree, 3);
			tree.ZamboniSegments();

			TextSegmentModel segment = Assert.IsType<TextSegmentModel>(Assert.Single(tree.WalkAllSegments()));
			Assert.Equal("ab", segment.Text);
		}

		[Fact]
		public void M14_ZamboniCoalescesSegmentsWithJsonEquivalentProperties()
		{
			MergeTreeModel tree = new();
			PropertySet jsonProps = new()
			{
				["config"] = JsonDocument.Parse("""{"color":"red","weight":1}""").RootElement.Clone(),
			};
			PropertySet clrProps = new()
			{
				["config"] = new PropertySet()
				{
					["color"] = "red",
					["weight"] = 1L,
				},
			};
			tree.InsertSegments(0, new ISegment[] { TextSegmentModel.Make("a", jsonProps) }, refSeq: 0, seq: 1, clientId: "seed");
			tree.InsertSegments(1, new ISegment[] { TextSegmentModel.Make("b", clrProps) }, refSeq: 1, seq: 2, clientId: "seed");

			AdvanceMinSeq(tree, 3);
			tree.ZamboniSegments();

			TextSegmentModel segment = Assert.IsType<TextSegmentModel>(Assert.Single(tree.WalkAllSegments()));
			Assert.Equal("ab", segment.Text);
		}

		private static MergeTreeModel CreateAckedTree(string text)
		{
			MergeTreeModel tree = new();
			tree.InsertSegments(0, SegmentArray(text), MergeTreeModel.UnassignedSequenceNumber, seq: 1, clientId: "seed");
			return tree;
		}

		private static SharedString CreateSharedStringWithText(string text)
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, text);
			return sharedString;
		}

		private static SharedString CreateSharedStringWithMarker(string markerId)
		{
			FakeFluidDataObjectSender sender = new();
			SharedString sharedString = new("doc", sender);
			sharedString.InsertMarker(0, ReferenceType.Tile, MarkerProps(markerId, "tile"));
			sharedString.ProcessDataObjectOp(LocalAck(sender.Sent[0].ClientSeq, refSeq: 0, seq: 1), sender.Sent[0].OpJson);
			return sharedString;
		}

		private static SharedString CreateSharedStringWithManyMarkers()
		{
			SharedString sharedString = new();
			for (int i = 0; i < 12; i++)
			{
				sharedString.InsertText(sharedString.GetLength(), "x");
				string tileLabel = i == 8 ? "target" : "other";
				sharedString.InsertMarker(
					sharedString.GetLength(),
					ReferenceType.Tile,
					MarkerProps($"{tileLabel}-{i}", tileLabel));
			}

			return sharedString;
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

		private static ISegment SegmentWithText(MergeTreeModel tree, string text)
		{
			return Assert.Single(
				tree.WalkAllSegments(),
				segment => segment is TextSegmentModel textSegment && textSegment.Text == text);
		}

		private static PropertySet MarkerProps(string markerId, params string[] tileLabels)
		{
			return new PropertySet()
			{
				[Marker.ReservedMarkerIdKey] = markerId,
				[Marker.ReservedTileLabelsKey] = tileLabels,
			};
		}

		private static void AdvanceMinSeq(MergeTreeModel tree, long minSeq)
		{
			if (tree.CurrentSeq < minSeq)
			{
				tree.CurrentSeq = minSeq;
			}

			tree.SetMinSeq(minSeq);
		}

		private static SequencedDocumentMessageDescriptor RemoteMessage(long refSeq, long seq)
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: 0, refSeq: refSeq, seq: seq),
				OpOrigin.Remote,
				"remote-client");
		}

		private static SequencedDocumentMessageDescriptor LocalAck(long clientSeq, long refSeq, long seq)
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: clientSeq, refSeq: refSeq, seq: seq),
				OpOrigin.Local);
		}
	}
}
