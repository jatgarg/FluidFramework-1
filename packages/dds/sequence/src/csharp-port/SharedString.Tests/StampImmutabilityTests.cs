// -----------------------------------------------------------------------------
// Immutability regressions for operation stamps shared by merge-tree metadata.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;
using TextSegmentModel = Microsoft.Office.Web.Fluid.MergeTree.TextSegment;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class StampImmutabilityTests
	{
		[Fact]
		public void OperationStamp_IsRecord_ValueEquality()
		{
			OperationStamp first = new()
			{
				Seq = 7,
				ClientId = 3,
				LocalSeq = 2,
			};
			OperationStamp second = new()
			{
				Seq = 7,
				ClientId = 3,
				LocalSeq = 2,
			};

			Assert.True(first == second);
			Assert.True(first.Equals(second));
			Assert.Equal(first, second);
		}

		[Fact]
		public void OperationStamp_WithExpression_CreatesNewInstance()
		{
			OperationStamp stamp = new()
			{
				Seq = MergeTreeModel.UnassignedSequenceNumber,
				ClientId = 4,
				LocalSeq = 9,
			};

			OperationStamp promoted = stamp with
			{
				Seq = 42,
				LocalSeq = null,
			};

			Assert.NotSame(stamp, promoted);
			Assert.Equal(MergeTreeModel.UnassignedSequenceNumber, stamp.Seq);
			Assert.Equal(9, stamp.LocalSeq);
			Assert.Equal(42, promoted.Seq);
			Assert.Null(promoted.LocalSeq);
		}

		[Fact]
		public void SharedStampBetweenSegmentAndPendingQueue_NotMutatedByPromotion()
		{
			Client client = CreateAckedClient("local", "abcdef");
			IMergeTreeRemoveMsg localRemove = client.RemoveText(1, 5);
			ISegment removed = SegmentWithText(client.MergeTree, "bcde");
			SetRemoveOperationStamp original = Assert.Single(removed.RemoveStamps.OfType<SetRemoveOperationStamp>());
			long? originalLocalSeq = original.LocalSeq;

			client.ApplyOp(localRemove, seq: 11, refSeq: 1, clientId: "local");

			SetRemoveOperationStamp promoted = Assert.Single(removed.RemoveStamps.OfType<SetRemoveOperationStamp>());
			Assert.NotSame(original, promoted);
			Assert.Equal(MergeTreeModel.UnassignedSequenceNumber, original.Seq);
			Assert.Equal(originalLocalSeq, original.LocalSeq);
			Assert.Equal(11, promoted.Seq);
			Assert.Null(promoted.LocalSeq);
		}

		[Fact]
		public void RemoveOperationStamp_IsImmutable()
		{
			RemoveOperationStamp first = new SetRemoveOperationStamp()
			{
				Seq = 3,
				ClientId = 2,
				LocalSeq = 1,
			};
			RemoveOperationStamp second = new SetRemoveOperationStamp()
			{
				Seq = 3,
				ClientId = 2,
				LocalSeq = 1,
			};

			RemoveOperationStamp changed = first with { Seq = 4 };

			Assert.True(first == second);
			Assert.True(first.Equals(second));
			Assert.NotSame(first, changed);
			Assert.Equal(3, first.Seq);
			Assert.Equal(4, changed.Seq);
			Assert.IsType<SetRemoveOperationStamp>(changed);
		}

		[Fact]
		public void SliceRemove_IsImmutable()
		{
			SliceRemoveOperationStamp first = new()
			{
				Seq = 3,
				ClientId = 2,
				LocalSeq = 1,
				StartSide = Side.Before,
				EndSide = Side.After,
			};
			SliceRemoveOperationStamp second = new()
			{
				Seq = 3,
				ClientId = 2,
				LocalSeq = 1,
				StartSide = Side.Before,
				EndSide = Side.After,
			};

			SliceRemoveOperationStamp changed = first with { StartSide = Side.After };

			Assert.True(first == second);
			Assert.True(first.Equals(second));
			Assert.NotSame(first, changed);
			Assert.Equal(Side.Before, first.StartSide);
			Assert.Equal(Side.After, changed.StartSide);
		}

		[Fact]
		public void AckPromotion_ReplacesStampInSegmentRemoveList_NotMutatesInPlace()
		{
			Client client = CreateAckedClient("client-a", "hello world");
			IMergeTreeRemoveMsg localRemove = client.RemoveText(0, 6);
			ISegment removed = SegmentWithText(client.MergeTree, "hello ");
			SetRemoveOperationStamp original = Assert.Single(removed.RemoveStamps.OfType<SetRemoveOperationStamp>());
			long? originalLocalSeq = original.LocalSeq;

			client.ApplyOp(new MergeTreeRemoveMsg() { Pos1 = 0, Pos2 = 6 }, seq: 10, refSeq: 1, clientId: "client-b");
			client.ApplyOp(localRemove, seq: 11, refSeq: 1, clientId: "client-a");

			SetRemoveOperationStamp promoted = Assert.Single(removed.RemoveStamps.OfType<SetRemoveOperationStamp>(), stamp => stamp.Seq == 11);
			Assert.NotSame(original, promoted);
			Assert.DoesNotContain(removed.RemoveStamps, stamp => ReferenceEquals(stamp, original));
			Assert.Equal(MergeTreeModel.UnassignedSequenceNumber, original.Seq);
			Assert.Equal(originalLocalSeq, original.LocalSeq);
			Assert.Null(promoted.LocalSeq);
			Assert.Contains(removed.RemoveStamps, stamp => stamp.Seq == 10 && stamp.ClientId != promoted.ClientId);
			Assert.Equal(new long[] { 10, 11 }, removed.RemoveStamps.Select(stamp => stamp.Seq));
		}

		[Fact]
		public void SnapshotLoad_StampsAreImmutable()
		{
			string snapshotJson = """
				{
					"version":"1",
					"segmentCount":1,
					"length":1,
					"startIndex":0,
					"segments":[
						{
							"json":"x",
							"seq":1,
							"client":"seed",
							"removedSeq":10,
							"movedSeq":12,
							"movedSeqs":[12],
							"movedClientIds":["move-client"]
						}
					],
					"headerMetadata":{
						"minSequenceNumber":0,
						"sequenceNumber":20,
						"orderedChunkMetadata":[{"id":"header"}],
						"totalLength":1,
						"totalSegmentCount":1
					}
				}
				""";
			Client client = new("loader");

			SharedStringSnapshotLoader.PopulateFromSnapshot(client, SharedStringSnapshotLoader.Load(snapshotJson));

			ISegment segment = Assert.Single(client.MergeTree.WalkAllSegments());
			Assert.All(segment.RemoveStamps, stamp =>
			{
				AssertInitOnlyProperty(stamp.GetType(), nameof(OperationStamp.Seq));
				AssertInitOnlyProperty(stamp.GetType(), nameof(OperationStamp.ClientId));
				AssertInitOnlyProperty(stamp.GetType(), nameof(OperationStamp.LocalSeq));
				if (stamp is SliceRemoveOperationStamp)
				{
					AssertInitOnlyProperty(stamp.GetType(), nameof(SliceRemoveOperationStamp.StartSide));
					AssertInitOnlyProperty(stamp.GetType(), nameof(SliceRemoveOperationStamp.EndSide));
				}
			});
		}

		[Fact]
		public void ComparePromotedVsOriginalStamp_ValuesDifferAsExpected()
		{
			SetRemoveOperationStamp original = new()
			{
				Seq = MergeTreeModel.UnassignedSequenceNumber,
				ClientId = 5,
				LocalSeq = 2,
			};
			SetRemoveOperationStamp expectedPromoted = new()
			{
				Seq = 9,
				ClientId = 5,
			};

			SetRemoveOperationStamp promoted = original with
			{
				Seq = 9,
				LocalSeq = null,
			};

			Assert.False(original == promoted);
			Assert.False(Stamps.Equal(original, promoted));
			Assert.Equal(expectedPromoted, promoted);
			Assert.True(Stamps.Equal(expectedPromoted, promoted));
		}

		private static Client CreateAckedClient(string clientId, string text)
		{
			Client client = new(clientId);
			IMergeTreeInsertMsg insert = client.InsertText(0, text);
			client.ApplyOp(insert, seq: 1, refSeq: 0, clientId: clientId);
			return client;
		}

		private static ISegment SegmentWithText(MergeTreeModel tree, string text)
		{
			return Assert.Single(
				tree.WalkAllSegments(),
				segment => segment is TextSegmentModel textSegment && textSegment.Text == text);
		}

		private static void AssertInitOnlyProperty(Type type, string propertyName)
		{
			PropertyInfo property = Assert.IsAssignableFrom<PropertyInfo>(type.GetProperty(propertyName));
			MethodInfo setter = Assert.IsAssignableFrom<MethodInfo>(property.SetMethod);
			Assert.Contains(typeof(IsExternalInit), setter.ReturnParameter.GetRequiredCustomModifiers());
		}
	}
}
