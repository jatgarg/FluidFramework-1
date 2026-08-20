#nullable enable

using System.Collections.Generic;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class StampsAndPropertiesTestsFromTS
	{
		// Ported from packages/dds/merge-tree/src/test/stamps.spec.ts — "returns true for reference equal stamps"
		[Fact]
		public void StampEquality_ReferenceEqualStampsMatch()
		{
			InsertOperationStamp stamp = AckedStamp(seq: 1, clientId: 1);

			Assert.True(Stamps.Equal(stamp, stamp));
		}

		// Ported from packages/dds/merge-tree/src/test/stamps.spec.ts — "returns true for equal stamps"
		[Fact]
		public void StampEquality_EqualValuesMatch()
		{
			Assert.True(Stamps.Equal(AckedStamp(seq: 1, clientId: 1), AckedStamp(seq: 1, clientId: 1)));
			Assert.True(Stamps.Equal(LocalStamp(localSeq: 1), LocalStamp(localSeq: 1)));
		}

		// Ported from packages/dds/merge-tree/src/test/stamps.spec.ts — "returns false for different stamps"
		[Fact]
		public void StampEquality_DifferentSeqClientOrLocalSeqDoNotMatch()
		{
			Assert.False(Stamps.Equal(AckedStamp(seq: 1, clientId: 1), AckedStamp(seq: 2, clientId: 1)));
			Assert.False(Stamps.Equal(AckedStamp(seq: 1, clientId: 1), AckedStamp(seq: 1, clientId: 2)));
			Assert.False(Stamps.Equal(AckedStamp(seq: 1, clientId: 1), LocalStamp(localSeq: 1)));
			Assert.False(Stamps.Equal(LocalStamp(localSeq: 1), LocalStamp(localSeq: 2)));
		}

		// Ported from packages/dds/merge-tree/src/test/stamps.spec.ts — "orders stamps correctly"
		[Fact]
		public void StampComparison_OrdersAckedBeforeLocalBySeqAndLocalSeq()
		{
			List<InsertOperationStamp> stamps = new()
			{
				AckedStamp(seq: 1, clientId: 1),
				AckedStamp(seq: 2, clientId: 2),
				AckedStamp(seq: 3, clientId: 1),
				LocalStamp(localSeq: 1),
				LocalStamp(localSeq: 2),
			};

			for (int i = 0; i < stamps.Count - 1; i++)
			{
				Assert.True(Stamps.LessThan(stamps[i], stamps[i + 1]));
				Assert.True(Stamps.GreaterThan(stamps[i + 1], stamps[i]));
				Assert.Equal(-1, Stamps.Compare(stamps[i], stamps[i + 1]));
				Assert.Equal(1, Stamps.Compare(stamps[i + 1], stamps[i]));
			}
		}

		// Ported from packages/dds/merge-tree/src/test/stamps.spec.ts — "compare can sort lists"
		[Fact]
		public void StampComparison_CompareCanSortLists()
		{
			InsertOperationStamp acked1 = AckedStamp(seq: 1, clientId: 1);
			InsertOperationStamp acked2 = AckedStamp(seq: 2, clientId: 2);
			InsertOperationStamp acked3 = AckedStamp(seq: 3, clientId: 1);
			InsertOperationStamp local1 = LocalStamp(localSeq: 1);
			InsertOperationStamp local2 = LocalStamp(localSeq: 2);
			List<InsertOperationStamp> stamps = new()
			{
				acked3,
				local1,
				acked1,
				local2,
				acked2,
			};

			Stamps.SortList(stamps);

			Assert.Equal(new[] { acked1, acked2, acked3, local1, local2 }, stamps);
		}

		// Ported from packages/dds/merge-tree/src/test/stamps.spec.ts — "inserts acked before unacked"
		[Fact]
		public void SpliceIntoList_InsertsAckedBeforeUnacked()
		{
			InsertOperationStamp acked1 = AckedStamp(seq: 1, clientId: 1);
			InsertOperationStamp acked2 = AckedStamp(seq: 2, clientId: 2);
			InsertOperationStamp acked3 = AckedStamp(seq: 3, clientId: 1);
			InsertOperationStamp local1 = LocalStamp(localSeq: 1);
			List<InsertOperationStamp> stamps = new() { acked1, acked2, local1 };

			Stamps.SpliceIntoList(stamps, acked3);

			Assert.Equal(new[] { acked1, acked2, acked3, local1 }, stamps);
		}

		// Ported from packages/dds/merge-tree/src/test/properties.spec.ts — "simple properties match"
		[Fact]
		public void MatchProperties_SimplePropertiesMatch()
		{
			Assert.True(PropertyMap.MatchProperties(
				new PropertySet() { ["a"] = "a" },
				new PropertySet() { ["a"] = "a" }));
		}

		// Ported from packages/dds/merge-tree/src/test/properties.spec.ts — "simple properties don't match"
		[Fact]
		public void MatchProperties_SimplePropertiesDoNotMatch()
		{
			Assert.False(PropertyMap.MatchProperties(
				new PropertySet() { ["a"] = "a" },
				new PropertySet() { ["a"] = "b" }));
		}

		// Ported from packages/dds/merge-tree/src/test/properties.spec.ts — "keys don't match"
		[Fact]
		public void MatchProperties_KeysDoNotMatch()
		{
			Assert.False(PropertyMap.MatchProperties(
				new PropertySet() { ["a"] = "a" },
				new PropertySet() { ["b"] = "a" }));
		}

		// Ported from packages/dds/merge-tree/src/test/properties.spec.ts — "complex properties match"
		[Fact]
		public void MatchProperties_ComplexPropertiesMatch()
		{
			Assert.True(PropertyMap.MatchProperties(
				new PropertySet()
				{
					["c"] = new Dictionary<string, object?>() { ["a"] = "a" },
				},
				new PropertySet()
				{
					["c"] = new Dictionary<string, object?>() { ["a"] = "a" },
				}));
		}

		// Ported from packages/dds/merge-tree/src/test/properties.spec.ts — "undefined and simple properties don't match"
		[Fact]
		public void MatchProperties_NullAndSimplePropertiesDoNotMatch()
		{
			Assert.False(PropertyMap.MatchProperties(null, new PropertySet() { ["a"] = "a" }));
		}

		// Ported from packages/dds/merge-tree/src/test/properties.spec.ts — "undefined and empty properties match"
		[Fact]
		public void MatchProperties_NullAndEmptyPropertiesMatch()
		{
			Assert.True(PropertyMap.MatchProperties(null, new PropertySet()));
		}

		private static InsertOperationStamp AckedStamp(long seq, int clientId)
		{
			return new InsertOperationStamp()
			{
				Seq = seq,
				ClientId = clientId,
			};
		}

		private static InsertOperationStamp LocalStamp(long localSeq)
		{
			return new InsertOperationStamp()
			{
				Seq = MergeTreeModel.UnassignedSequenceNumber,
				ClientId = 1,
				LocalSeq = localSeq,
			};
		}
	}
}
