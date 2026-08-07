// -----------------------------------------------------------------------------
// Wire-format parity regressions for O-series SharedString audit findings.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;
using System.Text.Json;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;
using IntervalSide = Microsoft.Office.Web.Fluid.Intervals.Side;
using MergeTreeSide = Microsoft.Office.Web.Fluid.MergeTree.Side;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class OpsWireParityTests
	{
		[Fact]
		public void O2_TsIntervalAddMapEnvelope_DeserializesToIntervalAdd()
		{
			const string json = """
				{"type":"act","key":"comments","value":{"opName":"add","value":{"sequenceNumber":7,"start":2,"end":5,"intervalType":2,"stickiness":3,"startSide":1,"endSide":0,"properties":{"intervalId":"i1","referenceRangeLabels":["comments"],"author":"ada"}}}}
				""";

			IntervalAddOpMsg op = Assert.IsType<IntervalAddOpMsg>(SharedStringOpSerializer.Deserialize(json));

			Assert.Equal("comments", op.CollectionName);
			Assert.Equal("i1", op.IntervalId);
			Assert.Equal(2, op.Start);
			Assert.Equal(5, op.End);
			Assert.Equal(IntervalType.SlideOnRemove, op.IntervalType);
			Assert.Equal(IntervalStickiness.Full, op.Stickiness);
			Assert.Equal(IntervalSide.After, op.StartSide);
			Assert.Equal(IntervalSide.Before, op.EndSide);
			Assert.Equal("ada", Assert.IsType<string>(op.Props!["author"]));
			Assert.False(op.Props.ContainsKey("intervalId"));
			Assert.False(op.Props.ContainsKey("referenceRangeLabels"));
		}

		[Fact]
		public void O2_IntervalPropertyChange_CanSerializeTsMapEnvelope()
		{
			IntervalPropertyChangedOpMsg op = new()
			{
				CollectionName = "comments",
				IntervalId = "i1",
				Props = new PropertySet()
				{
					["color"] = "blue",
				},
			};

			string json = SharedStringOpSerializer.SerializeIntervalCollectionOperation(op);
			using JsonDocument document = JsonDocument.Parse(json);
			JsonElement root = document.RootElement;
			JsonElement payload = root.GetProperty("value").GetProperty("value");

			Assert.Equal("act", root.GetProperty("type").GetString());
			Assert.Equal("comments", root.GetProperty("key").GetString());
			Assert.Equal("change", root.GetProperty("value").GetProperty("opName").GetString());
			Assert.Equal("i1", payload.GetProperty("properties").GetProperty("intervalId").GetString());
			Assert.Equal("comments", payload.GetProperty("properties").GetProperty("referenceRangeLabels")[0].GetString());
			Assert.Equal("blue", payload.GetProperty("properties").GetProperty("color").GetString());
			Assert.DoesNotContain("intervalOpKind", json);
		}

		[Fact]
		public void O4_RelativeInsertFields_RoundTripOnWire()
		{
			MergeTreeInsertMsg op = new()
			{
				RelativePos1 = new PropertySet()
				{
					["id"] = "marker-1",
					["before"] = true,
					["offset"] = 2,
				},
				Seg = "X",
			};

			string json = SharedStringOpSerializer.Serialize(op);
			using JsonDocument document = JsonDocument.Parse(json);
			JsonElement relativePos = document.RootElement.GetProperty("relativePos1");
			Assert.False(document.RootElement.TryGetProperty("pos1", out _));
			Assert.Equal("marker-1", relativePos.GetProperty("id").GetString());
			Assert.True(relativePos.GetProperty("before").GetBoolean());
			Assert.Equal(2, relativePos.GetProperty("offset").GetInt32());

			MergeTreeInsertMsg roundTripped = Assert.IsType<MergeTreeInsertMsg>(SharedStringOpSerializer.Deserialize(json));
			PropertySet roundTrippedRelativePos = Assert.IsType<PropertySet>(roundTripped.RelativePos1);
			Assert.Equal("marker-1", Assert.IsType<string>(roundTrippedRelativePos["id"]));
			Assert.True(Assert.IsType<bool>(roundTrippedRelativePos["before"]));
			Assert.Equal(2, Assert.IsType<int>(roundTrippedRelativePos["offset"]));
		}

		[Fact]
		public void O4_OpBuilderAnnotateMarker_UsesRelativePositions()
		{
			Marker marker = Marker.Make(
				ReferenceType.Simple,
				new PropertySet()
				{
					[Marker.ReservedMarkerIdKey] = "marker-1",
				});

			MergeTreeAnnotateMsg op = Assert.IsType<MergeTreeAnnotateMsg>(
				OpBuilder.CreateAnnotateMarkerOp(
					marker,
					new PropertySet()
					{
						["color"] = "red",
					}));

			string json = SharedStringOpSerializer.Serialize(op);
			using JsonDocument document = JsonDocument.Parse(json);
			Assert.False(document.RootElement.TryGetProperty("pos1", out _));
			Assert.False(document.RootElement.TryGetProperty("pos2", out _));
			Assert.True(document.RootElement.GetProperty("relativePos1").GetProperty("before").GetBoolean());
			Assert.Equal("marker-1", document.RootElement.GetProperty("relativePos2").GetProperty("id").GetString());

			MergeTreeAnnotateMsg roundTripped = Assert.IsType<MergeTreeAnnotateMsg>(SharedStringOpSerializer.Deserialize(json));
			Assert.NotNull(roundTripped.RelativePos1);
			Assert.NotNull(roundTripped.RelativePos2);
		}

		[Fact]
		public void O5_NestedGroup_SerializerRoundTripsTsShape()
		{
			MergeTreeGroupMsg nested = OpBuilder.CreateGroupOp(OpBuilder.CreateInsertOp(0, "a"));
			MergeTreeGroupMsg outer = OpBuilder.CreateGroupOp(nested, OpBuilder.CreateInsertOp(1, "b"));

			string json = SharedStringOpSerializer.Serialize(outer);
			MergeTreeGroupMsg roundTripped = Assert.IsType<MergeTreeGroupMsg>(SharedStringOpSerializer.Deserialize(json));

			Assert.IsType<MergeTreeGroupMsg>(roundTripped.Ops[0]);
			Assert.IsType<MergeTreeInsertMsg>(roundTripped.Ops[1]);
		}

		[Fact]
		public void O5_GroupRejectsIntervalOpMembers()
		{
			// TS-parity: MergeTreeGroupMsg (packages/dds/merge-tree/src/ops.ts IMergeTreeGroupMsg)
			// only carries merge-tree delta ops. Interval ops go on the wire individually as
			// IntervalCollectionMap "act" envelopes; they must never be batched into a group.
			MergeTreeGroupMsg group = OpBuilder.CreateGroupOp(
				new IntervalAddOpMsg()
				{
					CollectionName = "comments",
					IntervalId = "i1",
					Start = 1,
					End = 2,
				});

			Assert.Throws<System.Text.Json.JsonException>(
				() => SharedStringOpSerializer.Serialize(group));
		}

		[Fact]
		public void O6_ComplexProperties_DeserializeToClrValues()
		{
			const string json = """
				{"type":2,"pos1":0,"pos2":1,"props":{"count":1,"nested":{"flag":true},"items":["x",2,null]}}
				""";

			MergeTreeAnnotateMsg op = Assert.IsType<MergeTreeAnnotateMsg>(SharedStringOpSerializer.Deserialize(json));

			AssertNoJsonElements(op.Props);
			Assert.Equal(1, Assert.IsType<int>(op.Props!["count"]));
			PropertySet nested = Assert.IsType<PropertySet>(op.Props["nested"]);
			Assert.True(Assert.IsType<bool>(nested["flag"]));
			List<object?> items = Assert.IsType<List<object?>>(op.Props["items"]);
			Assert.Equal("x", Assert.IsType<string>(items[0]));
			Assert.Equal(2, Assert.IsType<int>(items[1]));
			Assert.Null(items[2]);
		}

		[Fact]
		public void O6_WriteJsonValue_SupportsClrDictionariesAndValueTypeArrays()
		{
			MergeTreeAnnotateMsg op = OpBuilder.CreateAnnotateRangeOp(
				0,
				1,
				new PropertySet()
				{
					["dict"] = new Dictionary<string, string>()
					{
						["mode"] = "plain",
					},
					["numbers"] = new int[] { 1, 2, 3 },
				});

			string json = SharedStringOpSerializer.Serialize(op);
			using JsonDocument document = JsonDocument.Parse(json);

			Assert.Equal("plain", document.RootElement.GetProperty("props").GetProperty("dict").GetProperty("mode").GetString());
			Assert.Equal(2, document.RootElement.GetProperty("props").GetProperty("numbers")[1].GetInt32());
		}

		[Fact]
		public void O7_SidedObliterateMissingBefore_DefaultsToAfterLikeTs()
		{
			const string json = "{\"type\":5,\"pos1\":{\"pos\":1},\"pos2\":{\"pos\":3}}";

			MergeTreeObliterateSidedMsg op = Assert.IsType<MergeTreeObliterateSidedMsg>(SharedStringOpSerializer.Deserialize(json));

			Assert.Equal(MergeTreeSide.After, op.Pos1.Side);
			Assert.Equal(MergeTreeSide.After, op.Pos2.Side);
		}

		[Fact]
		public void O7_OpBuilderNumericSidedObliterate_NormalizesEndInclusive()
		{
			MergeTreeObliterateSidedMsg op = OpBuilder.CreateObliterateRangeOpSided(2, 5);

			string json = SharedStringOpSerializer.Serialize(op);
			using JsonDocument document = JsonDocument.Parse(json);

			Assert.Equal(2, document.RootElement.GetProperty("pos1").GetProperty("pos").GetInt32());
			Assert.True(document.RootElement.GetProperty("pos1").GetProperty("before").GetBoolean());
			Assert.Equal(4, document.RootElement.GetProperty("pos2").GetProperty("pos").GetInt32());
			Assert.False(document.RootElement.GetProperty("pos2").GetProperty("before").GetBoolean());
		}

		[Fact]
		public void O8_OperationStampComparison_MatchesTsOrdering()
		{
			InsertOperationStamp acked = new()
			{
				Seq = 5,
				ClientId = 1,
			};
			InsertOperationStamp laterAcked = new()
			{
				Seq = 6,
				ClientId = 1,
			};
			InsertOperationStamp local = new()
			{
				Seq = Constants.UnassignedSequenceNumber,
				ClientId = 1,
				LocalSeq = 1,
			};

			Assert.True(Stamps.LessThan(acked, laterAcked));
			Assert.True(Stamps.LessThan(acked, local));
			Assert.True(Stamps.GreaterThan(local, acked));
			Assert.True(Stamps.Equal(acked, new InsertOperationStamp() { Seq = 5, ClientId = 1 }));
		}

		[Fact]
		public void O9_OpBuilderCoreOps_ProduceExpectedWire()
		{
			MergeTreeGroupMsg group = OpBuilder.CreateGroupOp(
				OpBuilder.CreateInsertOp(0, "a"),
				OpBuilder.CreateRemoveRangeOp(1, 2),
				OpBuilder.CreateAnnotateRangeOp(0, 1, new PropertySet() { ["bold"] = true }));

			string json = SharedStringOpSerializer.Serialize(group);
			using JsonDocument document = JsonDocument.Parse(json);
			JsonElement ops = document.RootElement.GetProperty("ops");

			Assert.Equal(0, ops[0].GetProperty("type").GetInt32());
			Assert.Equal("a", ops[0].GetProperty("seg").GetString());
			Assert.Equal(1, ops[1].GetProperty("type").GetInt32());
			Assert.Equal(2, ops[2].GetProperty("type").GetInt32());
			Assert.True(ops[2].GetProperty("props").GetProperty("bold").GetBoolean());
		}

		[Fact]
		public void O9_OpBuilderAnnotateMarker_ReturnsNullWithoutMarkerId()
		{
			Assert.Null(OpBuilder.CreateAnnotateMarkerOp(Marker.Make(ReferenceType.Simple), new PropertySet()));
		}

		private static void AssertNoJsonElements(object? value)
		{
			Assert.False(value is JsonElement);
			if (value is IReadOnlyDictionary<string, object?> dictionary)
			{
				foreach (object? child in dictionary.Values)
				{
					AssertNoJsonElements(child);
				}
			}
			else if (value is IEnumerable<object?> list)
			{
				foreach (object? child in list)
				{
					AssertNoJsonElements(child);
				}
			}
		}
	}
}
