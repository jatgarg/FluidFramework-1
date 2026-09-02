// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------

using System.Text.Json;
using Microsoft.Office.Web.Fluid;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class IntervalOpSerializerTests
	{
		[Fact]
		public void IntervalAddOp_JsonRoundTrip()
		{
			IntervalAddOpMsg op = new IntervalAddOpMsg()
			{
				CollectionName = "comments",
				IntervalId = "c1",
				Start = 5,
				End = 10,
				IntervalType = IntervalType.Transient,
				Props = new PropertySet()
				{
					["author"] = "ada",
					["resolved"] = false,
				},
			};

			IntervalAddOpMsg roundTripped = RoundTrip(op);

			Assert.Equal(IntervalOpKind.Add, roundTripped.IntervalOpKind);
			Assert.Equal("comments", roundTripped.CollectionName);
			Assert.Equal("c1", roundTripped.IntervalId);
			Assert.Equal(5, roundTripped.Start);
			Assert.Equal(10, roundTripped.End);
			Assert.Equal(IntervalType.Transient, roundTripped.IntervalType);
			PropertySet? props = roundTripped.Props;
			Assert.NotNull(props);
			Assert.Equal(2, props!.Count);
			Assert.Equal("ada", Assert.IsType<string>(props["author"]));
			Assert.False(Assert.IsType<bool>(props["resolved"]));
		}

		[Fact]
		public void IntervalDeleteOp_JsonRoundTrip()
		{
			IntervalDeleteOpMsg op = new IntervalDeleteOpMsg()
			{
				CollectionName = "comments",
				IntervalId = "c1",
			};

			IntervalDeleteOpMsg roundTripped = RoundTrip(op);

			Assert.Equal(IntervalOpKind.Delete, roundTripped.IntervalOpKind);
			Assert.Equal("comments", roundTripped.CollectionName);
			Assert.Equal("c1", roundTripped.IntervalId);
		}

		[Fact]
		public void IntervalChangeOp_NullEnd_RoundTrips()
		{
			IntervalChangeOpMsg op = new IntervalChangeOpMsg()
			{
				CollectionName = "comments",
				IntervalId = "c1",
				Start = 5,
				End = null,
			};

			string json = SharedStringOpSerializer.Serialize(op);
			IntervalChangeOpMsg roundTripped = Assert.IsType<IntervalChangeOpMsg>(
				SharedStringOpSerializer.Deserialize(json));

			Assert.Equal(IntervalOpKind.Change, roundTripped.IntervalOpKind);
			Assert.Equal("comments", roundTripped.CollectionName);
			Assert.Equal("c1", roundTripped.IntervalId);
			AssertNullableInt(5, roundTripped.Start);
			Assert.Null(roundTripped.End);
			// TS wire: op is wrapped in the IntervalCollectionMap "act" envelope, and
			// interval.serialize() (packages/dds/sequence/src/intervals/sequenceInterval.ts)
			// omits the endpoint when it's undefined. Assert absence, not null.
			using JsonDocument document = JsonDocument.Parse(json);
			JsonElement payload = document.RootElement.GetProperty("value").GetProperty("value");
			Assert.False(payload.TryGetProperty("end", out _));
		}

		[Fact]
		public void IntervalPropertyChangedOp_NullPropDeletes_RoundTrips()
		{
			IntervalPropertyChangedOpMsg op = new IntervalPropertyChangedOpMsg()
			{
				CollectionName = "comments",
				IntervalId = "c1",
				Props = new PropertySet()
				{
					["color"] = null,
					["resolved"] = true,
				},
			};

			IntervalPropertyChangedOpMsg roundTripped = RoundTrip(op);

			Assert.Equal(IntervalOpKind.PropertyChanged, roundTripped.IntervalOpKind);
			Assert.Equal("comments", roundTripped.CollectionName);
			Assert.Equal("c1", roundTripped.IntervalId);
			PropertySet props = roundTripped.Props;
			Assert.Equal(2, props.Count);
			Assert.True(props.ContainsKey("color"));
			Assert.Null(props["color"]);
			Assert.True(Assert.IsType<bool>(props["resolved"]));
		}

		[Fact]
		public void IntervalAddOp_UnknownKind_ThrowsLoggingError()
		{
			LoggingError exception = Assert.Throws<LoggingError>(
				() => SharedStringOpSerializer.Deserialize("{\"type\":10,\"intervalOpKind\":99,\"collection\":\"comments\",\"id\":\"c1\",\"start\":5,\"end\":10,\"props\":{}}"));

		}

		[Fact]
		public void IntervalOps_WireShape_MatchesTSIntervalCollectionMap()
		{
			// TS ref: packages/dds/sequence/src/intervalCollection.ts submitDelta emits
			// { opName: "add"/"delete"/"change", value: ... }, wrapped by
			// packages/dds/sequence/src/intervalCollectionMap.ts under
			// { type: "act", key: <collection>, value: { opName, value } }.
			AssertIntervalWireShape(
				new IntervalAddOpMsg()
				{
					CollectionName = "comments",
					IntervalId = "c1",
					Start = 5,
					End = 10,
					Props = new PropertySet(),
				},
				"add");
			AssertIntervalWireShape(
				new IntervalDeleteOpMsg()
				{
					CollectionName = "comments",
					IntervalId = "c1",
				},
				"delete");
			AssertIntervalWireShape(
				new IntervalChangeOpMsg()
				{
					CollectionName = "comments",
					IntervalId = "c1",
					Start = 5,
					End = null,
				},
				"change");
			AssertIntervalWireShape(
				new IntervalPropertyChangedOpMsg()
				{
					CollectionName = "comments",
					IntervalId = "c1",
					Props = new PropertySet(),
				},
				// TS-parity: propertyChanged ops go on the wire as opName "change" too.
				"change");
		}

		private static T RoundTrip<T>(T op)
			where T : IMergeTreeOp
		{
			return Assert.IsType<T>(
				SharedStringOpSerializer.Deserialize(SharedStringOpSerializer.Serialize(op)));
		}

		private static void AssertNullableInt(int expected, int? actual)
		{
			Assert.True(actual.HasValue);
			Assert.Equal(expected, actual.Value);
		}

		private static void AssertTypeCode(IMergeTreeOp op, int expectedType, int expectedIntervalOpKind)
		{
			string json = SharedStringOpSerializer.Serialize(op);
			using JsonDocument document = JsonDocument.Parse(json);
			Assert.Equal(expectedType, document.RootElement.GetProperty("type").GetInt32());
			Assert.Equal(expectedIntervalOpKind, document.RootElement.GetProperty("intervalOpKind").GetInt32());
		}

		private static void AssertIntervalWireShape(IMergeTreeOp op, string expectedOpName)
		{
			string json = SharedStringOpSerializer.Serialize(op);
			using JsonDocument document = JsonDocument.Parse(json);
			JsonElement root = document.RootElement;
			Assert.Equal("act", root.GetProperty("type").GetString());
			Assert.Equal("comments", root.GetProperty("key").GetString());
			JsonElement inner = root.GetProperty("value");
			Assert.Equal(expectedOpName, inner.GetProperty("opName").GetString());
		}

		[Fact]
		public void IntervalOp_WireCarriesClientCurrentSequenceNumber()
		{
			// TS-parity: intervals/sequenceInterval.ts sets sequenceNumber = client.getCurrentSeq()
			// on the serialized interval. Reconnect/rebase logic reads this.
			IntervalAddOpMsg op = new IntervalAddOpMsg()
			{
				CollectionName = "comments",
				IntervalId = "c1",
				Start = 2,
				End = 4,
				Props = new PropertySet(),
			};

			string json = SharedStringOpSerializer.Serialize(op, registry: null, currentSequenceNumber: 42);
			using JsonDocument document = JsonDocument.Parse(json);
			JsonElement payload = document.RootElement.GetProperty("value").GetProperty("value");
			Assert.Equal(42, payload.GetProperty("sequenceNumber").GetInt64());
		}

		[Fact]
		public void IntervalOp_WireDefaultsSequenceNumberToZeroWhenUnknown()
		{
			// Backwards compat: callers that don't have a client context (tests, ad-hoc serialization)
			// still get a valid wire message; TS treats missing sequenceNumber == 0.
			IntervalDeleteOpMsg op = new IntervalDeleteOpMsg()
			{
				CollectionName = "comments",
				IntervalId = "c1",
			};

			string json = SharedStringOpSerializer.Serialize(op);
			using JsonDocument document = JsonDocument.Parse(json);
			JsonElement payload = document.RootElement.GetProperty("value").GetProperty("value");
			Assert.Equal(0, payload.GetProperty("sequenceNumber").GetInt64());
		}

		[Fact]
		public void PropertySet_CLRSmallNumericTypes_SerializeAsNumbers()
		{
			// TS ref: packages/dds/merge-tree/src/properties.ts PropertySet — TS
			// treats all numeric properties as ordinary JSON numbers. CLR small
			// integer types (short, byte, sbyte, ushort, uint, ulong) are
			// numerically identical to their JSON representation. Previously the
			// serializer only handled int/long/double/float/decimal and threw for
			// the others AFTER the local tree had already been mutated.
			SharedString sharedString = new();
			sharedString.InsertText(0, "x", new PropertySet()
			{
				["s"] = (short)1,
				["us"] = (ushort)2,
				["b"] = (byte)3,
				["sb"] = (sbyte)4,
				["ui"] = (uint)5,
				["ul"] = (ulong)6,
			});

			// If any small-numeric type had thrown, InsertText would have thrown
			// too. Just reaching this point verifies the fix.
			Assert.Equal("x", sharedString.GetText());
		}

		[Fact]
		public void IntervalAddOp_SentinelEndpoints_Deserialize()
		{
			// TS ref: packages/dds/sequence/src/intervals/sequenceInterval.ts
			// createPositionReference — string endpoints "start" and "end"
			// anchor to the synthetic start-of-tree / end-of-tree segment.
			// The port routes these through EndpointSentinel on the op DTO;
			// the numeric Start/End field is meaningless when Sentinel is set.
			const string wire =
				"{\"type\":\"act\",\"key\":\"comments\",\"value\":{\"opName\":\"add\",\"value\":{" +
				"\"sequenceNumber\":0,\"intervalType\":2," +
				"\"start\":\"start\",\"end\":\"end\"," +
				"\"properties\":{\"intervalId\":\"whole\",\"referenceRangeLabels\":[\"comments\"]}}}}";

			IMergeTreeOp op = SharedStringOpSerializer.Deserialize(wire);
			IntervalAddOpMsg add = Assert.IsType<IntervalAddOpMsg>(op);
			Assert.Equal(EndpointSentinel.Start, add.StartSentinel);
			Assert.Equal(EndpointSentinel.End, add.EndSentinel);
			Assert.Equal(Intervals.Side.After, add.StartSide);
			Assert.Equal(Intervals.Side.Before, add.EndSide);
		}

		[Fact]
		public void IntervalAddOp_SentinelEndpoints_RoundTripSerializes()
		{
			// regression. Emitting an op DTO with sentinel-encoded
			// endpoints must serialize back as the string sentinels TS peers
			// understand.
			IntervalAddOpMsg add = new()
			{
				CollectionName = "comments",
				IntervalId = "whole",
				StartSentinel = EndpointSentinel.Start,
				EndSentinel = EndpointSentinel.End,
				StartSide = Intervals.Side.After,
				EndSide = Intervals.Side.Before,
			};

			string json = SharedStringOpSerializer.Serialize(add);
			using JsonDocument document = JsonDocument.Parse(json);
			JsonElement value = document.RootElement.GetProperty("value").GetProperty("value");
			Assert.Equal("start", value.GetProperty("start").GetString());
			Assert.Equal("end", value.GetProperty("end").GetString());
		}

		[Fact]
		public void SharedString_ApplyRemoteIntervalAdd_WithSentinelEndpoints_Succeeds()
		{
			// (reopened v1). The codec now decodes the wire
			// sentinels, and IntervalCollection routes them to
			// MergeTree.CreateReferencePositionAtEndpoint. Applying such an op
			// used to throw because the position validator rejected the
			// decoded -1.
			SharedString sharedString = new();
			sharedString.InsertText(0, "abcdef");
			const string wire =
				"{\"type\":\"act\",\"key\":\"comments\",\"value\":{\"opName\":\"add\",\"value\":{" +
				"\"sequenceNumber\":0,\"intervalType\":2," +
				"\"start\":\"start\",\"end\":\"end\"," +
				"\"properties\":{\"intervalId\":\"whole\",\"referenceRangeLabels\":[\"comments\"]}}}}";

			sharedString.ProcessDataObjectOp(
				new SequencedDocumentMessageDescriptor(
					SequenceNumber.ForTesting(clientSeq: 0, refSeq: 0, seq: 1),
					OpOrigin.Remote,
					"remote-client"),
				wire);

			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval? interval = collection.GetIntervalById("whole");
			Assert.NotNull(interval);
			// Interval anchored to start-of-tree and end-of-tree resolves to
			// positions 0 and length respectively.
			Assert.Equal(0, interval!.StartPosition);
			Assert.Equal(sharedString.GetLength(), interval.EndPosition);
		}

		[Fact]
		public void IntervalChangeOp_CombinedEndpointsAndProps_DeserializePreservesUserProps()
		{
			// TS ref: packages/dds/sequence/src/intervalCollection.ts
			// changeInterval — one op carries both endpoint delta and property
			// delta; user props on the combined form must round-trip.
			const string wire =
				"{\"type\":\"act\",\"key\":\"comments\",\"value\":{\"opName\":\"change\",\"value\":{" +
				"\"sequenceNumber\":0,\"intervalType\":2," +
				"\"start\":2,\"end\":4," +
				"\"properties\":{\"intervalId\":\"i1\",\"color\":\"red\",\"referenceRangeLabels\":[\"comments\"]}}}}";

			IMergeTreeOp op = SharedStringOpSerializer.Deserialize(wire);
			IntervalChangeOpMsg change = Assert.IsType<IntervalChangeOpMsg>(op);
			Assert.Equal(2, change.Start);
			Assert.Equal(4, change.End);
			Assert.NotNull(change.Props);
			Assert.Equal("red", change.Props!["color"]);
		}

		[Fact]
		public void IntervalChangeOp_CombinedEndpointsAndProps_SerializesAsSingleOp()
		{
			IntervalChangeOpMsg change = new()
			{
				CollectionName = "comments",
				IntervalId = "i1",
				Start = 2,
				End = 4,
				Props = new PropertySet()
				{
					["color"] = "red",
				},
			};

			string json = SharedStringOpSerializer.Serialize(change);

			using JsonDocument doc = JsonDocument.Parse(json);
			JsonElement value = doc.RootElement.GetProperty("value").GetProperty("value");
			Assert.Equal(2, value.GetProperty("start").GetInt32());
			Assert.Equal(4, value.GetProperty("end").GetInt32());
			Assert.Equal("red", value.GetProperty("properties").GetProperty("color").GetString());
		}
	}
}
