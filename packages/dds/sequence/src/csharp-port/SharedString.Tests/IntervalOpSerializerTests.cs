// -----------------------------------------------------------------------------
// Wave 12b interval op serialization tests for SharedString POC.
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
		public void IntervalAddOp_UnknownKind_ThrowsOcsException()
		{
			OcsException exception = Assert.Throws<OcsException>(
				() => SharedStringOpSerializer.Deserialize("{\"type\":10,\"intervalOpKind\":99,\"collection\":\"comments\",\"id\":\"c1\",\"start\":5,\"end\":10,\"props\":{}}"));

			Assert.Equal(OcsGateErrorCode.UnknownOp, exception.ErrorCode);
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
			// TS-parity: intervals/sequenceInterval.ts:468 sets sequenceNumber = client.getCurrentSeq()
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
	}
}
