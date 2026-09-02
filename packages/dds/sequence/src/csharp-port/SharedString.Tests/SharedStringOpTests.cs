// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------

using System.Collections.Generic;
using System.Text.Json;
using Microsoft.Office.Web.Fluid;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SharedStringOpTests
	{
		[Fact]
		public void InsertOp_JsonRoundTrip_PreservesFields()
		{
			MergeTreeInsertMsg op = new MergeTreeInsertMsg()
			{
				Pos1 = 5,
				Seg = new TextSegment("hello"),
			};

			MergeTreeInsertMsg roundTripped = Assert.IsType<MergeTreeInsertMsg>(
				SharedStringOpSerializer.Deserialize(SharedStringOpSerializer.Serialize(op)));

			AssertNullableInt(5, roundTripped.Pos1);
			AssertTextSegment(roundTripped.Seg, "hello");
		}

		[Fact]
		public void RemoveOp_JsonRoundTrip_PreservesFields()
		{
			MergeTreeRemoveMsg op = new MergeTreeRemoveMsg()
			{
				Pos1 = 3,
				Pos2 = 8,
			};

			MergeTreeRemoveMsg roundTripped = Assert.IsType<MergeTreeRemoveMsg>(
				SharedStringOpSerializer.Deserialize(SharedStringOpSerializer.Serialize(op)));

			AssertNullableInt(3, roundTripped.Pos1);
			AssertNullableInt(8, roundTripped.Pos2);
		}

		[Fact]
		public void InsertOp_SegmentIsPlainString_Deserializes()
		{
			MergeTreeInsertMsg op = Assert.IsType<MergeTreeInsertMsg>(
				SharedStringOpSerializer.Deserialize("{\"type\":0,\"pos1\":0,\"seg\":\"hello\"}"));

			AssertNullableInt(0, op.Pos1);
			AssertTextSegment(op.Seg, "hello");
		}

		[Fact]
		public void InsertOp_SegmentIsObject_Deserializes()
		{
			MergeTreeInsertMsg op = Assert.IsType<MergeTreeInsertMsg>(
				SharedStringOpSerializer.Deserialize("{\"type\":0,\"pos1\":0,\"seg\":{\"text\":\"hello\"}}"));

			AssertNullableInt(0, op.Pos1);
			AssertTextSegment(op.Seg, "hello");
		}

		[Fact]
		public void InsertOp_SegmentWithProps_RoundTripsProps()
		{
			PropertySet props = new PropertySet()
			{
				["color"] = "red",
				["font"] = "arial",
			};
			MergeTreeInsertMsg op = new MergeTreeInsertMsg()
			{
				Pos1 = 1,
				Seg = new TextSegment("hello", props),
			};

			MergeTreeInsertMsg roundTripped = Assert.IsType<MergeTreeInsertMsg>(
				SharedStringOpSerializer.Deserialize(SharedStringOpSerializer.Serialize(op)));

			TextSegment segment = AssertTextSegment(roundTripped.Seg, "hello");
			Assert.NotNull(segment.Properties);
			Assert.Equal("red", Assert.IsType<string>(segment.Properties!["color"]));
			Assert.Equal("arial", Assert.IsType<string>(segment.Properties["font"]));
		}

		[Fact]
		public void Deserialize_UnknownType_ThrowsLoggingError()
		{
			LoggingError exception = Assert.Throws<LoggingError>(
				() => SharedStringOpSerializer.Deserialize("{\"type\":99,\"pos1\":0}"));

		}

		[Fact]
		public void Serialize_MergeTreeInsertMsg_HasTypeFieldFirst()
		{
			string insertJson = SharedStringOpSerializer.Serialize(new MergeTreeInsertMsg()
			{
				Pos1 = 0,
				Seg = new TextSegment("hi"),
			});
			string removeJson = SharedStringOpSerializer.Serialize(new MergeTreeRemoveMsg()
			{
				Pos1 = 0,
				Pos2 = 2,
			});

			AssertFirstTypeProperty(insertJson, 0);
			AssertFirstTypeProperty(removeJson, 1);
		}

		[Fact]
		public void LocalInsert_EmitsInsertOpJson()
		{
			FakeFluidDataObjectSender fake = new FakeFluidDataObjectSender();
			SharedString sharedString = new SharedString("shared-string", fake);

			sharedString.InsertText(0, "hi");

			var sent = Assert.Single(fake.Sent);
			Assert.Equal("shared-string", sent.Address);
			Assert.Equal("insert", sent.OpTypeName);
			MergeTreeInsertMsg op = Assert.IsType<MergeTreeInsertMsg>(
				SharedStringOpSerializer.Deserialize(sent.OpJson));
			AssertNullableInt(0, op.Pos1);
			AssertTextSegment(op.Seg, "hi");
		}

		[Fact]
		public void LocalDelete_EmitsRemoveOpJson()
		{
			FakeFluidDataObjectSender fake = new FakeFluidDataObjectSender();
			SharedString sharedString = new SharedString("shared-string", fake);

			sharedString.InsertText(0, "hi");
			sharedString.DeleteText(0, 2);

			Assert.Equal(2, fake.Sent.Count);
			var sent = fake.Sent[1];
			Assert.Equal("shared-string", sent.Address);
			Assert.Equal("remove", sent.OpTypeName);
			MergeTreeRemoveMsg op = Assert.IsType<MergeTreeRemoveMsg>(
				SharedStringOpSerializer.Deserialize(sent.OpJson));
			AssertNullableInt(0, op.Pos1);
			AssertNullableInt(2, op.Pos2);
		}

		[Fact]
		public void MultipleLocalOps_HaveIncreasingClientSeq()
		{
			FakeFluidDataObjectSender fake = new FakeFluidDataObjectSender();
			SharedString sharedString = new SharedString("shared-string", fake);

			sharedString.InsertText(0, "a");
			sharedString.InsertText(1, "b");
			sharedString.InsertText(2, "c");

			Assert.Equal(3, fake.Sent.Count);
			Assert.True(fake.Sent[0].ClientSeq < fake.Sent[1].ClientSeq);
			Assert.True(fake.Sent[1].ClientSeq < fake.Sent[2].ClientSeq);
		}

		[Fact]
		public void LocalInsert_EventArgsMatchOpContent()
		{
			SharedString sharedString = new SharedString();
			SequenceDeltaEventArgs? captured = null;
			sharedString.OnSequenceDelta += (sender, args) => captured = args;

			sharedString.InsertText(0, "hello");

			Assert.NotNull(captured);
			Assert.Equal("insert", captured!.OpType);
			Assert.Equal(0, captured.Position);
			Assert.Equal(5, captured.Length);
			Assert.Equal("hello", captured.Text);
			Assert.True(captured.Local);
		}

		private static TextSegment AssertTextSegment(object? segmentSpec, string expectedText)
		{
			TextSegment? segment = TextSegment.FromJSONObject(segmentSpec);
			Assert.NotNull(segment);
			Assert.Equal(expectedText, segment!.Text);
			return segment;
		}

		private static void AssertNullableInt(int expected, int? actual)
		{
			Assert.True(actual.HasValue);
			Assert.Equal(expected, actual.Value);
		}

		private static void AssertFirstTypeProperty(string json, int expectedType)
		{
			using JsonDocument document = JsonDocument.Parse(json);
			JsonElement.ObjectEnumerator properties = document.RootElement.EnumerateObject();
			Assert.True(properties.MoveNext());
			JsonProperty firstProperty = properties.Current;
			Assert.Equal("type", firstProperty.Name);
			Assert.Equal(JsonValueKind.Number, firstProperty.Value.ValueKind);
			Assert.Equal(expectedType, firstProperty.Value.GetInt32());
		}

		private sealed class FakeFluidDataObjectSender : IFluidDataObjectSender
		{
			private long _nextClientSeq = 1;

			public List<(string Address, string OpTypeName, string OpJson, long ClientSeq)> Sent { get; } = new();

			public SequenceNumber QueueDataObjectMessage(string address, string opTypeName, string opJson)
			{
				long clientSeq = _nextClientSeq++;
				Sent.Add((address, opTypeName, opJson, clientSeq));
				return SequenceNumber.ForTesting(clientSeq);
			}
		}
	}
}
