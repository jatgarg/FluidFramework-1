// -----------------------------------------------------------------------------
// ReferenceType wire compatibility tests for TS merge-tree parity.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class ReferenceTypeWireCompatTests
	{
		private static readonly (string Name, ReferenceType Type, int WireValue)[] _tsReferenceTypes =
		{
			("Simple", ReferenceType.Simple, 0x0),
			("Tile", ReferenceType.Tile, 0x1),
			("RangeBegin", ReferenceType.RangeBegin, 0x10),
			("RangeEnd", ReferenceType.RangeEnd, 0x20),
			("SlideOnRemove", ReferenceType.SlideOnRemove, 0x40),
			("StayOnRemove", ReferenceType.StayOnRemove, 0x80),
			("Transient", ReferenceType.Transient, 0x100),
		};

		[Fact]
		public void ReferenceType_Enum_MatchesTSValues()
		{
			Assert.Equal(_tsReferenceTypes.Select(item => item.Name), Enum.GetNames<ReferenceType>());
			Assert.Equal(_tsReferenceTypes.Select(item => item.Type), Enum.GetValues<ReferenceType>());

			foreach ((string _, ReferenceType type, int wireValue) in _tsReferenceTypes)
			{
				Assert.Equal(wireValue, (int)type);
			}
		}

		[Fact]
		public void MarkerInsert_TileRefType_WireValueIs0x1()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("shared-string", sender);

			sharedString.InsertMarker(0, ReferenceType.Tile, MarkerProps("tile-marker", "tile"));

			var sent = Assert.Single(sender.Sent);
			Assert.Equal("insert", sent.OpTypeName);
			using JsonDocument document = JsonDocument.Parse(sent.OpJson);
			JsonElement marker = document.RootElement.GetProperty("seg").GetProperty("marker");
			Assert.Equal(0x1, marker.GetProperty("refType").GetInt32());
		}

		[Fact]
		public void RemoteMarker_TileRefType_ParsedCorrectly()
		{
			const string opJson = "{\"type\":0,\"pos1\":0,\"seg\":{\"marker\":{\"refType\":1},\"props\":{\"markerId\":\"remote-marker\",\"referenceTileLabels\":[\"tile\"]}}}";
			MergeTreeInsertMsg op = Assert.IsType<MergeTreeInsertMsg>(SharedStringOpSerializer.Deserialize(opJson));
			Marker parsedMarker = AssertMarker(op.Seg);
			Assert.Equal(ReferenceType.Tile, parsedMarker.RefType);

			var sharedString = new SharedString();
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq: 0, seq: 1), opJson);

			Marker? marker = sharedString.GetMarkerFromId("remote-marker");
			Assert.NotNull(marker);
			Assert.Equal(ReferenceType.Tile, marker!.RefType);
			Assert.True(marker.HasTileLabel("tile"));
		}

		[Fact]
		public void IntervalEndpoint_ReferenceType_MatchesTSSlideOnRemove()
		{
			ReferenceType startEndpoint = ReferenceType.RangeBegin | ReferenceType.SlideOnRemove;
			ReferenceType endEndpoint = ReferenceType.RangeEnd | ReferenceType.SlideOnRemove;

			AssertReferenceWireValue(startEndpoint, 0x50);
			AssertReferenceWireValue(endEndpoint, 0x60);

			var sharedString = new SharedString();
			sharedString.InsertText(0, "abcd");
			SequenceInterval interval = sharedString
				.GetIntervalCollection("comments", IntervalType.SlideOnRemove)
				.Add(1, 3, intervalId: "i1");

			Assert.Equal(0x40, (int)(interval.Start.RefType & ReferenceType.SlideOnRemove));
			Assert.Equal(0x40, (int)(interval.End.RefType & ReferenceType.SlideOnRemove));
		}

		[Fact]
		public void Reference_AllTypes_RoundTrip()
		{
			foreach ((string _, ReferenceType type, int wireValue) in _tsReferenceTypes)
			{
				AssertReferenceWireValue(type, wireValue);
			}
		}

		private static void AssertReferenceWireValue(ReferenceType refType, int expectedWireValue)
		{
			string json = SharedStringOpSerializer.Serialize(new MergeTreeInsertMsg()
			{
				Pos1 = 0,
				Seg = Marker.Make(refType),
			});

			using JsonDocument document = JsonDocument.Parse(json);
			JsonElement marker = document.RootElement.GetProperty("seg").GetProperty("marker");
			Assert.Equal(expectedWireValue, marker.GetProperty("refType").GetInt32());

			MergeTreeInsertMsg roundTripped = Assert.IsType<MergeTreeInsertMsg>(
				SharedStringOpSerializer.Deserialize(json));
			Assert.Equal(refType, AssertMarker(roundTripped.Seg).RefType);
		}

		private static Marker AssertMarker(object? segmentSpec)
		{
			Marker? marker = Marker.FromJSONObject(segmentSpec);
			Assert.NotNull(marker);
			return marker!;
		}

		private static PropertySet MarkerProps(string markerId, params string[] tileLabels)
		{
			return new PropertySet()
			{
				[Marker.reservedMarkerIdKey] = markerId,
				[Marker.reservedTileLabelsKey] = tileLabels,
			};
		}

		private static SequencedDocumentMessageDescriptor RemoteMessage(long refSeq, long seq)
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: 0, refSeq: refSeq, seq: seq),
				OpOrigin.Remote,
				"remote-client");
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
