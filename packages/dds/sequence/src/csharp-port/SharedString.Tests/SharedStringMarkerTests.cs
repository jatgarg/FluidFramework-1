// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SharedStringMarkerTests
	{
		[Fact]
		public void InsertMarker_AddsToStringWithLengthOne()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello");

			sharedString.InsertMarker(5, ReferenceType.Tile, MarkerProps("para1", "paragraph"));

			Assert.Equal("hello", sharedString.GetText());
			Assert.Equal(6, sharedString.GetLength());
		}

		[Fact]
		public void GetMarkerFromId_ReturnsInsertedMarker()
		{
			var sharedString = new SharedString();

			sharedString.InsertMarker(0, ReferenceType.Tile, MarkerProps("p1", "paragraph"));

			Marker? marker = sharedString.GetMarkerFromId("p1");
			Assert.NotNull(marker);
			Assert.Equal("p1", marker!.GetId());
			Assert.Equal(ReferenceType.Tile, marker.RefType);
		}

		[Fact]
		public void GetMarkerFromId_UnknownId_ReturnsNull()
		{
			var sharedString = new SharedString();

			Assert.Null(sharedString.GetMarkerFromId("missing"));
		}

		[Fact]
		public void SearchForMarker_ReturnsNextTileMarker()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello");
			sharedString.InsertMarker(5, ReferenceType.Tile, MarkerProps("p1", "para"));

			Marker? marker = sharedString.SearchForMarker(0, forwards: true, tileLabel: "para");

			Assert.NotNull(marker);
			Assert.Equal("p1", marker!.GetId());
		}

		[Fact]
		public void SearchForMarker_NoMatch_ReturnsNull()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello");
			sharedString.InsertMarker(5, ReferenceType.Tile, MarkerProps("p1", "para"));

			Assert.Null(sharedString.SearchForMarker(0, forwards: true, tileLabel: "doesnotexist"));
		}

		[Fact]
		public void MarkerInsert_JsonRoundTrip_PreservesRefTypeAndProps()
		{
			MergeTreeInsertMsg op = new MergeTreeInsertMsg()
			{
				Pos1 = 3,
				Seg = Marker.Make(
					ReferenceType.Tile | ReferenceType.SlideOnRemove,
					MarkerProps("m1", "paragraph")),
			};

			MergeTreeInsertMsg roundTripped = Assert.IsType<MergeTreeInsertMsg>(
				SharedStringOpSerializer.Deserialize(SharedStringOpSerializer.Serialize(op)));
			Marker marker = AssertMarker(roundTripped.Seg);

			Assert.Equal(ReferenceType.Tile | ReferenceType.SlideOnRemove, marker.RefType);
			Assert.Equal("m1", marker.GetId());
			Assert.True(marker.HasTileLabel("paragraph"));
		}

		[Fact]
		public void MarkerInsert_EmittedViaSharedString_HasCorrectJson()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("shared-string", sender);

			sharedString.InsertMarker(0, ReferenceType.Tile, MarkerProps("emit1", "tile"));

			var sent = Assert.Single(sender.Sent);
			Assert.Equal("insert", sent.OpTypeName);
			using JsonDocument document = JsonDocument.Parse(sent.OpJson);
			JsonElement segment = document.RootElement.GetProperty("seg");
			JsonElement marker = segment.GetProperty("marker");
			JsonElement props = segment.GetProperty("props");
			Assert.Equal((int)ReferenceType.Tile, marker.GetProperty("refType").GetInt32());
			Assert.Equal("emit1", props.GetProperty(Marker.reservedMarkerIdKey).GetString());
			JsonElement.ArrayEnumerator labels = props.GetProperty(Marker.reservedTileLabelsKey).EnumerateArray();
			Assert.True(labels.MoveNext());
			Assert.Equal("tile", labels.Current.GetString());
		}

		[Fact]
		public void RemoteInsertMarker_AddsToLocalIndex()
		{
			var sharedString = new SharedString();

			ProcessRemoteMarkerInsert(sharedString, 0, "remote1", "para", refSeq: 0, seq: 1);

			Marker? marker = sharedString.GetMarkerFromId("remote1");
			Assert.NotNull(marker);
			Assert.True(marker!.HasTileLabel("para"));
		}

		[Fact]
		public void RemoteInsertMarker_FiresEventWithLocalFalse()
		{
			var sharedString = new SharedString();
			var events = new List<SequenceDeltaEventArgs>();
			sharedString.OnSequenceDelta += (sender, args) => events.Add(args);

			ProcessRemoteMarkerInsert(sharedString, 0, "remote1", "para", refSeq: 0, seq: 1);

			SequenceDeltaEventArgs captured = Assert.Single(events);
			Assert.False(captured.Local);
			Assert.Equal("insert", captured.OpType);
			Assert.True(captured.IsMarker);
			Assert.Null(captured.Text);
			Assert.NotNull(captured.Marker);
			Assert.Equal("remote1", captured.Marker!.GetId());
		}

		[Fact]
		public void TwoClients_ConcurrentDifferentMarkers_BothInIndex()
		{
			var harness = new TwoClientHarness();
			LoadInitialSharedText(harness, "hello");
			harness.SenderA.Sent.Clear();
			harness.SenderB.Sent.Clear();
			long refSeq = harness.CurrentServerSeq;

			harness.ClientA.InsertMarker(5, ReferenceType.Tile, MarkerProps("A1", "para"));
			var sentA = Assert.Single(harness.SenderA.Sent);
			harness.ClientB.InsertMarker(5, ReferenceType.Tile, MarkerProps("B1", "para"));
			var sentB = Assert.Single(harness.SenderB.Sent);

			harness.DeliverAtoB(sentA, refSeq);
			harness.DeliverBtoA(sentB, refSeq);

			Assert.NotNull(harness.ClientA.GetMarkerFromId("A1"));
			Assert.NotNull(harness.ClientA.GetMarkerFromId("B1"));
			Assert.NotNull(harness.ClientB.GetMarkerFromId("A1"));
			Assert.NotNull(harness.ClientB.GetMarkerFromId("B1"));
			Assert.Equal(harness.ClientA.GetText(), harness.ClientB.GetText());
			Assert.Equal(harness.ClientA.GetLength(), harness.ClientB.GetLength());
		}

		[Fact]
		public void MarkerInsertedAfterSnapshotLoaded_MarkerPresent()
		{
			var sharedString = new SharedString();
			sharedString.LoadFromSnapshot(LoadSimpleHelloSnapshot());

			sharedString.InsertMarker(sharedString.GetLength(), ReferenceType.Tile, MarkerProps("snap1", "paragraph"));

			Marker? marker = sharedString.GetMarkerFromId("snap1");
			Assert.NotNull(marker);
			Assert.Equal(13, sharedString.GetPositionOfMarker(marker!));
			Assert.Equal("Hello, world!", sharedString.GetText());
		}

		[Fact]
		public void DeleteRange_ContainingMarker_RemovesFromIndex()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello");
			sharedString.InsertMarker(5, ReferenceType.Tile, MarkerProps("p1", "para"));

			sharedString.DeleteText(4, 6);

			Assert.Null(sharedString.GetMarkerFromId("p1"));
			Assert.Equal("hell", sharedString.GetText());
		}

		[Fact]
		public void AnnotateRange_OnMarker_UpdatesProperties()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hi");
			sharedString.InsertMarker(2, ReferenceType.Tile, MarkerProps("p1", "para"));

			sharedString.AnnotateRange(2, 3, new PropertySet()
			{
				["label"] = "bar",
			});

			Marker? marker = sharedString.GetMarkerFromId("p1");
			Assert.NotNull(marker);
			Assert.NotNull(marker!.Properties);
			Assert.Equal("bar", Assert.IsType<string>(marker.Properties!["label"]));
		}

		[Fact]
		public void MarkerIdIndex_ClearedThroughLocalRemoteAndAckedObliterate()
		{
			// TS ref: packages/dds/merge-tree/src/client.ts + test/client.searchForMarker.spec —
			// TS's marker id index cleans up when a marker is removed by local,
			// remote, or ACKed obliterate paths. The T02 audit finding flagged
			// this end-to-end path as untested.

			// Local remove path.
			var localString = new SharedString();
			localString.InsertText(0, "abc");
			localString.InsertMarker(3, ReferenceType.Tile, MarkerProps("p-local", "para"));
			Assert.NotNull(localString.GetMarkerFromId("p-local"));

			localString.DeleteText(2, 4);
			Assert.Null(localString.GetMarkerFromId("p-local"));

			// Remote remove path (marker + remove both come from remote).
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("s", sender);
			sharedString.InsertText(0, "abc");
			ProcessLocalAck(sharedString, Assert.Single(sender.Sent), refSeq: 0, seq: 1);
			sender.Sent.Clear();

			string markerInsertJson = SharedStringOpSerializer.Serialize(new MergeTreeInsertMsg()
			{
				Pos1 = 3,
				Seg = Marker.Make(ReferenceType.Tile, MarkerProps("p-remote", "para")),
			});
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq: 1, seq: 2), markerInsertJson);
			Assert.NotNull(sharedString.GetMarkerFromId("p-remote"));

			string removeJson = SharedStringOpSerializer.Serialize(new MergeTreeRemoveMsg()
			{
				Pos1 = 2,
				Pos2 = 4,
			});
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq: 2, seq: 3), removeJson);
			Assert.Null(sharedString.GetMarkerFromId("p-remote"));

			// Local -> ACKed remove path.
			var ackedSender = new FakeFluidDataObjectSender();
			var ackedString = new SharedString("s2", ackedSender);
			ackedString.InsertText(0, "abc");
			ackedString.InsertMarker(3, ReferenceType.Tile, MarkerProps("p-acked", "para"));
			ackedString.DeleteText(2, 4);
			// Ack all pending: insert, insertMarker, delete.
			long ackSeq = 1;
			foreach ((string Address, string OpTypeName, string OpJson, long ClientSeq) in ackedSender.Sent.ToList())
			{
				ProcessLocalAck(ackedString, (Address, OpTypeName, OpJson, ClientSeq), refSeq: 0, seq: ackSeq++);
			}

			Assert.Null(ackedString.GetMarkerFromId("p-acked"));
		}

		[Fact]
		public void NamedMarker_ClearedThroughObliterate()
		{
			// V2-T04 regression. TS specifies that a named-marker's id-index
			// entry is removed when the marker segment is obliterated (not
			// just when it's plain-removed). The existing marker-cleanup
			// test uses delete; the obliterate test uses an anonymous marker.
			// This fills the gap.
			SharedString sharedString = new();
			sharedString.InsertText(0, "abc");
			sharedString.InsertMarker(3, ReferenceType.Tile, MarkerProps("p-obliterate", "para"));
			Assert.NotNull(sharedString.GetMarkerFromId("p-obliterate"));

			sharedString.ObliterateRange(2, 4);

			Assert.Null(sharedString.GetMarkerFromId("p-obliterate"));
		}

		private static PropertySet MarkerProps(string markerId, params string[] tileLabels)
		{
			return new PropertySet()
			{
				[Marker.reservedMarkerIdKey] = markerId,
				[Marker.reservedTileLabelsKey] = tileLabels,
			};
		}

		private static Marker AssertMarker(object? segmentSpec)
		{
			Marker? marker = Marker.FromJSONObject(segmentSpec);
			Assert.NotNull(marker);
			return marker!;
		}

		private static void ProcessRemoteMarkerInsert(
			SharedString sharedString,
			int position,
			string markerId,
			string tileLabel,
			long refSeq,
			long seq,
			string clientId = "remote-client")
		{
			string opJson = SharedStringOpSerializer.Serialize(new MergeTreeInsertMsg()
			{
				Pos1 = position,
				Seg = Marker.Make(ReferenceType.Tile, MarkerProps(markerId, tileLabel)),
			});

			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq, seq, clientId), opJson);
		}

		private static void LoadInitialSharedText(TwoClientHarness harness, string text)
		{
			harness.ClientA.InsertText(0, text);
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: 0);
		}

		private static SequencedDocumentMessageDescriptor RemoteMessage(
			long refSeq,
			long seq,
			string clientId = "remote-client")
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: 0, refSeq: refSeq, seq: seq),
				OpOrigin.Remote,
				clientId);
		}

		private static void ProcessLocalAck(
			SharedString sharedString,
			(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
			long refSeq,
			long seq)
		{
			sharedString.ProcessDataObjectOp(
				new SequencedDocumentMessageDescriptor(
					SequenceNumber.ForTesting(clientSeq: sent.ClientSeq, refSeq: refSeq, seq: seq),
					OpOrigin.Local),
				sent.OpJson);
		}

		private static SharedStringSnapshotDto LoadSimpleHelloSnapshot()
		{
			string fixturePath = Path.Combine(
				AppContext.BaseDirectory,
				"Fixtures",
				"simple-hello.snapshot.json.txt");
			return SharedStringSnapshotLoader.Parse(File.ReadAllText(fixturePath));
		}

		private sealed class TwoClientHarness
		{
			private const string _clientAId = "client-a";
			private const string _clientBId = "client-b";
			private long _serverSeq;

			public TwoClientHarness()
			{
				SenderA = new FakeFluidDataObjectSender();
				SenderB = new FakeFluidDataObjectSender();
				ClientA = new SharedString("doc", SenderA);
				ClientB = new SharedString("doc", SenderB);
			}

			public SharedString ClientA { get; }

			public SharedString ClientB { get; }

			public FakeFluidDataObjectSender SenderA { get; }

			public FakeFluidDataObjectSender SenderB { get; }

			public long CurrentServerSeq => _serverSeq;

			public void DeliverAtoB(
				(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
				long refSeq)
			{
				Deliver(sent, fromClientId: _clientAId, from: ClientA, to: ClientB, refSeq: refSeq);
			}

			public void DeliverBtoA(
				(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
				long refSeq)
			{
				Deliver(sent, fromClientId: _clientBId, from: ClientB, to: ClientA, refSeq: refSeq);
			}

			private void Deliver(
				(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
				string fromClientId,
				SharedString from,
				SharedString to,
				long refSeq)
			{
				long serverSeq = ++_serverSeq;
				var descriptorForTo = new SequencedDocumentMessageDescriptor(
					SequenceNumber.ForTesting(clientSeq: sent.ClientSeq, refSeq: refSeq, seq: serverSeq),
					OpOrigin.Remote,
					fromClientId);
				to.ProcessDataObjectOp(descriptorForTo, sent.OpJson);

				var descriptorForFrom = new SequencedDocumentMessageDescriptor(
					SequenceNumber.ForTesting(clientSeq: sent.ClientSeq, refSeq: refSeq, seq: serverSeq),
					OpOrigin.Local,
					fromClientId);
				from.ProcessDataObjectOp(descriptorForFrom, sent.OpJson);
			}
		}
	}
}
