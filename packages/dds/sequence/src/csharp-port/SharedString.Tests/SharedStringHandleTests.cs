// -----------------------------------------------------------------------------
// Tests for Fluid handle wire-format values in SharedString property maps.
// -----------------------------------------------------------------------------

using System.Collections.Generic;
using System.Text.Json;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SharedStringHandleTests
	{
		[Fact]
		public void LocalAnnotate_HandleObject_EmitsSerializedHandleWireShape()
		{
			var sender = new FakeFluidDataObjectSender();
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("target");
			registry.Register(handle, "/dataObjects/target");
			var sharedString = new SharedString("shared-string", sender, registry);
			sharedString.InsertText(0, "hello");
			sender.Sent.Clear();

			sharedString.AnnotateRange(0, 5, new PropertySet()
			{
				["target"] = handle,
			});

			var sent = Assert.Single(sender.Sent);
			using JsonDocument document = JsonDocument.Parse(sent.OpJson);
			JsonElement handleWire = document.RootElement.GetProperty("props").GetProperty("target");
			Assert.Equal("__fluid_handle__", handleWire.GetProperty("type").GetString());
			Assert.Equal("/dataObjects/target", handleWire.GetProperty("url").GetString());
		}

		[Fact]
		public void RemoteAnnotate_HandleWire_WithRegistry_ReturnsRegisteredDataObject()
		{
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("target");
			registry.Register(handle, "/dataObjects/target");
			SharedString sharedString = CreateAckedSharedString("hello", registry);

			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq: 1, seq: 2), CreateRemoteHandleAnnotateOp("/dataObjects/target"));

			Assert.Same(handle, GetRequiredProperty(sharedString, 0, "target"));
		}

		[Fact]
		public void RemoteAnnotate_HandleWire_WithoutRegistry_ReturnsSerializedFluidHandle()
		{
			SharedString sharedString = CreateAckedSharedString("hello");

			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq: 1, seq: 2), CreateRemoteHandleAnnotateOp("/dataObjects/target"));

			SerializedFluidHandle handle = Assert.IsType<SerializedFluidHandle>(GetRequiredProperty(sharedString, 0, "target"));
			Assert.Equal("/dataObjects/target", handle.Url);
		}

		[Fact]
		public void MarkerProps_HandleObject_RoundTripsThroughWireShape()
		{
			var sender = new FakeFluidDataObjectSender();
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("target");
			registry.Register(handle, "/dataObjects/target");
			var senderString = new SharedString("shared-string", sender, registry);

			senderString.InsertMarker(0, ReferenceType.Tile, MarkerProps("m1", "paragraph", handle));

			var sent = Assert.Single(sender.Sent);
			using (JsonDocument document = JsonDocument.Parse(sent.OpJson))
			{
				JsonElement handleWire = document.RootElement.GetProperty("seg").GetProperty("props").GetProperty("target");
				Assert.Equal("__fluid_handle__", handleWire.GetProperty("type").GetString());
				Assert.Equal("/dataObjects/target", handleWire.GetProperty("url").GetString());
			}

			var receiverString = new SharedString(registry: registry);
			receiverString.ProcessDataObjectOp(RemoteMessage(refSeq: 0, seq: 1), sent.OpJson);
			Marker? marker = receiverString.GetMarkerFromId("m1");
			Assert.NotNull(marker);
			Assert.Same(handle, marker!.Properties!["target"]);
		}

		[Fact]
		public void RemoteInsert_HandleShape_MissingUrl_Rejected()
		{
			// TS ref: packages/runtime/runtime-utils/src/handles.ts isSerializedHandle
			// identifies a handle purely by `type === "__fluid_handle__"`. TS would
			// then fail at dereference when reading `value.url`. The port previously
			// accepted the malformed handle as ordinary property data and could
			// re-emit it, causing a TS peer to reject during processing. Match TS's
			// handle-identification but reject at ingestion to keep the wire clean.
			var receiverString = new SharedString();

			OcsException exception = Assert.Throws<OcsException>(() =>
				receiverString.ProcessDataObjectOp(
					RemoteMessage(refSeq: 0, seq: 1),
					"{\"type\":0,\"pos1\":0,\"seg\":{\"props\":{\"target\":{\"type\":\"__fluid_handle__\"}}}}"));

			Assert.Contains("Serialized Fluid handle is missing required 'url' property", exception.Message);
		}

		[Fact]
		public void IntervalAddAndPropertyChanged_HandleProps_RoundTrip()
		{
			var sender = new FakeFluidDataObjectSender();
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("target");
			var changedHandle = new TestFluidDataObject("changed");
			registry.Register(handle, "/dataObjects/target");
			registry.Register(changedHandle, "/dataObjects/changed");
			var sharedString = new SharedString("shared-string", sender, registry);
			sharedString.InsertText(0, "hello");
			sender.Sent.Clear();
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");

			SequenceInterval interval = collection.Add(1, 3, new PropertySet()
			{
				["target"] = handle,
			}, intervalId: "i1");

			var addSent = Assert.Single(sender.Sent);
			using (JsonDocument document = JsonDocument.Parse(addSent.OpJson))
			{
				// TS wire: op is wrapped in IntervalCollectionMap "act" envelope; the interval
				// properties bag lives at value.value.properties per intervalCollection.ts.
				JsonElement payload = document.RootElement.GetProperty("value").GetProperty("value");
				JsonElement handleWire = payload.GetProperty("properties").GetProperty("target");
				Assert.Equal("__fluid_handle__", handleWire.GetProperty("type").GetString());
				Assert.Equal("/dataObjects/target", handleWire.GetProperty("url").GetString());
			}

			// Inbound TS-compatible property-change op (opName "change" with only the
			// properties bag — mirrors intervalCollection.ts's propertyChanged emission).
			sharedString.ProcessDataObjectOp(
				RemoteMessage(refSeq: 0, seq: 1),
				"{\"type\":\"act\",\"key\":\"comments\",\"value\":{\"opName\":\"change\",\"value\":{\"sequenceNumber\":0,\"intervalType\":2,\"properties\":{\"changed\":{\"type\":\"__fluid_handle__\",\"url\":\"/dataObjects/changed\"},\"intervalId\":\"i1\",\"referenceRangeLabels\":[\"comments\"]}}}}");

			Assert.Same(changedHandle, interval.Properties!["changed"]);
		}

		[Fact]
		public void NestedHandleInsideArray_IsDetectedRecursively()
		{
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("target");
			registry.Register(handle, "/dataObjects/target");
			var op = new MergeTreeAnnotateMsg()
			{
				Pos1 = 0,
				Pos2 = 1,
				Props = new PropertySet()
				{
					["items"] = new object?[] { "plain", handle },
				},
			};

			string json = SharedStringOpSerializer.Serialize(op, registry);
			using (JsonDocument document = JsonDocument.Parse(json))
			{
				JsonElement nestedWire = document.RootElement.GetProperty("props").GetProperty("items")[1];
				Assert.Equal("__fluid_handle__", nestedWire.GetProperty("type").GetString());
				Assert.Equal("/dataObjects/target", nestedWire.GetProperty("url").GetString());
			}

			MergeTreeAnnotateMsg roundTripped = Assert.IsType<MergeTreeAnnotateMsg>(
				SharedStringOpSerializer.Deserialize(json, registry));
			List<object?> items = Assert.IsType<List<object?>>(roundTripped.Props!["items"]);
			Assert.Equal("plain", Assert.IsType<string>(items[0]));
			Assert.Same(handle, items[1]);
		}

		[Fact]
		public void NonHandleProps_PassThroughUnchanged()
		{
			var op = new MergeTreeAnnotateMsg()
			{
				Pos1 = 0,
				Pos2 = 1,
				Props = new PropertySet()
				{
					["text"] = "hello",
					["count"] = 3,
					["enabled"] = true,
					["nested"] = new PropertySet()
					{
						["type"] = "not_a_handle",
						["url"] = "/dataObjects/target",
					},
				},
			};

			MergeTreeAnnotateMsg roundTripped = Assert.IsType<MergeTreeAnnotateMsg>(
				SharedStringOpSerializer.Deserialize(SharedStringOpSerializer.Serialize(op)));

			PropertySet props = roundTripped.Props!;
			Assert.Equal("hello", Assert.IsType<string>(props["text"]));
			Assert.Equal(3, Assert.IsType<int>(props["count"]));
			Assert.True(Assert.IsType<bool>(props["enabled"]));
			PropertySet nested = Assert.IsType<PropertySet>(props["nested"]);
			Assert.Equal("not_a_handle", Assert.IsType<string>(nested["type"]));
			Assert.Equal("/dataObjects/target", Assert.IsType<string>(nested["url"]));
		}

		[Fact]
		public void SnapshotLoad_SegmentHandleProp_WithRegistry_ResolvesObject()
		{
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("target");
			registry.Register(handle, "/dataObjects/target");
			var sharedString = new SharedString(registry: registry);
			const string snapshotJson = "{\"version\":\"1\",\"segmentCount\":1,\"length\":5,\"startIndex\":0,\"segments\":[{\"text\":\"hello\",\"props\":{\"target\":{\"type\":\"__fluid_handle__\",\"url\":\"/dataObjects/target\"}}}],\"headerMetadata\":{\"minSequenceNumber\":0,\"sequenceNumber\":1,\"totalLength\":5,\"totalSegmentCount\":1,\"orderedChunkMetadata\":[{\"id\":\"header\"}]}}";

			sharedString.LoadFromSnapshot(snapshotJson);

			Assert.Same(handle, GetRequiredProperty(sharedString, 0, "target"));
		}

		[Fact]
		public void TwoClients_TsHandleAnnotateOp_ConvergesWithResolvedHandle()
		{
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("target");
			registry.Register(handle, "/dataObjects/target");
			SharedString clientA = CreateAckedSharedString("hello", registry, "client-a");
			SharedString clientB = CreateAckedSharedString("hello", registry, "client-b");
			string opJson = CreateRemoteHandleAnnotateOp("/dataObjects/target");

			clientA.ProcessDataObjectOp(RemoteMessage(refSeq: 1, seq: 2, clientId: "ts-client"), opJson);
			clientB.ProcessDataObjectOp(RemoteMessage(refSeq: 1, seq: 2, clientId: "ts-client"), opJson);

			Assert.Equal(clientA.GetText(), clientB.GetText());
			Assert.Same(handle, GetRequiredProperty(clientA, 0, "target"));
			Assert.Same(handle, GetRequiredProperty(clientB, 0, "target"));
		}

		[Fact]
		public void SerializedFluidHandle_RoundTripsWithoutRegistry()
		{
			var op = new MergeTreeAnnotateMsg()
			{
				Pos1 = 0,
				Pos2 = 1,
				Props = new PropertySet()
				{
					["target"] = new SerializedFluidHandle("/dataObjects/target"),
				},
			};

			string json = SharedStringOpSerializer.Serialize(op);
			using (JsonDocument document = JsonDocument.Parse(json))
			{
				JsonElement handleWire = document.RootElement.GetProperty("props").GetProperty("target");
				Assert.Equal("__fluid_handle__", handleWire.GetProperty("type").GetString());
				Assert.Equal("/dataObjects/target", handleWire.GetProperty("url").GetString());
			}

			MergeTreeAnnotateMsg roundTripped = Assert.IsType<MergeTreeAnnotateMsg>(
				SharedStringOpSerializer.Deserialize(json));
			SerializedFluidHandle handle = Assert.IsType<SerializedFluidHandle>(roundTripped.Props!["target"]);
			Assert.Equal("/dataObjects/target", handle.Url);
		}

		private static PropertySet MarkerProps(string markerId, string tileLabel, IFluidDataObject handle)
		{
			return new PropertySet()
			{
				[Marker.reservedMarkerIdKey] = markerId,
				[Marker.reservedTileLabelsKey] = new string[] { tileLabel },
				["target"] = handle,
			};
		}

		private static string CreateRemoteHandleAnnotateOp(string url)
		{
			return "{\"type\":2,\"pos1\":0,\"pos2\":5,\"props\":{\"target\":{\"type\":\"__fluid_handle__\",\"url\":\"" + url + "\"}}}";
		}

		private static SharedString CreateAckedSharedString(
			string text,
			IFluidDataObjectRegistry? registry = null,
			string id = "doc")
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString(id, sender, registry);
			sharedString.InsertText(0, text);
			var sent = Assert.Single(sender.Sent);
			sharedString.ProcessDataObjectOp(LocalAck(sent.ClientSeq, refSeq: 0, seq: 1), sent.OpJson);
			return sharedString;
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

		private static SequencedDocumentMessageDescriptor LocalAck(long clientSeq, long refSeq, long seq)
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: clientSeq, refSeq: refSeq, seq: seq),
				OpOrigin.Local);
		}

		private static object? GetRequiredProperty(SharedString sharedString, int position, string key)
		{
			PropertySet? properties = sharedString.GetPropertiesAtPosition(position);
			Assert.NotNull(properties);
			Assert.True(properties!.ContainsKey(key), $"Expected property '{key}' at position {position}.");
			return properties[key];
		}

		private sealed class TestFluidDataObject : IFluidDataObject
		{
			public TestFluidDataObject(string id)
			{
				Id = id;
			}

			public string Id { get; }
		}

		private sealed class FakeFluidDataObjectRegistry : IFluidDataObjectRegistry
		{
			private readonly Dictionary<string, IFluidDataObject> _objectsByUrl = new Dictionary<string, IFluidDataObject>();
			private readonly Dictionary<IFluidDataObject, string> _urlsByObject = new Dictionary<IFluidDataObject, string>();

			public void Register(IFluidDataObject obj, string url)
			{
				_objectsByUrl[url] = obj;
				_urlsByObject[obj] = url;
			}

			public IFluidDataObject? FindDataObject(string address)
			{
				_objectsByUrl.TryGetValue(address, out IFluidDataObject? obj);
				return obj;
			}

			public string GetDataObjectUrl(IFluidDataObject obj)
			{
				return _urlsByObject[obj];
			}
		}
	}
}
