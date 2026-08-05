// -----------------------------------------------------------------------------
// Tests for Fluid handle wire-format values in SharedDirectory.
// -----------------------------------------------------------------------------

using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class DirectoryHandleTests
	{
		[Fact]
		public void LocalSet_HandleObject_EmitsSerializedHandleWireShape()
		{
			var sender = new FakeFluidDataObjectSender();
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("target");
			registry.Register(handle, "/dataObjects/target");
			var directory = new SharedDirectory("dir", sender, registry);

			directory.Set("k", handle);

			var sent = Assert.Single(sender.Sent);
			using JsonDocument document = JsonDocument.Parse(sent.OpJson);
			JsonElement value = document.RootElement.GetProperty("value");
			Assert.Equal("Plain", value.GetProperty("type").GetString());
			JsonElement handleWire = value.GetProperty("value");
			Assert.Equal("__fluid_handle__", handleWire.GetProperty("type").GetString());
			Assert.Equal("/dataObjects/target", handleWire.GetProperty("url").GetString());
		}

		[Fact]
		public void RemoteSet_HandleWire_WithRegistry_ReturnsRegisteredDataObject()
		{
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("target");
			registry.Register(handle, "/dataObjects/target");
			var directory = new SharedDirectory(registry: registry);

			directory.ProcessDataObjectOp(RemoteMessage(), CreateRemoteHandleSetOp("/dataObjects/target"));

			Assert.Same(handle, directory.Get("k"));
		}

		[Fact]
		public void RemoteSet_HandleWire_WithoutRegistry_ReturnsSerializedFluidHandle()
		{
			var directory = new SharedDirectory();

			directory.ProcessDataObjectOp(RemoteMessage(), CreateRemoteHandleSetOp("/dataObjects/target"));

			SerializedFluidHandle handle = Assert.IsType<SerializedFluidHandle>(directory.Get("k"));
			Assert.Equal("/dataObjects/target", handle.Url);
		}

		[Fact]
		public void RemoteSet_HandleWire_UnresolvedByRegistry_ReturnsSerializedFluidHandle()
		{
			var directory = new SharedDirectory(registry: new FakeFluidDataObjectRegistry());

			directory.ProcessDataObjectOp(RemoteMessage(), CreateRemoteHandleSetOp("/dataObjects/target"));

			SerializedFluidHandle handle = Assert.IsType<SerializedFluidHandle>(directory.Get("k"));
			Assert.Equal("/dataObjects/target", handle.Url);
		}

		[Fact]
		public void DirectoryOpSerializer_HandleRoundTrip_WithRegistry_ReturnsSameObject()
		{
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("target");
			registry.Register(handle, "/dataObjects/target");
			var op = new DirectorySetOperation()
			{
				Path = "/",
				Key = "k",
				Value = new SerializableValue()
				{
					Type = "Plain",
					Value = handle,
				},
			};

			string json = DirectoryOpSerializer.Serialize(op, registry);
			DirectorySetOperation roundTripped = Assert.IsType<DirectorySetOperation>(DirectoryOpSerializer.Deserialize(json, registry));

			Assert.Same(handle, roundTripped.Value.Value);
		}

		[Fact]
		public void NestedHandles_SerializeAndDeserializeRecursively()
		{
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("target");
			registry.Register(handle, "/dataObjects/target");
			var op = new DirectorySetOperation()
			{
				Path = "/",
				Key = "nested",
				Value = new SerializableValue()
				{
					Type = "Plain",
					Value = new Dictionary<string, object?>()
					{
						["items"] = new object?[]
						{
							"plain",
							new Dictionary<string, object?>()
							{
								["handle"] = handle,
							},
						},
					},
				},
			};

			string json = DirectoryOpSerializer.Serialize(op, registry);
			using (JsonDocument document = JsonDocument.Parse(json))
			{
				JsonElement nestedWire = document.RootElement.GetProperty("value").GetProperty("value").GetProperty("items")[1].GetProperty("handle");
				Assert.Equal("__fluid_handle__", nestedWire.GetProperty("type").GetString());
				Assert.Equal("/dataObjects/target", nestedWire.GetProperty("url").GetString());
			}

			DirectorySetOperation roundTripped = Assert.IsType<DirectorySetOperation>(DirectoryOpSerializer.Deserialize(json, registry));
			Dictionary<string, object?> root = Assert.IsType<Dictionary<string, object?>>(roundTripped.Value.Value);
			List<object?> items = Assert.IsType<List<object?>>(root["items"]);
			Dictionary<string, object?> child = Assert.IsType<Dictionary<string, object?>>(items[1]);
			Assert.Same(handle, child["handle"]);
		}

		[Fact]
		public void LocalSet_HandleObject_WithoutRegistry_ThrowsInvalidOperation()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);
			var handle = new TestFluidDataObject("target");

			OcsException exception = Assert.Throws<OcsException>(() => directory.Set("k", handle));

			Assert.Equal(OcsGateErrorCode.InvalidOperation, exception.ErrorCode);
			Assert.Contains("SharedDirectory requires an IFluidDataObjectRegistry to serialize handle values.", exception.Message);
		}

		[Fact]
		public void LocalSet_SerializedFluidHandle_EmitsWireShapeWithoutRegistry()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);

			directory.Set("k", new SerializedFluidHandle("/dataObjects/target"));

			var sent = Assert.Single(sender.Sent);
			using JsonDocument document = JsonDocument.Parse(sent.OpJson);
			JsonElement handleWire = document.RootElement.GetProperty("value").GetProperty("value");
			Assert.Equal("__fluid_handle__", handleWire.GetProperty("type").GetString());
			Assert.Equal("/dataObjects/target", handleWire.GetProperty("url").GetString());
		}

		[Fact]
		public void NonHandleValues_MaterializeAsDictionary()
		{
			var op = new DirectorySetOperation()
			{
				Path = "/",
				Key = "notHandle",
				Value = new SerializableValue()
				{
					Type = "Plain",
					Value = new Dictionary<string, object?>()
					{
						["type"] = "not_a_handle",
						["url"] = "/dataObjects/target",
					},
				},
			};

			string json = DirectoryOpSerializer.Serialize(op);
			DirectorySetOperation roundTripped = Assert.IsType<DirectorySetOperation>(DirectoryOpSerializer.Deserialize(json));

			// TS-parity: JSON values now materialize to native types (Finding 18).
			Dictionary<string, object?> value = Assert.IsType<Dictionary<string, object?>>(roundTripped.Value.Value);
			Assert.Equal("not_a_handle", Assert.IsType<string>(value["type"]));
			Assert.Equal("/dataObjects/target", Assert.IsType<string>(value["url"]));
		}

		[Fact]
		public void SnapshotLoad_HandleWire_WithRegistry_ResolvesObject()
		{
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("target");
			registry.Register(handle, "/dataObjects/target");
			var directory = new SharedDirectory(registry: registry);
			const string snapshotJson = "{\"storage\":{\"k\":{\"type\":\"Plain\",\"value\":{\"type\":\"__fluid_handle__\",\"url\":\"/dataObjects/target\"}}}}";

			directory.LoadFromSnapshot(snapshotJson);

			Assert.Same(handle, directory.Get("k"));
		}

		[Fact]
		public void ValuesAndEntries_ReturnResolvedHandles()
		{
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("target");
			registry.Register(handle, "/dataObjects/target");
			var directory = new SharedDirectory(registry: registry);

			directory.ProcessDataObjectOp(RemoteMessage(), CreateRemoteHandleSetOp("/dataObjects/target"));

			Assert.Same(handle, Assert.Single(directory.Values));
			Assert.Same(handle, directory.Entries().Single().Value);
		}

		private static string CreateRemoteHandleSetOp(string url)
		{
			return "{\"type\":\"set\",\"path\":\"/\",\"key\":\"k\",\"value\":{\"type\":\"Plain\",\"value\":{\"type\":\"__fluid_handle__\",\"url\":\"" + url + "\"}}}";
		}

		private static SequencedDocumentMessageDescriptor RemoteMessage()
		{
			return new SequencedDocumentMessageDescriptor(SequenceNumber.ForTesting(0, seq: 1), OpOrigin.Remote);
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

		// -----------------------------------------------------------------
		// payloadPending regression tests
		// TS ref: runtime-utils/src/handles.ts encodeHandleForSerialization,
		// which emits `payloadPending: true` only for pending-payload handles
		// and omits the field entirely otherwise.
		// -----------------------------------------------------------------

		[Fact]
		public void RemoteSet_HandleWire_WithPayloadPending_PreservesFlagOnSerializedHandle()
		{
			var directory = new SharedDirectory();
			const string json = "{\"type\":\"set\",\"path\":\"/\",\"key\":\"k\",\"value\":{\"type\":\"Plain\",\"value\":{\"type\":\"__fluid_handle__\",\"url\":\"/pending\",\"payloadPending\":true}}}";

			directory.ProcessDataObjectOp(RemoteMessage(), json);

			SerializedFluidHandle handle = Assert.IsType<SerializedFluidHandle>(directory.Get("k"));
			Assert.Equal("/pending", handle.Url);
			Assert.True(handle.PayloadPending);
		}

		[Fact]
		public void LocalSet_SerializedFluidHandleWithPayloadPending_EmitsPayloadPendingOnWire()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);

			directory.Set("k", new SerializedFluidHandle("/pending", payloadPending: true));

			var sent = Assert.Single(sender.Sent);
			using JsonDocument document = JsonDocument.Parse(sent.OpJson);
			JsonElement handleWire = document.RootElement.GetProperty("value").GetProperty("value");
			Assert.Equal("__fluid_handle__", handleWire.GetProperty("type").GetString());
			Assert.Equal("/pending", handleWire.GetProperty("url").GetString());
			Assert.True(handleWire.GetProperty("payloadPending").GetBoolean());
		}

		[Fact]
		public void LocalSet_SerializedFluidHandleWithoutPayloadPending_OmitsPayloadPendingOnWire()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);

			directory.Set("k", new SerializedFluidHandle("/shared"));

			var sent = Assert.Single(sender.Sent);
			using JsonDocument document = JsonDocument.Parse(sent.OpJson);
			JsonElement handleWire = document.RootElement.GetProperty("value").GetProperty("value");
			// TS omits `payloadPending` entirely when the handle is not pending.
			Assert.False(handleWire.TryGetProperty("payloadPending", out _));
		}
	}
}
