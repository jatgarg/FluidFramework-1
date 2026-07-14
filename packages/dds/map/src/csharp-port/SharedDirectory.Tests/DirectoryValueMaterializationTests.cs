// -----------------------------------------------------------------------------
// Tests for TS-parity materialization of JSON values in SharedDirectory.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;
using System.Text.Json;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class DirectoryValueMaterializationTests
	{
		[Fact]
		public void RemoteSet_StringValue_ReturnedAsString()
		{
			var directory = new SharedDirectory();

			ProcessRemoteSet(directory, "k", "\"remote\"");

			Assert.Equal("remote", Assert.IsType<string>(directory.Get("k")));
		}

		[Fact]
		public void RemoteSet_NumberValue_ReturnedAsDouble()
		{
			var directory = new SharedDirectory();

			ProcessRemoteSet(directory, "k", "42");

			Assert.Equal(42d, Assert.IsType<double>(directory.Get("k")));
		}

		[Fact]
		public void RemoteSet_ObjectValue_ReturnedAsDictionary()
		{
			var directory = new SharedDirectory();

			ProcessRemoteSet(directory, "k", "{\"a\":1,\"nested\":{\"enabled\":true}}");

			object? value = directory.Get("k");
			Assert.IsNotType<JsonElement>(value);
			Dictionary<string, object?> dictionary = Assert.IsType<Dictionary<string, object?>>(value);
			Assert.Equal(1d, Assert.IsType<double>(dictionary["a"]));
			Dictionary<string, object?> nested = Assert.IsType<Dictionary<string, object?>>(dictionary["nested"]);
			Assert.True(Assert.IsType<bool>(nested["enabled"]));
		}

		[Fact]
		public void RemoteSet_ArrayValue_ReturnedAsList()
		{
			var directory = new SharedDirectory();

			ProcessRemoteSet(directory, "k", "[1,\"two\",{\"three\":3}]");

			List<object?> list = Assert.IsType<List<object?>>(directory.Get("k"));
			Assert.Equal(1d, Assert.IsType<double>(list[0]));
			Assert.Equal("two", Assert.IsType<string>(list[1]));
			Dictionary<string, object?> child = Assert.IsType<Dictionary<string, object?>>(list[2]);
			Assert.Equal(3d, Assert.IsType<double>(child["three"]));
		}

		[Fact]
		public void RemoteSet_NestedObjectContainingHandle_HandlePreservedAtDepth()
		{
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("target");
			registry.Register(handle, "/dataObjects/target");
			var directory = new SharedDirectory(registry: registry);

			ProcessRemoteSet(
				directory,
				"k",
				"{\"items\":[{\"nested\":{\"target\":{\"type\":\"__fluid_handle__\",\"url\":\"/dataObjects/target\"}}}]}");

			Dictionary<string, object?> root = Assert.IsType<Dictionary<string, object?>>(directory.Get("k"));
			List<object?> items = Assert.IsType<List<object?>>(root["items"]);
			Dictionary<string, object?> item = Assert.IsType<Dictionary<string, object?>>(items[0]);
			Dictionary<string, object?> nested = Assert.IsType<Dictionary<string, object?>>(item["nested"]);
			Assert.Same(handle, nested["target"]);
		}

		[Fact]
		public void Snapshot_LoadedValue_MaterializedNatively()
		{
			var directory = new SharedDirectory();
			const string snapshotJson = "{\"storage\":{\"k\":{\"type\":\"Plain\",\"value\":{\"a\":1,\"items\":[true,null,\"text\"]}}}}";

			directory.LoadFromSnapshot(snapshotJson);

			Dictionary<string, object?> root = Assert.IsType<Dictionary<string, object?>>(directory.Get("k"));
			Assert.Equal(1d, Assert.IsType<double>(root["a"]));
			List<object?> items = Assert.IsType<List<object?>>(root["items"]);
			Assert.True(Assert.IsType<bool>(items[0]));
			Assert.Null(items[1]);
			Assert.Equal("text", Assert.IsType<string>(items[2]));
		}

		private static void ProcessRemoteSet(SharedDirectory directory, string key, string valueJson)
		{
			directory.ProcessDataObjectOp(RemoteMessage(), CreateSetOpJson(key, valueJson));
		}

		private static string CreateSetOpJson(string key, string valueJson)
		{
			return "{\"type\":\"set\",\"path\":\"/\",\"key\":\""
				+ key
				+ "\",\"value\":{\"type\":\"Plain\",\"value\":"
				+ valueJson
				+ "}}";
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
	}
}
