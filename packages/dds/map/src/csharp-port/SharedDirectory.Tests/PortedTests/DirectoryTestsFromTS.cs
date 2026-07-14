// -----------------------------------------------------------------------------
// High-value SharedDirectory tests ported from TS directory.spec.ts.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class DirectoryTestsFromTS
	{
		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.spec.ts:511 — "Should populate the directory from an empty JSON object (old format)"
		public void ShouldPopulateTheDirectoryFromAnEmptyJsonObjectOldFormat()
		{
			var directory = new SharedDirectory();

			directory.LoadFromSnapshot("{}");

			Assert.Equal(0, directory.Count);
			directory.Set("testKey", "testValue");
			Assert.Equal("testValue", directory.Get("testKey"));
			directory.CreateSubDirectory("testSubDir").Set("testSubKey", "testSubValue");
			IDirectory subDir = Assert.IsAssignableFrom<IDirectory>(directory.GetWorkingDirectory("testSubDir"));
			Assert.Equal("testSubValue", subDir.Get("testSubKey"));
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.spec.ts:522 — "Should populate the directory from a basic JSON object (old format)"
		public void ShouldPopulateTheDirectoryFromABasicJsonObjectOldFormat()
		{
			const string json = "{\"storage\":{\"testKey\":{\"type\":\"Plain\",\"value\":\"testValue4\"},\"testKey2\":{\"type\":\"Plain\",\"value\":\"testValue5\"}},\"subdirectories\":{\"foo\":{\"storage\":{\"testKey\":{\"type\":\"Plain\",\"value\":\"testValue\"},\"testKey2\":{\"type\":\"Plain\",\"value\":\"testValue2\"}}},\"bar\":{\"storage\":{\"testKey3\":{\"type\":\"Plain\",\"value\":\"testValue3\"}}}}}";
			var directory = new SharedDirectory();

			directory.LoadFromSnapshot(json);

			Assert.Equal(2, directory.Count);
			Assert.Equal("testValue", directory.GetWorkingDirectory("/foo")!.Get("testKey"));
			Assert.Equal("testValue2", directory.GetWorkingDirectory("foo")!.Get("testKey2"));
			Assert.Equal("testValue3", directory.GetWorkingDirectory("/bar")!.Get("testKey3"));
			Assert.Equal("testValue4", directory.GetWorkingDirectory("")!.Get("testKey"));
			Assert.Equal("testValue5", directory.GetWorkingDirectory("/")!.Get("testKey2"));

			directory.Set("testKey", "newValue");
			Assert.Equal("newValue", directory.Get("testKey"));
			directory.CreateSubDirectory("testSubDir").Set("testSubKey", "newSubValue");
			Assert.Equal("newSubValue", directory.GetWorkingDirectory("testSubDir")!.Get("testSubKey"));
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.spec.ts:573 — "Should populate the directory with undefined values (old format)"
		public void ShouldPopulateTheDirectoryWithUndefinedValuesOldFormat()
		{
			const string json = "{\"storage\":{\"testKey\":{\"type\":\"Plain\",\"value\":\"testValue4\"},\"testKey2\":{\"type\":\"Plain\"}},\"subdirectories\":{\"foo\":{\"storage\":{\"testKey\":{\"type\":\"Plain\",\"value\":\"testValue\"},\"testKey2\":{\"type\":\"Plain\"}}},\"bar\":{\"storage\":{\"testKey3\":{\"type\":\"Plain\",\"value\":\"testValue3\"}}}}}";
			var directory = new SharedDirectory();

			directory.LoadFromSnapshot(json);

			Assert.Equal(2, directory.Count);
			Assert.Equal("testValue", directory.GetWorkingDirectory("/foo")!.Get("testKey"));
			Assert.Null(directory.GetWorkingDirectory("foo")!.Get("testKey2"));
			Assert.Equal("testValue3", directory.GetWorkingDirectory("/bar")!.Get("testKey3"));
			Assert.Equal("testValue4", directory.GetWorkingDirectory("")!.Get("testKey"));
			Assert.Null(directory.GetWorkingDirectory("/")!.Get("testKey2"));
			Assert.True(directory.Has("testKey2"));
			Assert.True(directory.GetWorkingDirectory("/foo")!.Has("testKey2"));
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.spec.ts:624 — "Should populate, serialize and de-serialize directory with long property values"
		public void ShouldPopulateSerializeAndDeserializeDirectoryWithLongPropertyValuesLoadOnlyEquivalent()
		{
			string longWord = "0123456789";
			for (int i = 0; i < 12; i++)
			{
				longWord += longWord;
			}

			string longWord2 = $"{longWord}_2";
			string json = "{\"blobs\":[\"blob0\"],\"content\":{\"storage\":{\"first\":{\"type\":\"Plain\",\"value\":\"second\"},\"long1\":{\"type\":\"Plain\",\"value\":"
				+ JsonSerializer.Serialize(longWord)
				+ "}},\"subdirectories\":{\"nested\":{\"storage\":{\"deepKey1\":{\"type\":\"Plain\",\"value\":\"deepValue1\"}}}}}}";
			var blobs = new Dictionary<string, string>()
			{
				["blob0"] = "{\"subdirectories\":{\"nested\":{\"storage\":{\"long2\":{\"type\":\"Plain\",\"value\":" + JsonSerializer.Serialize(longWord2) + "}}}}}",
			};
			var directory = new SharedDirectory();

			// TS-parity deviation: our port does not expose a serialize path, so this asserts load-only equivalence from the equivalent blob-split snapshot shape.
			directory.LoadFromSnapshot(json, blobName => blobs[blobName]);

			Assert.Equal("second", directory.Get("first"));
			Assert.Equal(longWord, directory.Get("long1"));
			IDirectory nested = Assert.IsAssignableFrom<IDirectory>(directory.GetWorkingDirectory("/nested"));
			Assert.Equal("deepValue1", nested.Get("deepKey1"));
			Assert.Equal(longWord2, nested.Get("long2"));
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.spec.ts:681 — "Should register subdirectory events on load"
		public void ShouldRegisterSubdirectoryEventsOnLoad()
		{
			const string json = "{\"subdirectories\":{\"child\":{}}}";
			var directory = new SharedDirectory();
			directory.LoadFromSnapshot(json);
			IDirectory child = Assert.IsAssignableFrom<IDirectory>(directory.GetSubDirectory("child"));
			int createdCount = 0;
			int deletedCount = 0;
			directory.OnSubDirectoryCreated += (s, e) =>
			{
				createdCount++;
				Assert.Equal("grandchild", e.SubdirName);
				Assert.Equal("/child", e.ParentPath);
				Assert.True(e.Local);
			};
			directory.OnSubDirectoryDeleted += (s, e) =>
			{
				deletedCount++;
				Assert.Equal("grandchild", e.SubdirName);
				Assert.Equal("/child", e.ParentPath);
				Assert.True(e.Local);
			};

			child.CreateSubDirectory("grandchild");
			child.DeleteSubDirectory("grandchild");

			Assert.Equal(1, createdCount);
			Assert.Equal(1, deletedCount);
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.spec.ts:728 — "Should lead to eventual consistency 1"
		public void ShouldLeadToEventualConsistency1DeleteRecreateNestedSubdirectory()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);

			IDirectory parent1 = directory.CreateSubDirectory("lists");
			ProcessLocalAck(directory, sender.Sent[0], seq: 10, refSeq: 0);
			IDirectory oldChild = parent1.CreateSubDirectory("ListLevels-0");
			ProcessLocalAck(directory, sender.Sent[1], seq: 11, refSeq: 10);
			oldChild.Set("random1", 1);
			ProcessLocalAck(directory, sender.Sent[2], seq: 12, refSeq: 11);
			oldChild.Set("random2", 2);
			ProcessLocalAck(directory, sender.Sent[3], seq: 13, refSeq: 12);

			oldChild.Set("random1", 3);
			var staleSet = sender.Sent[4];
			parent1.DeleteSubDirectory("ListLevels-0");
			ProcessLocalAck(directory, sender.Sent[5], seq: 20, refSeq: 13);
			directory.DeleteSubDirectory("lists");
			ProcessLocalAck(directory, sender.Sent[6], seq: 21, refSeq: 20);
			IDirectory parent2 = directory.CreateSubDirectory("lists");
			ProcessLocalAck(directory, sender.Sent[7], seq: 30, refSeq: 21);
			IDirectory newChild = parent2.CreateSubDirectory("ListLevels-0");
			ProcessLocalAck(directory, sender.Sent[8], seq: 31, refSeq: 30);
			newChild.Set("random1", 4);
			ProcessLocalAck(directory, sender.Sent[9], seq: 32, refSeq: 31);

			ProcessLocalAck(directory, staleSet, seq: 40, refSeq: 13);

			IDirectory loadedChild = Assert.IsAssignableFrom<IDirectory>(directory.GetSubDirectory("lists")!.GetSubDirectory("ListLevels-0"));
			Assert.Same(newChild, loadedChild);
			Assert.Equal(4, loadedChild.Get("random1"));
			Assert.False(loadedChild.Has("random2"));
			Assert.Null(loadedChild.Get("random2"));
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.spec.ts:1125 — "Shouldn't overwrite value if there is pending set"
		public void ShouldNotOverwriteValueIfThereIsPendingSet()
		{
			ConnectedDirectory client1 = CreateConnectedDirectory("directory1", "client-1");
			ConnectedDirectory client2 = CreateConnectedDirectory("directory2", "client-2");
			var clients = new[] { client1, client2 };

			client1.Directory.Set("test", "value1");
			client2.Directory.Set("test", "pending1");
			client2.Directory.Set("test", "pending2");

			ProcessSent(clients, client1, client1.Sender.Sent[0], seq: 1, refSeq: 0);
			Assert.Equal("value1", client1.Directory.Get("test"));
			Assert.Equal("pending2", client2.Directory.Get("test"));

			ProcessSent(clients, client2, client2.Sender.Sent[0], seq: 2, refSeq: 1);
			Assert.Equal("pending1", client1.Directory.Get("test"));
			Assert.Equal("pending2", client2.Directory.Get("test"));

			ProcessSent(clients, client2, client2.Sender.Sent[1], seq: 3, refSeq: 2);
			Assert.Equal("pending2", client1.Directory.Get("test"));
			Assert.Equal("pending2", client2.Directory.Get("test"));
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.spec.ts:1173 — "Shouldn't set values when pending clear"
		public void ShouldNotSetValuesWhenPendingClear()
		{
			ConnectedDirectory client1 = CreateConnectedDirectory("directory1", "client-1");
			ConnectedDirectory client2 = CreateConnectedDirectory("directory2", "client-2");
			var clients = new[] { client1, client2 };

			client1.Directory.Set("test", "directory1value1");
			client2.Directory.Set("test", "directory2value2");
			client2.Directory.Clear();
			client2.Directory.Set("test", "directory2value3");
			client2.Directory.Clear();

			ProcessSent(clients, client1, client1.Sender.Sent[0], seq: 1, refSeq: 0);
			Assert.Equal("directory1value1", client1.Directory.Get("test"));
			Assert.False(client2.Directory.Has("test"));

			ProcessSent(clients, client2, client2.Sender.Sent[0], seq: 2, refSeq: 1);
			Assert.Equal("directory2value2", client1.Directory.Get("test"));
			Assert.False(client2.Directory.Has("test"));

			ProcessSent(clients, client2, client2.Sender.Sent[1], seq: 3, refSeq: 2);
			Assert.False(client1.Directory.Has("test"));
			Assert.False(client2.Directory.Has("test"));

			ProcessSent(clients, client2, client2.Sender.Sent[2], seq: 4, refSeq: 3);
			Assert.Equal("directory2value3", client1.Directory.Get("test"));
			Assert.False(client2.Directory.Has("test"));

			ProcessSent(clients, client2, client2.Sender.Sent[3], seq: 5, refSeq: 4);
			Assert.False(client1.Directory.Has("test"));
			Assert.False(client2.Directory.Has("test"));

			client1.Directory.Set("test", "directory1value4");
			ProcessSent(clients, client1, client1.Sender.Sent[1], seq: 6, refSeq: 5);
			Assert.Equal("directory1value4", client1.Directory.Get("test"));
			Assert.Equal("directory1value4", client2.Directory.Get("test"));
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.spec.ts:1244 — "Directories should ensure eventual consistency using LWW approach 1: Test 1"
		public void DirectoriesShouldEnsureEventualConsistencyUsingLwwApproach1Test1()
		{
			ConnectedDirectory client1 = CreateConnectedDirectory("directory1", "client-1");
			ConnectedDirectory client2 = CreateConnectedDirectory("directory2", "client-2");
			var clients = new[] { client1, client2 };

			IDirectory root1SubDir = client1.Directory.CreateSubDirectory("testSubDir");
			root1SubDir.Set("key1", "testValue1");
			client1.Directory.DeleteSubDirectory("testSubDir");
			IDirectory root1SubDir2 = client1.Directory.CreateSubDirectory("testSubDir");
			root1SubDir2.Set("key2", "testValue2");
			client2.Directory.CreateSubDirectory("testSubDir");

			ProcessSent(clients, client1, client1.Sender.Sent[0], seq: 1, refSeq: 0);
			ProcessSent(clients, client1, client1.Sender.Sent[1], seq: 2, refSeq: 1);
			ProcessSent(clients, client1, client1.Sender.Sent[2], seq: 3, refSeq: 2);
			ProcessSent(clients, client1, client1.Sender.Sent[3], seq: 4, refSeq: 3);
			ProcessSent(clients, client1, client1.Sender.Sent[4], seq: 5, refSeq: 4);
			ProcessSent(clients, client2, client2.Sender.Sent[0], seq: 6, refSeq: 0);

			IDirectory directory1SubDir = Assert.IsAssignableFrom<IDirectory>(client1.Directory.GetSubDirectory("testSubDir"));
			IDirectory directory2SubDir = Assert.IsAssignableFrom<IDirectory>(client2.Directory.GetSubDirectory("testSubDir"));
			Assert.Equal(1, directory1SubDir.Count);
			Assert.Equal(1, directory2SubDir.Count);
			Assert.Equal("testValue2", directory1SubDir.Get("key2"));
			Assert.Equal("testValue2", directory2SubDir.Get("key2"));
			Assert.False(directory1SubDir.Has("key1"));
			Assert.False(directory2SubDir.Has("key1"));
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.spec.ts:1275 — "Directories should ensure eventual consistency using LWW approach 1: Test 2"
		public void DirectoriesShouldEnsureEventualConsistencyUsingLwwApproach1Test2()
		{
			ConnectedDirectory client1 = CreateConnectedDirectory("directory1", "client-1");
			ConnectedDirectory client2 = CreateConnectedDirectory("directory2", "client-2");
			var clients = new[] { client1, client2 };

			IDirectory root1SubDir = client1.Directory.CreateSubDirectory("testSubDir");
			client2.Directory.CreateSubDirectory("testSubDir");
			root1SubDir.Set("key1", "testValue1");
			client2.Directory.DeleteSubDirectory("testSubDir");
			client2.Directory.CreateSubDirectory("testSubDir");

			ProcessSent(clients, client1, client1.Sender.Sent[0], seq: 1, refSeq: 0);
			ProcessSent(clients, client2, client2.Sender.Sent[0], seq: 2, refSeq: 0);
			ProcessSent(clients, client1, client1.Sender.Sent[1], seq: 3, refSeq: 1);
			ProcessSent(clients, client2, client2.Sender.Sent[1], seq: 4, refSeq: 2);
			ProcessSent(clients, client2, client2.Sender.Sent[2], seq: 5, refSeq: 4);

			IDirectory directory1SubDir = Assert.IsAssignableFrom<IDirectory>(client1.Directory.GetSubDirectory("testSubDir"));
			IDirectory directory2SubDir = Assert.IsAssignableFrom<IDirectory>(client2.Directory.GetSubDirectory("testSubDir"));
			Assert.Equal(0, directory1SubDir.Count);
			Assert.Equal(0, directory2SubDir.Count);
			Assert.False(directory1SubDir.Has("key1"));
			Assert.False(directory2SubDir.Has("key1"));
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.spec.ts:454 — "Should serialize a directory without subdirectories as a JSON object"
		public void ShouldSerializeDirectoryWithoutSubdirectoriesHandleGetReturnsObject()
		{
			var sender = new FakeFluidDataObjectSender();
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("subMap");
			registry.Register(handle, "/dataObjects/subMap");
			var directory = new SharedDirectory("directory", sender, registry);

			// TS-parity deviation: our port does not expose a serialize path, so this verifies the local Get value and emitted handle wire shape.
			directory.Set("object", handle);

			Assert.Same(handle, directory.Get("object"));
			using JsonDocument document = JsonDocument.Parse(Assert.Single(sender.Sent).OpJson);
			JsonElement handleWire = document.RootElement.GetProperty("value").GetProperty("value");
			Assert.Equal("__fluid_handle__", handleWire.GetProperty("type").GetString());
			Assert.Equal("/dataObjects/subMap", handleWire.GetProperty("url").GetString());
		}

		private static ConnectedDirectory CreateConnectedDirectory(string id, string clientId)
		{
			var sender = new FakeFluidDataObjectSender()
			{
				LocalClientId = clientId,
			};
			return new ConnectedDirectory(new SharedDirectory(id, sender), sender, clientId);
		}

		private static void ProcessSent(
			IEnumerable<ConnectedDirectory> clients,
			ConnectedDirectory origin,
			(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
			long seq,
			long refSeq)
		{
			foreach (ConnectedDirectory client in clients)
			{
				OpOrigin opOrigin = ReferenceEquals(client, origin) ? OpOrigin.Local : OpOrigin.Remote;
				client.Directory.ProcessDataObjectOp(
					new SequencedDocumentMessageDescriptor(
						SequenceNumber.ForTesting(clientSeq: sent.ClientSeq, refSeq: refSeq, seq: seq),
						opOrigin,
						origin.ClientId),
					sent.OpJson);
			}
		}

		private static void ProcessLocalAck(
			SharedDirectory directory,
			(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
			long seq,
			long refSeq,
			string clientId = "local-client")
		{
			directory.ProcessDataObjectOp(
				new SequencedDocumentMessageDescriptor(
					SequenceNumber.ForTesting(clientSeq: sent.ClientSeq, refSeq: refSeq, seq: seq),
					OpOrigin.Local,
					clientId),
				sent.OpJson);
		}

		private sealed record ConnectedDirectory(SharedDirectory Directory, FakeFluidDataObjectSender Sender, string ClientId);

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
