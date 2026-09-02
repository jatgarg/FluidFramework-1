// -----------------------------------------------------------------------------
// Tests for op processing. Structure-inspired by TS
// packages/dds/map/src/test/mocha/directory.spec.ts.
// -----------------------------------------------------------------------------

using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class DirectoryOpProcessingTests
	{
		[Fact]
		public void LocalMutations_EmitExpectedOps_ToSender()
		{
			var setSender = new FakeFluidDataObjectSender();
			var setDirectory = new SharedDirectory("set-dir", setSender);
			setDirectory.Set("k", 42);
			var setSent = Assert.Single(setSender.Sent);
			Assert.Equal("set-dir", setSent.Address);
			Assert.Equal("set", setSent.OpTypeName);
			DirectorySetOperation setOp = Assert.IsType<DirectorySetOperation>(DirectoryOpSerializer.Deserialize(setSent.OpJson));
			Assert.Equal("/", setOp.Path);
			Assert.Equal("k", setOp.Key);
			Assert.Equal("Plain", setOp.Value.Type);
			AssertJsonNumber(42, setOp.Value.Value);

			var deleteSender = new FakeFluidDataObjectSender();
			var deleteDirectory = new SharedDirectory("delete-dir", deleteSender);
			deleteDirectory.Set("k", "value");
			deleteSender.Sent.Clear();
			deleteDirectory.Delete("k");
			var deleteSent = Assert.Single(deleteSender.Sent);
			Assert.Equal("delete", deleteSent.OpTypeName);
			DirectoryDeleteOperation deleteOp = Assert.IsType<DirectoryDeleteOperation>(DirectoryOpSerializer.Deserialize(deleteSent.OpJson));
			Assert.Equal("/", deleteOp.Path);
			Assert.Equal("k", deleteOp.Key);

			var clearSender = new FakeFluidDataObjectSender();
			var clearDirectory = new SharedDirectory("clear-dir", clearSender);
			clearDirectory.Set("k", "value");
			clearSender.Sent.Clear();
			clearDirectory.Clear();
			var clearSent = Assert.Single(clearSender.Sent);
			Assert.Equal("clear", clearSent.OpTypeName);
			DirectoryClearOperation clearOp = Assert.IsType<DirectoryClearOperation>(DirectoryOpSerializer.Deserialize(clearSent.OpJson));
			Assert.Equal("/", clearOp.Path);

			var createSender = new FakeFluidDataObjectSender();
			var createDirectory = new SharedDirectory("create-dir", createSender);
			createDirectory.CreateSubDirectory("foo");
			var createSent = Assert.Single(createSender.Sent);
			Assert.Equal("createSubDirectory", createSent.OpTypeName);
			DirectoryCreateSubDirectoryOperation createOp = Assert.IsType<DirectoryCreateSubDirectoryOperation>(DirectoryOpSerializer.Deserialize(createSent.OpJson));
			Assert.Equal("/", createOp.Path);
			Assert.Equal("foo", createOp.SubdirName);

			var deleteSubdirSender = new FakeFluidDataObjectSender();
			var deleteSubdirDirectory = new SharedDirectory("delete-subdir-dir", deleteSubdirSender);
			deleteSubdirDirectory.CreateSubDirectory("foo");
			deleteSubdirSender.Sent.Clear();
			deleteSubdirDirectory.DeleteSubDirectory("foo");
			var deleteSubdirSent = Assert.Single(deleteSubdirSender.Sent);
			Assert.Equal("deleteSubDirectory", deleteSubdirSent.OpTypeName);
			DirectoryDeleteSubDirectoryOperation deleteSubdirOp = Assert.IsType<DirectoryDeleteSubDirectoryOperation>(DirectoryOpSerializer.Deserialize(deleteSubdirSent.OpJson));
			Assert.Equal("/", deleteSubdirOp.Path);
			Assert.Equal("foo", deleteSubdirOp.SubdirName);
		}

		[Fact]
		public void NestedSet_UsesSubDirectoryPath()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);
			IDirectory foo = directory.CreateSubDirectory("foo");
			IDirectory bar = foo.CreateSubDirectory("bar");
			sender.Sent.Clear();

			bar.Set("k", 1);

			var sent = Assert.Single(sender.Sent);
			DirectorySetOperation op = Assert.IsType<DirectorySetOperation>(DirectoryOpSerializer.Deserialize(sent.OpJson));
			Assert.Equal("/foo/bar", op.Path);
			Assert.Equal("k", op.Key);
		}

		[Fact]
		public void LocalSet_PendingSkipsRemoteUntilAck()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);
			int eventCount = 0;
			directory.OnValueChanged += (s, e) => eventCount++;

			directory.Set("k", "local");
			var localSet = Assert.Single(sender.Sent);

			ProcessRemoteSet(directory, "/", "k", "remote");
			Assert.Equal("local", directory.Get("k"));
			Assert.Equal(1, eventCount);

			ProcessLocalAck(directory, localSet);
			ProcessRemoteSet(directory, "/", "k", "remote");

			AssertJsonString("remote", directory.Get("k"));
			Assert.Equal(2, eventCount);
		}

		[Fact]
		public void LocalClear_PendingSkipsRemoteSetUntilAck()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);
			directory.Set("k", "local");
			ProcessLocalAck(directory, Assert.Single(sender.Sent));
			sender.Sent.Clear();

			directory.Clear();
			var localClear = Assert.Single(sender.Sent);

			ProcessRemoteSet(directory, "/", "remote", "value");
			Assert.False(directory.Has("remote"));

			ProcessLocalAck(directory, localClear);
			ProcessRemoteSet(directory, "/", "remote", "value");

			AssertJsonString("value", directory.Get("remote"));
		}

		[Fact]
		public void OwnOpAck_DoesNotFireEvent_AndPopsPendingMarker()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);
			var events = new List<ValueChangedEventArgs>();
			directory.OnValueChanged += (s, e) => events.Add(e);

			directory.Set("k", 1);
			var localSet = Assert.Single(sender.Sent);
			Assert.Single(events);
			Assert.True(events[0].Local);

			ProcessLocalAck(directory, localSet);
			Assert.Single(events);

			ProcessRemoteSet(directory, "/", "k", 2);
			Assert.Equal(2, events.Count);
			Assert.False(events[1].Local);
			AssertJsonNumber(2, directory.Get("k"));
		}

		[Fact]
		public void OwnOpAck_InvalidPendingState_ThrowsOutOfOrderSequenceNumber()
		{
			var directory = new SharedDirectory("dir", new FakeFluidDataObjectSender());
			string missingPendingOp = CreateSetOpJson("/", "missing", "value");

			LoggingError noPending = Assert.Throws<LoggingError>(
				() => directory.ProcessDataObjectOp(LocalAck(1), missingPendingOp));

			var sender = new FakeFluidDataObjectSender(nextClientSeq: 5);
			var aheadDirectory = new SharedDirectory("ahead-dir", sender);
			aheadDirectory.Set("k", "value");
			var pending = Assert.Single(sender.Sent);
			Assert.Equal(5, pending.ClientSeq);

			LoggingError ackAhead = Assert.Throws<LoggingError>(
				() => aheadDirectory.ProcessDataObjectOp(LocalAck(10), pending.OpJson));
		}

		[Fact]
		public void RemoteSet_AppliesAndFiresEventWithLocalFalse()
		{
			var directory = new SharedDirectory();
			ValueChangedEventArgs? captured = null;
			int eventCount = 0;
			directory.OnValueChanged += (s, e) =>
			{
				captured = e;
				eventCount++;
			};

			ProcessRemoteSet(directory, "/", "k", 42);

			AssertJsonNumber(42, directory.Get("k"));
			Assert.Equal(1, eventCount);
			Assert.NotNull(captured);
			Assert.Equal("k", captured!.Key);
			Assert.Null(captured.PreviousValue);
			Assert.Equal("/", captured.Path);
			Assert.False(captured.Local);
		}

		[Fact]
		public void RemoteSet_SkippedWhenLocalPending()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);
			var events = new List<ValueChangedEventArgs>();
			directory.Set("k", "local");
			directory.OnValueChanged += (s, e) => events.Add(e);

			ProcessRemoteSet(directory, "/", "k", "remote");

			Assert.Equal("local", directory.Get("k"));
			Assert.Empty(events);
		}

		[Fact]
		public void RemoteClear_PreservesPendingKeys()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);
			ProcessRemoteSet(directory, "/", "stable", "remote");
			directory.Set("a", "localA");
			directory.Set("b", "localB");

			ProcessRemoteClear(directory, "/");

			Assert.False(directory.Has("stable"));
			Assert.Equal("localA", directory.Get("a"));
			Assert.Equal("localB", directory.Get("b"));
			Assert.Equal(new[] { "a", "b" }, directory.Keys.OrderBy(key => key).ToArray());
		}

		[Fact]
		public void RemoteDeleteSubDirectory_AppliesWhenNoPending()
		{
			var directory = new SharedDirectory();
			directory.CreateSubDirectory("gone");
			SubDirectoryEventArgs? captured = null;
			int eventCount = 0;
			directory.OnSubDirectoryDeleted += (s, e) =>
			{
				captured = e;
				eventCount++;
			};

			ProcessRemoteOp(directory, new DirectoryDeleteSubDirectoryOperation()
			{
				Path = "/",
				SubdirName = "gone",
			});

			Assert.False(directory.HasSubDirectory("gone"));
			Assert.Equal(1, eventCount);
			Assert.NotNull(captured);
			Assert.Equal("gone", captured!.SubdirName);
			Assert.Equal("/", captured.ParentPath);
			Assert.False(captured.Local);
		}

		[Fact]
		public void RemoteOp_UnknownInput_ThrowsExpectedOcsError()
		{
			var directory = new SharedDirectory();

			LoggingError unknownOp = Assert.Throws<LoggingError>(
				() => ProcessRemoteJson(directory, "{\"type\":\"gibberish\",\"path\":\"/\"}"));
		}

		[Fact]
		public void LoadFromSnapshot_FiresNoEvents()
		{
			const string json = "{\"storage\":{\"root\":{\"type\":\"Plain\",\"value\":\"r\"}},\"subdirectories\":{\"foo\":{\"storage\":{\"child\":{\"type\":\"Plain\",\"value\":\"c\"}}}}}";
			DirectorySnapshotDto snapshot = DirectorySnapshotLoader.Parse(json);
			var directory = new SharedDirectory();
			int valueChangedCount = 0;
			int createdCount = 0;
			int deletedCount = 0;
			directory.OnValueChanged += (s, e) => valueChangedCount++;
			directory.OnSubDirectoryCreated += (s, e) => createdCount++;
			directory.OnSubDirectoryDeleted += (s, e) => deletedCount++;

			directory.LoadFromSnapshot(snapshot);

			Assert.Equal(0, valueChangedCount);
			Assert.Equal(0, createdCount);
			Assert.Equal(0, deletedCount);
			Assert.Equal("r", directory.Get("root"));
			Assert.True(directory.HasSubDirectory("foo"));
		}

		private static void ProcessLocalAck(
			SharedDirectory directory,
			(string Address, string OpTypeName, string OpJson, long ClientSeq) sent)
		{
			directory.ProcessDataObjectOp(LocalAck(sent.ClientSeq), sent.OpJson);
		}

		private static SequencedDocumentMessageDescriptor LocalAck(long clientSeq)
		{
			return new SequencedDocumentMessageDescriptor(SequenceNumber.ForTesting(clientSeq), OpOrigin.Local);
		}

		private static SequencedDocumentMessageDescriptor RemoteMessage()
		{
			return new SequencedDocumentMessageDescriptor(SequenceNumber.ForTesting(0, seq: 1), OpOrigin.Remote);
		}

		private static void ProcessRemoteSet(SharedDirectory directory, string path, string key, object? value)
		{
			ProcessRemoteJson(directory, CreateSetOpJson(path, key, value));
		}

		private static void ProcessRemoteClear(SharedDirectory directory, string path)
		{
			ProcessRemoteOp(directory, new DirectoryClearOperation()
			{
				Path = path,
			});
		}

		private static void ProcessRemoteOp(SharedDirectory directory, DirectoryOperation op)
		{
			ProcessRemoteJson(directory, DirectoryOpSerializer.Serialize(op));
		}

		private static void ProcessRemoteJson(SharedDirectory directory, string opJson)
		{
			directory.ProcessDataObjectOp(RemoteMessage(), opJson);
		}

		private static string CreateSetOpJson(string path, string key, object? value)
		{
			return DirectoryOpSerializer.Serialize(new DirectorySetOperation()
			{
				Path = path,
				Key = key,
				Value = new SerializableValue()
				{
					Type = "Plain",
					Value = value,
				},
			});
		}

		private static void AssertJsonNumber(int expected, object? actual)
		{
			// TS-parity: JSON values now materialize to native types (Finding 18).
			Assert.Equal((double)expected, Assert.IsType<double>(actual));
		}

		private static void AssertJsonString(string expected, object? actual)
		{
			// TS-parity: JSON values now materialize to native types (Finding 18).
			Assert.Equal(expected, Assert.IsType<string>(actual));
		}
	}
}
