// -----------------------------------------------------------------------------
// Tests for SharedDirectory subdirectory incarnation filtering.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class DirectoryInstanceFilterTests
	{
		private static readonly BindingFlags PrivateInstance = BindingFlags.Instance | BindingFlags.NonPublic;

		[Fact]
		public void DeleteRecreate_StaleAckForOldInstance_Ignored()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);

			IDirectory oldFoo = directory.CreateSubDirectory("foo");
			var oldCreate = sender.Sent[0];
			ProcessLocalAck(directory, oldCreate, seq: 10, refSeq: 0);

			oldFoo.Set("stale", "old");
			var staleSet = sender.Sent[1];
			directory.DeleteSubDirectory("foo");
			var oldDelete = sender.Sent[2];
			ProcessLocalAck(directory, oldDelete, seq: 20, refSeq: 10);

			IDirectory newFoo = directory.CreateSubDirectory("foo");
			var newCreate = sender.Sent[3];
			ProcessLocalAck(directory, newCreate, seq: 30, refSeq: 20);

			ProcessLocalAck(directory, staleSet, seq: 40, refSeq: 10);

			Assert.Same(newFoo, directory.GetSubDirectory("foo"));
			Assert.False(newFoo.Has("stale"));
			Assert.Null(newFoo.Get("stale"));
		}

		[Fact]
		public void FilteredLocalAck_StaleInstance_ClearsRootAndSubdirQueues()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);

			IDirectory oldFoo = directory.CreateSubDirectory("foo");
			ProcessLocalAck(directory, sender.Sent[0], seq: 10, refSeq: 0);

			oldFoo.Set("stale", "old");
			var staleSet = sender.Sent[1];
			Assert.Single(GetPendingLocalOpSubdirectories(directory));
			Assert.Single(GetPendingStorageData(oldFoo));

			directory.DeleteSubDirectory("foo");
			ProcessLocalAck(directory, sender.Sent[2], seq: 20, refSeq: 10);
			IDirectory newFoo = directory.CreateSubDirectory("foo");
			ProcessLocalAck(directory, sender.Sent[3], seq: 30, refSeq: 20);

			Assert.Single(GetPendingLocalOpSubdirectories(directory));
			Assert.Single(GetPendingStorageData(oldFoo));

			ProcessLocalAck(directory, staleSet, seq: 40, refSeq: 10);

			Assert.Empty(GetPendingLocalOpSubdirectories(directory));
			Assert.Empty(GetPendingStorageData(oldFoo));
			Assert.Same(newFoo, directory.GetSubDirectory("foo"));
			Assert.False(newFoo.Has("stale"));
			Assert.Null(newFoo.Get("stale"));
		}

		[Fact]
		public void FilteredLocalAck_ThenNormalAcksContinue_NoLeak()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);

			IDirectory oldFoo = directory.CreateSubDirectory("foo");
			ProcessLocalAck(directory, sender.Sent[0], seq: 10, refSeq: 0);
			oldFoo.Set("stale", "old");
			var staleSet = sender.Sent[1];
			directory.DeleteSubDirectory("foo");
			ProcessLocalAck(directory, sender.Sent[2], seq: 20, refSeq: 10);
			IDirectory newFoo = directory.CreateSubDirectory("foo");
			ProcessLocalAck(directory, sender.Sent[3], seq: 30, refSeq: 20);

			ProcessLocalAck(directory, staleSet, seq: 40, refSeq: 10);
			newFoo.Set("fresh", "value");
			var freshSet = sender.Sent[4];

			Assert.Single(GetPendingLocalOpSubdirectories(directory));
			Assert.Single(GetPendingStorageData(newFoo));

			ProcessLocalAck(directory, freshSet, seq: 50, refSeq: 30);
			ProcessRemoteSet(directory, "/foo", "remote", "applied", seq: 60, refSeq: 50, clientId: "remote-client");

			Assert.Empty(GetPendingLocalOpSubdirectories(directory));
			Assert.Empty(GetPendingStorageData(newFoo));
			AssertJsonString("value", newFoo.Get("fresh"));
			AssertJsonString("applied", newFoo.Get("remote"));
			Assert.False(newFoo.Has("stale"));
		}

		[Fact]
		public void AcceptedLocalAck_PromotesStampsAsBefore()
		{
			var sender = new FakeFluidDataObjectSender()
			{
				LocalClientId = "local-creator",
			};
			var directory = new SharedDirectory("dir", sender);

			IDirectory foo = directory.CreateSubDirectory("foo");
			var create = sender.Sent[0];

			AssertSeqData(foo, seq: -1, clientSeq: 1);
			ProcessLocalAck(directory, create, seq: 10, refSeq: 0, clientId: "local-creator");

			AssertSeqData(foo, seq: 10, clientSeq: create.ClientSeq);
			Assert.Empty(GetPendingLocalOpSubdirectories(directory));
			Assert.Empty(GetPendingSubDirectoryData(GetRootSubDirectory(directory)));

			foo.Set("k", "value");
			var set = sender.Sent[1];
			Assert.Single(GetPendingStorageData(foo));

			ProcessLocalAck(directory, set, seq: 20, refSeq: 10, clientId: "local-creator");

			Assert.Empty(GetPendingLocalOpSubdirectories(directory));
			Assert.Empty(GetPendingStorageData(foo));
			AssertJsonString("value", foo.Get("k"));
		}

		[Fact]
		public void DeleteRecreate_StaleRemoteOpForOldInstance_Ignored()
		{
			var directory = new SharedDirectory();
			ProcessRemoteCreateSubDirectory(directory, "foo", seq: 10, refSeq: 0, clientId: "creator-a");
			ProcessRemoteDeleteSubDirectory(directory, "foo", seq: 20, refSeq: 10, clientId: "creator-a");
			ProcessRemoteCreateSubDirectory(directory, "foo", seq: 30, refSeq: 20, clientId: "creator-b");

			ProcessRemoteSet(directory, "/foo", "stale", "old", seq: 40, refSeq: 11, clientId: "creator-a");

			IDirectory? newFoo = directory.GetSubDirectory("foo");
			Assert.NotNull(newFoo);
			Assert.False(newFoo!.Has("stale"));
			Assert.Null(newFoo.Get("stale"));
		}

		[Fact]
		public void DeleteRecreate_OwnCreatesRemainVisible()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);

			directory.CreateSubDirectory("foo");
			ProcessLocalAck(directory, sender.Sent[0], seq: 10, refSeq: 0);
			directory.DeleteSubDirectory("foo");
			ProcessLocalAck(directory, sender.Sent[1], seq: 20, refSeq: 10);

			IDirectory newFoo = directory.CreateSubDirectory("foo");
			ProcessLocalAck(directory, sender.Sent[2], seq: 30, refSeq: 20);
			newFoo.Set("fresh", "value");
			ProcessLocalAck(directory, sender.Sent[3], seq: 40, refSeq: 30);

			Assert.Same(newFoo, directory.GetSubDirectory("foo"));
			Assert.Equal("value", newFoo.Get("fresh"));
		}

		[Fact]
		public void DetachedCreated_ThenClientJoins_OpsApply()
		{
			var directory = new SharedDirectory();
			IDirectory foo = directory.CreateSubDirectory("foo");

			ProcessRemoteSet(directory, "/foo", "remote", "value", seq: 1, refSeq: 0, clientId: "joining-client");

			AssertJsonString("value", foo.Get("remote"));
		}

		[Fact]
		public void RemoteRefSeqBeforeCreate_Filtered()
		{
			var directory = new SharedDirectory();
			ProcessRemoteCreateSubDirectory(directory, "foo", seq: 100, refSeq: 0, clientId: "creator");

			ProcessRemoteSet(directory, "/foo", "k", "stale", seq: 101, refSeq: 99, clientId: "other-client");

			IDirectory? foo = directory.GetSubDirectory("foo");
			Assert.NotNull(foo);
			Assert.False(foo!.Has("k"));
			Assert.Null(foo.Get("k"));
		}

		[Fact]
		public void RemoteRefSeqAfterCreate_Applied()
		{
			var directory = new SharedDirectory();
			ProcessRemoteCreateSubDirectory(directory, "foo", seq: 100, refSeq: 0, clientId: "creator");

			ProcessRemoteSet(directory, "/foo", "k", "current", seq: 101, refSeq: 100, clientId: "other-client");

			IDirectory? foo = directory.GetSubDirectory("foo");
			Assert.NotNull(foo);
			AssertJsonString("current", foo!.Get("k"));
		}

		[Fact]
		public void LocalOpMetadata_SubdirIdentity_MismatchIgnored()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);

			IDirectory oldFoo = directory.CreateSubDirectory("foo");
			ProcessLocalAck(directory, sender.Sent[0], seq: 10, refSeq: 0);
			oldFoo.Set("victim", "old");
			ProcessLocalAck(directory, sender.Sent[1], seq: 11, refSeq: 10);

			oldFoo.Delete("victim");
			var staleDelete = sender.Sent[2];
			ProcessRemoteDeleteSubDirectory(directory, "foo", seq: 20, refSeq: 11, clientId: "remote-a");
			ProcessRemoteCreateSubDirectory(directory, "foo", seq: 30, refSeq: 20, clientId: "remote-b");
			ProcessRemoteSet(directory, "/foo", "victim", "new", seq: 31, refSeq: 30, clientId: "remote-b");

			ProcessLocalAck(directory, staleDelete, seq: 40, refSeq: 11);

			IDirectory? newFoo = directory.GetSubDirectory("foo");
			Assert.NotNull(newFoo);
			AssertJsonString("new", newFoo!.Get("victim"));
		}

		[Fact]
		public void RootSubdir_HasEmptyClientIds_NotDetached()
		{
			var directory = new SharedDirectory();
			object root = GetRootSubDirectory(directory);

			Assert.Empty(GetClientIds(root));

			// TS parity: the root has an empty creator set, so it must not accept through the "detached" branch.
			ProcessRemoteSet(directory, "/", "root", "stale", seq: 1, refSeq: -1, clientId: null);
			Assert.False(directory.Has("root"));

			ProcessRemoteSet(directory, "/", "root", "value", seq: 2, refSeq: 0, clientId: null);
			AssertJsonString("value", directory.Get("root"));
		}

		[Fact]
		public void LocalCreatedSubdir_WhilePending_CreatorClientIdKnown()
		{
			var sender = new FakeFluidDataObjectSender()
			{
				LocalClientId = "local-creator",
			};
			var directory = new SharedDirectory("dir", sender);

			IDirectory foo = directory.CreateSubDirectory("foo");

			Assert.Equal(new[] { "local-creator" }, GetClientIds(foo).ToArray());
			Assert.True(IsMessageForCurrentInstance(foo, RemoteMessage(seq: 5, refSeq: -1, clientId: "local-creator")));
			Assert.False(IsMessageForCurrentInstance(foo, RemoteMessage(seq: 5, refSeq: -1, clientId: "other-client")));
		}

		[Fact]
		public void SnapshotLoad_ReadsCiField_SeedsSeqDataAndClientIds()
		{
			const string json = "{\"subdirectories\":{\"foo\":{\"ci\":{\"csn\":42,\"ccIds\":[\"snapshot-creator\"]}}}}";
			var directory = new SharedDirectory();

			directory.LoadFromSnapshot(json);

			IDirectory foo = Assert.IsAssignableFrom<IDirectory>(directory.GetSubDirectory("foo"));
			AssertSeqData(foo, seq: 42, clientSeq: 0);
			Assert.Equal(new[] { "snapshot-creator" }, GetClientIds(foo).ToArray());

			ProcessRemoteSet(directory, "/foo", "fromCreator", "applied", seq: 43, refSeq: -1, clientId: "snapshot-creator");
			AssertJsonString("applied", foo.Get("fromCreator"));
		}

		[Fact]
		public void SnapshotLoad_OldFormatWithoutCi_FallsBackGracefully()
		{
			const string json = "{\"subdirectories\":{\"foo\":{}}}";
			var directory = new SharedDirectory();

			directory.LoadFromSnapshot(json);

			IDirectory foo = Assert.IsAssignableFrom<IDirectory>(directory.GetSubDirectory("foo"));
			AssertSeqData(foo, seq: 0, clientSeq: 1);
			Assert.Empty(GetClientIds(foo));

			ProcessRemoteSet(directory, "/foo", "afterLoad", "value", seq: 2, refSeq: 0, clientId: "remote-client");
			AssertJsonString("value", foo.Get("afterLoad"));
		}

		[Fact]
		public void SubdirectoryOrdering_ConcurrentCreates_MatchesSeqDataComparator()
		{
			var directory = new SharedDirectory();

			ProcessRemoteCreateSubDirectory(directory, "second", seq: 10, refSeq: 0, clientId: "client-b", clientSeq: 2);
			ProcessRemoteCreateSubDirectory(directory, "first", seq: 10, refSeq: 0, clientId: "client-a", clientSeq: 1);

			Assert.Equal(new[] { "first", "second" }, directory.SubDirectories().Select(kvp => kvp.Key).ToArray());
			AssertSeqData(Assert.IsAssignableFrom<IDirectory>(directory.GetSubDirectory("first")), seq: 10, clientSeq: 1);
			AssertSeqData(Assert.IsAssignableFrom<IDirectory>(directory.GetSubDirectory("second")), seq: 10, clientSeq: 2);
		}

		[Fact]
		public void SubdirectoryOrdering_MixedSequencedAndPending_CorrectOrder()
		{
			var sender = new FakeFluidDataObjectSender()
			{
				LocalClientId = "local-creator",
			};
			var directory = new SharedDirectory("dir", sender);

			ProcessRemoteCreateSubDirectory(directory, "late", seq: 20, refSeq: 0, clientId: "remote-late", clientSeq: 1);
			IDirectory pending = directory.CreateSubDirectory("pending");
			ProcessRemoteCreateSubDirectory(directory, "early", seq: 10, refSeq: 0, clientId: "remote-early", clientSeq: 1);

			Assert.Equal(new[] { "early", "late", "pending" }, directory.SubDirectories().Select(kvp => kvp.Key).ToArray());
			AssertSeqData(pending, seq: -1, clientSeq: 1);
			Assert.Equal(new[] { "local-creator" }, GetClientIds(pending).ToArray());
		}

		[Fact]
		public void RemoteOp_MissingSubdirPath_Ignored()
		{
			var directory = new SharedDirectory();
			directory.Set("existing", "value");

			ProcessRemoteSet(directory, "/missing", "k", "remote", seq: 1, refSeq: 0, clientId: "remote-client");

			Assert.False(directory.HasSubDirectory("missing"));
			Assert.Equal("value", directory.Get("existing"));
			Assert.Equal(new[] { "existing" }, directory.Keys);
		}

		[Fact]
		public void RemoteOp_MissingNestedPath_Ignored()
		{
			var directory = new SharedDirectory();
			IDirectory foo = directory.CreateSubDirectory("foo");
			foo.Set("existing", "value");

			ProcessRemoteDelete(directory, "/foo/bar/baz", "k", seq: 1, refSeq: 0, clientId: "remote-client");

			Assert.Same(foo, directory.GetSubDirectory("foo"));
			Assert.False(foo.HasSubDirectory("bar"));
			Assert.Equal("value", foo.Get("existing"));
		}

		[Fact]
		public void LocalOp_MissingSubdirPath_StillThrows()
		{
			var directory = new SharedDirectory();
			string opJson = DirectoryOpSerializer.Serialize(new DirectorySetOperation()
			{
				Path = "/missing",
				Key = "k",
				Value = new SerializableValue()
				{
					Type = "Plain",
					Value = "local",
				},
			});

			OcsException exception = Assert.Throws<OcsException>(
				() => directory.ProcessDataObjectOp(
					new SequencedDocumentMessageDescriptor(
						SequenceNumber.ForTesting(clientSeq: 1, refSeq: 0, seq: 1),
						OpOrigin.Local,
						"local-client"),
					opJson));
			Assert.Equal(OcsGateErrorCode.InvalidOperation, exception.ErrorCode);
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

		private static void ProcessRemoteCreateSubDirectory(
			SharedDirectory directory,
			string subdirName,
			long seq,
			long refSeq,
			string? clientId,
			long clientSeq = 0)
		{
			directory.ProcessDataObjectOp(
				RemoteMessage(seq, refSeq, clientId, clientSeq),
				DirectoryOpSerializer.Serialize(new DirectoryCreateSubDirectoryOperation()
				{
					Path = "/",
					SubdirName = subdirName,
				}));
		}

		private static void ProcessRemoteDeleteSubDirectory(
			SharedDirectory directory,
			string subdirName,
			long seq,
			long refSeq,
			string? clientId)
		{
			directory.ProcessDataObjectOp(
				RemoteMessage(seq, refSeq, clientId),
				DirectoryOpSerializer.Serialize(new DirectoryDeleteSubDirectoryOperation()
				{
					Path = "/",
					SubdirName = subdirName,
				}));
		}

		private static void ProcessRemoteSet(
			SharedDirectory directory,
			string path,
			string key,
			object? value,
			long seq,
			long refSeq,
			string? clientId)
		{
			directory.ProcessDataObjectOp(
				RemoteMessage(seq, refSeq, clientId),
				DirectoryOpSerializer.Serialize(new DirectorySetOperation()
				{
					Path = path,
					Key = key,
					Value = new SerializableValue()
					{
						Type = "Plain",
						Value = value,
					},
				}));
		}

		private static void ProcessRemoteDelete(
			SharedDirectory directory,
			string path,
			string key,
			long seq,
			long refSeq,
			string? clientId)
		{
			directory.ProcessDataObjectOp(
				RemoteMessage(seq, refSeq, clientId),
				DirectoryOpSerializer.Serialize(new DirectoryDeleteOperation()
				{
					Path = path,
					Key = key,
				}));
		}

		private static SequencedDocumentMessageDescriptor RemoteMessage(long seq, long refSeq, string? clientId, long clientSeq = 0)
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: clientSeq, refSeq: refSeq, seq: seq),
				OpOrigin.Remote,
				clientId);
		}

		private static object GetRootSubDirectory(SharedDirectory directory)
		{
			return GetField<object>(directory, "_root");
		}

		private static IEnumerable<string> GetClientIds(object subdir)
		{
			return GetProperty<IEnumerable<string>>(subdir, "ClientIds");
		}

		private static void AssertSeqData(object subdir, long seq, long clientSeq)
		{
			object seqData = GetProperty<object>(subdir, "SeqData");
			Assert.Equal(seq, GetProperty<long>(seqData, "Seq"));
			Assert.Equal(clientSeq, GetProperty<long>(seqData, "ClientSeq"));
		}

		private static bool IsMessageForCurrentInstance(object subdir, SequencedDocumentMessageDescriptor msg)
		{
			MethodInfo method = subdir.GetType().GetMethod("IsMessageForCurrentInstanceOfSubDirectory", PrivateInstance)!;
			return Assert.IsType<bool>(method.Invoke(subdir, new object?[] { msg, null }));
		}

		private static IDictionary GetPendingLocalOpSubdirectories(SharedDirectory directory)
		{
			return GetField<IDictionary>(directory, "_pendingLocalOpSubdirectories");
		}

		private static IList GetPendingStorageData(object subdir)
		{
			return GetField<IList>(subdir, "_pendingStorageData");
		}

		private static IList GetPendingSubDirectoryData(object subdir)
		{
			return GetField<IList>(subdir, "_pendingSubDirectoryData");
		}

		private static T GetField<T>(object instance, string name)
		{
			object? value = instance.GetType().GetField(name, PrivateInstance)!.GetValue(instance);
			return Assert.IsAssignableFrom<T>(value);
		}

		private static T GetProperty<T>(object instance, string name)
		{
			object? value = instance.GetType().GetProperty(name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)!.GetValue(instance);
			return Assert.IsAssignableFrom<T>(value);
		}

		private static void AssertJsonString(string expected, object? actual)
		{
			// TS-parity: JSON values now materialize to native types (Finding 18).
			Assert.Equal(expected, Assert.IsType<string>(actual));
		}
	}
}
