using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class DirectoryPendingLifetimeTests
	{
		private static readonly BindingFlags PrivateInstance = BindingFlags.Instance | BindingFlags.NonPublic;

		[Fact]
		public void TwoSetsToSameKey_ShareOneLifetime_AndFirstAckLeavesSecondPending()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);

			directory.Set("k", "first");
			directory.Set("k", "second");

			Assert.Equal(2, sender.Sent.Count);
			AssertSinglePendingKeyLifetime(directory, "k", 2);

			ProcessLocalAck(directory, sender.Sent[0]);

			AssertSinglePendingKeyLifetime(directory, "k", 1);
			Assert.Equal("second", directory.Get("k"));

			ProcessLocalAck(directory, sender.Sent[1]);

			Assert.Empty(GetPendingStorageData(directory));
			Assert.Equal("second", directory.Get("k"));
		}

		[Fact]
		public void GetDuringMultiSet_ReturnsLatestPendingValue_NotAckedValue()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);

			directory.Set("k", "first");
			directory.Set("k", "second");
			ProcessLocalAck(directory, sender.Sent[0]);

			Assert.Equal("second", directory.Get("k"));
			Assert.True(directory.Has("k"));
		}

		[Fact]
		public void EntriesKeysValuesAndCount_UseOptimisticPendingView()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);
			ProcessRemoteSet(directory, "stable", "base");

			directory.Set("stable", "local");
			directory.Set("new", "pending");

			Dictionary<string, object?> entries = directory.Entries().ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
			Assert.Equal(2, directory.Count);
			Assert.Equal(new[] { "stable", "new" }, directory.Keys.ToArray());
			Assert.Equal("local", entries["stable"]);
			Assert.Equal("pending", entries["new"]);
			Assert.Contains("local", directory.Values);
			Assert.Contains("pending", directory.Values);
		}

		[Fact]
		public void RemoteSetToKeyWithLocalPending_LocalPendingWinsUntilAck()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);
			int eventCount = 0;
			directory.OnValueChanged += (s, e) => eventCount++;

			directory.Set("k", "local");
			var localSet = Assert.Single(sender.Sent);

			// Behavior change from Option B → Option A: remote ops still update
			// sequenced storage, but the pending lifetime owns the optimistic view.
			ProcessRemoteSet(directory, "k", "remote");

			Assert.Equal("local", directory.Get("k"));
			Assert.Equal("local", Assert.Single(directory.Entries()).Value);
			Assert.Equal(1, eventCount);

			ProcessLocalAck(directory, localSet);
			Assert.Equal("local", directory.Get("k"));
		}

		[Fact]
		public void PendingClearFollowedByPendingSet_SetSurvivesAfterClearAck()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);
			ProcessRemoteSet(directory, "removed", "base");

			directory.Clear();
			var clear = sender.Sent[0];
			directory.Set("survivor", "pending");
			var set = sender.Sent[1];

			Assert.False(directory.Has("removed"));
			Assert.True(directory.Has("survivor"));
			Assert.Equal(new[] { "survivor" }, directory.Keys.ToArray());

			ProcessLocalAck(directory, clear);

			Assert.False(directory.Has("removed"));
			Assert.Equal("pending", directory.Get("survivor"));
			Assert.Equal(new[] { "survivor" }, directory.Keys.ToArray());

			ProcessLocalAck(directory, set);
			Assert.Equal("pending", directory.Get("survivor"));
		}

		[Fact]
		public void PendingCreateThenDeleteSubDirectory_OptimisticIterationReflectsQueueOrder()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);

			directory.CreateSubDirectory("child");
			var create = sender.Sent[0];
			Assert.True(directory.HasSubDirectory("child"));

			directory.DeleteSubDirectory("child");
			var delete = sender.Sent[1];

			Assert.False(directory.HasSubDirectory("child"));
			Assert.Equal(0, directory.CountSubDirectory());
			Assert.Empty(directory.SubDirectories());
			Assert.Equal(
				new[] { "PendingSubDirectoryCreate", "PendingSubDirectoryDelete" },
				GetPendingSubDirectoryData(directory).Cast<object>().Select(entry => entry.GetType().Name).ToArray());

			ProcessLocalAck(directory, create);
			Assert.False(directory.HasSubDirectory("child"));
			Assert.Empty(directory.SubDirectories());

			ProcessLocalAck(directory, delete);
			Assert.False(directory.HasSubDirectory("child"));
			Assert.Empty(GetPendingSubDirectoryData(directory));
		}

		[Fact]
		public void AckingSecondSetBeforeFirst_ThrowsAndLeavesLifetimeIntact()
		{
			var sender = new FakeFluidDataObjectSender();
			var directory = new SharedDirectory("dir", sender);

			directory.Set("k", "first");
			directory.Set("k", "second");

			LoggingError ex = Assert.Throws<LoggingError>(() => ProcessLocalAck(directory, sender.Sent[1]));
			AssertSinglePendingKeyLifetime(directory, "k", 2);
			Assert.Equal("second", directory.Get("k"));

			ProcessLocalAck(directory, sender.Sent[0]);
			ProcessLocalAck(directory, sender.Sent[1]);
			Assert.Empty(GetPendingStorageData(directory));
		}

		private static void AssertSinglePendingKeyLifetime(SharedDirectory directory, string key, int keySetCount)
		{
			IList pendingStorageData = GetPendingStorageData(directory);
			object lifetime = Assert.Single(pendingStorageData.Cast<object>());
			Assert.Equal("PendingKeyLifetime", lifetime.GetType().Name);
			Assert.Equal(key, GetProperty<string>(lifetime, "Key"));
			Assert.Equal(keySetCount, GetProperty<IList>(lifetime, "KeySets").Count);
		}

		private static IList GetPendingStorageData(SharedDirectory directory)
		{
			object root = GetRootSubDirectory(directory);
			return GetField<IList>(root, "_pendingStorageData");
		}

		private static IList GetPendingSubDirectoryData(SharedDirectory directory)
		{
			object root = GetRootSubDirectory(directory);
			return GetField<IList>(root, "_pendingSubDirectoryData");
		}

		private static object GetRootSubDirectory(SharedDirectory directory)
		{
			return GetField<object>(directory, "_root");
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

		private static void ProcessRemoteSet(SharedDirectory directory, string key, object? value)
		{
			directory.ProcessDataObjectOp(
				RemoteMessage(),
				DirectoryOpSerializer.Serialize(new DirectorySetOperation()
				{
					Path = "/",
					Key = key,
					Value = new SerializableValue()
					{
						Type = "Plain",
						Value = value,
					},
				}));
		}
	}
}
