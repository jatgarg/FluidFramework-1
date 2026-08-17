// -----------------------------------------------------------------------------
// High-value SharedDirectory iteration tests ported from TS directory.iteration.spec.ts.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class DirectoryIterationTestsFromTS
	{
		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.iteration.spec.ts — "should have eventually consistent iteration order between clients when simultaneous set"
		public void ShouldHaveEventuallyConsistentIterationOrderBetweenClientsWhenSimultaneousSet()
		{
			ConnectedDirectory client1 = CreateConnectedDirectory("shared-directory-1", "client-1");
			ConnectedDirectory client2 = CreateConnectedDirectory("shared-directory-2", "client-2");
			var clients = new[] { client1, client2 };

			client1.Directory.Set("key1", "value1");
			client1.Directory.Set("key2", "value2");
			ProcessSent(clients, client1, client1.Sender.Sent[0], seq: 1, refSeq: 0);
			ProcessSent(clients, client1, client1.Sender.Sent[1], seq: 2, refSeq: 1);

			client1.Directory.Set("key3", "value3");
			client2.Directory.Set("key4", "value4");
			ProcessSent(clients, client1, client1.Sender.Sent[2], seq: 3, refSeq: 2);
			ProcessSent(clients, client2, client2.Sender.Sent[0], seq: 4, refSeq: 2);

			// TS-parity deviation: our port returns a point-in-time list, not a live iterator.
			AssertKeys(client1.Directory, "key1", "key2", "key3", "key4");
			AssertKeys(client2.Directory, "key1", "key2", "key3", "key4");
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.iteration.spec.ts — "should have eventually consistent iteration order between clients when suppressed delete"
		public void ShouldHaveEventuallyConsistentIterationOrderBetweenClientsWhenSuppressedDelete()
		{
			ConnectedDirectory client1 = CreateConnectedDirectory("shared-directory-1", "client-1");
			ConnectedDirectory client2 = CreateConnectedDirectory("shared-directory-2", "client-2");
			var clients = new[] { client1, client2 };

			client1.Directory.Set("key1", "value1");
			client1.Directory.Set("key2", "value2");
			ProcessSent(clients, client1, client1.Sender.Sent[0], seq: 1, refSeq: 0);
			ProcessSent(clients, client1, client1.Sender.Sent[1], seq: 2, refSeq: 1);

			client1.Directory.Delete("key1");
			client2.Directory.Set("key1", "otherValue1");
			ProcessSent(clients, client1, client1.Sender.Sent[2], seq: 3, refSeq: 2);
			ProcessSent(clients, client2, client2.Sender.Sent[0], seq: 4, refSeq: 2);

			// TS-parity deviation: our port returns a point-in-time list backed by Dictionary slot order, not TS Map lifetime order after delete/re-add.
			AssertKeys(client1.Directory, "key1", "key2");
			AssertKeys(client2.Directory, "key1", "key2");
			Assert.Equal("otherValue1", client1.Directory.Get("key1"));
			Assert.Equal("otherValue1", client2.Directory.Get("key1"));
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.iteration.spec.ts — "should have eventually consistent iteration order between clients when clear"
		public void ShouldHaveEventuallyConsistentIterationOrderBetweenClientsWhenClear()
		{
			ConnectedDirectory client1 = CreateConnectedDirectory("shared-directory-1", "client-1");
			ConnectedDirectory client2 = CreateConnectedDirectory("shared-directory-2", "client-2");
			var clients = new[] { client1, client2 };

			client1.Directory.Set("key1", "value1");
			client1.Directory.Set("key2", "value2");
			ProcessSent(clients, client1, client1.Sender.Sent[0], seq: 1, refSeq: 0);
			ProcessSent(clients, client1, client1.Sender.Sent[1], seq: 2, refSeq: 1);

			client1.Directory.Clear();
			client2.Directory.Set("key3", "value3");
			client2.Directory.Set("key1", "otherValue1");
			client2.Directory.Set("key4", "value4");
			ProcessSent(clients, client1, client1.Sender.Sent[2], seq: 3, refSeq: 2);
			ProcessSent(clients, client2, client2.Sender.Sent[0], seq: 4, refSeq: 2);
			ProcessSent(clients, client2, client2.Sender.Sent[1], seq: 5, refSeq: 2);
			ProcessSent(clients, client2, client2.Sender.Sent[2], seq: 6, refSeq: 2);

			// TS-parity deviation: our port returns a point-in-time list, not a live iterator.
			AssertKeys(client1.Directory, "key3", "key1", "key4");
			AssertKeys(client2.Directory, "key3", "key1", "key4");
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.iteration.spec.ts — "should have eventually consistent iteration order with nested subdirectory operations"
		public void ShouldHaveEventuallyConsistentIterationOrderWithNestedSubdirectoryOperations()
		{
			ConnectedDirectory client1 = CreateConnectedDirectory("shared-directory-1", "client-1");
			ConnectedDirectory client2 = CreateConnectedDirectory("shared-directory-2", "client-2");
			var clients = new[] { client1, client2 };

			client1.Directory.Set("rootKey1", "rootValue1");
			IDirectory subDir1 = client1.Directory.CreateSubDirectory("subdir");
			subDir1.Set("subKey1", "subValue1");
			client1.Directory.Set("rootKey2", "rootValue2");
			ProcessSent(clients, client1, client1.Sender.Sent[0], seq: 1, refSeq: 0);
			ProcessSent(clients, client1, client1.Sender.Sent[1], seq: 2, refSeq: 1);
			ProcessSent(clients, client1, client1.Sender.Sent[2], seq: 3, refSeq: 2);
			ProcessSent(clients, client1, client1.Sender.Sent[3], seq: 4, refSeq: 3);

			IDirectory subDir2 = Assert.IsAssignableFrom<IDirectory>(client2.Directory.GetSubDirectory("subdir"));
			client1.Directory.Set("rootKey3", "rootValue3");
			subDir1.Set("subKey2", "subValue2");
			client2.Directory.Set("rootKey4", "rootValue4");
			subDir2.Set("subKey3", "subValue3");
			client1.Directory.Delete("rootKey1");
			subDir1.Delete("subKey1");
			client2.Directory.Set("rootKey1", "newRootValue1");

			ProcessSent(clients, client1, client1.Sender.Sent[4], seq: 5, refSeq: 4);
			ProcessSent(clients, client1, client1.Sender.Sent[5], seq: 6, refSeq: 4);
			ProcessSent(clients, client2, client2.Sender.Sent[0], seq: 7, refSeq: 4);
			ProcessSent(clients, client2, client2.Sender.Sent[1], seq: 8, refSeq: 4);
			ProcessSent(clients, client1, client1.Sender.Sent[6], seq: 9, refSeq: 4);
			ProcessSent(clients, client1, client1.Sender.Sent[7], seq: 10, refSeq: 4);
			ProcessSent(clients, client2, client2.Sender.Sent[2], seq: 11, refSeq: 4);

			// TS-parity deviation: our port returns a point-in-time list backed by Dictionary slot order, not TS Map lifetime order after delete/re-add.
			AssertKeys(client1.Directory, "rootKey1", "rootKey2", "rootKey3", "rootKey4");
			AssertKeys(client2.Directory, "rootKey1", "rootKey2", "rootKey3", "rootKey4");
			AssertKeys(subDir1, "subKey2", "subKey3");
			AssertKeys(subDir2, "subKey2", "subKey3");
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.iteration.spec.ts — "should have eventually consistent subdirectory iteration order with multiple create/delete"
		public void ShouldHaveEventuallyConsistentSubdirectoryIterationOrderWithMultipleCreateDelete()
		{
			ConnectedDirectory client1 = CreateConnectedDirectory("shared-directory-1", "client-1");
			ConnectedDirectory client2 = CreateConnectedDirectory("shared-directory-2", "client-2");
			var clients = new[] { client1, client2 };

			client1.Directory.CreateSubDirectory("dir1");
			client1.Directory.CreateSubDirectory("dir2");
			client1.Directory.CreateSubDirectory("dir3");
			ProcessSent(clients, client1, client1.Sender.Sent[0], seq: 1, refSeq: 0);
			ProcessSent(clients, client1, client1.Sender.Sent[1], seq: 2, refSeq: 1);
			ProcessSent(clients, client1, client1.Sender.Sent[2], seq: 3, refSeq: 2);

			AssertSubdirectories(client1.Directory, "dir1", "dir2", "dir3");
			AssertSubdirectories(client2.Directory, "dir1", "dir2", "dir3");

			client1.Directory.DeleteSubDirectory("dir2");
			client2.Directory.CreateSubDirectory("dir4");
			ProcessSent(clients, client1, client1.Sender.Sent[3], seq: 4, refSeq: 3);
			ProcessSent(clients, client2, client2.Sender.Sent[0], seq: 5, refSeq: 3);

			client2.Directory.DeleteSubDirectory("dir1");
			client2.Directory.CreateSubDirectory("dir2");
			client1.Directory.CreateSubDirectory("dir5");
			ProcessSent(clients, client1, client1.Sender.Sent[4], seq: 6, refSeq: 5);
			ProcessSent(clients, client2, client2.Sender.Sent[1], seq: 7, refSeq: 5);
			ProcessSent(clients, client2, client2.Sender.Sent[2], seq: 8, refSeq: 5);

			// TS-parity deviation: our port returns a point-in-time list, not a live iterator.
			AssertSubdirectories(client1.Directory, "dir3", "dir4", "dir5", "dir2");
			AssertSubdirectories(client2.Directory, "dir3", "dir4", "dir5", "dir2");
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

		private static void AssertKeys(IDirectory directory, params string[] expectedKeys)
		{
			Assert.Equal(expectedKeys, directory.Keys.ToArray());
		}

		private static void AssertSubdirectories(IDirectory directory, params string[] expectedSubdirectories)
		{
			Assert.Equal(expectedSubdirectories, directory.SubDirectories().Select(kvp => kvp.Key).ToArray());
		}

		private sealed record ConnectedDirectory(SharedDirectory Directory, FakeFluidDataObjectSender Sender, string ClientId);
	}
}
