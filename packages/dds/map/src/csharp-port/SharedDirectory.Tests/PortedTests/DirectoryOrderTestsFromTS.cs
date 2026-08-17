// -----------------------------------------------------------------------------
// High-value SharedDirectory ordering tests ported from TS directory.order.spec.ts.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class DirectoryOrderTestsFromTS
	{
		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.order.spec.ts — "create subdirectories"
		public void CreateSubdirectories()
		{
			var directory = new SharedDirectory();

			directory.CreateSubDirectory("b");
			directory.CreateSubDirectory("a");
			directory.CreateSubDirectory("c");

			AssertDirectoryIterationOrder(directory, "b", "a", "c");
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.order.spec.ts — "create nested subdirectories"
		public void CreateNestedSubdirectories()
		{
			var directory = new SharedDirectory();

			directory.CreateSubDirectory("c").CreateSubDirectory("c-a");
			directory.CreateSubDirectory("a").CreateSubDirectory("a-b");
			directory.CreateSubDirectory("a").CreateSubDirectory("a-a");
			directory.CreateSubDirectory("b");
			directory.CreateSubDirectory("c").CreateSubDirectory("c-c");
			directory.CreateSubDirectory("a").CreateSubDirectory("a-c");
			directory.CreateSubDirectory("c").CreateSubDirectory("c-b");

			AssertDirectoryIterationOrder(directory, "c", "a", "b");
			AssertDirectoryIterationOrder(directory.GetWorkingDirectory("/a")!, "a-b", "a-a", "a-c");
			AssertDirectoryIterationOrder(directory.GetWorkingDirectory("/b")!);
			AssertDirectoryIterationOrder(directory.GetWorkingDirectory("/c")!, "c-a", "c-c", "c-b");
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.order.spec.ts — "delete subdirectories"
		public void DeleteSubdirectories()
		{
			var directory = new SharedDirectory();

			directory.CreateSubDirectory("c").CreateSubDirectory("c-a");
			directory.CreateSubDirectory("a").CreateSubDirectory("a-b");
			directory.CreateSubDirectory("a").CreateSubDirectory("a-a");
			directory.CreateSubDirectory("b");
			directory.CreateSubDirectory("c").CreateSubDirectory("c-c");
			directory.CreateSubDirectory("a").CreateSubDirectory("a-c");
			directory.CreateSubDirectory("c").CreateSubDirectory("c-b");

			directory.DeleteSubDirectory("a");
			AssertDirectoryIterationOrder(directory, "c", "b");

			directory.CreateSubDirectory("a");
			AssertDirectoryIterationOrder(directory, "c", "b", "a");

			directory.GetWorkingDirectory("/c")!.CreateSubDirectory("c-d");
			directory.GetWorkingDirectory("/c")!.DeleteSubDirectory("c-c");

			AssertDirectoryIterationOrder(directory.GetWorkingDirectory("/c")!, "c-a", "c-b", "c-d");
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.order.spec.ts — "Remote messages have conflicts with the local pending ops"
		public void RemoteMessagesHaveConflictsWithTheLocalPendingOps()
		{
			ConnectedDirectory directory1 = CreateConnectedDirectory("directory1", "client-1");
			ConnectedDirectory directory2 = CreateConnectedDirectory("directory2", "client-2");
			var clients = new[] { directory1, directory2 };

			directory1.Directory.CreateSubDirectory("b");
			directory2.Directory.CreateSubDirectory("a");
			directory2.Directory.CreateSubDirectory("b");

			ProcessSent(clients, directory1, directory1.Sender.Sent[0], seq: 1, refSeq: 0);
			ProcessSent(clients, directory2, directory2.Sender.Sent[0], seq: 2, refSeq: 0);
			ProcessSent(clients, directory2, directory2.Sender.Sent[1], seq: 3, refSeq: 0);
			AssertDirectoryIterationOrder(directory1.Directory, "b", "a");
			AssertDirectoryIterationOrder(directory2.Directory, "b", "a");

			directory1.Directory.CreateSubDirectory("d").CreateSubDirectory("d-b");
			directory2.Directory.CreateSubDirectory("d").CreateSubDirectory("d-a");

			ProcessSent(clients, directory1, directory1.Sender.Sent[1], seq: 4, refSeq: 3);
			ProcessSent(clients, directory1, directory1.Sender.Sent[2], seq: 5, refSeq: 4);
			ProcessSent(clients, directory2, directory2.Sender.Sent[2], seq: 6, refSeq: 3);
			ProcessSent(clients, directory2, directory2.Sender.Sent[3], seq: 7, refSeq: 6);
			AssertDirectoryIterationOrder(directory1.Directory, "b", "a", "d");
			AssertDirectoryIterationOrder(directory2.Directory, "b", "a", "d");
			AssertDirectoryIterationOrder(directory1.Directory.GetWorkingDirectory("/d")!, "d-b", "d-a");
			AssertDirectoryIterationOrder(directory2.Directory.GetWorkingDirectory("/d")!, "d-b", "d-a");

			directory1.Directory.DeleteSubDirectory("d");
			directory2.Directory.GetWorkingDirectory("/d")!.CreateSubDirectory("d-c");
			AssertDirectoryIterationOrder(directory1.Directory, "b", "a");
			AssertDirectoryIterationOrder(directory2.Directory, "b", "a", "d");
			AssertDirectoryIterationOrder(directory2.Directory.GetWorkingDirectory("/d")!, "d-b", "d-a", "d-c");

			ProcessSent(clients, directory1, directory1.Sender.Sent[3], seq: 8, refSeq: 7);
			ProcessSent(clients, directory2, directory2.Sender.Sent[4], seq: 9, refSeq: 7);
			AssertDirectoryIterationOrder(directory1.Directory, "b", "a");
			AssertDirectoryIterationOrder(directory2.Directory, "b", "a");
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.order.spec.ts — "can be compatible with the old format summary"
		public void CanBeCompatibleWithTheOldFormatSummary()
		{
			const string json = "{\"storage\":{\"key1\":{\"type\":\"Plain\",\"value\":\"val1\"},\"key2\":{\"type\":\"Plain\",\"value\":\"val2\"}},\"subdirectories\":{\"b\":{\"storage\":{\"testKey\":{\"type\":\"Plain\",\"value\":\"testValue\"},\"testKey2\":{\"type\":\"Plain\",\"value\":\"testValue2\"}}},\"c\":{\"storage\":{\"testKey3\":{\"type\":\"Plain\",\"value\":\"testValue3\"}}},\"a\":{\"storage\":{}}}}";
			var directory = new SharedDirectory();

			directory.LoadFromSnapshot(json);

			AssertDirectoryIterationOrder(directory, "b", "c", "a");
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.order.spec.ts — "can be compatible with the new format summary"
		public void CanBeCompatibleWithTheNewFormatSummary()
		{
			const string json = "{\"storage\":{\"key1\":{\"type\":\"Plain\",\"value\":\"val1\"},\"key2\":{\"type\":\"Plain\",\"value\":\"val2\"}},\"subdirectories\":{\"b\":{\"storage\":{\"testKey\":{\"type\":\"Plain\",\"value\":\"testValue\"},\"testKey2\":{\"type\":\"Plain\",\"value\":\"testValue2\"}},\"ci\":{\"csn\":4,\"ccIds\":[\"client1\"]}},\"c\":{\"storage\":{\"testKey3\":{\"type\":\"Plain\",\"value\":\"testValue3\"}},\"subdirectories\":{\"c_b\":{\"storage\":{},\"ci\":{\"csn\":5,\"ccIds\":[\"client1\"]}},\"c_a\":{\"storage\":{},\"ci\":{\"csn\":6,\"ccIds\":[\"client1\"]}}},\"ci\":{\"csn\":2,\"ccIds\":[\"client2\"],\"ccsn\":1}},\"a\":{\"storage\":{},\"ci\":{\"csn\":4,\"ccIds\":[\"client2\"]}}},\"ci\":{\"csn\":1,\"ccIds\":[\"client1\",\"client2\"]}}";
			var directory = new SharedDirectory();

			directory.LoadFromSnapshot(json);

			AssertDirectoryIterationOrder(directory, "c", "b", "a");
			AssertDirectoryIterationOrder(directory.GetWorkingDirectory("/c")!, "c_b", "c_a");
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.order.spec.ts — "serialize the contents, load it into another directory and maintain the order"
		public void SerializeTheContentsLoadItIntoAnotherDirectoryAndMaintainTheOrderLoadOnlyEquivalent()
		{
			const string json = "{\"subdirectories\":{\"c\":{\"ci\":{\"csn\":0,\"ccIds\":[]}},\"b\":{\"ci\":{\"csn\":0,\"ccIds\":[]}},\"a\":{\"ci\":{\"csn\":0,\"ccIds\":[]}}}}";
			var directory = new SharedDirectory();

			// TS-parity deviation: our port does not expose a serialize path, so this loads the equivalent serialized content shape.
			directory.LoadFromSnapshot(json);

			AssertDirectoryIterationOrder(directory, "c", "b", "a");
		}

		[Fact]
		// Ported from packages/dds/map/src/test/mocha/directory.order.spec.ts — "can be compatible with the detached scenario"
		public void CanBeCompatibleWithTheDetachedScenario()
		{
			const string json = "{\"blobs\":[],\"content\":{\"ci\":{\"csn\":0,\"ccIds\":[]},\"subdirectories\":{\"detached1\":{\"ci\":{\"csn\":0,\"ccIds\":[\"97cd0b77-34b1-46a8-bbe2-5fbefb3e014b\"]}},\"detached2\":{\"ci\":{\"csn\":0,\"ccIds\":[\"97cd0b77-34b1-46a8-bbe2-5fbefb3e014b\"]}},\"detached3\":{\"ci\":{\"csn\":-1,\"ccIds\":[\"97cd0b77-34b1-46a8-bbe2-5fbefb3e014b\"]}}}}}";
			var directory = new SharedDirectory();

			directory.LoadFromSnapshot(json, blobName => throw new KeyNotFoundException(blobName));

			AssertDirectoryIterationOrder(directory, "detached1", "detached2", "detached3");
			Assert.NotNull(directory.GetSubDirectory("detached3"));
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

		private static void AssertDirectoryIterationOrder(IDirectory directory, params string[] expectedDirNames)
		{
			Assert.Equal(expectedDirNames, directory.SubDirectories().Select(kvp => kvp.Key).ToArray());
		}

		private sealed record ConnectedDirectory(SharedDirectory Directory, FakeFluidDataObjectSender Sender, string ClientId);
	}
}
