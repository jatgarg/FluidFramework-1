// -----------------------------------------------------------------------------
// Snapshot fallback regression tests for tolerant TS load semantics.
// -----------------------------------------------------------------------------

#nullable enable

using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SnapshotFallbackTests
	{
		[Fact]
		public void Load_MissingMinSequenceNumber_FallsBackToSequenceNumber()
		{
			string snapshotJson = CreateV1Snapshot(
				"[\"A\"]",
				segmentCount: 1,
				length: 1,
				minSequenceNumber: null,
				sequenceNumber: 42);

			Client client = LoadIntoClient(snapshotJson);

			Assert.Equal("A", client.GetText(0, client.GetLength()));
			Assert.Equal(42, client.CollabWindowMinSeq);
			Assert.Equal(42, client.CollabWindowCurrentSeq);
		}

		[Fact]
		public void Load_MissingMinSequenceNumber_MergeTreeMinSeqSet()
		{
			string snapshotJson = CreateV1Snapshot(
				"[\"AB\"]",
				segmentCount: 1,
				length: 2,
				minSequenceNumber: null,
				sequenceNumber: 42);

			Client client = LoadIntoClient(snapshotJson);

			Assert.Equal(42, client.MergeTree.MinSeq);
			Assert.Equal(42, client.MergeTree.CurrentSeq);
			Assert.Equal("AB", client.GetText(0, client.GetLength()));
		}

		[Fact]
		public void Load_LegacyMissingChunkMinSequenceNumber_FallsBackToChunkSequenceNumber()
		{
			const string snapshotJson = """
				{
					"chunkStartSegmentIndex":0,
					"chunkSegmentCount":1,
					"chunkLengthChars":1,
					"totalLengthChars":1,
					"totalSegmentCount":1,
					"chunkSequenceNumber":42,
					"segmentTexts":["A"]
				}
				""";

			Client client = LoadIntoClient(snapshotJson);

			Assert.Equal(42, client.CollabWindowMinSeq);
			Assert.Equal(42, client.MergeTree.MinSeq);
			Assert.Equal("A", client.GetText(0, client.GetLength()));
		}

		[Fact]
		public void Load_AllOptionalFieldsPresent_UsesActualValues()
		{
			string snapshotJson = CreateV1Snapshot(
				"[\"A\"]",
				segmentCount: 1,
				length: 1,
				minSequenceNumber: 7,
				sequenceNumber: 42);

			Client client = LoadIntoClient(snapshotJson);

			Assert.Equal(7, client.CollabWindowMinSeq);
			Assert.Equal(7, client.MergeTree.MinSeq);
			Assert.Equal(42, client.CollabWindowCurrentSeq);
			Assert.Equal(42, client.MergeTree.CurrentSeq);
		}

		[Fact]
		public void Load_MissingSegmentSeq_UsesUniversalSequenceNumber()
		{
			string snapshotJson = CreateV1Snapshot(
				"[{\"json\":\"A\",\"client\":\"seed\"}]",
				segmentCount: 1,
				length: 1,
				minSequenceNumber: 0,
				sequenceNumber: 42);

			Client client = LoadIntoClient(snapshotJson);

			ISegment segment = Assert.Single(client.MergeTree.WalkAllSegments());
			Assert.Equal(Constants.UniversalSequenceNumber, segment.Seq);
			Assert.NotEqual(Constants.NonCollabClient, segment.ClientId);
		}

		[Fact]
		public void Load_LegacyRemovedClient_FillsRemovedClientIds()
		{
			string snapshotJson = CreateV1Snapshot(
				"[{\"json\":\"x\",\"removedSeq\":10,\"removedClient\":\"remove-client\"}]",
				segmentCount: 1,
				length: 1,
				minSequenceNumber: 0,
				sequenceNumber: 20);

			Client client = LoadIntoClient(snapshotJson);

			ISegment segment = Assert.Single(client.MergeTree.WalkAllSegments());
			RemoveOperationStamp stamp = Assert.Single(segment.RemoveStamps);
			Assert.IsType<SetRemoveOperationStamp>(stamp);
			Assert.Equal(10, stamp.Seq);
		}

		[Fact]
		public void Load_MissingOrderedChunkMetadata_Throws()
		{
			// TS-parity: MergeTreeHeaderMetadata.orderedChunkMetadata is required
			// (packages/dds/merge-tree/src/snapshotChunks.ts). The writer always emits
			// at least [{id: "header"}], so absence indicates a malformed snapshot.
			string snapshotJson = CreateV1Snapshot(
				"[\"A\"]",
				segmentCount: 1,
				length: 1,
				minSequenceNumber: 1,
				sequenceNumber: 1,
				orderedChunkMetadataJson: null);

			LoggingError exception = Assert.Throws<LoggingError>(() => LoadIntoClient(snapshotJson));

			Assert.Contains("orderedChunkMetadata", exception.Message);
		}

		private static Client LoadIntoClient(string snapshotJson)
		{
			Client client = new("loader");
			SharedStringSnapshotLoader.PopulateFromSnapshot(client, SharedStringSnapshotLoader.Load(snapshotJson));
			return client;
		}

		private static string CreateV1Snapshot(
			string segmentsJson,
			int segmentCount,
			int length,
			long? minSequenceNumber = 0,
			long sequenceNumber = 1,
			string? orderedChunkMetadataJson = "[{\"id\":\"header\"}]")
		{
			string minSequenceNumberJson = minSequenceNumber is null
				? string.Empty
				: ",\"minSequenceNumber\":" + minSequenceNumber.Value;
			string orderedChunkMetadata = orderedChunkMetadataJson is null
				? string.Empty
				: ",\"orderedChunkMetadata\":" + orderedChunkMetadataJson;
			return "{\"version\":\"1\",\"segmentCount\":"
				+ segmentCount
				+ ",\"length\":"
				+ length
				+ ",\"segments\":"
				+ segmentsJson
				+ ",\"startIndex\":0,\"headerMetadata\":{\"sequenceNumber\":"
				+ sequenceNumber
				+ minSequenceNumberJson
				+ orderedChunkMetadata
				+ ",\"totalLength\":"
				+ length
				+ ",\"totalSegmentCount\":"
				+ segmentCount
				+ "}}";
		}
	}
}
