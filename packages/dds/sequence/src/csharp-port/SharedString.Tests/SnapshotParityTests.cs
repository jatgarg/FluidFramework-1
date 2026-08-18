// -----------------------------------------------------------------------------
// Snapshot parity regression tests for S-series audit findings.
// -----------------------------------------------------------------------------

#nullable enable

using System.Linq;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SnapshotParityTests
	{
		[Fact]
		public void Load_MultiStampArrays_PreservesSetRemoveAndSliceRemoveSequences()
		{
			string snapshotJson = CreateV1Snapshot(
				"""
				[
					{
						"json":"x",
						"seq":1,
						"client":"seed",
						"removedSeq":10,
						"removedSeqs":[10,11],
						"removedClientIds":["remove-a","remove-b"],
						"movedSeq":12,
						"movedSeqs":[12,13],
						"movedClientIds":["move-a","move-b"]
					}
				]
				""",
				segmentCount: 1,
				length: 1,
				minSequenceNumber: 0,
				sequenceNumber: 20);
			Client client = new("loader");

			SharedStringSnapshotLoader.PopulateFromSnapshot(client, SharedStringSnapshotLoader.Load(snapshotJson));

			ISegment segment = Assert.Single(client.MergeTree.WalkAllSegments());
			// TS-parity: V1 merge-info arrays preserve every loaded tombstone stamp (Finding S1).
			Assert.Equal(new[] { "setRemove", "setRemove", "sliceRemove", "sliceRemove" }, segment.RemoveStamps.Select(stamp => stamp.Type));
			Assert.Equal(new long[] { 10, 11, 12, 13 }, segment.RemoveStamps.Select(stamp => stamp.Seq));
			Assert.Equal(string.Empty, client.GetText(0, client.GetLength()));
		}

		[Fact]
		public void Load_CatchupOps_SnapshotSequenceEqualsCatchupSequence_Accepted()
		{
			// TS ref: packages/dds/sequence/src/sequence.ts (loadCatchupOps).
			// TS validates `sequenceNumber < collabWindow.currentSeq` (strict). A
			// catchup message whose sequence number equals the snapshot's own
			// sequence number is legal — it represents the last message of the
			// grouped batch captured by the snapshot being replayed. The port
			// previously used `<=` and rejected this valid case.
			string snapshotJson = CreateV1Snapshot(
				"[\"A\"]",
				segmentCount: 1,
				length: 1,
				minSequenceNumber: 0,
				sequenceNumber: 2,
				catchupOpsBlobNamesJson: "[\"catchup\"]");
			const string catchupOpsJson = "[{\"sequenceNumber\":2,\"referenceSequenceNumber\":2,\"minimumSequenceNumber\":0,\"clientId\":\"remote\",\"contents\":{\"type\":0,\"pos1\":1,\"seg\":\"!\"}}]";
			SharedString sharedString = new();

			sharedString.LoadFromSnapshot(snapshotJson, _ => catchupOpsJson);

			Assert.Equal("A!", sharedString.GetText());
		}

		[Fact]
		public void Load_LegacyHeaderChunk_HydratesText()
		{
			const string snapshotJson = """
				{
					"chunkStartSegmentIndex":0,
					"chunkSegmentCount":1,
					"chunkLengthChars":5,
					"totalLengthChars":5,
					"totalSegmentCount":1,
					"chunkSequenceNumber":1,
					"chunkMinSequenceNumber":1,
					"segmentTexts":["Hello"]
				}
				""";
			SharedString sharedString = new();

			sharedString.LoadFromSnapshot(snapshotJson);

			// TS-parity: undefined-version legacy chunks are upgraded to V1 at load (Finding S4).
			Assert.Equal("Hello", sharedString.GetText());
			Assert.Equal(5, sharedString.GetLength());
		}

		[Fact]
		public void Load_LegacyHeaderAndBodyChunks_HydratesText()
		{
			const string snapshotJson = """
				{
					"chunkStartSegmentIndex":0,
					"chunkSegmentCount":1,
					"chunkLengthChars":1,
					"totalLengthChars":2,
					"totalSegmentCount":2,
					"chunkSequenceNumber":1,
					"chunkMinSequenceNumber":1,
					"segmentTexts":["A"]
				}
				""";
			const string bodyJson = """
				{
					"chunkStartSegmentIndex":1,
					"chunkSegmentCount":1,
					"chunkLengthChars":1,
					"segmentTexts":["B"]
				}
				""";
			SharedString sharedString = new();

			sharedString.LoadFromSnapshot(snapshotJson, blobName =>
			{
				Assert.Equal("body", blobName);
				return bodyJson;
			});

			// TS-parity: legacy headers name the overflow chunk "body" (Finding S4).
			Assert.Equal("AB", sharedString.GetText());
			Assert.Equal(2, sharedString.GetLength());
		}

		[Fact]
		public void Load_MarkerSegment_CountsMarkerLengthAndIndexesMarker()
		{
			string snapshotJson = CreateV1Snapshot(
				"""
				[
					"A",
					{"marker":{"refType":1},"props":{"markerId":"m1","referenceTileLabels":["tile"]}},
					"B"
				]
				""",
				segmentCount: 3,
				length: 3);
			SharedString sharedString = new();

			sharedString.LoadFromSnapshot(snapshotJson);

			Marker? marker = sharedString.GetMarkerFromId("m1");
			// TS-parity: snapshot length is sequence length, not rendered text length (Finding S5).
			Assert.NotNull(marker);
			Assert.Equal("AB", sharedString.GetText());
			Assert.Equal(3, sharedString.GetLength());
			Assert.Equal(1, sharedString.GetPositionOfMarker(marker!));
			Assert.True(marker!.HasTileLabel("tile"));
		}

		[Fact]
		public void Load_CatchupInsertAtRemovedLoadedTombstone_UsesSnapshotInsertionMetadata()
		{
			string snapshotJson = CreateV1Snapshot(
				"""
				[
					"0",
					{"json":"x","removedSeq":2,"removedClientIds":["remove-client"]},
					"2"
				]
				""",
				segmentCount: 3,
				length: 3,
				minSequenceNumber: 1,
				sequenceNumber: 3,
				catchupOpsBlobNamesJson: "[\"catchup\"]");
			const string catchupOpsJson = "[{\"sequenceNumber\":4,\"referenceSequenceNumber\":1,\"minimumSequenceNumber\":1,\"clientId\":\"insert-client\",\"contents\":{\"type\":0,\"pos1\":1,\"seg\":\"1\"}}]";
			SharedString sharedString = new();

			sharedString.LoadFromSnapshot(snapshotJson, _ => catchupOpsJson);

			// TS-parity: below-MSN loaded tombstones remain visible to older refSeq catchup ops (Finding S6).
			Assert.Equal("012", sharedString.GetText());
			Assert.Equal(3, sharedString.GetLength());
		}

		[Fact]
		public void Load_CatchupInsertAtObliteratedLoadedTombstone_UsesSnapshotInsertionMetadata()
		{
			string snapshotJson = CreateV1Snapshot(
				"""
				[
					"0",
					{"json":"x","movedSeq":2,"movedSeqs":[2],"movedClientIds":["obliterate-client"]},
					"2"
				]
				""",
				segmentCount: 3,
				length: 3,
				minSequenceNumber: 1,
				sequenceNumber: 3,
				catchupOpsBlobNamesJson: "[\"catchup\"]");
			const string catchupOpsJson = "[{\"sequenceNumber\":4,\"referenceSequenceNumber\":1,\"minimumSequenceNumber\":1,\"clientId\":\"insert-client\",\"contents\":{\"type\":0,\"pos1\":1,\"seg\":\"1\"}}]";
			SharedString sharedString = new();

			sharedString.LoadFromSnapshot(snapshotJson, _ => catchupOpsJson);

			// TS-parity: obliterated loaded tombstones preserve reference-sequence placement (Finding S6).
			Assert.Equal("012", sharedString.GetText());
			Assert.Equal(3, sharedString.GetLength());
		}

		[Fact]
		public void Load_UnknownZeroLengthSegment_IsIgnored()
		{
			string snapshotJson = CreateV1Snapshot(
				"[{}]",
				segmentCount: 1,
				length: 0,
				minSequenceNumber: 0,
				sequenceNumber: 1);
			SharedString sharedString = new();

			sharedString.LoadFromSnapshot(snapshotJson);

			// TS-parity: unrecognized segment specs deserialize to no segment instead of crashing (Finding S8).
			Assert.Equal(string.Empty, sharedString.GetText());
			Assert.Equal(0, sharedString.GetLength());
		}

		private static string CreateV1Snapshot(
			string segmentsJson,
			int segmentCount,
			int length,
			long minSequenceNumber = 0,
			long sequenceNumber = 1,
			string? catchupOpsBlobNamesJson = null)
		{
			string catchupOpsBlobNames = catchupOpsBlobNamesJson is null
				? string.Empty
				: $",\"catchupOpsBlobNames\":{catchupOpsBlobNamesJson}";
			return "{\"version\":\"1\",\"segmentCount\":"
				+ segmentCount
				+ ",\"length\":"
				+ length
				+ ",\"segments\":"
				+ segmentsJson
				+ ",\"startIndex\":0,\"headerMetadata\":{\"minSequenceNumber\":"
				+ minSequenceNumber
				+ ",\"sequenceNumber\":"
				+ sequenceNumber
				+ ",\"orderedChunkMetadata\":[{\"id\":\"header\"}]"
				+ catchupOpsBlobNames
				+ ",\"totalLength\":"
				+ length
				+ ",\"totalSegmentCount\":"
				+ segmentCount
				+ "}}";
		}
	}
}
