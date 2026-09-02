// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class SharedStringSnapshotTests
	{
		[Fact]
		public void LoadFromSnapshot_SimpleHello_HydratesCorrectText()
		{
			SharedStringSnapshotDto dto = LoadFixture("simple-hello.snapshot.bin");
			var sharedString = new SharedString();

			sharedString.LoadFromSnapshot(dto);

			Assert.Equal("Hello, world!", sharedString.GetText());
			Assert.Equal(13, sharedString.GetLength());
		}

		[Fact]
		public void LoadFromSnapshot_TwoInsertions_HydratesFinalText()
		{
			SharedStringSnapshotDto dto = LoadFixture("two-insertions.snapshot.bin");
			var sharedString = new SharedString();

			sharedString.LoadFromSnapshot(dto);

			Assert.Equal("Hello beautiful world!", sharedString.GetText());
			Assert.Equal(22, sharedString.GetLength());
		}

		[Fact]
		public void LoadFromSnapshot_InsertThenDelete_HydratesFinalText()
		{
			SharedStringSnapshotDto dto = LoadFixture("insert-then-delete.snapshot.bin");
			var sharedString = new SharedString();

			sharedString.LoadFromSnapshot(dto);

			Assert.Equal("Hell world!", sharedString.GetText());
			Assert.Equal(11, sharedString.GetLength());
		}

		[Fact]
		public void LoadFromSnapshot_FiresNoEvents()
		{
			SharedStringSnapshotDto dto = LoadFixture("simple-hello.snapshot.bin");
			var sharedString = new SharedString();
			int eventCount = 0;
			sharedString.OnSequenceDelta += (s, e) => eventCount++;

			sharedString.LoadFromSnapshot(dto);

			Assert.Equal(0, eventCount);
		}

		[Fact]
		public void LoadFromSnapshot_OnNonEmptyString_Throws()
		{
			SharedStringSnapshotDto dto = LoadFixture("simple-hello.snapshot.bin");
			var sharedString = new SharedString();
			sharedString.InsertText(0, "existing");

			LoggingError exception = Assert.Throws<LoggingError>(() => sharedString.LoadFromSnapshot(dto));

		}

		[Fact]
		public void Load_SingleChunk_StillWorks()
		{
			var sharedString = new SharedString();
			string snapshotJson = CreateHeaderSnapshotJson(
				segmentsJson: "[\"Hello\"]",
				segmentCount: 1,
				length: 5,
				orderedChunkMetadataJson: "[{\"id\":\"header\"}]",
				totalSegmentCount: 1,
				totalLength: 5);

			sharedString.LoadFromSnapshot(snapshotJson);

			Assert.Equal("Hello", sharedString.GetText());
			Assert.Equal(5, sharedString.GetLength());
		}

		[Fact]
		public void Load_DetectsMultiChunk()
		{
			var sharedString = new SharedString();
			var resolved = new HashSet<string>(StringComparer.Ordinal);
			string snapshotJson = CreateHeaderSnapshotJson(
				segmentsJson: "[\"Hello\"]",
				segmentCount: 1,
				length: 5,
				orderedChunkMetadataJson: "[{\"id\":\"header\"},{\"id\":\"body_0\"}]",
				totalSegmentCount: 2,
				totalLength: 11);

			sharedString.LoadFromSnapshot(snapshotJson, blobName =>
			{
				resolved.Add(blobName);
				return CreateChunkJson("[\" world\"]", segmentCount: 1, length: 6, startIndex: 1);
			});

			Assert.Contains("body_0", resolved);
			Assert.Equal("Hello world", sharedString.GetText());
		}

		[Fact]
		public void Load_MergesChunksInHeaderOrder()
		{
			var sharedString = new SharedString();
			var blobs = new Dictionary<string, string>(StringComparer.Ordinal)
			{
				["chunk-one"] = CreateChunkJson("[\"C\"]", segmentCount: 1, length: 1, startIndex: 2),
				["chunk-two"] = CreateChunkJson("[\"B\"]", segmentCount: 1, length: 1, startIndex: 1),
			};
			string snapshotJson = CreateHeaderSnapshotJson(
				segmentsJson: "[\"A\"]",
				segmentCount: 1,
				length: 1,
				orderedChunkMetadataJson: "[{\"id\":\"header\"},{\"id\":\"chunk-two\"},{\"id\":\"chunk-one\"}]",
				totalSegmentCount: 3,
				totalLength: 3);

			sharedString.LoadFromSnapshot(snapshotJson, blobName => blobs[blobName]);

			Assert.Equal("ABC", sharedString.GetText());
		}

		[Fact]
		public void Load_NullResolver_MultiChunk_Throws()
		{
			var sharedString = new SharedString();
			string snapshotJson = CreateHeaderSnapshotJson(
				segmentsJson: "[\"A\"]",
				segmentCount: 1,
				length: 1,
				orderedChunkMetadataJson: "[{\"id\":\"header\"},{\"id\":\"body_0\"}]",
				totalSegmentCount: 2,
				totalLength: 2);

			LoggingError exception = Assert.Throws<LoggingError>(() => sharedString.LoadFromSnapshot(snapshotJson));

			Assert.Contains("blobResolver", exception.Message);
		}

		[Fact]
		public void Load_CatchupOps_ReplayedAfterBase()
		{
			var sharedString = new SharedString();
			string snapshotJson = CreateHeaderSnapshotJson(
				segmentsJson: "[\"Hello\"]",
				segmentCount: 1,
				length: 5,
				orderedChunkMetadataJson: "[{\"id\":\"header\"}]",
				totalSegmentCount: 1,
				totalLength: 5,
				catchupOpsBlobNamesJson: "[\"catchup_0\"]");
			const string catchupOpsJson = "[{\"sequenceNumber\":2,\"referenceSequenceNumber\":1,\"minimumSequenceNumber\":1,\"clientId\":\"remote-client\",\"contents\":{\"type\":0,\"pos1\":5,\"seg\":\"!\"}}]";

			sharedString.LoadFromSnapshot(snapshotJson, blobName => catchupOpsJson);

			Assert.Equal("Hello!", sharedString.GetText());
		}

		[Fact]
		public void Load_TsLegacyCatchupOpsBlob_DiscoveredByWellKnownName()
		{
			// TS ref: merge-tree/src/snapshotlegacy.ts SnapshotLegacy.catchupOps —
			// TS's canonical legacy layout stores catchup ops in an unnamed
			// blob (default name "catchupOps") and TS discovers it via
			// storage.list. The header does NOT carry a catchupOpsBlobNames
			// list. The port has no list API, so it probes the well-known
			// name via the resolver.
			var sharedString = new SharedString();
			string snapshotJson = CreateHeaderSnapshotJson(
				segmentsJson: "[\"Hello\"]",
				segmentCount: 1,
				length: 5,
				orderedChunkMetadataJson: "[{\"id\":\"header\"}]",
				totalSegmentCount: 1,
				totalLength: 5);
			const string catchupOpsJson = "[{\"sequenceNumber\":2,\"referenceSequenceNumber\":1,\"minimumSequenceNumber\":1,\"clientId\":\"remote-client\",\"contents\":{\"type\":0,\"pos1\":5,\"seg\":\"!\"}}]";

			sharedString.LoadFromSnapshot(snapshotJson, blobName => blobName == "catchupOps" ? catchupOpsJson : null!);

			Assert.Equal("Hello!", sharedString.GetText());
		}

		[Fact]
		public void Load_NoCatchupBlob_ResolverReturnsNullForProbe_LoadsBase()
		{
			// The TS-legacy probe must be silent — if the resolver has no
			// "catchupOps" blob, snapshot load must complete with just the
			// header content and not throw.
			var sharedString = new SharedString();
			string snapshotJson = CreateHeaderSnapshotJson(
				segmentsJson: "[\"Hello\"]",
				segmentCount: 1,
				length: 5,
				orderedChunkMetadataJson: "[{\"id\":\"header\"}]",
				totalSegmentCount: 1,
				totalLength: 5);

			sharedString.LoadFromSnapshot(snapshotJson, blobName => null!);

			Assert.Equal("Hello", sharedString.GetText());
		}

		[Fact]
		public void Load_CatchupOps_OrderPreserved()
		{
			var sharedString = new SharedString();
			string snapshotJson = CreateHeaderSnapshotJson(
				segmentsJson: "[\"A\"]",
				segmentCount: 1,
				length: 1,
				orderedChunkMetadataJson: "[{\"id\":\"header\"}]",
				totalSegmentCount: 1,
				totalLength: 1,
				catchupOpsBlobNamesJson: "[\"catchup_0\"]");
			// TS-parity: catchup ops are a serialized ISequencedDocumentMessage[]
			// (see packages/dds/merge-tree/src/snapshotlegacy.ts). Each entry must
			// carry the full sequence numbering + clientId + contents fields.
			const string catchupOpsJson = "[" +
				"{\"sequenceNumber\":2,\"referenceSequenceNumber\":1,\"minimumSequenceNumber\":1,\"clientId\":\"remote-client\",\"contents\":{\"type\":0,\"pos1\":1,\"seg\":\"B\"}}," +
				"{\"sequenceNumber\":3,\"referenceSequenceNumber\":2,\"minimumSequenceNumber\":1,\"clientId\":\"remote-client\",\"contents\":{\"type\":0,\"pos1\":2,\"seg\":\"C\"}}" +
				"]";

			sharedString.LoadFromSnapshot(snapshotJson, blobName => catchupOpsJson);

			Assert.Equal("ABC", sharedString.GetText());
		}

		[Fact]
		public void Load_CatchupOps_GroupedBatchSharesSequenceNumber()
		{
			// TS ref: packages/dds/sequence/src/sequence.ts loadCatchupOps —
			// validation uses `sequenceNumber < collabWindow.currentSeq`
			// (strict), permitting equality for messages in the same grouped
			// batch.
			var sharedString = new SharedString();
			string snapshotJson = CreateHeaderSnapshotJson(
				segmentsJson: "[\"A\"]",
				segmentCount: 1,
				length: 1,
				orderedChunkMetadataJson: "[{\"id\":\"header\"}]",
				totalSegmentCount: 1,
				totalLength: 1,
				catchupOpsBlobNamesJson: "[\"catchup_0\"]");
			// Both catchup ops carry sequenceNumber=2 (same grouped batch). They
			// use pos:0 so both are valid against the shared refSeq=1 perspective.
			const string catchupOpsJson = "[" +
				"{\"sequenceNumber\":2,\"referenceSequenceNumber\":1,\"minimumSequenceNumber\":1,\"clientId\":\"remote-client\",\"contents\":{\"type\":0,\"pos1\":0,\"seg\":\"B\"}}," +
				"{\"sequenceNumber\":2,\"referenceSequenceNumber\":1,\"minimumSequenceNumber\":1,\"clientId\":\"remote-client\",\"contents\":{\"type\":0,\"pos1\":0,\"seg\":\"C\"}}" +
				"]";

			sharedString.LoadFromSnapshot(snapshotJson, blobName => catchupOpsJson);

			// The important assertion: no InvalidSnapshot thrown, and both segments
			// were applied. Merge-tree ordering under identical-seq inserts is a
			// separate concern; here we care only that the sequence check accepted.
			Assert.Equal(3, sharedString.GetLength());
		}

		[Fact]
		public void Load_MissingChunk_ResolverReturnsNull_Throws()
		{
			var sharedString = new SharedString();
			string snapshotJson = CreateHeaderSnapshotJson(
				segmentsJson: "[\"A\"]",
				segmentCount: 1,
				length: 1,
				orderedChunkMetadataJson: "[{\"id\":\"header\"},{\"id\":\"missing_chunk\"}]",
				totalSegmentCount: 2,
				totalLength: 2);

			LoggingError exception = Assert.Throws<LoggingError>(
				() => sharedString.LoadFromSnapshot(snapshotJson, blobName => null!));

			Assert.Contains("missing_chunk", exception.Message);
		}

		// -----------------------------------------------------------------
		// Strict catchup ops regression tests.
		// TS ref: packages/dds/merge-tree/src/snapshotlegacy.ts —
		//   builder.addBlob("catchupOps", JSON.stringify(catchUpMsgs));
		// Each entry is an ISequencedDocumentMessage whose sequence numbering
		// + clientId + contents fields are all required.
		// -----------------------------------------------------------------

		[Fact]
		public void Load_CatchupOps_MissingSequenceNumber_Throws()
		{
			var sharedString = new SharedString();
			string snapshotJson = CreateHeaderSnapshotJson(
				segmentsJson: "[\"A\"]",
				segmentCount: 1,
				length: 1,
				orderedChunkMetadataJson: "[{\"id\":\"header\"}]",
				totalSegmentCount: 1,
				totalLength: 1,
				catchupOpsBlobNamesJson: "[\"catchup\"]");
			const string catchupOpsJson = "[{\"referenceSequenceNumber\":1,\"minimumSequenceNumber\":1,\"clientId\":\"r\",\"contents\":{\"type\":0,\"pos1\":1,\"seg\":\"B\"}}]";

			LoggingError exception = Assert.Throws<LoggingError>(
				() => sharedString.LoadFromSnapshot(snapshotJson, _ => catchupOpsJson));

			Assert.Contains("sequenceNumber", exception.Message);
		}

		[Fact]
		public void Load_CatchupOps_MissingReferenceSequenceNumber_Throws()
		{
			var sharedString = new SharedString();
			string snapshotJson = CreateHeaderSnapshotJson(
				segmentsJson: "[\"A\"]",
				segmentCount: 1,
				length: 1,
				orderedChunkMetadataJson: "[{\"id\":\"header\"}]",
				totalSegmentCount: 1,
				totalLength: 1,
				catchupOpsBlobNamesJson: "[\"catchup\"]");
			const string catchupOpsJson = "[{\"sequenceNumber\":2,\"minimumSequenceNumber\":1,\"clientId\":\"r\",\"contents\":{\"type\":0,\"pos1\":1,\"seg\":\"B\"}}]";

			LoggingError exception = Assert.Throws<LoggingError>(
				() => sharedString.LoadFromSnapshot(snapshotJson, _ => catchupOpsJson));

			Assert.Contains("referenceSequenceNumber", exception.Message);
		}

		[Fact]
		public void Load_CatchupOps_MissingMinimumSequenceNumber_Throws()
		{
			var sharedString = new SharedString();
			string snapshotJson = CreateHeaderSnapshotJson(
				segmentsJson: "[\"A\"]",
				segmentCount: 1,
				length: 1,
				orderedChunkMetadataJson: "[{\"id\":\"header\"}]",
				totalSegmentCount: 1,
				totalLength: 1,
				catchupOpsBlobNamesJson: "[\"catchup\"]");
			const string catchupOpsJson = "[{\"sequenceNumber\":2,\"referenceSequenceNumber\":1,\"clientId\":\"r\",\"contents\":{\"type\":0,\"pos1\":1,\"seg\":\"B\"}}]";

			LoggingError exception = Assert.Throws<LoggingError>(
				() => sharedString.LoadFromSnapshot(snapshotJson, _ => catchupOpsJson));

			Assert.Contains("minimumSequenceNumber", exception.Message);
		}

		[Fact]
		public void Load_CatchupOps_MissingClientId_Throws()
		{
			var sharedString = new SharedString();
			string snapshotJson = CreateHeaderSnapshotJson(
				segmentsJson: "[\"A\"]",
				segmentCount: 1,
				length: 1,
				orderedChunkMetadataJson: "[{\"id\":\"header\"}]",
				totalSegmentCount: 1,
				totalLength: 1,
				catchupOpsBlobNamesJson: "[\"catchup\"]");
			const string catchupOpsJson = "[{\"sequenceNumber\":2,\"referenceSequenceNumber\":1,\"minimumSequenceNumber\":1,\"contents\":{\"type\":0,\"pos1\":1,\"seg\":\"B\"}}]";

			LoggingError exception = Assert.Throws<LoggingError>(
				() => sharedString.LoadFromSnapshot(snapshotJson, _ => catchupOpsJson));

			Assert.Contains("clientId", exception.Message);
		}

		[Fact]
		public void Load_CatchupOps_MissingContents_Throws()
		{
			var sharedString = new SharedString();
			string snapshotJson = CreateHeaderSnapshotJson(
				segmentsJson: "[\"A\"]",
				segmentCount: 1,
				length: 1,
				orderedChunkMetadataJson: "[{\"id\":\"header\"}]",
				totalSegmentCount: 1,
				totalLength: 1,
				catchupOpsBlobNamesJson: "[\"catchup\"]");
			const string catchupOpsJson = "[{\"sequenceNumber\":2,\"referenceSequenceNumber\":1,\"minimumSequenceNumber\":1,\"clientId\":\"r\"}]";

			LoggingError exception = Assert.Throws<LoggingError>(
				() => sharedString.LoadFromSnapshot(snapshotJson, _ => catchupOpsJson));

			Assert.Contains("contents", exception.Message);
		}

		[Fact]
		public void Load_CatchupOps_StringEntry_Rejected()
		{
			var sharedString = new SharedString();
			string snapshotJson = CreateHeaderSnapshotJson(
				segmentsJson: "[\"A\"]",
				segmentCount: 1,
				length: 1,
				orderedChunkMetadataJson: "[{\"id\":\"header\"}]",
				totalSegmentCount: 1,
				totalLength: 1,
				catchupOpsBlobNamesJson: "[\"catchup\"]");
			// TS never emits strings inside the catchup array; each entry must be an
			// ISequencedDocumentMessage object.
			const string catchupOpsJson = "[\"{\\\"type\\\":0,\\\"pos1\\\":1,\\\"seg\\\":\\\"B\\\"}\"]";

			LoggingError exception = Assert.Throws<LoggingError>(
				() => sharedString.LoadFromSnapshot(snapshotJson, _ => catchupOpsJson));

		}

		private static SharedStringSnapshotDto LoadFixture(string fileName)
		{
			string fixturesDir = Path.Combine(AppContext.BaseDirectory, "Fixtures");
			byte[] bytes = File.ReadAllBytes(Path.Combine(fixturesDir, fileName));
			return SharedStringSnapshotLoader.Parse(bytes);
		}

		private static string CreateHeaderSnapshotJson(
			string segmentsJson,
			int segmentCount,
			int length,
			string orderedChunkMetadataJson,
			int totalSegmentCount,
			int totalLength,
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
				+ ",\"startIndex\":0,\"headerMetadata\":{\"minSequenceNumber\":1,\"sequenceNumber\":1,\"orderedChunkMetadata\":"
				+ orderedChunkMetadataJson
				+ catchupOpsBlobNames
				+ ",\"totalLength\":"
				+ totalLength
				+ ",\"totalSegmentCount\":"
				+ totalSegmentCount
				+ "}}";
		}

		private static string CreateChunkJson(string segmentsJson, int segmentCount, int length, int startIndex)
		{
			return "{\"version\":\"1\",\"segmentCount\":"
				+ segmentCount
				+ ",\"length\":"
				+ length
				+ ",\"segments\":"
				+ segmentsJson
				+ ",\"startIndex\":"
				+ startIndex
				+ "}";
		}
	}
}
