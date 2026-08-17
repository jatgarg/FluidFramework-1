// -----------------------------------------------------------------------------
// Regression tests for wire + snapshot fixes bundled into Commit 3.
//
// Covers findings SD-W01, SD-W07, SD-T01, and SD-T03 from the independent
// audit (INDEPENDENT-AUDIT.md).
//
// SD-W01: blob-split snapshot with an empty blobs array must not require a
//         resolver (matching TS Promise.all over an empty array in
//         directory.ts:700-715).
// SD-W07: payloadPending handles must NOT be resolved to a live
//         IFluidDataObject through the registry — the flag would be
//         silently dropped on re-emission. Matches TS RemoteFluidObjectHandle
//         retention of payloadPending in serializer.ts:151-162.
// SD-T01: LoadFromSnapshot into a non-empty directory must throw
//         InvalidState and leave pre-existing content unchanged.
// SD-T03: A local-op ack whose client sequence is the zero/missing
//         sentinel must throw InvalidSequenceNumber with the pending
//         entry left unchanged.
// -----------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Text.Json;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class DirectoryAuditRegressionTests
	{
		// -------------------------------------------------------------
		// SD-W01: empty blobs array does not require a resolver
		// -------------------------------------------------------------

		[Fact]
		public void LoadFromSnapshot_BlobSplitFormat_EmptyBlobsArray_WithoutResolver_Succeeds()
		{
			// TS ref: directory.ts:700-715. Promise.all over an empty blobs array
			// resolves immediately; storage.readBlob is never called. C# previously
			// threw regardless. Fixes SD-W01.
			var dir = new SharedDirectory();
			const string snapshotJson = "{\"blobs\":[],\"content\":{}}";

			// Should NOT throw, even though blobResolver is null.
			dir.LoadFromSnapshot(snapshotJson);

			Assert.Equal(0, dir.Count);
			Assert.Equal(0, dir.CountSubDirectory());
		}

		[Fact]
		public void LoadFromSnapshot_BlobSplitFormat_EmptyBlobs_PopulatesContent()
		{
			// The empty-blobs relaxation must still populate the inline content.
			var dir = new SharedDirectory();
			const string snapshotJson = "{\"blobs\":[],\"content\":{\"storage\":{\"k\":{\"type\":\"Plain\",\"value\":\"v\"}}}}";

			dir.LoadFromSnapshot(snapshotJson);

			Assert.Equal("v", dir.Get("k"));
		}

		[Fact]
		public void LoadFromSnapshot_BlobSplitFormat_NonEmptyBlobs_StillRequiresResolver()
		{
			// Ensure the relaxation didn't accidentally allow non-empty blobs through.
			var dir = new SharedDirectory();
			const string snapshotJson = "{\"blobs\":[\"blob0\"],\"content\":{}}";

			Assert.Throws<OcsException>(() => dir.LoadFromSnapshot(snapshotJson));
		}

		// -------------------------------------------------------------
		// SD-W07: payloadPending handles are not resolved to live objects
		// -------------------------------------------------------------

		[Fact]
		public void RemoteSet_PayloadPendingHandle_ThroughRegistry_KeepsAsSerializedHandle()
		{
			// TS ref: serializer.ts:151-162 constructs RemoteFluidObjectHandle with
			// payloadPending preserved. Previous C# resolved through the registry to a
			// live IFluidDataObject, dropping the flag. Fixes SD-W07.
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("target");
			registry.Register(handle, "/dataObjects/target");
			var dir = new SharedDirectory(registry: registry);

			// Wire op with payloadPending: true. Even though the registry can resolve
			// "/dataObjects/target", the pending flag means the payload is not ready
			// and we must NOT dereference to the live object.
			const string op = "{\"type\":\"set\",\"path\":\"/\",\"key\":\"k\","
				+ "\"value\":{\"type\":\"Plain\","
				+ "\"value\":{\"type\":\"__fluid_handle__\",\"url\":\"/dataObjects/target\",\"payloadPending\":true}}}";
			dir.ProcessDataObjectOp(RemoteMessage(refSeq: 0, seq: 1), op);

			object? stored = dir.Get("k");
			SerializedFluidHandle deferredHandle = Assert.IsType<SerializedFluidHandle>(stored);
			Assert.Equal("/dataObjects/target", deferredHandle.Url);
			Assert.True(deferredHandle.PayloadPending);
		}

		[Fact]
		public void RemoteSet_NonPendingHandle_ThroughRegistry_StillResolvesToLiveObject()
		{
			// Sanity: the SD-W07 fix must not break the happy path — non-pending
			// handles still resolve to their live IFluidDataObject.
			var registry = new FakeFluidDataObjectRegistry();
			var handle = new TestFluidDataObject("target");
			registry.Register(handle, "/dataObjects/target");
			var dir = new SharedDirectory(registry: registry);

			const string op = "{\"type\":\"set\",\"path\":\"/\",\"key\":\"k\","
				+ "\"value\":{\"type\":\"Plain\","
				+ "\"value\":{\"type\":\"__fluid_handle__\",\"url\":\"/dataObjects/target\"}}}";
			dir.ProcessDataObjectOp(RemoteMessage(refSeq: 0, seq: 1), op);

			Assert.Same(handle, dir.Get("k"));
		}

		// -------------------------------------------------------------
		// SD-T01: LoadFromSnapshot into a non-empty directory throws
		// -------------------------------------------------------------

		[Fact]
		public void LoadFromSnapshot_NonEmptyDirectory_ThrowsAndLeavesContentUnchanged()
		{
			// TS ref: directory.ts:700-715, 722-724. TS's loadCore is called once at
			// construction; loading into an already-populated directory would be a
			// program error. C# raises InvalidState. Verify that pre-existing content
			// is unchanged after the throw.
			var dir = new SharedDirectory();
			dir.Set("pre-existing", "keep-me");
			dir.CreateSubDirectory("pre-existing-subdir");

			const string snapshotJson = "{\"storage\":{\"new-key\":{\"type\":\"Plain\",\"value\":\"new-value\"}}}";

			OcsException exception = Assert.Throws<OcsException>(() => dir.LoadFromSnapshot(snapshotJson));
			Assert.Equal(OcsGateErrorCode.InvalidState, exception.ErrorCode);

			// Pre-existing content untouched.
			Assert.Equal("keep-me", dir.Get("pre-existing"));
			Assert.True(dir.HasSubDirectory("pre-existing-subdir"));
			// New snapshot content not applied.
			Assert.Null(dir.Get("new-key"));
		}

		[Fact]
		public void LoadFromSnapshot_NonEmptyDirectory_OnlyFromSubdirectory_Throws()
		{
			// Even one subdirectory (no keys at root) should block LoadFromSnapshot.
			var dir = new SharedDirectory();
			dir.CreateSubDirectory("just-a-subdir");

			const string snapshotJson = "{\"storage\":{}}";

			OcsException exception = Assert.Throws<OcsException>(() => dir.LoadFromSnapshot(snapshotJson));
			Assert.Equal(OcsGateErrorCode.InvalidState, exception.ErrorCode);
			Assert.True(dir.HasSubDirectory("just-a-subdir"));
		}

		// -------------------------------------------------------------
		// SD-T03: local ack with zero/missing client sequence throws
		// -------------------------------------------------------------

		[Fact]
		public void LocalAck_MissingClientSequenceNumber_ThrowsInvalidSequenceNumber()
		{
			// TS ref: directory.ts:76-83, 792-815. A local ack must carry a client
			// sequence so the port can locate the matching pending entry. A zero /
			// missing sentinel indicates a malformed ack from the runtime; C# throws
			// InvalidSequenceNumber.
			var sender = new FakeFluidDataObjectSender();
			var dir = new SharedDirectory("dir", sender);

			// Enqueue a local set — this creates a pending entry.
			dir.Set("k", "v");
			var sentOp = Assert.Single(sender.Sent);

			// Simulate an ack that lacks a client sequence number. HasClientSequenceNumber
			// on SequenceNumber returns false when clientSequenceNumber == 0, which is
			// the sentinel for "runtime didn't stamp a client seq on this ack".
			SequenceNumber ackWithoutClientSeq = SequenceNumber.ForTesting(
				clientSeq: 0,
				refSeq: 0,
				seq: 1);
			var descriptor = new SequencedDocumentMessageDescriptor(
				ackWithoutClientSeq,
				OpOrigin.Local,
				clientId: "self");

			OcsException exception = Assert.Throws<OcsException>(
				() => dir.ProcessDataObjectOp(descriptor, sentOp.OpJson));
			Assert.Equal(OcsGateErrorCode.InvalidSequenceNumber, exception.ErrorCode);

			// The pending entry is unchanged — the key is still visible optimistically.
			Assert.Equal("v", dir.Get("k"));
		}

		// -------------------------------------------------------------
		// Helpers (shared fakes replicated from other test files)
		// -------------------------------------------------------------

		private static SequencedDocumentMessageDescriptor RemoteMessage(long refSeq, long seq)
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: 0, refSeq: refSeq, seq: seq),
				OpOrigin.Remote,
				clientId: "remote-client");
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
			private readonly Dictionary<string, IFluidDataObject> _objectsByUrl = new();
			private readonly Dictionary<IFluidDataObject, string> _urlsByObject = new();

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
