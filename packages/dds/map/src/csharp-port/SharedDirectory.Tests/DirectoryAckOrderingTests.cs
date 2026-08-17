// -----------------------------------------------------------------------------
// SD-A01 repro test for SharedDirectory.
//
// Tests the clientSeq preservation on the "already-optimistic child gets
// sequenced" path — TS ref: directory.ts:2124-2167. When ApplyRemoteCreate
// runs against a subdir that already exists (either from a local pending
// create or from a previous remote create), TS unconditionally assigns
//   subDir.seqData.clientSeq = clientSequenceNumber
// from the incoming message. C# previously stored -1 for any non-local
// origin, losing the message's clientSeq. That matters for grouped-batch
// sibling ordering, which SeqDataComparator breaks by clientSeq when seq
// matches — the two orderings diverge whenever one sibling comes through
// the MarkCreatedSubDirectorySequencedNoLock branch and another comes
// through CreateRemoteSeqData with a smaller clientSeq than the "already
// optimistic" sibling should have had.
// -----------------------------------------------------------------------------

using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class DirectoryAckOrderingTests
	{
		[Fact]
		public void RemoteCreate_AgainstOptimisticLocalCreate_PreservesRemoteClientSeqForOrdering()
		{
			// Setup: attached SharedDirectory with a local sender. Local client creates
			// "Y" — subdir goes into pending storage with seqData.seq == -1 (unacked).
			var sender = new FakeFluidDataObjectSender();
			var dir = new SharedDirectory("dir", sender);
			dir.CreateSubDirectory("Y");
			Assert.Single(sender.Sent);

			// A remote message creates a fresh sibling "X" with (seq=42, clientSeq=3).
			// This goes through CreateRemoteSeqData → X ends up with SeqData(42, 3).
			SequencedDocumentMessageDescriptor remoteX = new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: 3, refSeq: 0, seq: 42),
				OpOrigin.Remote,
				clientId: "remote-client");
			dir.ProcessDataObjectOp(remoteX, "{\"type\":\"createSubDirectory\",\"path\":\"/\",\"subdirName\":\"X\"}");

			// A remote message creates "Y" (already optimistically present). Goes through
			// the else branch → MarkCreatedSubDirectorySequencedNoLock. The message has
			// (seq=42, clientSeq=5). TS assigns Y.clientSeq = 5. Previous C# assigned -1.
			SequencedDocumentMessageDescriptor remoteY = new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: 5, refSeq: 0, seq: 42),
				OpOrigin.Remote,
				clientId: "remote-client");
			dir.ProcessDataObjectOp(remoteY, "{\"type\":\"createSubDirectory\",\"path\":\"/\",\"subdirName\":\"Y\"}");

			// TS iteration order (both siblings at seq=42, sort by clientSeq):
			//   X (clientSeq=3), then Y (clientSeq=5).
			// With the SD-A01 bug (Y.clientSeq stored as -1):
			//   Y (clientSeq=-1), then X (clientSeq=3). Different order — observable.
			IReadOnlyList<string> order = dir.SubDirectories().Select(kvp => kvp.Key).ToList();
			Assert.Equal(new[] { "X", "Y" }, order);
		}
	}
}

