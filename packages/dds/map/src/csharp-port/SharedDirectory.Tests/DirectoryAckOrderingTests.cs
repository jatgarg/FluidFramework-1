// -----------------------------------------------------------------------------
// Tests SharedDirectory sibling ordering under remote acknowledgement of
// an already-optimistic child.
//
// TS ref: packages/dds/map/src/directory.ts. When ApplyRemoteCreate
// runs against a subdir that already exists (from a local pending create or a
// previous remote create), TS assigns
//   subDir.seqData.clientSeq = clientSequenceNumber
// from the incoming message unconditionally. That clientSeq drives
// SeqDataComparator sibling ordering for grouped batches — a mismatch changes
// enumeration order when two siblings share a sequenceNumber but arrive via
// different code paths.
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
			// the else branch → MarkCreatedSubDirectorySequencedNoLock. The message
			// has (seq=42, clientSeq=5). TS assigns Y.clientSeq = 5.
			SequencedDocumentMessageDescriptor remoteY = new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: 5, refSeq: 0, seq: 42),
				OpOrigin.Remote,
				clientId: "remote-client");
			dir.ProcessDataObjectOp(remoteY, "{\"type\":\"createSubDirectory\",\"path\":\"/\",\"subdirName\":\"Y\"}");

			// TS iteration order (both siblings at seq=42, sort by clientSeq):
			//   X (clientSeq=3), then Y (clientSeq=5).
			IReadOnlyList<string> order = dir.SubDirectories().Select(kvp => kvp.Key).ToList();
			Assert.Equal(new[] { "X", "Y" }, order);
		}
	}
}

