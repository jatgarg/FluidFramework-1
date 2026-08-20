#nullable enable

using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;
using TextSegmentModel = Microsoft.Office.Web.Fluid.MergeTree.TextSegment;
using static Microsoft.Office.Web.Fluid.Tests.PortedTestUtilities;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class ClientApplyMsgTestsFromTS
	{
		// Ported from packages/dds/merge-tree/src/test/client.applyMsg.spec.ts — "Interleaved inserts, annotates, and deletes"
		[Fact]
		public void InterleavedInsertsAnnotatesAndDeletes_ConvergeAndPreserveAnnotations()
		{
			PortedTestUtilities.TwoClientHarness harness = new();
			harness.LoadInitialText("hello world");

			harness.ClientA.AnnotateRange(0, 5, new PropertySet() { ["style"] = "greeting" });
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), harness.CurrentServerSeq);
			harness.SenderA.Sent.Clear();
			harness.ClientB.DeleteText(5, 6);
			harness.DeliverBtoA(Assert.Single(harness.SenderB.Sent), harness.CurrentServerSeq);
			harness.SenderB.Sent.Clear();
			harness.ClientA.InsertText(harness.ClientA.GetLength(), "!");
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), harness.CurrentServerSeq);

			AssertConverged(harness.ClientA, harness.ClientB);
			Assert.Equal("helloworld!", harness.ClientA.GetText());
			Assert.Equal("greeting", Assert.IsType<string>(harness.ClientB.GetPropertiesAtPosition(0)!["style"]));
		}

		// Ported from packages/dds/merge-tree/src/test/client.applyMsg.spec.ts — "overlapping deletes"
		[Fact]
		public void OverlappingDeletes_FromDifferentClients_KeepSingleVisibleRemainder()
		{
			PortedTestUtilities.TwoClientHarness harness = new();
			harness.LoadInitialText("hello world");
			long refSeq = harness.CurrentServerSeq;

			harness.ClientA.DeleteText(0, 6);
			var deleteA = Assert.Single(harness.SenderA.Sent);
			harness.ClientB.DeleteText(2, 8);
			var deleteB = Assert.Single(harness.SenderB.Sent);

			harness.DeliverAtoB(deleteA, refSeq);
			harness.DeliverBtoA(deleteB, refSeq);

			AssertConverged(harness.ClientA, harness.ClientB);
			Assert.Equal("rld", harness.ClientA.GetText());
		}

		// Ported from packages/dds/merge-tree/src/test/client.applyMsg.spec.ts — "overlapping insert and delete"
		[Fact]
		public void OverlappingInsertAndDelete_PreservesConcurrentInsert()
		{
			PortedTestUtilities.TwoClientHarness harness = new();
			harness.LoadInitialText("abcdef");
			long refSeq = harness.CurrentServerSeq;

			harness.ClientA.InsertText(3, "XX");
			var insert = Assert.Single(harness.SenderA.Sent);
			harness.ClientB.DeleteText(2, 5);
			var delete = Assert.Single(harness.SenderB.Sent);

			harness.DeliverAtoB(insert, refSeq);
			harness.DeliverBtoA(delete, refSeq);

			AssertConverged(harness.ClientA, harness.ClientB);
			Assert.Equal("abXXf", harness.ClientA.GetText());
		}

		// Ported from packages/dds/merge-tree/src/test/client.applyMsg.spec.ts — "intersecting insert after local delete"
		[Fact]
		public void IntersectingInsertAfterDelete_IsVisibleAtDeletedPosition()
		{
			SharedString sharedString = new();

			ProcessRemoteInsert(sharedString, 0, "c", refSeq: 0, seq: 1, clientId: "client-c");
			ProcessRemoteRemove(sharedString, 0, 1, refSeq: 1, seq: 2, clientId: "client-c");
			ProcessRemoteInsert(sharedString, 0, "b", refSeq: 2, seq: 3, clientId: "client-b");
			ProcessRemoteInsert(sharedString, 0, "c", refSeq: 3, seq: 4, clientId: "client-c");

			Assert.Equal("cb", sharedString.GetText());
		}

		// Ported from packages/dds/merge-tree/src/test/client.applyMsg.spec.ts — "conflicting insert over local delete"
		[Fact]
		public void ConflictingInsertOverLocalDelete_ConvergesAfterAck()
		{
			SharedString sharedString = CreateAckedSharedString("client-a", "abcdef").SharedString;
			sharedString.DeleteText(2, 4);
			List<SequenceDeltaEventArgs> events = CaptureEvents(sharedString);

			ProcessRemoteInsert(sharedString, 3, "X", refSeq: 1, seq: 2, clientId: "client-b");

			Assert.Contains(events, e => !e.Local && e.OpType == "insert");
			Assert.Equal("abXef", sharedString.GetText());
		}

		// Ported from packages/dds/merge-tree/src/test/client.applyMsg.spec.ts — "Local insert after acked local delete"
		[Fact]
		public void LocalInsertAfterAckedLocalDelete_UsesCurrentVisiblePosition()
		{
			var (sharedString, sender) = CreateAckedSharedString("client-a", "ZZ");

			sharedString.DeleteText(0, 1);
			ProcessLocalAck(sharedString, Assert.Single(sender.Sent), refSeq: 1, seq: 2);
			sender.Sent.Clear();
			sharedString.InsertText(0, "C");
			ProcessLocalAck(sharedString, Assert.Single(sender.Sent), refSeq: 2, seq: 3);

			Assert.Equal("CZ", sharedString.GetText());
		}

		// Ported from packages/dds/merge-tree/src/test/client.applyMsg.spec.ts — "Remote Remove before conflicting insert"
		[Fact]
		public void RemoteRemoveBeforeConflictingInsert_AllowsLaterInsertAtRemovedPosition()
		{
			SharedString sharedString = CreateAckedSharedString("client-a", "Z").SharedString;

			ProcessRemoteRemove(sharedString, 0, 1, refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteInsert(sharedString, 0, "B", refSeq: 2, seq: 3, clientId: "client-b");
			ProcessRemoteInsert(sharedString, 0, "C", refSeq: 2, seq: 4, clientId: "client-c");

			Assert.Equal("CB", sharedString.GetText());
		}

		// Ported from packages/dds/merge-tree/src/test/client.applyMsg.spec.ts — "Concurrent insert into removed segment across block boundary"
		[Fact]
		public void ConcurrentInsertIntoRemovedSegmentAcrossBlockBoundary_RemainsVisibleAndKeepsInvariants()
		{
			MergeTreeModel tree = CreateSingleCharacterTree("abcdefghijklmnopqrst");

			tree.MarkRangeRemoved(6, 14, refSeq: 20, seq: 30, clientId: "remover");
			tree.InsertSegments(10, new ISegment[] { TextSegmentModel.Make("X") }, refSeq: 20, seq: 31, clientId: "inserter");

			Assert.Contains("X", tree.GetText());
			Assert.Equal(13, tree.GetLength());
			PartialLengths.ValidateBlockPartialLengthInvariants(tree.Root);
		}
	}
}
