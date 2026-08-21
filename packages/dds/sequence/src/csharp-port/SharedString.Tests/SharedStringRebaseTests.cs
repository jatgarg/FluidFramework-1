// -----------------------------------------------------------------------------
// Rebase-on-reconnect tests for SharedString C# port.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;
using MergeTreeSide = Microsoft.Office.Web.Fluid.MergeTree.Side;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SharedStringRebaseTests
	{
		[Fact]
		public void Rebase_PendingInsert_AfterInterveningInsertBefore_PositionAdjusted()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdefghij");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdefghij").SharedString;

			local.InsertText(5, "abc");
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 2, "XX", refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 2, "XX", refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			MergeTreeInsertMsg op = Assert.IsType<MergeTreeInsertMsg>(SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Equal(7, op.Pos1);

			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 2, seq: 3);
			Assert.Equal("abXXcdeabcfghij", local.GetText());
			Assert.Equal(local.GetText(), serverVisible.GetText());
		}

		[Fact]
		public void Rebase_PendingInsert_AfterInterveningRemoveBefore_PositionAdjusted()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdefghij");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdefghij").SharedString;

			local.InsertText(5, "abc");
			sender.Sent.Clear();
			ProcessRemoteRemove(local, 1, 4, refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteRemove(serverVisible, 1, 4, refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			MergeTreeInsertMsg op = Assert.IsType<MergeTreeInsertMsg>(SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Equal(2, op.Pos1);

			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 2, seq: 3);
			Assert.Equal("aeabcfghij", local.GetText());
			Assert.Equal(local.GetText(), serverVisible.GetText());
		}

		[Fact]
		public void Rebase_PendingRemove_AfterInterveningInsertInside_RangeAdjusted()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdefghi");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdefghi").SharedString;

			local.DeleteText(3, 7);
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 5, "X", refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 5, "X", refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			MergeTreeGroupMsg group = Assert.IsType<MergeTreeGroupMsg>(SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Collection(
				group.Ops,
				first => AssertRemove(first, 3, 5),
				second => AssertRemove(second, 4, 6));

			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 2, seq: 3);
			Assert.Equal("abcXhi", local.GetText());
			Assert.Equal(local.GetText(), serverVisible.GetText());
		}

		[Fact]
		public void Rebase_PendingAnnotate_RangeAdjusted()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdefghi");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdefghi").SharedString;
			PropertySet props = new()
			{
				["color"] = "red",
			};

			local.AnnotateRange(3, 7, props);
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 5, "X", refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 5, "X", refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			MergeTreeGroupMsg group = Assert.IsType<MergeTreeGroupMsg>(SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Collection(
				group.Ops,
				first => AssertAnnotate(first, 3, 5),
				second => AssertAnnotate(second, 6, 8));

			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 2, seq: 3);
			AssertColor(serverVisible, 3, "red");
			AssertColor(serverVisible, 4, "red");
			AssertNoColor(serverVisible, 5);
			AssertColor(serverVisible, 6, "red");
			AssertColor(serverVisible, 7, "red");
			Assert.Equal(local.GetText(), serverVisible.GetText());
		}

		[Fact]
		public void Rebase_PendingGroup_AllSubOpsRebased()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdefghi");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdefghi").SharedString;

			local.RunInBatch(() =>
			{
				local.InsertText(5, "I");
				local.DeleteText(1, 3);
				local.AnnotateRange(2, 4, new PropertySet()
				{
					["tag"] = "group",
				});
			});
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 0, "XX", refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 0, "XX", refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			MergeTreeGroupMsg group = Assert.IsType<MergeTreeGroupMsg>(SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Collection(
				group.Ops,
				first => AssertInsert(first, 7),
				second => AssertRemove(second, 3, 5),
				third => AssertAnnotate(third, 4, 6));

			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 2, seq: 3);
			Assert.Equal("XXadeIfghi", local.GetText());
			Assert.Equal(local.GetText(), serverVisible.GetText());
			AssertTag(serverVisible, 4, "group");
			AssertTag(serverVisible, 5, "group");
		}

		[Fact]
		public void Rebase_PendingIntervalAdd_AfterInterveningInsertBefore_PositionsAdjusted()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdefghij");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdefghij").SharedString;

			local.GetIntervalCollection("comments").Add(5, 10, intervalId: "abc");
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 2, "XX", refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 2, "XX", refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			IntervalAddOpMsg op = Assert.IsType<IntervalAddOpMsg>(SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Equal(7, op.Start);
			Assert.Equal(12, op.End);

			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 2, seq: 3);
			AssertIntervalPositions(local.GetIntervalCollection("comments").GetIntervalById("abc"), 7, 12);
			AssertIntervalPositions(serverVisible.GetIntervalCollection("comments").GetIntervalById("abc"), 7, 12);
		}

		[Fact]
		public void Rebase_PendingIntervalChange_PositionsAdjusted()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdefghij");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdefghij").SharedString;
			AddAckedInterval(local, sender, serverVisible, "comments", "abc", 1, 3, refSeq: 1, seq: 2);

			local.GetIntervalCollection("comments").Change("abc", newStart: 5, newEnd: 10);
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 2, "XX", refSeq: 2, seq: 3, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 2, "XX", refSeq: 2, seq: 3, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			IntervalChangeOpMsg op = Assert.IsType<IntervalChangeOpMsg>(SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Equal("abc", op.IntervalId);
			Assert.Equal(7, op.Start);
			Assert.Equal(12, op.End);

			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 3, seq: 4);
			AssertIntervalPositions(local.GetIntervalCollection("comments").GetIntervalById("abc"), 7, 12);
			AssertIntervalPositions(serverVisible.GetIntervalCollection("comments").GetIntervalById("abc"), 7, 12);
		}

		[Fact]
		public void Rebase_PendingRemoveById_NoPositionDependency()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdefghij");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdefghij").SharedString;
			AddAckedInterval(local, sender, serverVisible, "comments", "abc", 1, 3, refSeq: 1, seq: 2);

			local.GetIntervalCollection("comments").RemoveIntervalById("abc");
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 2, "XX", refSeq: 2, seq: 3, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 2, "XX", refSeq: 2, seq: 3, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			IntervalDeleteOpMsg op = Assert.IsType<IntervalDeleteOpMsg>(SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Equal("abc", op.IntervalId);

			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 3, seq: 4);
			Assert.Null(local.GetIntervalCollection("comments").GetIntervalById("abc"));
			Assert.Null(serverVisible.GetIntervalCollection("comments").GetIntervalById("abc"));
		}

		[Fact]
		public void Rebase_PendingIntervalAdd_TargetSegmentRemoved_SlidesForward()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdefghij");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdefghij").SharedString;

			local.GetIntervalCollection("comments").Add(5, 5, intervalId: "abc");
			sender.Sent.Clear();
			ProcessRemoteRemove(local, 4, 7, refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteRemove(serverVisible, 4, 7, refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			IntervalAddOpMsg op = Assert.IsType<IntervalAddOpMsg>(SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Equal(4, op.Start);
			Assert.Equal(4, op.End);

			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 2, seq: 3);
			Assert.Equal("abcdhij", local.GetText());
			Assert.Equal(local.GetText(), serverVisible.GetText());
			AssertIntervalPositions(serverVisible.GetIntervalCollection("comments").GetIntervalById("abc"), 4, 4);
		}

		[Fact]
		public void Rebase_PendingIntervalAdd_BothEndpointsDetached_DropsInterval()
		{
			// regression. TS's rebaseLocalInterval returns undefined
			// (drops the pending op) AND calls deleteExistingInterval when
			// the endpoint slides to 'detached'. The port previously only
			// dropped Transient intervals; normal SlideOnRemove intervals
			// whose endpoints both slid off were clamped and resubmitted at
			// invalid coordinates. TS ref: intervalCollection.ts
			// rebaseLocalInterval — `rebasedEndpoint === "detached"` branch.
			var (local, sender) = CreateAckedSharedString("client-a", "abcdef");

			local.GetIntervalCollection("comments").Add(2, 4, intervalId: "abc");
			sender.Sent.Clear();

			// Remote peer wipes the whole document — both endpoints of "abc"
			// now anchor to sequenced-removed segments with no visible
			// content to slide to.
			ProcessRemoteRemove(local, 0, 6, refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			// Regeneration must drop the interval add and NOT emit an op with
			// nonsensical coordinates.
			Assert.Empty(sender.Sent);
			Assert.Null(local.GetIntervalCollection("comments").GetIntervalById("abc"));
		}

		[Fact]
		public void Rebase_PendingIntervalAdd_InGroup_HandledCorrectly()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdefghi");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdefghi").SharedString;

			local.RunInBatch(() =>
			{
				local.InsertText(5, "I");
				local.GetIntervalCollection("comments").Add(6, 8, intervalId: "abc");
			});
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 0, "XX", refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 0, "XX", refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			// TS-parity: interval ops always go on the wire as individual "act" envelopes,
			// never inside a MergeTreeGroupMsg. So the rebase emits the merge-tree op
			// (insert, single member => not grouped) and the interval add separately.
			Assert.Equal(2, sender.Sent.Count);
			MergeTreeInsertMsg insert = Assert.IsType<MergeTreeInsertMsg>(SharedStringOpSerializer.Deserialize(sender.Sent[0].OpJson));
			AssertInsert(insert, 7);
			IntervalAddOpMsg intervalAdd = Assert.IsType<IntervalAddOpMsg>(SharedStringOpSerializer.Deserialize(sender.Sent[1].OpJson));
			AssertIntervalAdd(intervalAdd, 8, 10);

			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, sender.Sent[0], refSeq: 2, seq: 3);
			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, sender.Sent[1], refSeq: 2, seq: 4);
			Assert.Equal("XXabcdeIfghi", local.GetText());
			Assert.Equal(local.GetText(), serverVisible.GetText());
			AssertIntervalPositions(serverVisible.GetIntervalCollection("comments").GetIntervalById("abc"), 8, 10);
		}

		[Fact]
		public void Rebase_NoInterveningOps_PositionsUnchanged()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdef");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdef").SharedString;

			local.InsertText(3, "X");
			sender.Sent.Clear();

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			MergeTreeInsertMsg op = Assert.IsType<MergeTreeInsertMsg>(SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Equal(3, op.Pos1);

			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 1, seq: 2);
			Assert.Equal("abcXdef", local.GetText());
			Assert.Equal(local.GetText(), serverVisible.GetText());
		}

		[Fact]
		public void Rebase_ThenAck_PromotesStamps()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdefghij");

			local.InsertText(5, "abc");
			ISegment insertedSegment = GetRequiredContainingSegment(local, 5);
			Assert.Equal(MergeTree.MergeTree.UnassignedSequenceNumber, insertedSegment.Seq);
			Assert.NotNull(insertedSegment.LocalSeq);
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 2, "XX", refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();
			ProcessLocalAck(local, Assert.Single(sender.Sent), refSeq: 2, seq: 3);

			Assert.Equal(3, insertedSegment.Seq);
			Assert.Null(insertedSegment.LocalSeq);
			Assert.Equal("abXXcdeabcfghij", local.GetText());
		}

		[Fact]
		public void Rebase_TwoClientConvergence()
		{
			var senderA = new FakeFluidDataObjectSender();
			var senderB = new FakeFluidDataObjectSender();
			var clientA = new SharedString("client-a", senderA);
			var clientB = new SharedString("client-b", senderB);

			clientA.InsertText(0, "abcdefghi");
			var initial = Assert.Single(senderA.Sent);
			clientB.ProcessDataObjectOp(RemoteMessage(refSeq: 0, seq: 1, clientId: "client-a"), initial.OpJson);
			ProcessLocalAck(clientA, initial, refSeq: 0, seq: 1);
			senderA.Sent.Clear();
			senderB.Sent.Clear();

			clientA.InsertText(5, "A");
			clientA.RegeneratePendingOps();
			senderA.Sent.Clear();
			clientB.InsertText(2, "BB");
			var bOp = Assert.Single(senderB.Sent);
			clientA.ProcessDataObjectOp(RemoteMessage(refSeq: 1, seq: 2, clientId: "client-b"), bOp.OpJson);
			ProcessLocalAck(clientB, bOp, refSeq: 1, seq: 2);

			clientA.RegeneratePendingOps();
			var rebasedA = Assert.Single(senderA.Sent);
			clientB.ProcessDataObjectOp(RemoteMessage(refSeq: 2, seq: 3, clientId: "client-a"), rebasedA.OpJson);
			ProcessLocalAck(clientA, rebasedA, refSeq: 2, seq: 3);

			Assert.Equal("abBBcdeAfghi", clientA.GetText());
			Assert.Equal(clientA.GetText(), clientB.GetText());
		}

		[Fact]
		public void Rebase_PendingInsert_MultiHop_ConvergesCorrectly()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdefghij");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdefghij").SharedString;

			local.InsertText(5, "I");
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 0, "XX", refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 0, "XX", refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();
			MergeTreeInsertMsg firstRebase = Assert.IsType<MergeTreeInsertMsg>(
				SharedStringOpSerializer.Deserialize(Assert.Single(sender.Sent).OpJson));
			Assert.Equal(7, firstRebase.Pos1);

			sender.Sent.Clear();
			ProcessRemoteInsert(local, 4, "YY", refSeq: 2, seq: 3, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 4, "YY", refSeq: 2, seq: 3, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			MergeTreeInsertMsg secondRebase = Assert.IsType<MergeTreeInsertMsg>(
				SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Equal(9, secondRebase.Pos1);
			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 3, seq: 4);
			Assert.Equal("XXabYYcdeIfghij", local.GetText());
			Assert.Equal(local.GetText(), serverVisible.GetText());
		}

		[Fact]
		public void Rebase_PendingObliterate_NoExpand_DuringReconnect()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "0123456789AB");
			SharedString serverVisible = CreateAckedSharedString("server", "0123456789AB").SharedString;

			local.ObliterateRange(5, 10);
			sender.Sent.Clear();
			// Regression for C12: this insert lands inside the original obliterate range,
			// but the reconnect-generated obliterate must split rather than expand over it.
			ProcessRemoteInsert(local, 7, "XXX", refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 7, "XXX", refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			MergeTreeGroupMsg group = Assert.IsType<MergeTreeGroupMsg>(SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Collection(
				group.Ops,
				first => AssertObliterate(first, 5, 7),
				second => AssertObliterate(second, 8, 11));
			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 2, seq: 3);
			Assert.Equal("01234XXXAB", local.GetText());
			Assert.Equal(local.GetText(), serverVisible.GetText());
		}

		[Fact]
		public void Rebase_PendingObliterate_ConcurrentInsertBeforeRange_Unaffected()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "0123456789AB");
			SharedString serverVisible = CreateAckedSharedString("server", "0123456789AB").SharedString;

			local.ObliterateRange(5, 10);
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 2, "BB", refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 2, "BB", refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			AssertObliterate(SharedStringOpSerializer.Deserialize(rebased.OpJson), 7, 12);
			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 2, seq: 3);
			Assert.Equal("01BB234AB", local.GetText());
			Assert.Equal(local.GetText(), serverVisible.GetText());
		}

		[Fact]
		public void Rebase_PendingObliterate_ConcurrentInsertAfterRange_Unaffected()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "0123456789AB");
			SharedString serverVisible = CreateAckedSharedString("server", "0123456789AB").SharedString;

			local.ObliterateRange(5, 10);
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 10, "ZZ", refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 10, "ZZ", refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			AssertObliterate(SharedStringOpSerializer.Deserialize(rebased.OpJson), 5, 10);
			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 2, seq: 3);
			Assert.Equal("01234ZZAB", local.GetText());
			Assert.Equal(local.GetText(), serverVisible.GetText());
		}

		[Fact]
		public void Rebase_GroupOp_MixedTypes_EachSubOpRebased()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdefghij");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdefghij").SharedString;

			local.RunInBatch(() =>
			{
				local.InsertText(5, "I");
				local.DeleteText(8, 9);
				local.AnnotateRange(6, 8, new PropertySet()
				{
					["tag"] = "mixed",
				});
				local.ObliterateRange(2, 4);
			});
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 0, "XX", refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 0, "XX", refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			MergeTreeGroupMsg group = Assert.IsType<MergeTreeGroupMsg>(SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Collection(
				group.Ops,
				first => AssertInsert(first, 7),
				second => AssertRemove(second, 10, 11),
				third => AssertAnnotate(third, 8, 10),
				fourth => AssertObliterate(fourth, 4, 6));
			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 2, seq: 3);
			Assert.Equal("XXabeIfgij", local.GetText());
			Assert.Equal(local.GetText(), serverVisible.GetText());
			AssertTag(serverVisible, 6, "mixed");
			AssertTag(serverVisible, 7, "mixed");
		}

		[Fact]
		public void Rebase_SidedObliterate_PreservesSidesAcrossRebase()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "0123456789AB");
			SharedString serverVisible = CreateAckedSharedString("server", "0123456789AB").SharedString;

			local.ObliterateRange(SequencePlace.At(5, MergeTreeSide.After), SequencePlace.At(10, MergeTreeSide.Before));
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 0, "XX", refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 0, "XX", refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			AssertSidedObliterate(SharedStringOpSerializer.Deserialize(rebased.OpJson), 7, MergeTreeSide.After, 12, MergeTreeSide.Before);
			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 2, seq: 3);
			ProcessRemoteInsert(local, 7, "S", refSeq: 2, seq: 4, clientId: "client-c");
			ProcessRemoteInsert(serverVisible, 7, "S", refSeq: 2, seq: 4, clientId: "client-c");
			Assert.Equal("XX01234AB", local.GetText());
			Assert.Equal(local.GetText(), serverVisible.GetText());
		}

		[Fact]
		public void Rebase_IntervalAdd_MultiHop_ConvergesCorrectly()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdefghij");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdefghij").SharedString;

			local.GetIntervalCollection("comments").Add(5, 8, intervalId: "abc");
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 0, "XX", refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 0, "XX", refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();
			IntervalAddOpMsg firstRebase = Assert.IsType<IntervalAddOpMsg>(
				SharedStringOpSerializer.Deserialize(Assert.Single(sender.Sent).OpJson));
			Assert.Equal(7, firstRebase.Start);
			Assert.Equal(10, firstRebase.End);

			sender.Sent.Clear();
			ProcessRemoteInsert(local, 4, "YY", refSeq: 2, seq: 3, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 4, "YY", refSeq: 2, seq: 3, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			IntervalAddOpMsg secondRebase = Assert.IsType<IntervalAddOpMsg>(SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Equal(9, secondRebase.Start);
			Assert.Equal(12, secondRebase.End);
			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 3, seq: 4);
			AssertIntervalPositions(local.GetIntervalCollection("comments").GetIntervalById("abc"), 9, 12);
			AssertIntervalPositions(serverVisible.GetIntervalCollection("comments").GetIntervalById("abc"), 9, 12);
		}

		[Fact]
		public void Rebase_LocalOpDuringDisconnect_ArrivesAfter_RemoteAppliedFirst()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdef");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdef").SharedString;

			local.DeleteText(2, 4);
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 1, "R", refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 1, "R", refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			AssertRemove(SharedStringOpSerializer.Deserialize(rebased.OpJson), 3, 5);
			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 2, seq: 3);
			Assert.Equal("aRbef", local.GetText());
			Assert.Equal(local.GetText(), serverVisible.GetText());
		}

		[Fact]
		public void Rebase_ObliterateOfInsertedText_HandlesCorrectly()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "AB");
			SharedString serverVisible = CreateAckedSharedString("server", "AB").SharedString;

			local.RunInBatch(() =>
			{
				local.InsertText(1, "xxx");
				local.ObliterateRange(1, 4);
			});
			sender.Sent.Clear();

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			MergeTreeGroupMsg group = Assert.IsType<MergeTreeGroupMsg>(SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Collection(
				group.Ops,
				first => AssertInsert(first, 1),
				second => AssertObliterate(second, 1, 4));
			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 1, seq: 2);
			Assert.Equal("AB", local.GetText());
			Assert.Equal(local.GetText(), serverVisible.GetText());
		}

		[Fact]
		public void Rebase_ManyPendingOps_OrderPreserved()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdef");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdef").SharedString;

			local.InsertText(6, "X");
			local.DeleteText(0, 1);
			local.InsertText(0, "Y");
			sender.Sent.Clear();
			ProcessRemoteInsert(local, 0, "R", refSeq: 1, seq: 2, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 0, "R", refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			Assert.Collection(
				sender.Sent,
				first => Assert.IsType<MergeTreeInsertMsg>(SharedStringOpSerializer.Deserialize(first.OpJson)),
				second => Assert.IsType<MergeTreeRemoveMsg>(SharedStringOpSerializer.Deserialize(second.OpJson)),
				third => Assert.IsType<MergeTreeInsertMsg>(SharedStringOpSerializer.Deserialize(third.OpJson)));
			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, sender.Sent[0], refSeq: 2, seq: 3);
			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, sender.Sent[1], refSeq: 3, seq: 4);
			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, sender.Sent[2], refSeq: 4, seq: 5);
			Assert.Equal(local.GetText(), serverVisible.GetText());
		}

		[Fact]
		public void Rebase_PendingCombinedIntervalChange_PreservesProperties()
		{
			// regression. TS ref: packages/dds/sequence/src/
			// intervalCollection.ts rebasePositionalOp — TS spreads the entire
			// change op (`{...op, value: {...op.value}}`) so `properties`
			// survives alongside recomputed endpoint fields. The port's
			// CloneOp for IntervalChangeOpMsg must copy Props to match.
			var (local, sender) = CreateAckedSharedString("client-a", "abcdefghij");
			SharedString serverVisible = CreateAckedSharedString("server", "abcdefghij").SharedString;

			AddAckedInterval(local, sender, serverVisible, "comments", "abc", 2, 5, refSeq: 1, seq: 2);
			sender.Sent.Clear();

			// Combined endpoint + property local change.
			local.GetIntervalCollection("comments").Change(
				"abc",
				newStart: 3,
				newEnd: 5,
				props: new PropertySet() { ["color"] = "blue" });
			sender.Sent.Clear();

			// A remote insert lands during the disconnected window.
			ProcessRemoteInsert(local, 0, "R", refSeq: 2, seq: 3, clientId: "client-b");
			ProcessRemoteInsert(serverVisible, 0, "R", refSeq: 2, seq: 3, clientId: "client-b");

			local.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			IntervalChangeOpMsg op = Assert.IsType<IntervalChangeOpMsg>(SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Equal("abc", op.IntervalId);
			// Endpoint fields are rebased (shifted +1 by the remote insert).
			Assert.Equal(4, op.Start);
			Assert.Equal(6, op.End);
			// Properties MUST survive the clone.
			Assert.NotNull(op.Props);
			Assert.Equal("blue", Assert.IsType<string>(op.Props!["color"]));

			// Peer picks up both endpoint AND property.
			ApplyRebasedOpToServerAndAckLocal(local, serverVisible, rebased, refSeq: 3, seq: 4);
			SequenceInterval? serverSide = serverVisible.GetIntervalCollection("comments").GetIntervalById("abc");
			Assert.NotNull(serverSide);
			Assert.Equal(4, serverSide!.StartPosition);
			Assert.Equal(6, serverSide.EndPosition);
			Assert.NotNull(serverSide.Properties);
			Assert.Equal("blue", Assert.IsType<string>(serverSide.Properties!["color"]));
		}

		private static (SharedString SharedString, FakeFluidDataObjectSender Sender) CreateAckedSharedString(string id, string text)
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString(id, sender);
			sharedString.InsertText(0, text);
			ProcessLocalAck(sharedString, Assert.Single(sender.Sent), refSeq: 0, seq: 1);
			sender.Sent.Clear();
			return (sharedString, sender);
		}

		private static void AddAckedInterval(
			SharedString local,
			FakeFluidDataObjectSender sender,
			SharedString serverVisible,
			string collectionName,
			string intervalId,
			int start,
			int end,
			long refSeq,
			long seq)
		{
			local.GetIntervalCollection(collectionName).Add(start, end, intervalId: intervalId);
			var sent = Assert.Single(sender.Sent);
			serverVisible.ProcessDataObjectOp(RemoteMessage(refSeq, seq, clientId: "client-a"), sent.OpJson);
			ProcessLocalAck(local, sent, refSeq, seq);
			sender.Sent.Clear();
		}

		private static void ApplyRebasedOpToServerAndAckLocal(
			SharedString local,
			SharedString serverVisible,
			(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
			long refSeq,
			long seq)
		{
			serverVisible.ProcessDataObjectOp(RemoteMessage(refSeq, seq, clientId: "client-a"), sent.OpJson);
			ProcessLocalAck(local, sent, refSeq, seq);
		}

		private static void ProcessRemoteInsert(
			SharedString sharedString,
			int position,
			string text,
			long refSeq,
			long seq,
			string clientId)
		{
			string opJson = SharedStringOpSerializer.Serialize(new MergeTreeInsertMsg()
			{
				Pos1 = position,
				Seg = text,
			});
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq, seq, clientId), opJson);
		}

		private static void ProcessRemoteRemove(
			SharedString sharedString,
			int start,
			int end,
			long refSeq,
			long seq,
			string clientId)
		{
			string opJson = SharedStringOpSerializer.Serialize(new MergeTreeRemoveMsg()
			{
				Pos1 = start,
				Pos2 = end,
			});
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq, seq, clientId), opJson);
		}

		private static void ProcessLocalAck(
			SharedString sharedString,
			(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
			long refSeq,
			long seq)
		{
			sharedString.ProcessDataObjectOp(LocalAck(sent.ClientSeq, refSeq, seq), sent.OpJson);
		}

		private static SequencedDocumentMessageDescriptor RemoteMessage(
			long refSeq,
			long seq,
			string clientId)
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: 0, refSeq: refSeq, seq: seq),
				OpOrigin.Remote,
				clientId);
		}

		private static SequencedDocumentMessageDescriptor LocalAck(long clientSeq, long refSeq, long seq)
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: clientSeq, refSeq: refSeq, seq: seq),
				OpOrigin.Local);
		}

		private static ISegment GetRequiredContainingSegment(SharedString sharedString, int position)
		{
			(ISegment segment, int offsetInSegment)? result = sharedString.GetContainingSegment(position);
			Assert.NotNull(result);
			return result.Value.segment;
		}

		private static void AssertInsert(IMergeTreeOp op, int expectedPosition)
		{
			MergeTreeInsertMsg insert = Assert.IsType<MergeTreeInsertMsg>(op);
			Assert.Equal(expectedPosition, insert.Pos1);
		}

		private static void AssertRemove(IMergeTreeOp op, int expectedStart, int expectedEnd)
		{
			MergeTreeRemoveMsg remove = Assert.IsType<MergeTreeRemoveMsg>(op);
			Assert.Equal(expectedStart, remove.Pos1);
			Assert.Equal(expectedEnd, remove.Pos2);
		}

		private static void AssertAnnotate(IMergeTreeOp op, int expectedStart, int expectedEnd)
		{
			MergeTreeAnnotateMsg annotate = Assert.IsType<MergeTreeAnnotateMsg>(op);
			Assert.Equal(expectedStart, annotate.Pos1);
			Assert.Equal(expectedEnd, annotate.Pos2);
		}

		private static void AssertObliterate(IMergeTreeOp op, int expectedStart, int expectedEnd)
		{
			MergeTreeObliterateMsg obliterate = Assert.IsType<MergeTreeObliterateMsg>(op);
			Assert.Equal(expectedStart, obliterate.Pos1);
			Assert.Equal(expectedEnd, obliterate.Pos2);
		}

		private static void AssertSidedObliterate(
			IMergeTreeOp op,
			int expectedStart,
			MergeTreeSide expectedStartSide,
			int expectedEnd,
			MergeTreeSide expectedEndSide)
		{
			MergeTreeObliterateSidedMsg obliterate = Assert.IsType<MergeTreeObliterateSidedMsg>(op);
			Assert.Equal(expectedStart, obliterate.Pos1.Position);
			Assert.Equal(expectedStartSide, obliterate.Pos1.Side);
			Assert.Equal(expectedEnd, obliterate.Pos2.Position);
			Assert.Equal(expectedEndSide, obliterate.Pos2.Side);
		}

		private static void AssertIntervalAdd(IMergeTreeOp op, int expectedStart, int expectedEnd)
		{
			IntervalAddOpMsg intervalAdd = Assert.IsType<IntervalAddOpMsg>(op);
			Assert.Equal(expectedStart, intervalAdd.Start);
			Assert.Equal(expectedEnd, intervalAdd.End);
		}

		private static void AssertIntervalPositions(SequenceInterval? interval, int expectedStart, int expectedEnd)
		{
			Assert.NotNull(interval);
			Assert.Equal(expectedStart, interval!.StartPosition);
			Assert.Equal(expectedEnd, interval.EndPosition);
		}

		private static void AssertColor(SharedString sharedString, int position, string expectedColor)
		{
			PropertySet properties = GetRequiredProperties(sharedString, position);
			Assert.Equal(expectedColor, Assert.IsType<string>(properties["color"]));
		}

		private static void AssertNoColor(SharedString sharedString, int position)
		{
			PropertySet? properties = sharedString.GetPropertiesAtPosition(position);
			Assert.True(properties is null || !properties.ContainsKey("color"), $"Expected no color at position {position}.");
		}

		private static void AssertTag(SharedString sharedString, int position, string expectedTag)
		{
			PropertySet properties = GetRequiredProperties(sharedString, position);
			Assert.Equal(expectedTag, Assert.IsType<string>(properties["tag"]));
		}

		private static PropertySet GetRequiredProperties(SharedString sharedString, int position)
		{
			PropertySet? properties = sharedString.GetPropertiesAtPosition(position);
			Assert.NotNull(properties);
			return properties!;
		}
	}
}
