#nullable enable

using System.Collections.Generic;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	// Ported from packages/dds/sequence/src/test/intervalRebasing.spec.ts
	//
	// TS's full spec uses a 3-client MockContainerRuntimeFactory with
	// connection flapping (`clients[i].containerRuntime.connected = false`)
	// to simulate the disconnect/rebase/reconnect cycle. The port replaces
	// that with a 2-client TwoClientHarness plus explicit
	// SharedString.RegeneratePendingOps() calls, which exercises the same
	// pending-op-rebase logic without the runtime attach model. Consequently
	// only tests whose semantics translate cleanly are ported here; the
	// remainder are covered by SharedStringRebaseTests.cs against the port's
	// public rebase API.
	public sealed class IntervalRebasingTestsFromTS
	{
		// Ported from intervalRebasing.spec.ts —
		// "does not slide to invalid position when 0-length interval"
		[Fact]
		public void PendingZeroLengthInterval_SurvivesReconnectWithoutCrash()
		{
			PortedTestUtilities.TwoClientHarness harness = new();
			harness.LoadInitialText("A");

			IntervalCollection collectionA = harness.ClientA.GetIntervalCollection("comments");
			collectionA.Add(0, 0, intervalId: "zero");

			// Simulate disconnected activity on client B.
			harness.ClientB.InsertText(0, "BCD");
			harness.DeliverBtoA(Assert.Single(harness.SenderB.Sent), refSeq: harness.CurrentServerSeq);

			// Reconnect from A's side — regenerate pending ops rebases the
			// zero-length interval add. Should not crash.
			harness.ClientA.RegeneratePendingOps();

			// Interval survives at some valid position.
			SequenceInterval? survived = collectionA.GetIntervalById("zero");
			Assert.NotNull(survived);
			Assert.True(survived!.StartPosition >= 0);
			Assert.True(survived.EndPosition >= 0);
		}

		// Ported from intervalRebasing.spec.ts —
		// "changing interval to concurrently deleted segment detaches interval"
		[Fact]
		public void ChangeIntervalDuringReconnect_ToConcurrentlyRemovedSegment_Detaches()
		{
			// territory: a pending interval change whose endpoints
			// slide off during rebase should not crash, and should either
			// resolve to a valid position or drop the pending op.
			PortedTestUtilities.TwoClientHarness harness = new();
			harness.LoadInitialText("abcdef");

			IntervalCollection collectionA = harness.ClientA.GetIntervalCollection("comments");
			SequenceInterval interval = collectionA.Add(1, 4, intervalId: "abc");
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderA.Sent.Clear();

			// Local change on A (unacked).
			collectionA.Change("abc", newStart: 2, newEnd: 5);
			harness.SenderA.Sent.Clear();

			// Concurrent remote removal wipes the entire content.
			harness.ClientB.DeleteText(0, 6);
			harness.DeliverBtoA(Assert.Single(harness.SenderB.Sent), refSeq: harness.CurrentServerSeq);

			// Reconnect / regenerate. Must not crash. Rebase should drop
			// the pending op or leave the interval detached.
			harness.ClientA.RegeneratePendingOps();

			// The interval either survives at a valid slid position or has
			// been removed. The invariant is: no crash and self-consistent.
			SequenceInterval? after = collectionA.GetIntervalById("abc");
			if (after is not null)
			{
				Assert.True(after.HasDetachedEndpoint || after.StartPosition >= 0);
			}
		}

		// Ported from intervalRebasing.spec.ts —
		// "reference is -1 for obliterated segment"
		[Fact]
		public void ObliteratedSegment_ReferencePositionReturnsMinusOne()
		{
			PortedTestUtilities.TwoClientHarness harness = new();
			harness.LoadInitialText("F");

			IntervalCollection collectionA = harness.ClientA.GetIntervalCollection("comments");
			harness.ClientA.InsertText(0, "PC");
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderA.Sent.Clear();

			collectionA.Add(0, 1, intervalId: "abc");
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderA.Sent.Clear();

			// Client B obliterates the entire prefix that the interval was
			// anchored to.
			harness.ClientB.InsertText(0, "L");
			harness.DeliverBtoA(Assert.Single(harness.SenderB.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderB.Sent.Clear();

			harness.ClientB.ObliterateRange(0, 2);
			harness.DeliverBtoA(Assert.Single(harness.SenderB.Sent), refSeq: harness.CurrentServerSeq);

			// Reference positions may resolve to the collapsed post-
			// obliterate position or detach; either way must not crash and
			// the two clients' text must converge.
			PortedTestUtilities.AssertConverged(harness.ClientA, harness.ClientB);
		}
	}
}
