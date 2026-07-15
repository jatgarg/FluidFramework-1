// -----------------------------------------------------------------------------
// Transient interval semantics for SharedString.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SharedStringTransientIntervalTests
	{
		[Fact]
		public void Transient_Interval_Survives_NonTouchingMutation()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdef");
			IntervalCollection collection = sharedString.GetIntervalCollection("transient", IntervalType.Transient);
			SequenceInterval interval = collection.Add(2, 4, intervalId: "t1");

			sharedString.InsertText(6, "ZZ");

			Assert.Same(interval, collection.GetIntervalById("t1"));
			AssertIntervalPositions(interval, 2, 4);
		}

		[Fact]
		public void Transient_EndpointRemoved_Detaches()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdef");
			IntervalCollection collection = sharedString.GetIntervalCollection("transient", IntervalType.Transient);
			SequenceInterval interval = collection.Add(2, 4, intervalId: "t1");

			sharedString.DeleteText(2, 3);

			Assert.True(interval.Start.IsDetached);
			Assert.Null(collection.GetIntervalById("t1"));
		}

		[Fact]
		public void Transient_Iterator_SkipsDetached()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdefghij");
			IntervalCollection collection = sharedString.GetIntervalCollection("mixed", IntervalType.SlideOnRemove);
			collection.Add(2, 4, intervalId: "transient-removed", intervalType: IntervalType.Transient);
			collection.Add(7, 8, intervalId: "transient-survivor", intervalType: IntervalType.Transient);
			collection.Add(2, 4, intervalId: "slide-survivor");

			sharedString.DeleteText(2, 3);

			AssertIds(collection.CreateForwardIteratorWithStartPosition(0), "slide-survivor", "transient-survivor");
		}

		[Fact]
		public void Transient_vs_SlideOnRemove_SameCollection_DifferentBehavior()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdef");
			IntervalCollection collection = sharedString.GetIntervalCollection("mixed", IntervalType.SlideOnRemove);
			collection.Add(2, 2, intervalId: "transient", intervalType: IntervalType.Transient);
			collection.Add(2, 2, intervalId: "slide");

			sharedString.DeleteText(2, 3);

			Assert.Null(collection.GetIntervalById("transient"));
			AssertIntervalPositions(collection.GetIntervalById("slide"), 2, 2);
		}

		[Fact]
		public void Transient_EndpointObliterated_Detaches()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdef");
			IntervalCollection collection = sharedString.GetIntervalCollection("transient", IntervalType.Transient);
			SequenceInterval interval = collection.Add(2, 4, intervalId: "t1");

			sharedString.ObliterateRange(2, 3);

			Assert.True(interval.Start.IsDetached);
			Assert.Null(collection.GetIntervalById("t1"));
		}

		[Fact]
		public void Transient_EndpointDetached_DoesNotSurviveRebase()
		{
			var (local, sender) = CreateAckedSharedString("client-a", "abcdefghij");
			IntervalCollection collection = local.GetIntervalCollection("transient", IntervalType.Transient);
			var deleted = new List<IntervalDeletedEventArgs>();
			collection.OnDeleteInterval += (_, args) => deleted.Add(args);
			collection.Add(5, 5, intervalId: "t1");
			sender.Sent.Clear();
			ProcessRemoteRemove(local, 4, 7, refSeq: 1, seq: 2, clientId: "client-b");

			local.RegeneratePendingOps();

			Assert.Empty(sender.Sent);
			Assert.Null(collection.GetIntervalById("t1"));
			IntervalDeletedEventArgs deletedEvent = Assert.Single(deleted);
			Assert.True(deletedEvent.Local);
			Assert.Equal("t1", deletedEvent.Interval.Id);

			local.RegeneratePendingOps();

			Assert.Empty(sender.Sent);
		}

		[Fact]
		public void Transient_OnDeleteInterval_DoesNotFireOnEndpointDetach()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdef");
			IntervalCollection collection = sharedString.GetIntervalCollection("transient", IntervalType.Transient);
			var deleted = new List<IntervalDeletedEventArgs>();
			collection.OnDeleteInterval += (_, args) => deleted.Add(args);
			collection.Add(2, 4, intervalId: "t1");

			sharedString.DeleteText(2, 3);

			Assert.Empty(deleted);
			Assert.Null(collection.GetIntervalById("t1"));
			Assert.Empty(deleted);
		}

		private static SharedString CreateSharedStringWithText(string text)
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, text);
			return sharedString;
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

		private static void AssertIntervalPositions(SequenceInterval? interval, int expectedStart, int expectedEnd)
		{
			Assert.NotNull(interval);
			Assert.Equal(expectedStart, interval!.StartPosition);
			Assert.Equal(expectedEnd, interval.EndPosition);
		}

		private static void AssertIds(IEnumerable<SequenceInterval> intervals, params string[] expectedIds)
		{
			Assert.Equal(expectedIds, intervals.Select(interval => interval.Id).ToArray());
		}
	}
}
