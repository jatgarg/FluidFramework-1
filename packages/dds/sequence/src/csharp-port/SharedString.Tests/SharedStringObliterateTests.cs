// -----------------------------------------------------------------------------
// Basic obliterate tests for SharedString C# port.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Text.Json;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SharedStringObliterateTests
	{
		[Fact]
		public void Local_Obliterate_HidesRange()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");

			sharedString.ObliterateRange(0, 6);

			Assert.Equal("world", sharedString.GetText());
			Assert.Equal(5, sharedString.GetLength());
		}

		[Fact]
		public void Local_Obliterate_EmitsWireOp()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("shared-string", sender);
			sharedString.InsertText(0, "hello world");

			sharedString.ObliterateRange(0, 6);

			Assert.Equal(2, sender.Sent.Count);
			var sent = sender.Sent[1];
			Assert.Equal("shared-string", sent.Address);
			Assert.Equal("obliterate", sent.OpTypeName);
			AssertObliterateWire(sent.OpJson, 0, 6);
		}

		[Fact]
		public void Remote_Obliterate_HidesRange()
		{
			SharedString sharedString = CreateSharedStringWithAckedText("hello world");

			ProcessRemoteObliterate(sharedString, 0, 6, refSeq: 1, seq: 2);

			Assert.Equal("world", sharedString.GetText());
			Assert.Equal(5, sharedString.GetLength());
		}

		[Fact]
		public void Local_Then_Remote_Ack_MatchesClientSeq()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			sharedString.InsertText(0, "hello world");
			ProcessLocalAck(sharedString, sender.Sent[0], refSeq: 0, seq: 1);
			ISegment leftSegment = GetRequiredContainingSegment(sharedString, 0);

			sharedString.ObliterateRange(0, 6);

			Assert.Equal(MergeTree.MergeTree.UnassignedSequenceNumber, leftSegment.ObliteratedSeq);
			Assert.Equal(sender.Sent[1].ClientSeq, leftSegment.ObliteratedLocalSeq);
			ProcessLocalAck(sharedString, sender.Sent[1], refSeq: 1, seq: 2);
			Assert.Equal(2, leftSegment.ObliteratedSeq);
			Assert.Null(leftSegment.ObliteratedLocalSeq);
			Assert.Equal("world", sharedString.GetText());
		}

		[Fact]
		public void Obliterate_ThenInsertAfter_DoesNotAffectInsertion()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");
			sharedString.ObliterateRange(0, 6);

			sharedString.InsertText(sharedString.GetLength(), "!");

			Assert.Equal("world!", sharedString.GetText());
		}

		[Fact]
		public void Obliterate_Nested_InsideRemoved_NoDoubleCount()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			sharedString.InsertText(0, "hello world");
			ProcessLocalAck(sharedString, sender.Sent[0], refSeq: 0, seq: 1);
			sharedString.DeleteText(0, 6);

			ProcessRemoteObliterate(sharedString, 0, 6, refSeq: 1, seq: 2);

			Assert.Equal("world", sharedString.GetText());
			Assert.Equal(5, sharedString.GetLength());
		}

		[Fact]
		public void Obliterate_EmptyRange_NoOp()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			sharedString.InsertText(0, "hello");

			sharedString.ObliterateRange(5, 5);

			Assert.Equal("hello", sharedString.GetText());
			Assert.Single(sender.Sent);
		}

		[Fact]
		public void Obliterate_WireRoundTrip()
		{
			var op = new MergeTreeObliterateMsg()
			{
				Pos1 = 0,
				Pos2 = 6,
			};

			string json = SharedStringOpSerializer.Serialize(op);
			MergeTreeObliterateMsg roundTripped = Assert.IsType<MergeTreeObliterateMsg>(
				SharedStringOpSerializer.Deserialize(json));
			var sharedString = CreateSharedStringWithAckedText("hello world");
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq: 1, seq: 2), json);

			AssertNullableInt(0, roundTripped.Pos1);
			AssertNullableInt(6, roundTripped.Pos2);
			AssertObliterateWire(json, 0, 6);
			Assert.Equal("world", sharedString.GetText());
		}

		[Fact]
		public void Obliterate_PositionAfter_Correct()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");
			SequenceInterval interval = sharedString.GetIntervalCollection("refs").Add(6, 6);

			sharedString.ObliterateRange(0, 6);

			Assert.Equal(0, sharedString.LocalReferencePositionToPosition(interval.Start));
			Assert.Equal("world", sharedString.GetText());
		}

		[Fact]
		public void ObliteratedRange_ConcurrentInsert_IsEaten()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			sharedString.InsertText(0, "hello world");
			ProcessLocalAck(sharedString, sender.Sent[0], refSeq: 0, seq: 1);
			sharedString.ObliterateRange(0, 6);
			ProcessLocalAck(sharedString, sender.Sent[1], refSeq: 1, seq: 10);

			ProcessRemoteInsert(sharedString, 3, "X", refSeq: 5, seq: 11);

			Assert.Equal("world", sharedString.GetText());
			Assert.Equal(5, sharedString.GetLength());
		}

		[Fact]
		public void RemoteObliterate_ThenLocalInsertAck_InsertEaten()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			sharedString.InsertText(0, "hello world");
			ProcessLocalAck(sharedString, sender.Sent[0], refSeq: 0, seq: 1);
			sender.Sent.Clear();
			sharedString.InsertText(3, "X");
			var pendingInsert = Assert.Single(sender.Sent);

			ProcessRemoteObliterate(sharedString, 0, 6, refSeq: 1, seq: 10);

			Assert.Equal("Xworld", sharedString.GetText());
			ProcessLocalAck(sharedString, pendingInsert, refSeq: 1, seq: 11);
			Assert.Equal("world", sharedString.GetText());
		}

		[Fact]
		public void ConcurrentInsertOutsideRange_NotEaten()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			sharedString.InsertText(0, "hello world");
			ProcessLocalAck(sharedString, sender.Sent[0], refSeq: 0, seq: 1);
			sharedString.ObliterateRange(0, 6);
			ProcessLocalAck(sharedString, sender.Sent[1], refSeq: 1, seq: 10);

			ProcessRemoteInsert(sharedString, 8, "X", refSeq: 5, seq: 11);

			Assert.Equal("woXrld", sharedString.GetText());
		}

		[Fact]
		public void ConcurrentInsertAtBoundary_NotEaten()
		{
			SharedString startBoundary = CreateObliteratedSharedString();
			ProcessRemoteInsert(startBoundary, 0, "S", refSeq: 5, seq: 11);
			Assert.Equal("Sworld", startBoundary.GetText());

			SharedString endBoundary = CreateObliteratedSharedString();
			ProcessRemoteInsert(endBoundary, 6, "E", refSeq: 5, seq: 11);
			Assert.Equal("Eworld", endBoundary.GetText());
		}

		[Fact]
		public void SequentialObliterates_Overlap()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			sharedString.InsertText(0, "hello ");
			ProcessLocalAck(sharedString, sender.Sent[0], refSeq: 0, seq: 1);
			sharedString.InsertText(6, "wo");
			ProcessLocalAck(sharedString, sender.Sent[1], refSeq: 1, seq: 2);
			sharedString.InsertText(8, "rl");
			ProcessLocalAck(sharedString, sender.Sent[2], refSeq: 2, seq: 3);
			sharedString.InsertText(10, "d");
			ProcessLocalAck(sharedString, sender.Sent[3], refSeq: 3, seq: 4);
			ISegment helloSegment = GetRequiredContainingSegment(sharedString, 0);

			ProcessRemoteObliterate(sharedString, 0, 6, refSeq: 1, seq: 10);
			ISegment overlappedSegment = GetRequiredContainingSegment(sharedString, 2);

			ProcessRemoteObliterate(sharedString, 2, 4, refSeq: 10, seq: 11);

			Assert.Equal(10, helloSegment.ObliteratedSeq);
			Assert.Equal(11, overlappedSegment.ObliteratedSeq);
			Assert.Equal("wod", sharedString.GetText());
		}

		[Fact]
		public void TwoClientConvergence_ObliterateVsInsert()
		{
			var harness = new TwoClientHarness();
			LoadInitialSharedText(harness, "hello world");
			harness.SenderA.Sent.Clear();
			harness.SenderB.Sent.Clear();
			long refSeq = harness.CurrentServerSeq;

			harness.ClientA.ObliterateRange(0, 6);
			var obliterate = Assert.Single(harness.SenderA.Sent);
			harness.ClientB.InsertText(3, "X");
			var insert = Assert.Single(harness.SenderB.Sent);

			harness.DeliverAtoB(obliterate, refSeq);
			harness.DeliverBtoA(insert, refSeq);

			Assert.Equal("world", harness.ClientA.GetText());
			Assert.Equal("world", harness.ClientB.GetText());
			AssertConverged(harness.ClientA, harness.ClientB);
		}

		[Fact]
		public void Reconnect_PendingObliterate_ReplayedCorrectly()
		{
			var sender = new FakeFluidDataObjectSender();
			var local = new SharedString("client-a", sender);
			local.InsertText(0, "hello world");
			ProcessLocalAck(local, sender.Sent[0], refSeq: 0, seq: 1);
			SharedString serverVisible = CreateSharedStringWithAckedText("hello world");

			local.ObliterateRange(0, 6);
			var replayed = sender.Sent[1];
			serverVisible.ProcessDataObjectOp(RemoteMessage(refSeq: 1, seq: 2, clientId: "client-a"), replayed.OpJson);
			ProcessLocalAck(local, replayed, refSeq: 1, seq: 2);

			Assert.Equal("world", local.GetText());
			Assert.Equal("world", serverVisible.GetText());
		}

		[Fact]
		public void Reconnect_PendingInsertIntoRemoteObliterate_InsertEaten()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("client-a", sender);
			sharedString.InsertText(0, "hello world");
			ProcessLocalAck(sharedString, sender.Sent[0], refSeq: 0, seq: 1);
			sender.Sent.Clear();
			sharedString.InsertText(3, "X");
			var pendingInsert = Assert.Single(sender.Sent);

			ProcessRemoteObliterate(sharedString, 0, 6, refSeq: 1, seq: 10, clientId: "client-b");
			ProcessLocalAck(sharedString, pendingInsert, refSeq: 1, seq: 11);

			Assert.Equal("world", sharedString.GetText());
		}

		private static SharedString CreateSharedStringWithAckedText(string text)
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			sharedString.InsertText(0, text);
			ProcessLocalAck(sharedString, sender.Sent[0], refSeq: 0, seq: 1);
			return sharedString;
		}

		private static SharedString CreateObliteratedSharedString()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			sharedString.InsertText(0, "hello world");
			ProcessLocalAck(sharedString, sender.Sent[0], refSeq: 0, seq: 1);
			sharedString.ObliterateRange(0, 6);
			ProcessLocalAck(sharedString, sender.Sent[1], refSeq: 1, seq: 10);
			return sharedString;
		}

		private static void ProcessRemoteInsert(
			SharedString sharedString,
			int position,
			string text,
			long refSeq,
			long seq,
			string clientId = "remote-client")
		{
			string opJson = SharedStringOpSerializer.Serialize(new MergeTreeInsertMsg()
			{
				Pos1 = position,
				Seg = text,
			});
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq, seq, clientId), opJson);
		}

		private static void ProcessRemoteObliterate(
			SharedString sharedString,
			int start,
			int end,
			long refSeq,
			long seq,
			string clientId = "remote-client")
		{
			string opJson = $"{{\"type\":\"obliterate\",\"pos1\":{start},\"pos2\":{end}}}";
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
			string clientId = "remote-client")
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

		private static void AssertObliterateWire(string json, int expectedStart, int expectedEnd)
		{
			using JsonDocument document = JsonDocument.Parse(json);
			JsonElement root = document.RootElement;
			Assert.Equal(4, root.GetProperty("type").GetInt32());
			Assert.Equal(expectedStart, root.GetProperty("pos1").GetInt32());
			Assert.Equal(expectedEnd, root.GetProperty("pos2").GetInt32());
			MergeTreeObliterateMsg op = Assert.IsType<MergeTreeObliterateMsg>(
				SharedStringOpSerializer.Deserialize(json));
			AssertNullableInt(expectedStart, op.Pos1);
			AssertNullableInt(expectedEnd, op.Pos2);
		}

		private static void AssertNullableInt(int expected, int? actual)
		{
			Assert.True(actual.HasValue);
			Assert.Equal(expected, actual.Value);
		}

		private static void LoadInitialSharedText(TwoClientHarness harness, string text)
		{
			harness.ClientA.InsertText(0, text);
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: 0);
			AssertConverged(harness.ClientA, harness.ClientB);
		}

		private static void AssertConverged(SharedString clientA, SharedString clientB)
		{
			string textA = clientA.GetText();
			string textB = clientB.GetText();
			Assert.True(
				string.Equals(textA, textB, StringComparison.Ordinal),
				$"Expected clients to converge. Client A: '{textA}'. Client B: '{textB}'.");
		}

		private sealed class TwoClientHarness
		{
			private const string _clientAId = "client-a";
			private const string _clientBId = "client-b";
			private long _serverSeq;

			public TwoClientHarness()
			{
				SenderA = new FakeFluidDataObjectSender();
				SenderB = new FakeFluidDataObjectSender();
				ClientA = new SharedString(_clientAId, SenderA);
				ClientB = new SharedString(_clientBId, SenderB);
			}

			public SharedString ClientA { get; }

			public SharedString ClientB { get; }

			public FakeFluidDataObjectSender SenderA { get; }

			public FakeFluidDataObjectSender SenderB { get; }

			public long CurrentServerSeq => _serverSeq;

			public void DeliverAtoB(
				(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
				long refSeq)
			{
				Deliver(sent, fromClientId: _clientAId, from: ClientA, to: ClientB, refSeq: refSeq);
			}

			public void DeliverBtoA(
				(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
				long refSeq)
			{
				Deliver(sent, fromClientId: _clientBId, from: ClientB, to: ClientA, refSeq: refSeq);
			}

			private void Deliver(
				(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
				string fromClientId,
				SharedString from,
				SharedString to,
				long refSeq)
			{
				long serverSeq = ++_serverSeq;
				var descriptorForTo = new SequencedDocumentMessageDescriptor(
					SequenceNumber.ForTesting(clientSeq: sent.ClientSeq, refSeq: refSeq, seq: serverSeq),
					OpOrigin.Remote,
					fromClientId);
				to.ProcessDataObjectOp(descriptorForTo, sent.OpJson);

				var descriptorForFrom = new SequencedDocumentMessageDescriptor(
					SequenceNumber.ForTesting(clientSeq: sent.ClientSeq, refSeq: refSeq, seq: serverSeq),
					OpOrigin.Local,
					fromClientId);
				from.ProcessDataObjectOp(descriptorForFrom, sent.OpJson);
			}
		}
	}
}
