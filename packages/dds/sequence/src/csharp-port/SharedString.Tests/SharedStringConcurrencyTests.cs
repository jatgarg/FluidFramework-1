// -----------------------------------------------------------------------------
// Wave 8 op reception + concurrency tests for SharedString POC.
//
// The CORRECTNESS-CRITICAL wave 8 file: proves two-client convergence
// for the merge-tree operational-transformation implementation.
// -----------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class SharedStringConcurrencyTests
	{
		[Fact]
		public void RemoteInsert_AppliesToState()
		{
			var sharedString = new SharedString();

			ProcessRemoteInsert(sharedString, 0, "hello", seq: 1);

			Assert.Equal("hello", sharedString.GetText());
			Assert.Equal(5, sharedString.GetLength());
		}

		[Fact]
		public void RemoteInsert_FiresEventWithLocalFalse()
		{
			var sharedString = new SharedString();
			var events = new List<SequenceDeltaEventArgs>();
			sharedString.OnSequenceDelta += (sender, e) => events.Add(e);

			ProcessRemoteInsert(sharedString, 0, "hello", seq: 1);

			SequenceDeltaEventArgs captured = Assert.Single(events);
			Assert.False(captured.Local);
			Assert.Equal("insert", captured.OpType);
			Assert.Equal(0, captured.Position);
			Assert.Equal(5, captured.Length);
			Assert.Equal("hello", captured.Text);
		}

		[Fact]
		public void RemoteDelete_AppliesToState()
		{
			var sharedString = CreateSharedStringWithAckedText("hello world");

			ProcessRemoteDelete(sharedString, 5, 6, refSeq: 1, seq: 2);

			Assert.Equal("helloworld", sharedString.GetText());
		}

		[Fact]
		public void RemoteDelete_FiresEventWithLocalFalse()
		{
			var sharedString = CreateSharedStringWithAckedText("hello world");
			var events = new List<SequenceDeltaEventArgs>();
			sharedString.OnSequenceDelta += (sender, e) => events.Add(e);

			ProcessRemoteDelete(sharedString, 5, 6, refSeq: 1, seq: 2);

			SequenceDeltaEventArgs captured = Assert.Single(events);
			Assert.False(captured.Local);
			Assert.Equal("remove", captured.OpType);
			Assert.Equal(5, captured.Position);
			Assert.Equal(1, captured.Length);
			Assert.Null(captured.Text);
		}

		[Fact]
		public void UnknownOpType_ThrowsOcsException()
		{
			var sharedString = new SharedString();

			OcsException exception = Assert.Throws<OcsException>(
				() => sharedString.ProcessDataObjectOp(RemoteMessage(refSeq: 0, seq: 1), "{\"type\":99,\"pos1\":0}"));
			Assert.Equal(OcsGateErrorCode.UnknownOp, exception.ErrorCode);
		}

		[Fact]
		public void OwnOpAck_DoesNotFireEvent()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			var events = new List<SequenceDeltaEventArgs>();
			sharedString.OnSequenceDelta += (eventSender, e) => events.Add(e);

			sharedString.InsertText(0, "hi");
			var sent = Assert.Single(sender.Sent);

			ProcessLocalAck(sharedString, sent, refSeq: 0, seq: 1);

			SequenceDeltaEventArgs captured = Assert.Single(events);
			Assert.True(captured.Local);
			Assert.Equal("insert", captured.OpType);
		}

		[Fact]
		public void OwnOpAck_DoesNotDoubleApplyState()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);

			sharedString.InsertText(0, "hello");
			var sent = Assert.Single(sender.Sent);

			ProcessLocalAck(sharedString, sent, refSeq: 0, seq: 1);

			Assert.Equal("hello", sharedString.GetText());
		}

		[Fact]
		public void MixedLocalRemoteEvents_FireInCorrectOrderWithCorrectLocalFlag()
		{
			var harness = new TwoClientHarness();
			var eventsA = new List<SequenceDeltaEventArgs>();
			var eventsB = new List<SequenceDeltaEventArgs>();
			harness.ClientA.OnSequenceDelta += (sender, e) => eventsA.Add(e);
			harness.ClientB.OnSequenceDelta += (sender, e) => eventsB.Add(e);

			harness.ClientA.InsertText(0, "hello");
			var sentA = Assert.Single(harness.SenderA.Sent);
			harness.DeliverAtoB(sentA, refSeq: 0);

			SequenceDeltaEventArgs eventA = Assert.Single(eventsA);
			SequenceDeltaEventArgs eventB = Assert.Single(eventsB);
			Assert.True(eventA.Local);
			Assert.Equal("insert", eventA.OpType);
			Assert.False(eventB.Local);
			Assert.Equal("insert", eventB.OpType);
			Assert.Equal("hello", eventB.Text);
		}

		[Fact]
		public void TwoClients_SequentialInserts_Converge()
		{
			var harness = new TwoClientHarness();

			harness.ClientA.InsertText(0, "hello");
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: 0);
			Assert.Equal("hello", harness.ClientA.GetText());
			Assert.Equal("hello", harness.ClientB.GetText());

			harness.SenderB.Sent.Clear();
			harness.ClientB.InsertText(5, " world");
			harness.DeliverBtoA(Assert.Single(harness.SenderB.Sent), refSeq: harness.CurrentServerSeq);

			Assert.Equal("hello world", harness.ClientA.GetText());
			Assert.Equal("hello world", harness.ClientB.GetText());
			AssertConverged(harness.ClientA, harness.ClientB);
		}

		[Fact]
		public void TwoClients_ConcurrentInserts_ConvergeToSameState()
		{
			var harness = new TwoClientHarness();
			LoadInitialSharedText(harness, "hello");

			harness.SenderA.Sent.Clear();
			harness.SenderB.Sent.Clear();
			long refSeq = harness.CurrentServerSeq;
			harness.ClientA.InsertText(5, " world");
			var sentA = Assert.Single(harness.SenderA.Sent);
			harness.ClientB.InsertText(5, "!");
			var sentB = Assert.Single(harness.SenderB.Sent);

			harness.DeliverAtoB(sentA, refSeq);
			harness.DeliverBtoA(sentB, refSeq);

			AssertConverged(harness.ClientA, harness.ClientB);
		}

		[Fact]
		public void TwoClients_ConcurrentDeletes_Converge()
		{
			var harness = new TwoClientHarness();
			LoadInitialSharedText(harness, "hello world");

			harness.SenderA.Sent.Clear();
			harness.SenderB.Sent.Clear();
			long refSeq = harness.CurrentServerSeq;
			harness.ClientA.DeleteText(0, 5);
			var sentA = Assert.Single(harness.SenderA.Sent);
			harness.ClientB.DeleteText(6, 11);
			var sentB = Assert.Single(harness.SenderB.Sent);

			harness.DeliverAtoB(sentA, refSeq);
			harness.DeliverBtoA(sentB, refSeq);

			AssertConverged(harness.ClientA, harness.ClientB);
		}

		[Fact]
		public void TwoClients_InsertVsDelete_Converge()
		{
			var harness = new TwoClientHarness();
			LoadInitialSharedText(harness, "hello");

			harness.SenderA.Sent.Clear();
			harness.SenderB.Sent.Clear();
			long refSeq = harness.CurrentServerSeq;
			harness.ClientA.InsertText(5, " world");
			var sentA = Assert.Single(harness.SenderA.Sent);
			harness.ClientB.DeleteText(0, 5);
			var sentB = Assert.Single(harness.SenderB.Sent);

			harness.DeliverAtoB(sentA, refSeq);
			harness.DeliverBtoA(sentB, refSeq);

			AssertConverged(harness.ClientA, harness.ClientB);
		}

		[Fact]
		public void TwoClients_ManyInterleaved_Converge()
		{
			var harness = new TwoClientHarness();

			harness.ClientA.InsertText(0, "a");
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: 0);
			harness.SenderA.Sent.Clear();

			harness.ClientB.InsertText(1, "b");
			harness.DeliverBtoA(Assert.Single(harness.SenderB.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderB.Sent.Clear();

			harness.ClientA.InsertText(0, "0");
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderA.Sent.Clear();

			harness.ClientB.DeleteText(1, 2);
			harness.DeliverBtoA(Assert.Single(harness.SenderB.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderB.Sent.Clear();

			harness.ClientA.InsertText(harness.ClientA.GetLength(), "X");
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderA.Sent.Clear();

			harness.ClientB.InsertText(1, "Y");
			harness.DeliverBtoA(Assert.Single(harness.SenderB.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderB.Sent.Clear();

			harness.ClientA.DeleteText(0, 1);
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderA.Sent.Clear();

			harness.ClientB.InsertText(harness.ClientB.GetLength(), "Z");
			harness.DeliverBtoA(Assert.Single(harness.SenderB.Sent), refSeq: harness.CurrentServerSeq);

			AssertConverged(harness.ClientA, harness.ClientB);
		}

		[Fact]
		public void TwoClients_FromEmpty_ManyInserts_Converge()
		{
			var harness = new TwoClientHarness();

			InsertAndDeliverAtoB(harness, 0, "A1");
			InsertAndDeliverBtoA(harness, harness.ClientB.GetLength(), "B1");
			InsertAndDeliverAtoB(harness, harness.ClientA.GetLength(), "A2");
			InsertAndDeliverBtoA(harness, harness.ClientB.GetLength(), "B2");
			InsertAndDeliverAtoB(harness, harness.ClientA.GetLength(), "A3");
			InsertAndDeliverBtoA(harness, harness.ClientB.GetLength(), "B3");

			AssertConverged(harness.ClientA, harness.ClientB);
			string finalText = harness.ClientA.GetText();
			Assert.Contains("A1", finalText);
			Assert.Contains("B1", finalText);
			Assert.Contains("A2", finalText);
			Assert.Contains("B2", finalText);
			Assert.Contains("A3", finalText);
			Assert.Contains("B3", finalText);
		}

		[Fact]
		public void LocalUnackedInsert_RemoteInsertBefore_BothPreserved()
		{
			var harness = new TwoClientHarness();

			harness.ClientA.InsertText(0, "hello");
			var sentA = Assert.Single(harness.SenderA.Sent);
			harness.ClientB.InsertText(0, "world");
			var sentB = Assert.Single(harness.SenderB.Sent);

			harness.DeliverBtoA(sentB, refSeq: 0);
			harness.DeliverAtoB(sentA, refSeq: 0);

			AssertConverged(harness.ClientA, harness.ClientB);
			string finalText = harness.ClientA.GetText();
			Assert.Contains("hello", finalText);
			Assert.Contains("world", finalText);
		}

		[Fact]
		public void LocalUnackedDelete_RemoteInsertInRange_HandlesCorrectly()
		{
			var harness = new TwoClientHarness();
			LoadInitialSharedText(harness, "abcdef");

			harness.SenderA.Sent.Clear();
			harness.SenderB.Sent.Clear();
			long refSeq = harness.CurrentServerSeq;
			harness.ClientA.DeleteText(2, 4);
			var sentA = Assert.Single(harness.SenderA.Sent);
			harness.ClientB.InsertText(3, "X");
			var sentB = Assert.Single(harness.SenderB.Sent);

			harness.DeliverBtoA(sentB, refSeq);
			harness.DeliverAtoB(sentA, refSeq);

			AssertConverged(harness.ClientA, harness.ClientB);
		}

		[Fact]
		public void LoadSnapshot_ThenLocalOps_WorksCorrectly()
		{
			var sharedString = new SharedString();
			sharedString.LoadFromSnapshot(LoadSimpleHelloSnapshot());

			sharedString.InsertText(13, "!");

			Assert.Equal("Hello, world!!", sharedString.GetText());
		}

		[Fact]
		public void LoadSnapshot_ThenRemoteOp_AppliesCorrectly()
		{
			var sharedString = new SharedString();
			var events = new List<SequenceDeltaEventArgs>();
			sharedString.LoadFromSnapshot(LoadSimpleHelloSnapshot());
			sharedString.OnSequenceDelta += (sender, e) => events.Add(e);

			ProcessRemoteInsert(sharedString, 13, "!", refSeq: 1, seq: 2);

			Assert.Equal("Hello, world!!", sharedString.GetText());
			SequenceDeltaEventArgs captured = Assert.Single(events);
			Assert.False(captured.Local);
			Assert.Equal("insert", captured.OpType);
			Assert.Equal(13, captured.Position);
			Assert.Equal("!", captured.Text);
		}

		private static SharedString CreateSharedStringWithAckedText(string text)
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			sharedString.InsertText(0, text);
			ProcessLocalAck(sharedString, Assert.Single(sender.Sent), refSeq: 0, seq: 1);
			return sharedString;
		}

		private static void LoadInitialSharedText(TwoClientHarness harness, string text)
		{
			harness.ClientA.InsertText(0, text);
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: 0);
			AssertConverged(harness.ClientA, harness.ClientB);
		}

		private static void InsertAndDeliverAtoB(TwoClientHarness harness, int position, string text)
		{
			harness.SenderA.Sent.Clear();
			harness.ClientA.InsertText(position, text);
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
		}

		private static void InsertAndDeliverBtoA(TwoClientHarness harness, int position, string text)
		{
			harness.SenderB.Sent.Clear();
			harness.ClientB.InsertText(position, text);
			harness.DeliverBtoA(Assert.Single(harness.SenderB.Sent), refSeq: harness.CurrentServerSeq);
		}

		private static void ProcessRemoteInsert(
			SharedString sharedString,
			int position,
			string text,
			long refSeq = 0,
			long seq = 1,
			string clientId = "remote-client")
		{
			string opJson = SharedStringOpSerializer.Serialize(new MergeTreeInsertMsg()
			{
				Pos1 = position,
				Seg = text,
			});

			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq, seq, clientId), opJson);
		}

		private static void ProcessRemoteDelete(
			SharedString sharedString,
			int start,
			int end,
			long refSeq,
			long seq,
			string clientId = "remote-client")
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

		private static SharedStringSnapshotDto LoadSimpleHelloSnapshot()
		{
			string fixturePath = Path.Combine(
				AppContext.BaseDirectory,
				"Fixtures",
				"simple-hello.snapshot.json.txt");
			return SharedStringSnapshotLoader.Parse(File.ReadAllText(fixturePath));
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
				ClientA = new SharedString("doc", SenderA);
				ClientB = new SharedString("doc", SenderB);
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
