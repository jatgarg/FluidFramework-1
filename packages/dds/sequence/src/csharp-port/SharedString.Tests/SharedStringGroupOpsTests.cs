// -----------------------------------------------------------------------------
// Group op tests for SharedString C# port.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;
using System.Text.Json;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SharedStringGroupOpsTests
	{
		[Fact]
		public void RunInBatch_TwoInserts_EmitsSingleGroupOp()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("shared-string", sender);

			sharedString.RunInBatch(() =>
			{
				sharedString.InsertText(0, "hello");
				sharedString.InsertText(5, " world");
			});

			var sent = Assert.Single(sender.Sent);
			Assert.Equal("shared-string", sent.Address);
			Assert.Equal("group", sent.OpTypeName);
			MergeTreeGroupMsg groupOp = AssertGroupWire(sent.OpJson, expectedCount: 2);
			Assert.All(groupOp.Ops, op => Assert.Equal(MergeTreeDeltaType.Insert, op.Type));
			Assert.Equal("hello world", sharedString.GetText());
		}

		[Fact]
		public void RunInBatch_SingleInsert_EmitsPlainOp()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("shared-string", sender);

			sharedString.RunInBatch(() => sharedString.InsertText(0, "hello"));

			var sent = Assert.Single(sender.Sent);
			Assert.Equal("insert", sent.OpTypeName);
			using JsonDocument document = JsonDocument.Parse(sent.OpJson);
			Assert.Equal(0, document.RootElement.GetProperty("type").GetInt32());
			Assert.IsType<MergeTreeInsertMsg>(SharedStringOpSerializer.Deserialize(sent.OpJson));
		}

		[Fact]
		public void RunInBatch_Empty_EmitsNothing()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("shared-string", sender);

			sharedString.RunInBatch(() =>
			{
			});

			Assert.Empty(sender.Sent);
		}

		[Fact]
		public void RunInBatch_Mixed_InsertRemoveAnnotate_ApplyInOrder()
		{
			SharedString sharedString = CreateSharedStringWithAckedText("hello world");

			sharedString.RunInBatch(() =>
			{
				sharedString.InsertText(5, " there");
				sharedString.DeleteText(0, 5);
				sharedString.AnnotateRange(0, 6, new PropertySet()
				{
					["bold"] = true,
				});
			});

			Assert.Equal(" there world", sharedString.GetText());
			PropertySet? properties = sharedString.GetPropertiesAtPosition(0);
			Assert.NotNull(properties);
			Assert.True(Assert.IsType<bool>(properties["bold"]));
		}

		[Fact]
		public void RunInBatch_Nested_Throws()
		{
			var sharedString = new SharedString();

			Assert.Throws<LoggingError>(() =>
				sharedString.RunInBatch(() =>
					sharedString.RunInBatch(() => sharedString.InsertText(0, "nope"))));
		}

		[Fact]
		public void Remote_GroupOp_AppliedInOrder()
		{
			var sharedString = new SharedString();
			const string groupJson = "{\"type\":\"group\",\"ops\":[{\"type\":\"insert\",\"pos1\":0,\"seg\":\"hello\"},{\"type\":\"insert\",\"pos1\":5,\"seg\":\" world\"}]}";

			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq: 0, seq: 1), groupJson);

			Assert.Equal("hello world", sharedString.GetText());
		}

		[Fact]
		public void GroupOp_WireRoundTrip()
		{
			MergeTreeGroupMsg groupOp = new();
			groupOp.Ops.Add(new MergeTreeInsertMsg()
			{
				Pos1 = 0,
				Seg = "hello",
			});
			groupOp.Ops.Add(new MergeTreeInsertMsg()
			{
				Pos1 = 5,
				Seg = " world",
			});

			string json = SharedStringOpSerializer.Serialize(groupOp);
			MergeTreeGroupMsg roundTripped = AssertGroupWire(json, expectedCount: 2);
			var sharedString = new SharedString();
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq: 0, seq: 1), json);

			Assert.Equal(MergeTreeDeltaType.Insert, roundTripped.Ops[0].Type);
			Assert.Equal(MergeTreeDeltaType.Insert, roundTripped.Ops[1].Type);
			Assert.Equal("hello world", sharedString.GetText());
		}

		[Fact]
		public void GroupOp_LocalAck_PromotesAllSubOpStamps()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			ISegment? firstSegment = null;
			ISegment? secondSegment = null;

			sharedString.RunInBatch(() =>
			{
				sharedString.InsertText(0, "a");
				firstSegment = GetRequiredContainingSegment(sharedString, 0);
				sharedString.InsertText(1, "b");
				secondSegment = GetRequiredContainingSegment(sharedString, 1);
				sharedString.DeleteText(0, 1);
			});

			var sent = Assert.Single(sender.Sent);
			AssertGroupWire(sent.OpJson, expectedCount: 3);
			Assert.NotNull(firstSegment);
			Assert.NotNull(secondSegment);
			Assert.Equal(MergeTree.MergeTree.UnassignedSequenceNumber, firstSegment!.Seq);
			Assert.Equal(MergeTree.MergeTree.UnassignedSequenceNumber, secondSegment!.Seq);
			Assert.Equal(MergeTree.MergeTree.UnassignedSequenceNumber, firstSegment.RemovedSeq);

			ProcessLocalAck(sharedString, sent, refSeq: 0, seq: 10);

			Assert.Equal(10, firstSegment.Seq);
			Assert.Equal(10, secondSegment.Seq);
			Assert.Equal(10, firstSegment.RemovedSeq);
			Assert.Null(firstSegment.LocalSeq);
			Assert.Null(secondSegment.LocalSeq);
			Assert.Null(firstSegment.RemovedLocalSeq);
			Assert.Equal("b", sharedString.GetText());
		}

		[Fact]
		public void GroupOp_TwoClientConvergence()
		{
			var harness = new TwoClientHarness();
			long refSeq = harness.CurrentServerSeq;

			harness.ClientA.RunInBatch(() =>
			{
				harness.ClientA.InsertText(0, "A");
				harness.ClientA.InsertText(1, "B");
			});
			var group = Assert.Single(harness.SenderA.Sent);
			harness.ClientB.InsertText(0, "x");
			var insert = Assert.Single(harness.SenderB.Sent);

			harness.DeliverAtoB(group, refSeq);
			harness.DeliverBtoA(insert, refSeq);

			AssertConverged(harness.ClientA, harness.ClientB);
		}

		[Fact]
		public void GroupOp_WithObliterate_Concurrent()
		{
			var harness = new TwoClientHarness();
			LoadInitialSharedText(harness, "hello world");
			harness.SenderA.Sent.Clear();
			harness.SenderB.Sent.Clear();
			long refSeq = harness.CurrentServerSeq;

			harness.ClientA.RunInBatch(() =>
			{
				harness.ClientA.InsertText(11, "!");
				harness.ClientA.ObliterateRange(0, 6);
			});
			var group = Assert.Single(harness.SenderA.Sent);
			harness.ClientB.InsertText(3, "X");
			var insert = Assert.Single(harness.SenderB.Sent);

			harness.DeliverAtoB(group, refSeq);
			harness.DeliverBtoA(insert, refSeq);

			Assert.Equal("world!", harness.ClientA.GetText());
			Assert.Equal("world!", harness.ClientB.GetText());
			AssertConverged(harness.ClientA, harness.ClientB);
		}

		[Fact]
		public void RunInBatch_TextThenInterval_PreservesLocalOrderOnWire()
		{
			// regression. TS runtime queues submitLocalMessage calls in
			// call order (both merge-tree ops via sequence.ts submitDelta and
			// interval ops via intervalCollectionMap.submitMessage go through
			// the same runtime queue). The port must not send interval ops
			// while merge-tree ops in the same batch are still buffered —
			// peers would see the interval endpoint before the text it
			// anchors to.
			SharedString sharedString = CreateSharedStringWithAckedText("abc");
			FakeFluidDataObjectSender sender = GetSenderFor(sharedString);
			sender.Sent.Clear();

			sharedString.RunInBatch(() =>
			{
				sharedString.InsertText(3, "XYZ");
				sharedString.GetIntervalCollection("comments").Add(3, 6, intervalId: "i1");
			});

			Assert.Equal(2, sender.Sent.Count);
			IMergeTreeOp first = SharedStringOpSerializer.Deserialize(sender.Sent[0].OpJson);
			IMergeTreeOp second = SharedStringOpSerializer.Deserialize(sender.Sent[1].OpJson);
			Assert.IsType<MergeTreeInsertMsg>(first);
			Assert.IsType<IntervalAddOpMsg>(second);
		}

		[Fact]
		public void RunInBatch_IntervalBetweenInserts_SplitsGroupPreservingOrder()
		{
			// regression. Interval ops in the middle of a batch split
			// any group of merge-tree ops around them so wire order matches
			// local order: insert, interval, insert.
			SharedString sharedString = CreateSharedStringWithAckedText("abc");
			FakeFluidDataObjectSender sender = GetSenderFor(sharedString);
			sender.Sent.Clear();

			sharedString.RunInBatch(() =>
			{
				sharedString.InsertText(3, "X");
				sharedString.GetIntervalCollection("comments").Add(0, 4, intervalId: "i1");
				sharedString.InsertText(4, "Y");
			});

			Assert.Equal(3, sender.Sent.Count);
			Assert.IsType<MergeTreeInsertMsg>(SharedStringOpSerializer.Deserialize(sender.Sent[0].OpJson));
			Assert.IsType<IntervalAddOpMsg>(SharedStringOpSerializer.Deserialize(sender.Sent[1].OpJson));
			Assert.IsType<MergeTreeInsertMsg>(SharedStringOpSerializer.Deserialize(sender.Sent[2].OpJson));
		}

		[Fact]
		public void RunInBatch_TwoMergeTreeOpsAfterInterval_StillGroupsTail()
		{
			// regression. Batching still coalesces adjacent merge-tree
			// ops into a single group. An interval op breaks the group only
			// at its position — subsequent merge-tree ops start a fresh
			// group.
			SharedString sharedString = CreateSharedStringWithAckedText("abc");
			FakeFluidDataObjectSender sender = GetSenderFor(sharedString);
			sender.Sent.Clear();

			sharedString.RunInBatch(() =>
			{
				sharedString.GetIntervalCollection("comments").Add(0, 3, intervalId: "i1");
				sharedString.InsertText(3, "X");
				sharedString.InsertText(4, "Y");
			});

			Assert.Equal(2, sender.Sent.Count);
			Assert.IsType<IntervalAddOpMsg>(SharedStringOpSerializer.Deserialize(sender.Sent[0].OpJson));
			MergeTreeGroupMsg group = AssertGroupWire(sender.Sent[1].OpJson, expectedCount: 2);
			Assert.All(group.Ops, op => Assert.Equal(MergeTreeDeltaType.Insert, op.Type));
		}

		private static FakeFluidDataObjectSender GetSenderFor(SharedString sharedString)
		{
			// The test SharedString was created via CreateSharedStringWithAckedText,
			// which passes a FakeFluidDataObjectSender we can retrieve via
			// the private _sender field. Reflection-free approach: re-derive
			// via a known channel by creating a fresh sender-inspection scheme.
			// Simpler: use reflection here — this is test-only glue.
			System.Reflection.FieldInfo? field = typeof(SharedString).GetField(
				"_sender",
				System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
			return (FakeFluidDataObjectSender)(field!.GetValue(sharedString)!);
		}

		private static SharedString CreateSharedStringWithAckedText(string text)
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			sharedString.InsertText(0, text);
			ProcessLocalAck(sharedString, sender.Sent[0], refSeq: 0, seq: 1);
			sender.Sent.Clear();
			return sharedString;
		}

		private static void LoadInitialSharedText(TwoClientHarness harness, string text)
		{
			harness.ClientA.InsertText(0, text);
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: 0);
			AssertConverged(harness.ClientA, harness.ClientB);
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

		private static MergeTreeGroupMsg AssertGroupWire(string json, int expectedCount)
		{
			using JsonDocument document = JsonDocument.Parse(json);
			JsonElement root = document.RootElement;
			Assert.Equal(3, root.GetProperty("type").GetInt32());
			JsonElement ops = root.GetProperty("ops");
			Assert.Equal(JsonValueKind.Array, ops.ValueKind);
			Assert.Equal(expectedCount, ops.GetArrayLength());
			MergeTreeGroupMsg groupOp = Assert.IsType<MergeTreeGroupMsg>(
				SharedStringOpSerializer.Deserialize(json));
			Assert.Equal(expectedCount, groupOp.Ops.Count);
			return groupOp;
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
