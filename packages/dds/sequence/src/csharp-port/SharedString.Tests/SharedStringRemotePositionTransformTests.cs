// -----------------------------------------------------------------------------
// Remote SequenceDelta position transform tests for SharedString C# port.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SharedStringRemotePositionTransformTests
	{
		[Fact]
		public void RemoteInsert_WithLocalPendingInsertBefore_EventFiresTransformedPos()
		{
			SharedString sharedString = CreateSharedStringWithAckedText("abcdefghij", out FakeFluidDataObjectSender sender);
			sharedString.InsertText(2, "XXX");
			List<SequenceDeltaEventArgs> events = CaptureEvents(sharedString);

			ProcessRemoteInsert(sharedString, 5, "R", refSeq: 1, seq: 2);

			SequenceDeltaEventArgs captured = AssertSingleRemoteEvent(events, "insert");
			Assert.Equal(8, captured.Position);
			Assert.Equal("R", captured.Text);
			AssertRange(Assert.Single(captured.Ranges), 8, 1);
			Assert.Equal("abXXXcdeRfghij", sharedString.GetText());
			Assert.Single(sender.Sent);
		}

		[Fact]
		public void RemoteInsert_WithLocalPendingInsertAfterRemote_UnaffectedByLocal()
		{
			SharedString sharedString = CreateSharedStringWithAckedText("abcdefghij", out _);
			sharedString.InsertText(8, "XXX");
			List<SequenceDeltaEventArgs> events = CaptureEvents(sharedString);

			ProcessRemoteInsert(sharedString, 3, "R", refSeq: 1, seq: 2);

			SequenceDeltaEventArgs captured = AssertSingleRemoteEvent(events, "insert");
			Assert.Equal(3, captured.Position);
			AssertRange(Assert.Single(captured.Ranges), 3, 1);
			Assert.Equal("abcRdefghXXXij", sharedString.GetText());
		}

		[Fact]
		public void RemoteRemove_AfterLocalPendingInsert_PositionShifted()
		{
			SharedString sharedString = CreateSharedStringWithAckedText("abcdefghij", out _);
			sharedString.InsertText(5, "XXX");
			List<SequenceDeltaEventArgs> events = CaptureEvents(sharedString);

			ProcessRemoteRemove(sharedString, 6, 9, refSeq: 1, seq: 2);

			SequenceDeltaEventArgs captured = AssertSingleRemoteEvent(events, "remove");
			Assert.Equal(9, captured.Position);
			Assert.Equal(3, captured.Length);
			AssertRange(Assert.Single(captured.Ranges), 9, 3);
			Assert.Equal("abcdeXXXfj", sharedString.GetText());
		}

		[Fact]
		public void RemoteInsert_InGroupOp_EachSubOpTransformed()
		{
			SharedString sharedString = CreateSharedStringWithAckedText("abcdef", out _);
			sharedString.InsertText(2, "LL");
			List<SequenceDeltaEventArgs> events = CaptureEvents(sharedString);
			MergeTreeGroupMsg group = new();
			group.Ops.Add(new MergeTreeInsertMsg() { Pos1 = 4, Seg = "X" });
			group.Ops.Add(new MergeTreeInsertMsg() { Pos1 = 1, Seg = "Y" });
			group.Ops.Add(new MergeTreeInsertMsg() { Pos1 = 8, Seg = "Z" });

			ProcessRemoteOp(sharedString, group, refSeq: 1, seq: 2);

			Assert.Collection(
				events,
				first =>
				{
					Assert.False(first.Local);
					Assert.Equal("insert", first.OpType);
					Assert.Equal(6, first.Position);
					Assert.Equal("X", first.Text);
				},
				second =>
				{
					Assert.False(second.Local);
					Assert.Equal("insert", second.OpType);
					Assert.Equal(1, second.Position);
					Assert.Equal("Y", second.Text);
				},
				third =>
				{
					Assert.False(third.Local);
					Assert.Equal("insert", third.OpType);
					Assert.Equal(10, third.Position);
					Assert.Equal("Z", third.Text);
				});
			Assert.Equal("aYbLLcdXefZ", sharedString.GetText());
		}

		[Fact]
		public void RemoteAnnotate_AfterLocalRemove_RangePartiallyShifted()
		{
			SharedString sharedString = CreateSharedStringWithAckedText("abcdefghij", out _);
			sharedString.DeleteText(0, 3);
			List<SequenceDeltaEventArgs> events = CaptureEvents(sharedString);

			ProcessRemoteAnnotate(
				sharedString,
				2,
				8,
				new PropertySet() { ["color"] = "red" },
				refSeq: 1,
				seq: 2);

			SequenceDeltaEventArgs captured = AssertSingleRemoteEvent(events, "annotate");
			Assert.Equal(0, captured.Position);
			Assert.Equal(5, captured.Length);
			AssertRange(Assert.Single(captured.Ranges), 0, 5);
			Assert.Equal("red", Assert.IsType<string>(GetRequiredProperty(sharedString, 0, "color")));
			Assert.Equal("red", Assert.IsType<string>(GetRequiredProperty(sharedString, 4, "color")));
			Assert.Null(sharedString.GetPropertiesAtPosition(5));
		}

		[Fact]
		public void RemoteObliterate_TransformedForLocalPendingInserts()
		{
			SharedString sharedString = CreateSharedStringWithAckedText("hello world", out FakeFluidDataObjectSender sender);
			sharedString.InsertText(3, "X");
			var pendingInsert = Assert.Single(sender.Sent);
			List<SequenceDeltaEventArgs> events = CaptureEvents(sharedString);

			ProcessRemoteObliterate(sharedString, 0, 6, refSeq: 1, seq: 10);

			SequenceDeltaEventArgs captured = AssertSingleRemoteEvent(events, "obliterate");
			Assert.Equal(0, captured.Position);
			Assert.Equal(6, captured.Length);
			Assert.Collection(
				captured.Ranges,
				first => AssertRange(first, 0, 3),
				second => AssertRange(second, 1, 3));
			Assert.Equal("Xworld", sharedString.GetText());

			ProcessLocalAck(sharedString, pendingInsert, refSeq: 1, seq: 11);

			Assert.Equal("world", sharedString.GetText());
		}

		[Fact]
		public void SequenceDelta_TransformedPos_ConsistentWith_GetText()
		{
			SharedString sharedString = CreateSharedStringWithAckedText("abcdefghij", out _);
			sharedString.InsertText(2, "LL");
			List<SequenceDeltaEventArgs> events = CaptureEvents(sharedString);

			ProcessRemoteAnnotate(
				sharedString,
				4,
				8,
				new PropertySet() { ["style"] = "remote" },
				refSeq: 1,
				seq: 2);

			SequenceDeltaEventArgs captured = AssertSingleRemoteEvent(events, "annotate");
			Assert.Equal(6, captured.Position);
			Assert.Equal(4, captured.Length);
			Assert.Equal("efgh", sharedString.GetText(captured.Position, captured.Position + captured.Length));
		}

		[Fact]
		public void NoLocalPending_RemoteOpPositionUnchanged()
		{
			SharedString sharedString = CreateSharedStringWithAckedText("abcdefghij", out _);
			List<SequenceDeltaEventArgs> events = CaptureEvents(sharedString);

			ProcessRemoteRemove(sharedString, 2, 5, refSeq: 1, seq: 2);

			SequenceDeltaEventArgs captured = AssertSingleRemoteEvent(events, "remove");
			Assert.Equal(2, captured.Position);
			Assert.Equal(3, captured.Length);
			AssertRange(Assert.Single(captured.Ranges), 2, 3);
			Assert.Equal("abfghij", sharedString.GetText());
		}

		private static SharedString CreateSharedStringWithAckedText(string text, out FakeFluidDataObjectSender sender)
		{
			sender = new FakeFluidDataObjectSender();
			SharedString sharedString = new("doc", sender);
			sharedString.InsertText(0, text);
			ProcessLocalAck(sharedString, Assert.Single(sender.Sent), refSeq: 0, seq: 1);
			sender.Sent.Clear();
			return sharedString;
		}

		private static List<SequenceDeltaEventArgs> CaptureEvents(SharedString sharedString)
		{
			List<SequenceDeltaEventArgs> events = new();
			sharedString.OnSequenceDelta += (sender, args) => events.Add(args);
			return events;
		}

		private static void ProcessRemoteInsert(
			SharedString sharedString,
			int position,
			string text,
			long refSeq,
			long seq)
		{
			ProcessRemoteOp(
				sharedString,
				new MergeTreeInsertMsg()
				{
					Pos1 = position,
					Seg = text,
				},
				refSeq,
				seq);
		}

		private static void ProcessRemoteRemove(
			SharedString sharedString,
			int start,
			int end,
			long refSeq,
			long seq)
		{
			ProcessRemoteOp(
				sharedString,
				new MergeTreeRemoveMsg()
				{
					Pos1 = start,
					Pos2 = end,
				},
				refSeq,
				seq);
		}

		private static void ProcessRemoteObliterate(
			SharedString sharedString,
			int start,
			int end,
			long refSeq,
			long seq)
		{
			ProcessRemoteOp(
				sharedString,
				new MergeTreeObliterateMsg()
				{
					Pos1 = start,
					Pos2 = end,
				},
				refSeq,
				seq);
		}

		private static void ProcessRemoteAnnotate(
			SharedString sharedString,
			int start,
			int end,
			PropertySet props,
			long refSeq,
			long seq)
		{
			ProcessRemoteOp(
				sharedString,
				new MergeTreeAnnotateMsg()
				{
					Pos1 = start,
					Pos2 = end,
					Props = props,
				},
				refSeq,
				seq);
		}

		private static void ProcessRemoteOp(
			SharedString sharedString,
			IMergeTreeOp op,
			long refSeq,
			long seq,
			string clientId = "remote-client")
		{
			sharedString.ProcessDataObjectOp(
				RemoteMessage(refSeq, seq, clientId),
				SharedStringOpSerializer.Serialize(op));
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

		private static SequenceDeltaEventArgs AssertSingleRemoteEvent(
			IReadOnlyList<SequenceDeltaEventArgs> events,
			string opType)
		{
			SequenceDeltaEventArgs captured = Assert.Single(events);
			Assert.False(captured.Local);
			Assert.Equal(opType, captured.OpType);
			return captured;
		}

		private static void AssertRange(SequenceDeltaRange range, int expectedPosition, int expectedLength)
		{
			Assert.Equal(expectedPosition, range.Position);
			Assert.Equal(expectedLength, range.Length);
		}

		private static object? GetRequiredProperty(SharedString sharedString, int position, string key)
		{
			PropertySet? properties = sharedString.GetPropertiesAtPosition(position);
			Assert.NotNull(properties);
			Assert.True(properties!.ContainsKey(key), $"Expected property '{key}' at position {position}.");
			return properties[key];
		}
	}
}
