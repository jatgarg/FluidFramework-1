// -----------------------------------------------------------------------------
// Sided obliterate tests for SharedString C# port.
// -----------------------------------------------------------------------------

#nullable enable

using System.Text.Json;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SharedStringSidedObliterateTests
	{
		[Fact]
		public void Sided_BasicObliterate_HidesRange()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");

			sharedString.ObliterateRange(SequencePlace.At(0, Side.Before), SequencePlace.At(6, Side.After));

			Assert.Equal("world", sharedString.GetText());
			Assert.Equal(5, sharedString.GetLength());
		}

		[Fact]
		public void Sided_ConcurrentInsertAtStart_SideBefore_NotEaten()
		{
			SharedString sharedString = CreateSidedObliteratedSharedString(Side.Before, Side.After);

			ProcessRemoteInsert(sharedString, 5, "S", refSeq: 5, seq: 11);

			Assert.Equal("01234SAB", sharedString.GetText());
		}

		[Fact]
		public void Sided_ConcurrentInsertAtStart_SideAfter_Eaten()
		{
			SharedString sharedString = CreateSidedObliteratedSharedString(Side.After, Side.After);

			ProcessRemoteInsert(sharedString, 5, "S", refSeq: 5, seq: 11);

			Assert.Equal("01234AB", sharedString.GetText());
		}

		[Fact]
		public void Sided_ConcurrentInsertAtEnd_SideAfter_Eaten()
		{
			SharedString sharedString = CreateSidedObliteratedSharedString(Side.Before, Side.After);

			ProcessRemoteInsert(sharedString, 10, "E", refSeq: 5, seq: 11);

			Assert.Equal("01234AB", sharedString.GetText());
		}

		[Fact]
		public void Sided_ConcurrentInsertAtEnd_SideBefore_NotEaten()
		{
			SharedString sharedString = CreateSidedObliteratedSharedString(Side.Before, Side.Before);

			ProcessRemoteInsert(sharedString, 10, "E", refSeq: 5, seq: 11);

			Assert.Equal("01234EAB", sharedString.GetText());
		}

		[Fact]
		public void Sided_ConcurrentInsertInMiddle_AlwaysEaten()
		{
			SharedString sharedString = CreateSidedObliteratedSharedString(Side.Before, Side.Before);

			ProcessRemoteInsert(sharedString, 7, "M", refSeq: 5, seq: 11);

			Assert.Equal("01234AB", sharedString.GetText());
		}

		[Fact]
		public void Sided_WireRoundTrip()
		{
			var op = new MergeTreeObliterateSidedMsg()
			{
				Pos1 = SequencePlace.At(0, Side.Before),
				Pos2 = SequencePlace.At(6, Side.After),
			};

			string json = SharedStringOpSerializer.Serialize(op);
			using JsonDocument document = JsonDocument.Parse(json);
			JsonElement root = document.RootElement;
			Assert.Equal(5, root.GetProperty("type").GetInt32());
			AssertPlace(root.GetProperty("pos1"), expectedPosition: 0, expectedBefore: true);
			AssertPlace(root.GetProperty("pos2"), expectedPosition: 6, expectedBefore: false);

			MergeTreeObliterateSidedMsg roundTripped = Assert.IsType<MergeTreeObliterateSidedMsg>(
				SharedStringOpSerializer.Deserialize(json));
			Assert.Equal(0, roundTripped.Pos1.Position);
			Assert.Equal(Side.Before, roundTripped.Pos1.Side);
			Assert.Equal(6, roundTripped.Pos2.Position);
			Assert.Equal(Side.After, roundTripped.Pos2.Side);

			SharedString sharedString = CreateSharedStringWithAckedText("hello world");
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq: 1, seq: 2), json);
			Assert.Equal("world", sharedString.GetText());
		}

		[Fact]
		public void NonSided_And_Sided_Coexist()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "abcdefghij");

			sharedString.ObliterateRange(1, 3);
			sharedString.ObliterateRange(SequencePlace.At(2, Side.Before), SequencePlace.At(5, Side.Before));

			Assert.Equal("adhij", sharedString.GetText());
		}

		[Fact]
		public void Sided_LocalPendingInsert_OnAck_EatenIfSideDictates()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			sharedString.InsertText(0, "0123456789AB");
			ProcessLocalAck(sharedString, sender.Sent[0], refSeq: 0, seq: 1);
			sender.Sent.Clear();
			sharedString.InsertText(5, "X");
			var pendingInsert = Assert.Single(sender.Sent);

			ProcessRemoteSidedObliterate(sharedString, 5, Side.After, 10, Side.Before, refSeq: 1, seq: 10);

			Assert.Equal("01234XAB", sharedString.GetText());
			ProcessLocalAck(sharedString, pendingInsert, refSeq: 1, seq: 11);
			Assert.Equal("01234AB", sharedString.GetText());
		}

		private static SharedString CreateSidedObliteratedSharedString(Side startSide, Side endSide)
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			sharedString.InsertText(0, "0123456789AB");
			ProcessLocalAck(sharedString, sender.Sent[0], refSeq: 0, seq: 1);
			sharedString.ObliterateRange(SequencePlace.At(5, startSide), SequencePlace.At(10, endSide));
			ProcessLocalAck(sharedString, sender.Sent[1], refSeq: 1, seq: 10);
			Assert.Equal("01234AB", sharedString.GetText());
			return sharedString;
		}

		private static SharedString CreateSharedStringWithAckedText(string text)
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("doc", sender);
			sharedString.InsertText(0, text);
			ProcessLocalAck(sharedString, sender.Sent[0], refSeq: 0, seq: 1);
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

		private static void ProcessRemoteSidedObliterate(
			SharedString sharedString,
			int start,
			Side startSide,
			int end,
			Side endSide,
			long refSeq,
			long seq,
			string clientId = "remote-client")
		{
			string opJson = $"{{\"type\":5,\"pos1\":{{\"pos\":{start},\"before\":{BoolJson(startSide == Side.Before)}}},\"pos2\":{{\"pos\":{end},\"before\":{BoolJson(endSide == Side.Before)}}}}}";
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

		private static void AssertPlace(JsonElement element, int expectedPosition, bool expectedBefore)
		{
			Assert.Equal(expectedPosition, element.GetProperty("pos").GetInt32());
			Assert.Equal(expectedBefore, element.GetProperty("before").GetBoolean());
		}

		private static string BoolJson(bool value)
		{
			return value ? "true" : "false";
		}
	}
}
