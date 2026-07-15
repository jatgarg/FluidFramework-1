#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;
using TextSegmentModel = Microsoft.Office.Web.Fluid.MergeTree.TextSegment;

namespace Microsoft.Office.Web.Fluid.Tests
{
	internal static class PortedTestUtilities
	{
		public static SharedString CreateSharedStringWithText(string text)
		{
			SharedString sharedString = new();
			if (text.Length > 0)
			{
				sharedString.InsertText(0, text);
			}

			return sharedString;
		}

		public static (SharedString SharedString, FakeFluidDataObjectSender Sender) CreateAckedSharedString(string clientId, string text)
		{
			FakeFluidDataObjectSender sender = new();
			SharedString sharedString = new(clientId, sender);
			if (text.Length > 0)
			{
				sharedString.InsertText(0, text);
				ProcessLocalAck(sharedString, Assert.Single(sender.Sent), refSeq: 0, seq: 1);
				sender.Sent.Clear();
			}

			return (sharedString, sender);
		}

		public static MergeTreeModel CreateAckedTree(string text, long seq = 1, string clientId = "seed")
		{
			MergeTreeModel tree = new();
			if (text.Length > 0)
			{
				tree.InsertSegments(0, SegmentArray(text), MergeTreeModel.UnassignedSequenceNumber, seq, clientId);
			}

			return tree;
		}

		public static MergeTreeModel CreateSingleCharacterTree(string text)
		{
			MergeTreeModel tree = new();
			long seq = 1;
			for (int i = 0; i < text.Length; i++)
			{
				tree.InsertSegments(i, SegmentArray(text[i].ToString()), refSeq: seq - 1, seq: seq, clientId: "seed");
				seq++;
			}

			return tree;
		}

		public static ISegment[] SegmentArray(string text, PropertySet? props = null)
		{
			return new ISegment[] { TextSegmentModel.Make(text, props) };
		}

		public static ISegment SegmentWithText(MergeTreeModel tree, string text)
		{
			return Assert.Single(
				tree.WalkAllSegments(),
				segment => segment is TextSegmentModel textSegment && string.Equals(textSegment.Text, text, StringComparison.Ordinal));
		}

		public static PropertySet MarkerProps(string markerId, params string[] tileLabels)
		{
			return new PropertySet()
			{
				[Marker.ReservedMarkerIdKey] = markerId,
				[Marker.ReservedTileLabelsKey] = tileLabels,
			};
		}

		public static List<SequenceDeltaEventArgs> CaptureEvents(SharedString sharedString)
		{
			List<SequenceDeltaEventArgs> events = new();
			sharedString.OnSequenceDelta += (_, args) => events.Add(args);
			return events;
		}

		public static void ProcessRemoteInsert(
			SharedString sharedString,
			int position,
			string text,
			long refSeq,
			long seq,
			string clientId = "remote-client")
		{
			ProcessRemoteOp(
				sharedString,
				new MergeTreeInsertMsg()
				{
					Pos1 = position,
					Seg = text,
				},
				refSeq,
				seq,
				clientId);
		}

		public static void ProcessRemoteMarkerInsert(
			SharedString sharedString,
			int position,
			string markerId,
			string tileLabel,
			long refSeq,
			long seq,
			string clientId = "remote-client")
		{
			ProcessRemoteOp(
				sharedString,
				new MergeTreeInsertMsg()
				{
					Pos1 = position,
					Seg = Marker.Make(ReferenceType.Tile, MarkerProps(markerId, tileLabel)),
				},
				refSeq,
				seq,
				clientId);
		}

		public static void ProcessRemoteRemove(
			SharedString sharedString,
			int start,
			int end,
			long refSeq,
			long seq,
			string clientId = "remote-client")
		{
			ProcessRemoteOp(
				sharedString,
				new MergeTreeRemoveMsg()
				{
					Pos1 = start,
					Pos2 = end,
				},
				refSeq,
				seq,
				clientId);
		}

		public static void ProcessRemoteAnnotate(
			SharedString sharedString,
			int start,
			int end,
			PropertySet props,
			long refSeq,
			long seq,
			string clientId = "remote-client")
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
				seq,
				clientId);
		}

		public static void ProcessRemoteObliterate(
			SharedString sharedString,
			int start,
			int end,
			long refSeq,
			long seq,
			string clientId = "remote-client")
		{
			ProcessRemoteOp(
				sharedString,
				new MergeTreeObliterateMsg()
				{
					Pos1 = start,
					Pos2 = end,
				},
				refSeq,
				seq,
				clientId);
		}

		public static void ProcessRemoteOp(
			SharedString sharedString,
			IMergeTreeOp op,
			long refSeq,
			long seq,
			string clientId = "remote-client")
		{
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq, seq, clientId), SharedStringOpSerializer.Serialize(op));
		}

		public static void ProcessLocalAck(
			SharedString sharedString,
			(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
			long refSeq,
			long seq)
		{
			sharedString.ProcessDataObjectOp(LocalAck(sent.ClientSeq, refSeq, seq), sent.OpJson);
		}

		public static SequencedDocumentMessageDescriptor RemoteMessage(long refSeq, long seq, string clientId = "remote-client")
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: 0, refSeq: refSeq, seq: seq),
				OpOrigin.Remote,
				clientId);
		}

		public static SequencedDocumentMessageDescriptor LocalAck(long clientSeq, long refSeq, long seq)
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: clientSeq, refSeq: refSeq, seq: seq),
				OpOrigin.Local);
		}

		public static (ISegment Segment, int OffsetInSegment) GetRequiredContainingSegment(SharedString sharedString, int position)
		{
			(ISegment segment, int offsetInSegment)? result = sharedString.GetContainingSegment(position);
			Assert.NotNull(result);
			return (result.Value.segment, result.Value.offsetInSegment);
		}

		public static void AssertIntervalPositions(SequenceInterval? interval, int start, int end)
		{
			Assert.NotNull(interval);
			Assert.Equal((int?)start, interval!.StartPosition);
			Assert.Equal((int?)end, interval.EndPosition);
		}

		public static void AssertIds(IEnumerable<SequenceInterval> intervals, params string[] expectedIds)
		{
			Assert.Equal(expectedIds, intervals.Select(interval => interval.Id).ToArray());
		}

		public static void AssertConverged(SharedString first, SharedString second)
		{
			Assert.Equal(first.GetText(), second.GetText());
			Assert.Equal(first.GetLength(), second.GetLength());
		}

		public sealed class TwoClientHarness
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

			public void LoadInitialText(string text)
			{
				ClientA.InsertText(0, text);
				DeliverAtoB(Assert.Single(SenderA.Sent), refSeq: 0);
				SenderA.Sent.Clear();
				SenderB.Sent.Clear();
				AssertConverged(ClientA, ClientB);
			}

			public void DeliverAtoB((string Address, string OpTypeName, string OpJson, long ClientSeq) sent, long refSeq)
			{
				Deliver(sent, _clientAId, ClientA, ClientB, refSeq);
			}

			public void DeliverBtoA((string Address, string OpTypeName, string OpJson, long ClientSeq) sent, long refSeq)
			{
				Deliver(sent, _clientBId, ClientB, ClientA, refSeq);
			}

			private void Deliver(
				(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
				string fromClientId,
				SharedString from,
				SharedString to,
				long refSeq)
			{
				long serverSeq = ++_serverSeq;
				to.ProcessDataObjectOp(
					new SequencedDocumentMessageDescriptor(
						SequenceNumber.ForTesting(clientSeq: sent.ClientSeq, refSeq: refSeq, seq: serverSeq),
						OpOrigin.Remote,
						fromClientId),
					sent.OpJson);
				from.ProcessDataObjectOp(
					new SequencedDocumentMessageDescriptor(
						SequenceNumber.ForTesting(clientSeq: sent.ClientSeq, refSeq: refSeq, seq: serverSeq),
						OpOrigin.Local,
						fromClientId),
					sent.OpJson);
			}
		}
	}
}
