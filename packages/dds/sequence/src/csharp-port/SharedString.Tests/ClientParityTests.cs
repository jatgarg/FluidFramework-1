// -----------------------------------------------------------------------------
// Client parity regressions for C-series audit findings.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class ClientParityTests
	{
		[Fact]
		public void ClientEmptyInsert_DoesNotCreatePendingOp()
		{
			Client client = new("local");

			IMergeTreeInsertMsg op = client.InsertText(0, string.Empty);

			Assert.Equal(0, client.GetLength());
			Assert.Equal(0, client.PendingOpCount);
			Assert.Null(op.ClientSeq);
			Assert.Equal(string.Empty, Assert.IsType<string>(op.Seg));
		}

		[Fact]
		public void RemoteInsert_RelativeToMarker_ResolvesPosition()
		{
			Client client = CreateClientWithAckedMarker();

			client.ApplyOp(
				new MergeTreeInsertMsg()
				{
					RelativePos1 = new PropertySet()
					{
						["id"] = "marker-1",
						["before"] = true,
					},
					Seg = "X",
				},
				seq: 3,
				refSeq: 2,
				clientId: "remote");

			Assert.Equal("aXb", client.GetText(0, client.GetLength()));
			Assert.Equal(4, client.GetLength());
		}

		[Fact]
		public void RemoteRemove_RelativeToMarker_ResolvesRange()
		{
			Client client = CreateClientWithAckedMarker();

			client.ApplyOp(
				new MergeTreeRemoveMsg()
				{
					RelativePos1 = new PropertySet()
					{
						["id"] = "marker-1",
						["before"] = true,
					},
					RelativePos2 = new PropertySet()
					{
						["id"] = "marker-1",
					},
				},
				seq: 3,
				refSeq: 2,
				clientId: "remote");

			Assert.Equal("ab", client.GetText(0, client.GetLength()));
			Assert.Equal(2, client.GetLength());
			Assert.Null(client.GetMarkerFromId("marker-1"));
		}

		[Fact]
		public void RemoteGroup_MixedOps_EmitsPerMemberEventsInOrder()
		{
			SharedString sharedString = CreateSharedStringWithAckedText("abcdef", out _);
			List<SequenceDeltaEventArgs> events = CaptureEvents(sharedString);
			MergeTreeGroupMsg group = new();
			group.Ops.Add(new MergeTreeRemoveMsg()
			{
				Pos1 = 1,
				Pos2 = 3,
			});
			group.Ops.Add(new MergeTreeInsertMsg()
			{
				Pos1 = 2,
				Seg = "X",
			});

			ProcessRemoteOp(sharedString, group, refSeq: 1, seq: 2);

			Assert.Equal("adXef", sharedString.GetText());
			Assert.Collection(
				events,
				first =>
				{
					Assert.Equal("remove", first.OpType);
					Assert.Equal(1, first.Position);
					Assert.Equal(2, first.Length);
				},
				second =>
				{
					Assert.Equal("insert", second.OpType);
					Assert.Equal(2, second.Position);
					Assert.Equal("X", second.Text);
				});
		}

		[Fact]
		public void AckLocalRemove_PromotesSplitPendingSegments()
		{
			Client client = CreateAckedClient("local", "abcdef");
			IMergeTreeRemoveMsg localRemove = client.RemoveText(1, 5);

			client.ApplyOp(
				new MergeTreeRemoveMsg()
				{
					Pos1 = 2,
					Pos2 = 4,
				},
				seq: 10,
				refSeq: 1,
				clientId: "remote");
			client.ApplyOp(localRemove, seq: 11, refSeq: 1, clientId: "local");

			Assert.Equal("af", client.GetText(0, client.GetLength()));
			foreach (TextSegment segment in client.MergeTree.WalkAllSegments().OfType<TextSegment>())
			{
				if (segment.Text is "b" or "cd" or "e")
				{
					SetRemoveOperationStamp stamp = Assert.Single(
						segment.RemoveStamps.OfType<SetRemoveOperationStamp>(),
						candidate => candidate.Seq == 11);
					Assert.Null(stamp.LocalSeq);
				}
			}
		}

		[Fact]
		public void MinimumSequence_DoesNotAdvancePastInFlightLocalRefSeq()
		{
			Client client = new("local");
			IMergeTreeInsertMsg pendingInsert = client.InsertText(0, "hello");

			client.ApplyOp(
				new MergeTreeInsertMsg()
				{
					Pos1 = 0,
					Seg = "abc",
				},
				seq: 17,
				refSeq: 16,
				clientId: "remote",
				minimumSequenceNumber: 16);

			Assert.Equal(0, client.CollabWindowMinSeq);

			client.ApplyOp(
				pendingInsert,
				seq: 18,
				refSeq: 16,
				clientId: "local",
				minimumSequenceNumber: 16);

			Assert.Equal(16, client.CollabWindowMinSeq);
		}

		[Fact]
		public void RemoteInsert_TraversesHierarchyWithOnlyHistoricalLength()
		{
			Client client = CreateAckedCharacterClient("local", 12);

			client.ApplyOp(
				new MergeTreeRemoveMsg()
				{
					Pos1 = 0,
					Pos2 = 12,
				},
				seq: 20,
				refSeq: 12,
				clientId: "remote-remove");
			client.ApplyOp(
				new MergeTreeInsertMsg()
				{
					Pos1 = 6,
					Seg = "X",
				},
				seq: 21,
				refSeq: 12,
				clientId: "remote-insert");

			Assert.Equal("X", client.GetText(0, client.GetLength()));
		}

		[Fact]
		public void LocalInsertAck_TraversesObliteratedHierarchy()
		{
			Client client = CreateAckedCharacterClient("local", 12);
			IMergeTreeInsertMsg localInsert = client.InsertText(6, "X");

			client.ApplyOp(
				new MergeTreeObliterateMsg()
				{
					Pos1 = 0,
					Pos2 = 12,
				},
				seq: 20,
				refSeq: 12,
				clientId: "remote-obliterate");
			Assert.Equal("X", client.GetText(0, client.GetLength()));

			client.ApplyOp(localInsert, seq: 21, refSeq: 12, clientId: "local");

			Assert.Equal(string.Empty, client.GetText(0, client.GetLength()));
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

		private static Client CreateClientWithAckedMarker()
		{
			Client client = CreateAckedClient("local", "ab");
			IMergeTreeInsertMsg markerOp = client.InsertMarker(
				1,
				ReferenceType.Tile,
				new PropertySet()
				{
					[Marker.reservedMarkerIdKey] = "marker-1",
				});
			client.ApplyOp(markerOp, seq: 2, refSeq: 1, clientId: "local");
			return client;
		}

		private static Client CreateAckedClient(string clientId, string text)
		{
			Client client = new(clientId);
			IMergeTreeInsertMsg op = client.InsertText(0, text);
			client.ApplyOp(op, seq: 1, refSeq: 0, clientId: clientId);
			return client;
		}

		private static Client CreateAckedCharacterClient(string clientId, int length)
		{
			Client client = new(clientId);
			for (int i = 0; i < length; i++)
			{
				IMergeTreeInsertMsg op = client.InsertText(client.GetLength(), "x");
				client.ApplyOp(op, seq: i + 1, refSeq: i, clientId: clientId);
			}

			return client;
		}

		private static List<SequenceDeltaEventArgs> CaptureEvents(SharedString sharedString)
		{
			List<SequenceDeltaEventArgs> events = new();
			sharedString.OnSequenceDelta += (sender, args) => events.Add(args);
			return events;
		}

		private static void ProcessRemoteOp(SharedString sharedString, IMergeTreeOp op, long refSeq, long seq)
		{
			sharedString.ProcessDataObjectOp(
				RemoteMessage(refSeq, seq),
				SharedStringOpSerializer.Serialize(op));
		}

		private static void ProcessLocalAck(
			SharedString sharedString,
			(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
			long refSeq,
			long seq)
		{
			sharedString.ProcessDataObjectOp(
				new SequencedDocumentMessageDescriptor(
					SequenceNumber.ForTesting(clientSeq: sent.ClientSeq, refSeq: refSeq, seq: seq),
					OpOrigin.Local),
				sent.OpJson);
		}

		private static SequencedDocumentMessageDescriptor RemoteMessage(long refSeq, long seq)
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: 0, refSeq: refSeq, seq: seq),
				OpOrigin.Remote,
				"remote-client");
		}
	}
}
