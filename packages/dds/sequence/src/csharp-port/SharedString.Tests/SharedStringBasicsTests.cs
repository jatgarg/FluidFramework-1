// -----------------------------------------------------------------------------
// SharedString basics tests.
// -----------------------------------------------------------------------------

using System.Collections.Generic;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class SharedStringBasicsTests
	{
		[Fact]
		public void Ctor_AssignsId_WhenNotProvided()
		{
			var sharedString = new SharedString();

			Assert.False(string.IsNullOrEmpty(sharedString.Id));
		}

		[Fact]
		public void Ctor_AcceptsExplicitId()
		{
			var sharedString = new SharedString("test-id");

			Assert.Equal("test-id", sharedString.Id);
		}

		[Fact]
		public void EmptyString_HasZeroLength()
		{
			var sharedString = new SharedString();

			Assert.Equal(0, sharedString.GetLength());
			Assert.Equal(string.Empty, sharedString.GetText());
		}

		[Fact]
		public void InsertText_At0_AppendsAndReadsBack()
		{
			var sharedString = new SharedString();

			sharedString.InsertText(0, "hello");

			Assert.Equal("hello", sharedString.GetText());
			Assert.Equal(5, sharedString.GetLength());
		}

		[Fact]
		public void InsertText_AtMiddle_SplitsCorrectly()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");

			sharedString.InsertText(5, " beautiful");

			Assert.Equal("hello beautiful world", sharedString.GetText());
		}

		[Fact]
		public void InsertText_AtEnd_Appends()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello");

			sharedString.InsertText(sharedString.GetLength(), " world");

			Assert.Equal("hello world", sharedString.GetText());
		}

		[Fact]
		public void GetText_WithRange_ReturnsSubstring()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");

			Assert.Equal("hello", sharedString.GetText(0, 5));
			Assert.Equal("world", sharedString.GetText(6, 11));
		}

		[Fact]
		public void DeleteText_RemovesRange()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");

			sharedString.DeleteText(5, 6);

			Assert.Equal("helloworld", sharedString.GetText());
			Assert.Equal(10, sharedString.GetLength());
		}

		[Fact]
		public void DeleteText_FullContent_LeavesEmpty()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello");

			sharedString.DeleteText(0, sharedString.GetLength());

			Assert.Equal(string.Empty, sharedString.GetText());
			Assert.Equal(0, sharedString.GetLength());
		}

		[Fact]
		public void InsertText_MultipleSequentially_Concatenates()
		{
			var sharedString = new SharedString();

			sharedString.InsertText(0, "one");
			sharedString.InsertText(3, " two");
			sharedString.InsertText(7, " three");

			Assert.Equal("one two three", sharedString.GetText());
		}

		[Fact]
		public void OnSequenceDelta_FiresOnLocalInsert_WithLocalTrue()
		{
			var sharedString = new SharedString();
			SequenceDeltaEventArgs? captured = null;
			sharedString.OnSequenceDelta += (s, e) => captured = e;

			sharedString.InsertText(0, "hello");

			Assert.NotNull(captured);
			Assert.True(captured!.Local);
			Assert.Equal("insert", captured.OpType);
			Assert.Equal(0, captured.Position);
			Assert.Equal(5, captured.Length);
			Assert.Equal("hello", captured.Text);
		}

		[Fact]
		public void OnSequenceDelta_FiresOnLocalDelete_WithLocalTrue()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");
			var events = new List<SequenceDeltaEventArgs>();
			sharedString.OnSequenceDelta += (s, e) => events.Add(e);

			sharedString.DeleteText(5, 6);

			SequenceDeltaEventArgs captured = Assert.Single(events);
			Assert.True(captured.Local);
			Assert.Equal("remove", captured.OpType);
			Assert.Equal(5, captured.Position);
			Assert.Equal(1, captured.Length);
			Assert.Null(captured.Text);
		}

		[Fact]
		public void OnSequenceDelta_LocalEvent_ExposesLocalClientId()
		{
			// TS ref: packages/dds/sequence/src/sequenceDeltaEvent.ts — the
			// event carries the local client id on local events once a client
			// id has been assigned.
			var sharedString = new SharedString("local-client");
			SequenceDeltaEventArgs? captured = null;
			sharedString.OnSequenceDelta += (_, e) => captured = e;

			sharedString.InsertText(0, "hi");

			Assert.NotNull(captured);
			Assert.True(captured!.Local);
			Assert.Equal("local-client", captured.ClientId);
		}

		[Fact]
		public void OnSequenceDelta_RemoteAnnotate_PropertyDeltas_CarryPreviousValues()
		{
			// TS ref: sequenceDeltaEvent.ts — propertyDeltas on a remote
			// annotate carry the values that were REPLACED so listeners can
			// compute what changed.
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("s", sender);
			sharedString.InsertText(0, "abc");
			ProcessLocalAck(sharedString, Assert.Single(sender.Sent), refSeq: 0, seq: 1);
			sender.Sent.Clear();

			// Seed the range with an initial color.
			sharedString.AnnotateRange(0, 3, new PropertySet() { ["color"] = "blue" });
			ProcessLocalAck(sharedString, Assert.Single(sender.Sent), refSeq: 1, seq: 2);
			sender.Sent.Clear();

			SequenceDeltaEventArgs? captured = null;
			sharedString.OnSequenceDelta += (_, e) =>
			{
				if (!e.Local && e.OpType == "annotate")
				{
					captured = e;
				}
			};

			// Remote annotate replaces color: blue -> red on the whole range.
			string wire = "{\"type\":2,\"pos1\":0,\"pos2\":3,\"props\":{\"color\":\"red\"}}";
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq: 2, seq: 3), wire);

			Assert.NotNull(captured);
			SequenceDeltaRange range = Assert.Single(captured!.Ranges);
			Assert.NotNull(range.PropertyDeltas);
			// Previous value was "blue" (not the new "red" from the op).
			Assert.Equal("blue", Assert.IsType<string>(range.PropertyDeltas!["color"]));
		}

		private static SequencedDocumentMessageDescriptor RemoteMessage(long refSeq, long seq, string clientId = "remote-client")
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: 0, refSeq: refSeq, seq: seq),
				OpOrigin.Remote,
				clientId);
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

		[Fact]
		public void OnSequenceDelta_LocalListenerMutatingOp_DoesNotAffectWire()
		{
			// TS ref: sequence.ts / opBuilder.ts — TS's SequenceDeltaEvent
			// exposes a live op reference; a listener that mutates it corrupts
			// the on-wire message. The port hands the listener a defensive
			// clone so mutations stay inside the event args.
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("s", sender);
			sharedString.InsertText(0, "seed");
			sender.Sent.Clear();

			sharedString.OnSequenceDelta += (_, e) =>
			{
				if (e.Local && e.Op is MergeTreeInsertMsg insertMsg)
				{
					// A malicious/buggy listener rewrites the position.
					insertMsg.Pos1 = 999;
				}
			};

			sharedString.RunInBatch(() =>
			{
				sharedString.InsertText(4, "X");
			});

			var sent = Assert.Single(sender.Sent);
			using System.Text.Json.JsonDocument doc = System.Text.Json.JsonDocument.Parse(sent.OpJson);
			// Wire op preserves the original pos1 (4) despite the listener's mutation.
			Assert.Equal(4, doc.RootElement.GetProperty("pos1").GetInt32());
		}

		[Fact]
		public void GetPropertiesAtPosition_OutsideContent_ReturnsNull()
		{
			// TS ref: packages/dds/merge-tree/src/client.ts
			// getPropertiesAtPosition — returns undefined for positions
			// outside content.
			var sharedString = new SharedString();
			sharedString.InsertText(0, "abc");

			Assert.Null(sharedString.GetPropertiesAtPosition(-1));
			Assert.Null(sharedString.GetPropertiesAtPosition(3));   // end position
			Assert.Null(sharedString.GetPropertiesAtPosition(100));
		}

		[Fact]
		public void InsertText_WithSender_EmitsOp()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("shared-string", sender);

			sharedString.InsertText(0, "hello");

			var sent = Assert.Single(sender.Sent);
			Assert.Equal("shared-string", sent.Address);
			Assert.Equal("insert", sent.OpTypeName);
			Assert.False(string.IsNullOrEmpty(sent.OpJson));
			Assert.Equal(1, sent.ClientSeq);
		}

		[Fact]
		public void InsertText_WithoutSender_DoesNotEmit()
		{
			var sharedString = new SharedString(sender: null);

			sharedString.InsertText(0, "hello");

			Assert.Equal("hello", sharedString.GetText());
		}
	}
}
