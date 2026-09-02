// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SharedStringAnnotateTests
	{
		[Fact]
		public void AnnotateRange_SetsPropsOnSegment()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");

			sharedString.AnnotateRange(0, 5, new PropertySet()
			{
				["color"] = "red",
			});

			Assert.Equal("red", Assert.IsType<string>(GetRequiredProperty(sharedString, 0, "color")));
		}

		[Fact]
		public void AnnotateRange_MultipleProps_AllApplied()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");

			sharedString.AnnotateRange(0, 5, new PropertySet()
			{
				["color"] = "red",
				["font"] = "arial",
				["bold"] = true,
				["weight"] = 700,
			});

			PropertySet properties = GetRequiredProperties(sharedString, 2);
			Assert.Equal("red", Assert.IsType<string>(properties["color"]));
			Assert.Equal("arial", Assert.IsType<string>(properties["font"]));
			Assert.True(Assert.IsType<bool>(properties["bold"]));
			Assert.Equal(700, Assert.IsType<int>(properties["weight"]));
		}

		[Fact]
		public void AnnotateRange_OverExistingProps_MergesAndOverwrites()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");
			sharedString.AnnotateRange(0, 5, new PropertySet()
			{
				["color"] = "red",
			});

			sharedString.AnnotateRange(0, 5, new PropertySet()
			{
				["color"] = "blue",
				["bold"] = true,
			});

			PropertySet properties = GetRequiredProperties(sharedString, 1);
			Assert.Equal("blue", Assert.IsType<string>(properties["color"]));
			Assert.True(Assert.IsType<bool>(properties["bold"]));
		}

		[Fact]
		public void AnnotateRange_NullValue_DeletesProperty()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");
			sharedString.AnnotateRange(0, 5, new PropertySet()
			{
				["color"] = "red",
			});

			sharedString.AnnotateRange(0, 5, new PropertySet()
			{
				["color"] = null,
			});

			AssertNoProperty(sharedString, 1, "color");
		}

		[Fact]
		public void AnnotateRange_NullValue_WirePreservesDeletionKey()
		{
			// TS ref: packages/dds/merge-tree/src/opBuilder.ts createAnnotateRangeOp —
			// TS emits the null property value on the wire so remote peers can
			// apply the deletion. The T03 audit finding flagged this end-to-end
			// path as untested.
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("s", sender);
			sharedString.InsertText(0, "hello world");
			sender.Sent.Clear();

			sharedString.AnnotateRange(0, 5, new PropertySet()
			{
				["color"] = null,
			});

			var sent = Assert.Single(sender.Sent);
			using System.Text.Json.JsonDocument doc = System.Text.Json.JsonDocument.Parse(sent.OpJson);
			System.Text.Json.JsonElement props = doc.RootElement.GetProperty("props");
			Assert.True(props.TryGetProperty("color", out System.Text.Json.JsonElement colorValue));
			Assert.Equal(System.Text.Json.JsonValueKind.Null, colorValue.ValueKind);
		}

		[Fact]
		public void AnnotateRange_SplitsSegmentsAtBoundaries()
		{
			MergeTree.MergeTree tree = new MergeTree.MergeTree();
			tree.InsertSegments(
				0,
				new ISegment[] { new TextSegment("hello world") },
				MergeTree.MergeTree.UnassignedSequenceNumber,
				MergeTree.MergeTree.UnassignedSequenceNumber,
				"client");

			tree.AnnotateRange(2, 7, new PropertySet()
			{
				["color"] = "red",
			}, MergeTree.MergeTree.UnassignedSequenceNumber, MergeTree.MergeTree.UnassignedSequenceNumber, "client");

			List<ISegment> segments = tree.WalkAllSegments().ToList();
			Assert.Equal(3, segments.Count);
			Assert.Equal("he", Assert.IsType<TextSegment>(segments[0]).Text);
			Assert.Equal("llo w", Assert.IsType<TextSegment>(segments[1]).Text);
			Assert.Equal("orld", Assert.IsType<TextSegment>(segments[2]).Text);
			Assert.Null(segments[0].Properties);
			Assert.Equal("red", Assert.IsType<string>(segments[1].Properties!["color"]));
			Assert.Null(segments[2].Properties);
		}

		[Fact]
		public void LocalAnnotate_EmitsAnnotateOp()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("shared-string", sender);
			sharedString.InsertText(0, "hello world");
			sender.Sent.Clear();

			sharedString.AnnotateRange(2, 7, new PropertySet()
			{
				["color"] = "red",
			});

			var sent = Assert.Single(sender.Sent);
			Assert.Equal("shared-string", sent.Address);
			Assert.Equal("annotate", sent.OpTypeName);
			MergeTreeAnnotateMsg op = Assert.IsType<MergeTreeAnnotateMsg>(
				SharedStringOpSerializer.Deserialize(sent.OpJson));
			Assert.Equal(2, op.Pos1);
			Assert.Equal(7, op.Pos2);
			Assert.NotNull(op.Props);
			Assert.Equal("red", Assert.IsType<string>(op.Props!["color"]));
		}

		[Fact]
		public void LocalAnnotate_WithoutSender_DoesNotEmit()
		{
			var sharedString = new SharedString(sender: null);
			sharedString.InsertText(0, "hello world");
			SequenceDeltaEventArgs? captured = null;
			sharedString.OnSequenceDelta += (sender, args) => captured = args;

			sharedString.AnnotateRange(0, 5, new PropertySet()
			{
				["color"] = "red",
			});

			Assert.Equal("red", Assert.IsType<string>(GetRequiredProperty(sharedString, 0, "color")));
			Assert.NotNull(captured);
			Assert.Equal("annotate", captured!.OpType);
		}

		[Fact]
		public void LocalAnnotate_FiresEventWithLocalTrue()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");
			var events = new List<SequenceDeltaEventArgs>();
			sharedString.OnSequenceDelta += (sender, args) => events.Add(args);

			sharedString.AnnotateRange(2, 7, new PropertySet()
			{
				["color"] = "red",
			});

			SequenceDeltaEventArgs captured = Assert.Single(events);
			Assert.True(captured.Local);
			Assert.Equal("annotate", captured.OpType);
			Assert.Equal(2, captured.Position);
			Assert.Equal(5, captured.Length);
			Assert.Null(captured.Text);
			Assert.NotNull(captured.AnnotatedProperties);
			Assert.Equal("red", Assert.IsType<string>(captured.AnnotatedProperties!["color"]));
		}

		[Fact]
		public void RemoteAnnotate_FiresEventWithLocalFalse()
		{
			SharedString sharedString = CreateSharedStringWithAckedText("hello world");
			var events = new List<SequenceDeltaEventArgs>();
			sharedString.OnSequenceDelta += (sender, args) => events.Add(args);

			ProcessRemoteAnnotate(sharedString, 0, 5, new PropertySet()
			{
				["color"] = "red",
			}, refSeq: 1, seq: 2);

			SequenceDeltaEventArgs captured = Assert.Single(events);
			Assert.False(captured.Local);
			Assert.Equal("annotate", captured.OpType);
			Assert.Equal(0, captured.Position);
			Assert.Equal(5, captured.Length);
			Assert.Null(captured.Text);
			Assert.Equal("red", Assert.IsType<string>(GetRequiredProperty(sharedString, 0, "color")));
		}

		[Fact]
		public void RemoteAnnotate_WithLocalPendingInsertBefore_PreviousPropsFromRemotePerspective()
		{
			// Remote annotate walks in the message's (refSeq, clientId)
			// perspective; pre-annotate property snapshotting must use the
			// same perspective so PreviousProperties align with the segments
			// actually mutated. Walking the current local view misaligns
			// after local pending inserts before the range.
			var sender = new FakeFluidDataObjectSender();
			SharedString sharedString = new("doc", sender);
			sharedString.InsertText(0, "abcde");
			ProcessLocalAck(sharedString, Assert.Single(sender.Sent), refSeq: 0, seq: 1);
			sender.Sent.Clear();

			// Pre-existing color=blue on the whole string, acked.
			sharedString.AnnotateRange(0, 5, new PropertySet() { ["color"] = "blue" });
			ProcessLocalAck(sharedString, Assert.Single(sender.Sent), refSeq: 1, seq: 2);
			sender.Sent.Clear();

			// Local pending insert of 5 chars at the front. Local view
			// shifts by 5; remote's perspective still sees the original 5.
			sharedString.InsertText(0, "XXXXX");
			sender.Sent.Clear();

			var events = new List<SequenceDeltaEventArgs>();
			sharedString.OnSequenceDelta += (_, e) => events.Add(e);

			// Remote annotate on the ORIGINAL 5 chars in remote perspective.
			ProcessRemoteAnnotate(
				sharedString,
				0,
				5,
				new PropertySet() { ["color"] = "red" },
				refSeq: 2,
				seq: 3);

			SequenceDeltaEventArgs captured = Assert.Single(events);
			Assert.False(captured.Local);
			// Previous color must be "blue" — the value on the segments the
			// remote actually mutates. Prior to the fix the snapshot walked
			// the local view and covered only the pending "XXXXX" segment
			// (no color), so PropertyDeltas came back null.
			SequenceDeltaRange range = Assert.Single(captured.Ranges);
			Assert.Equal("blue", Assert.IsType<string>(range.PropertyDeltas["color"]));
		}

		[Fact]
		public void TwoClients_SequentialAnnotate_Converges()
		{
			var harness = new TwoClientHarness();
			LoadInitialSharedText(harness, "hello world");
			harness.SenderA.Sent.Clear();

			harness.ClientA.AnnotateRange(0, 5, new PropertySet()
			{
				["color"] = "red",
			});
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);

			AssertConvergedText(harness.ClientA, harness.ClientB);
			Assert.Equal("red", Assert.IsType<string>(GetRequiredProperty(harness.ClientA, 0, "color")));
			Assert.Equal("red", Assert.IsType<string>(GetRequiredProperty(harness.ClientB, 0, "color")));
		}

		[Fact]
		public void TwoClients_ConcurrentNonOverlappingAnnotate_Converges()
		{
			var harness = new TwoClientHarness();
			LoadInitialSharedText(harness, "abcdefghij");
			harness.SenderA.Sent.Clear();
			harness.SenderB.Sent.Clear();
			long refSeq = harness.CurrentServerSeq;

			harness.ClientA.AnnotateRange(0, 3, new PropertySet()
			{
				["color"] = "red",
			});
			var sentA = Assert.Single(harness.SenderA.Sent);
			harness.ClientB.AnnotateRange(5, 8, new PropertySet()
			{
				["bold"] = true,
			});
			var sentB = Assert.Single(harness.SenderB.Sent);

			harness.DeliverAtoB(sentA, refSeq);
			harness.DeliverBtoA(sentB, refSeq);

			AssertConvergedText(harness.ClientA, harness.ClientB);
			Assert.Equal("red", Assert.IsType<string>(GetRequiredProperty(harness.ClientA, 1, "color")));
			Assert.Equal("red", Assert.IsType<string>(GetRequiredProperty(harness.ClientB, 1, "color")));
			Assert.True(Assert.IsType<bool>(GetRequiredProperty(harness.ClientA, 6, "bold")));
			Assert.True(Assert.IsType<bool>(GetRequiredProperty(harness.ClientB, 6, "bold")));
		}

		[Fact]
		public void TwoClients_ConcurrentOverlappingAnnotate_LastWriteWins()
		{
			var harness = new TwoClientHarness();
			LoadInitialSharedText(harness, "hello world");
			harness.SenderA.Sent.Clear();
			harness.SenderB.Sent.Clear();
			long refSeq = harness.CurrentServerSeq;

			harness.ClientA.AnnotateRange(0, 5, new PropertySet()
			{
				["color"] = "red",
			});
			var sentA = Assert.Single(harness.SenderA.Sent);
			harness.ClientB.AnnotateRange(0, 5, new PropertySet()
			{
				["color"] = "blue",
			});
			var sentB = Assert.Single(harness.SenderB.Sent);

			harness.DeliverAtoB(sentA, refSeq);
			harness.DeliverBtoA(sentB, refSeq);

			object? colorA = GetRequiredProperty(harness.ClientA, 0, "color");
			object? colorB = GetRequiredProperty(harness.ClientB, 0, "color");
			AssertConvergedText(harness.ClientA, harness.ClientB);
			Assert.Equal(colorA, colorB);
			Assert.Equal("blue", Assert.IsType<string>(colorA));
		}

		[Fact]
		public void LoadSnapshot_ThenAnnotate_WorksCorrectly()
		{
			var sharedString = new SharedString();
			sharedString.LoadFromSnapshot(LoadSimpleHelloSnapshot());

			sharedString.AnnotateRange(7, 12, new PropertySet()
			{
				["color"] = "green",
			});

			Assert.Equal("Hello, world!", sharedString.GetText());
			Assert.Equal("green", Assert.IsType<string>(GetRequiredProperty(sharedString, 7, "color")));
			AssertNoProperty(sharedString, 0, "color");
		}

		[Fact]
		public void AnnotateRange_ZeroWidth_Throws()
		{
			// regression. TS Client.getValidOpRange rejects zero-width
			// annotate as RangeOutOfBounds because `end <= start` is invalid
			// for local ops. The port must not silently queue an empty
			// annotate op.
			SharedString sharedString = new();
			sharedString.InsertText(0, "hello world");

			Assert.Throws<UsageError>(() =>
				sharedString.AnnotateRange(3, 3, new PropertySet() { ["color"] = "red" }));
		}

		[Fact]
		public void AnnotateRange_InvertedRange_Throws()
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, "hello world");

			Assert.Throws<UsageError>(() =>
				sharedString.AnnotateRange(5, 2, new PropertySet() { ["color"] = "red" }));
		}

		[Fact]
		public void AnnotateRange_StartAtLength_Throws()
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, "hello world");

			Assert.Throws<UsageError>(() =>
				sharedString.AnnotateRange(sharedString.GetLength(), sharedString.GetLength(), new PropertySet() { ["color"] = "red" }));
		}

		[Fact]
		public void SequenceDelta_InsertEvent_HasEmptyPropertyDeltas()
		{
			// regression. TS SequenceDeltaEvent.ranges[i].propertyDeltas
			// is always a PropertySet — empty for non-annotate ops, populated
			// for annotate. The port must not surface `null` here; listeners
			// should be able to enumerate without null-guarding.
			SharedString sharedString = new();
			List<SequenceDeltaEventArgs> events = new();
			sharedString.OnSequenceDelta += (_, e) => events.Add(e);

			sharedString.InsertText(0, "hello");

			SequenceDeltaEventArgs args = Assert.Single(events);
			Assert.NotEmpty(args.Ranges);
			foreach (SequenceDeltaRange range in args.Ranges)
			{
				Assert.NotNull(range.PropertyDeltas);
				Assert.Empty(range.PropertyDeltas);
			}
		}

		[Fact]
		public void SequenceDelta_AnnotateEvent_NestedPropertyClonedDeeply()
		{
			// regression. Event listeners must not be able to mutate
			// nested property values that also live inside the enqueued
			// canonical op / segment property map.
			SharedString sharedString = new();
			sharedString.InsertText(0, "hello");

			Dictionary<string, object?> originalMeta = new()
			{
				["author"] = "alice",
				["ts"] = 1234,
			};

			List<SequenceDeltaEventArgs> events = new();
			sharedString.OnSequenceDelta += (_, e) =>
			{
				events.Add(e);
				// Mutate the nested dict that the listener received via the
				// event. This must not reach through to the segment or the
				// batched op.
				if (e.AnnotatedProperties is not null
					&& e.AnnotatedProperties["meta"] is IDictionary<string, object?> nested)
				{
					nested["author"] = "eve";
				}
			};

			sharedString.AnnotateRange(0, 5, new PropertySet() { ["meta"] = originalMeta });

			// The caller's original dict must be untouched.
			Assert.Equal("alice", originalMeta["author"]);
			// The segment's stored nested dict must be untouched.
			PropertySet? props = sharedString.GetPropertiesAtPosition(0);
			Assert.NotNull(props);
			IDictionary<string, object?> storedMeta = Assert.IsAssignableFrom<IDictionary<string, object?>>(props!["meta"]);
			Assert.Equal("alice", storedMeta["author"]);
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
			AssertConvergedText(harness.ClientA, harness.ClientB);
		}

		private static void ProcessRemoteAnnotate(
			SharedString sharedString,
			int start,
			int end,
			PropertySet props,
			long refSeq,
			long seq,
			string clientId = "remote-client")
		{
			string opJson = SharedStringOpSerializer.Serialize(new MergeTreeAnnotateMsg()
			{
				Pos1 = start,
				Pos2 = end,
				Props = props,
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

		private static PropertySet GetRequiredProperties(SharedString sharedString, int position)
		{
			PropertySet? properties = sharedString.GetPropertiesAtPosition(position);
			Assert.NotNull(properties);
			return properties!;
		}

		private static object? GetRequiredProperty(SharedString sharedString, int position, string key)
		{
			PropertySet properties = GetRequiredProperties(sharedString, position);
			Assert.True(properties.ContainsKey(key), $"Expected property '{key}' at position {position}.");
			return properties[key];
		}

		private static void AssertNoProperty(SharedString sharedString, int position, string key)
		{
			PropertySet? properties = sharedString.GetPropertiesAtPosition(position);
			Assert.True(properties is null || !properties.ContainsKey(key), $"Expected property '{key}' to be absent at position {position}.");
		}

		private static void AssertConvergedText(SharedString clientA, SharedString clientB)
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
