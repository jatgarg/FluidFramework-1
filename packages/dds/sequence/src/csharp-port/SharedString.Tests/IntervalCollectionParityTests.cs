// -----------------------------------------------------------------------------
// IntervalCollection parity regressions for IC-series audit findings.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using IntervalSide = Microsoft.Office.Web.Fluid.Intervals.Side;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class IntervalCollectionParityTests
	{
		[Fact]
		public void AddChangeDeleteAndEnumerate_MatchesTsHappyPath()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdef");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			collection.Add(1, 3, intervalId: "a");
			collection.Add(4, 5, intervalId: "b");

			AssertIds(collection, "a", "b");

			SequenceInterval? changed = collection.Change("a", newStart: 2, newEnd: 4);
			SequenceInterval? removed = collection.RemoveIntervalById("b");

			Assert.NotNull(changed);
			AssertIntervalPositions(changed!, 2, 4);
			Assert.NotNull(removed);
			Assert.Null(collection.GetIntervalById("b"));
			AssertIds(collection, "a");
		}

		[Fact]
		public void QueryWrappers_ReturnTsShapedResults()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdefgh");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			collection.Add(1, 1, intervalId: "a");
			collection.Add(1, 3, intervalId: "b");
			collection.Add(4, 6, intervalId: "c");
			collection.Add(6, 6, intervalId: "d");

			AssertIds(collection.FindOverlappingIntervals(3, 4), "b", "c");
			AssertIds(collection.FindIntervalsWithStartpointInRange(1, 4), "a", "b", "c");
			Assert.Empty(collection.FindIntervalsWithStartpointInRange(0, 4));
			AssertIds(collection.FindIntervalsWithEndpointInRange(3, 6), "b", "c", "d");
			Assert.Equal("b", collection.PreviousInterval(4)?.Id);
			Assert.Equal("c", collection.NextInterval(4)?.Id);

			List<string?> mappedIds = new();
			collection.Map(interval => mappedIds.Add(interval.Id));
			Assert.Equal(new[] { "a", "b", "c", "d" }, mappedIds);
		}

		[Fact]
		public void Change_WithProperties_EmitsPropertyDeltasAndEndpointChange()
		{
			SharedString sharedString = CreateSharedStringWithText("hello world");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval interval = collection.Add(
				0,
				3,
				new PropertySet()
				{
					["a"] = 1,
					["remove"] = "old",
				},
				intervalId: "i1");
			List<IntervalPropertyChangedEventArgs> propertyEvents = new();
			List<IntervalChangedEventArgs> changeEvents = new();
			List<IntervalUpdatedEventArgs> changedEvents = new();
			collection.OnPropertyChanged += (_, args) => propertyEvents.Add(args);
			collection.OnChangeInterval += (_, args) => changeEvents.Add(args);
			collection.OnChanged += (_, args) => changedEvents.Add(args);

			SequenceInterval? changed = collection.Change(
				"i1",
				newStart: 2,
				newEnd: 5,
				props: new PropertySet()
				{
					["a"] = 2,
					["remove"] = null,
					["added"] = "new",
				});

			Assert.Same(interval, changed);
			AssertIntervalPositions(interval, 2, 5);
			Assert.Equal(2, Assert.IsType<int>(interval.Properties!["a"]));
			Assert.Equal("new", Assert.IsType<string>(interval.Properties["added"]));
			Assert.False(interval.Properties.ContainsKey("remove"));

			IntervalPropertyChangedEventArgs propertyEvent = Assert.Single(propertyEvents);
			Assert.True(propertyEvent.Local);
			Assert.Null(propertyEvent.Operation);
			Assert.Equal(2, Assert.IsType<int>(propertyEvent.ChangedProps["a"]));
			Assert.Null(propertyEvent.ChangedProps["remove"]);
			Assert.Equal("old", Assert.IsType<string>(propertyEvent.PropertyDeltas["remove"]));
			Assert.Equal(1, Assert.IsType<int>(propertyEvent.PropertyDeltas["a"]));
			Assert.Null(propertyEvent.PropertyDeltas["added"]);

			IntervalChangedEventArgs changeEvent = Assert.Single(changeEvents);
			Assert.True(changeEvent.Local);
			Assert.False(changeEvent.Slide);
			Assert.Equal(0, changeEvent.PreviousStart);
			Assert.Equal(3, changeEvent.PreviousEnd);

			// TS ref: packages/dds/sequence/src/intervalCollection.ts changeInterval —
			// TS emits a single combined "changed" event with both endpoint delta
			// and property delta when both change together.
			IntervalUpdatedEventArgs combinedChanged = Assert.Single(changedEvents);
			Assert.Equal(0, combinedChanged.PreviousStart);
			Assert.Equal(3, combinedChanged.PreviousEnd);
			Assert.Equal(1, Assert.IsType<int>(combinedChanged.PropertyDeltas["a"]));
			Assert.Equal("old", Assert.IsType<string>(combinedChanged.PropertyDeltas["remove"]));
			Assert.Null(combinedChanged.PropertyDeltas["added"]);
		}

		[Fact]
		public void ChangeProperties_NoActualChange_DoesNotEmitPropertyEvent()
		{
			SharedString sharedString = CreateSharedStringWithText("hello");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			collection.Add(
				0,
				1,
				new PropertySet()
				{
					["color"] = "blue",
				},
				intervalId: "i1");
			List<IntervalPropertyChangedEventArgs> propertyEvents = new();
			collection.OnPropertyChanged += (_, args) => propertyEvents.Add(args);

			collection.ChangeProperties(
				"i1",
				new PropertySet()
				{
					["color"] = "blue",
				});

			Assert.Empty(propertyEvents);
		}

		[Fact]
		public void RemotePropertyChange_NullRemovesPropertyAndReportsPreviousValue()
		{
			SharedString sharedString = CreateSharedStringWithText("hello");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval interval = collection.Add(
				0,
				1,
				new PropertySet()
				{
					["color"] = "blue",
				},
				intervalId: "i1");
			List<IntervalPropertyChangedEventArgs> propertyEvents = new();
			collection.OnPropertyChanged += (_, args) => propertyEvents.Add(args);

			ProcessRemoteIntervalPropertyChanged(
				sharedString,
				"comments",
				"i1",
				new PropertySet()
				{
					["color"] = null,
					["size"] = 2,
				},
				refSeq: 0,
				seq: 1);

			Assert.NotNull(interval.Properties);
			Assert.False(interval.Properties!.ContainsKey("color"));
			Assert.Equal(2, Assert.IsType<int>(interval.Properties["size"]));

			IntervalPropertyChangedEventArgs propertyEvent = Assert.Single(propertyEvents);
			Assert.False(propertyEvent.Local);
			Assert.IsType<IntervalPropertyChangedOpMsg>(propertyEvent.Operation);
			Assert.Null(propertyEvent.ChangedProps["color"]);
			Assert.Equal("blue", Assert.IsType<string>(propertyEvent.PropertyDeltas["color"]));
			Assert.Null(propertyEvent.PropertyDeltas["size"]);
		}

		[Fact]
		public void RemoteAddChangeDelete_EventsCarryRemoteOperation()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdef");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			List<IntervalAddedEventArgs> addEvents = new();
			List<IntervalChangedEventArgs> changeEvents = new();
			List<IntervalDeletedEventArgs> deleteEvents = new();
			collection.OnAddInterval += (_, args) => addEvents.Add(args);
			collection.OnChangeInterval += (_, args) => changeEvents.Add(args);
			collection.OnDeleteInterval += (_, args) => deleteEvents.Add(args);

			ProcessRemoteIntervalAdd(sharedString, "comments", "remote", 1, 2, refSeq: 0, seq: 1);
			IntervalAddedEventArgs addEvent = Assert.Single(addEvents);
			Assert.False(addEvent.Local);
			Assert.IsType<IntervalAddOpMsg>(addEvent.Operation);
			AssertIntervalPositions(collection.GetIntervalById("remote")!, 1, 2);

			ProcessRemoteIntervalChange(sharedString, "comments", "remote", 2, 4, refSeq: 1, seq: 2);
			IntervalChangedEventArgs changeEvent = Assert.Single(changeEvents);
			Assert.False(changeEvent.Local);
			Assert.IsType<IntervalChangeOpMsg>(changeEvent.Operation);
			Assert.Equal(1, changeEvent.PreviousStart);
			Assert.Equal(2, changeEvent.PreviousEnd);
			AssertIntervalPositions(collection.GetIntervalById("remote")!, 2, 4);

			ProcessRemoteIntervalDelete(sharedString, "comments", "remote", refSeq: 2, seq: 3);
			IntervalDeletedEventArgs deleteEvent = Assert.Single(deleteEvents);
			Assert.False(deleteEvent.Local);
			Assert.IsType<IntervalDeleteOpMsg>(deleteEvent.Operation);
			Assert.Equal(2, deleteEvent.PreviousStart);
			Assert.Equal(4, deleteEvent.PreviousEnd);
			Assert.Null(collection.GetIntervalById("remote"));
		}

		[Fact]
		public void Rebase_PendingIntervalAdd_AfterRemoteInsert_ResubmitsAdjustedEndpoints()
		{
			var (sharedString, sender) = CreateAckedSharedString("client-a", "abcdefghij");
			sharedString.GetIntervalCollection("comments").Add(5, 10, intervalId: "abc");
			sender.Sent.Clear();
			ProcessRemoteInsert(sharedString, 2, "XX", refSeq: 1, seq: 2, clientId: "client-b");

			sharedString.RegeneratePendingOps();

			var rebased = Assert.Single(sender.Sent);
			IntervalAddOpMsg op = Assert.IsType<IntervalAddOpMsg>(SharedStringOpSerializer.Deserialize(rebased.OpJson));
			Assert.Equal(7, op.Start);
			Assert.Equal(12, op.End);
		}

		[Fact]
		public void Rebase_DetachedTransientPendingAdd_DropsIntervalAndEmitsDelete()
		{
			var (sharedString, sender) = CreateAckedSharedString("client-a", "abcdefghij");
			IntervalCollection collection = sharedString.GetIntervalCollection("transient", IntervalType.Transient);
			List<IntervalDeletedEventArgs> deleteEvents = new();
			collection.OnDeleteInterval += (_, args) => deleteEvents.Add(args);
			collection.Add(5, 5, intervalId: "t1");
			sender.Sent.Clear();
			ProcessRemoteRemove(sharedString, 4, 7, refSeq: 1, seq: 2, clientId: "client-b");

			sharedString.RegeneratePendingOps();

			Assert.Empty(sender.Sent);
			Assert.Null(collection.GetIntervalById("t1"));
			IntervalDeletedEventArgs deleteEvent = Assert.Single(deleteEvents);
			Assert.True(deleteEvent.Local);
			Assert.Equal("t1", deleteEvent.Interval.Id);
		}

		[Fact]
		public void Change_SideOnly_WithoutPositions_Throws()
		{
			// TS ref: packages/dds/sequence/src/intervalCollection.ts change() —
			// TS's change API is defined in terms of positions; side-only changes
			// would not have a coherent wire shape (positions are required for
			// the receiver's dispatch). Reject at the API boundary.
			SharedString sharedString = CreateSharedStringWithText("abcdef");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			collection.Add(1, 3, intervalId: "i1");

			OcsException ex = Assert.Throws<OcsException>(
				() => collection.Change("i1", newStart: null, newEnd: null, newStartSide: Intervals.Side.After, newEndSide: Intervals.Side.Before));
			Assert.Contains("positions", ex.Message);
		}

		[Fact]
		public void Change_ReferenceRangeLabels_Rejected()
		{
			// TS ref: packages/dds/sequence/src/intervalCollection.ts
			// changeProperties. Mutating reservedRangeLabelsKey is rejected
			// because the wire path canonicalizes to the collection name.
			SharedString sharedString = CreateSharedStringWithText("abcdef");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			collection.Add(1, 3, intervalId: "i1");

			OcsException ex = Assert.Throws<OcsException>(
				() => collection.ChangeProperties("i1", new PropertySet() { ["referenceRangeLabels"] = new[] { "evil" } }));
			Assert.Contains("referenceRangeLabels", ex.Message);
		}

		[Fact]
		public void Change_CombinedEndpointsAndProps_EmitsSingleWireOp()
		{
			// TS ref: packages/dds/sequence/src/intervalCollection.ts
			// changeInterval — one op carrying both endpoint delta and
			// property delta. Receivers of the two-op form would drop user
			// properties.
			var sender = new FakeFluidDataObjectSender();
			SharedString sharedString = new("s", sender);
			sharedString.InsertText(0, "abcdef");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			collection.Add(1, 3, intervalId: "i1", properties: new PropertySet() { ["color"] = "blue" });
			sender.Sent.Clear();

			collection.Change("i1", newStart: 2, newEnd: 4, props: new PropertySet() { ["color"] = "red" });

			// Exactly one wire op — the combined change form.
			var sent = Assert.Single(sender.Sent);
			using JsonDocument doc = JsonDocument.Parse(sent.OpJson);
			JsonElement value = doc.RootElement.GetProperty("value").GetProperty("value");
			Assert.Equal("change", doc.RootElement.GetProperty("value").GetProperty("opName").GetString());
			Assert.Equal(2, value.GetProperty("start").GetInt32());
			Assert.Equal(4, value.GetProperty("end").GetInt32());
			Assert.Equal("red", value.GetProperty("properties").GetProperty("color").GetString());
		}

		[Fact]
		public void ConcurrentChange_RemoteChangeDuringPending_DoesNotOverwriteLocalView()
		{
			// TS ref: packages/dds/sequence/src/intervalCollection.ts
			// ackChange — remote changes arriving during a local pending
			// change reconcile against pending[id].consensus rather than the
			// locally-mutated live interval, so concurrent editors converge.
			(SharedString local, FakeFluidDataObjectSender sender) = CreateAckedSharedString("client-a", "abcdefghij");
			IntervalCollection collection = local.GetIntervalCollection("comments");

			// Seed a shared interval known to both clients (add is local + ACKed).
			SequenceInterval interval = collection.Add(2, 5, intervalId: "i1");
			ProcessLocalAck(local, Assert.Single(sender.Sent), refSeq: 1, seq: 2);
			sender.Sent.Clear();

			// Local pending change: [2,5] -> [1,5]
			collection.Change("i1", newStart: 1, newEnd: 5);
			var pendingLocalSend = Assert.Single(sender.Sent);
			sender.Sent.Clear();

			// Concurrent remote change moves i1 to [3,7]. The local view
			// stays at [1,5] because the remote reconciles into consensus.
			ProcessRemoteIntervalChange(local, "comments", "i1", 3, 7, refSeq: 2, seq: 3);

			SequenceInterval? current = collection.GetIntervalById("i1");
			Assert.NotNull(current);
			Assert.Equal(1, current!.StartPosition);
			Assert.Equal(5, current.EndPosition);
			Assert.Same(interval, current);

			// Now ACK the local pending change. The consensus at ACK is [3,7]
			// (from the concurrent remote change), and our local mutation on
			// top brought it back to what we want ([1,5]). The port's live
			// interval still shows [1,5] — that's the caller's local intent.
			ProcessLocalAck(local, pendingLocalSend, refSeq: 3, seq: 4);
			SequenceInterval? afterAck = collection.GetIntervalById("i1");
			Assert.NotNull(afterAck);
			Assert.Equal(1, afterAck!.StartPosition);
			Assert.Equal(5, afterAck.EndPosition);
		}

		[Fact]
		public void ConcurrentChange_SegmentRemovedByRemoteWhilePending_FiresSlideEvent()
		{
			// TS ref: packages/dds/sequence/src/intervalCollection.ts
			// ackInterval — a "changed" event with slide: true fires at ACK
			// time when a segment holding one of the interval's endpoints was
			// sequenced-removed while our change was pending.
			(SharedString local, FakeFluidDataObjectSender sender) = CreateAckedSharedString("client-a", "abcdefghij");
			IntervalCollection collection = local.GetIntervalCollection("comments");

			SequenceInterval interval = collection.Add(2, 5, intervalId: "i1");
			ProcessLocalAck(local, Assert.Single(sender.Sent), refSeq: 1, seq: 2);
			sender.Sent.Clear();

			// Local pending change on i1 endpoints.
			collection.Change("i1", newStart: 3, newEnd: 6);
			var pendingLocalSend = Assert.Single(sender.Sent);
			sender.Sent.Clear();

			// A remote peer deletes the segment currently holding the start
			// endpoint (position 3), causing an eventual endpoint slide at
			// ACK time.
			ProcessRemoteRemove(local, 3, 4, refSeq: 2, seq: 3, clientId: "client-b");

			int slideCount = 0;
			collection.OnChanged += (_, e) =>
			{
				if (e.Slide)
				{
					slideCount++;
				}
			};

			// ACK of our local change — under TS semantics fires a "changed"
			// event with slide: true because the endpoint segment was
			// sequenced-removed during the pending window.
			ProcessLocalAck(local, pendingLocalSend, refSeq: 3, seq: 4);

			Assert.Equal(1, slideCount);
		}

		[Fact]
		public void ChangeProperties_EqualValue_StillSubmitsWireOp()
		{
			// V2-W04 regression. TS ref: packages/dds/sequence/src/
			// intervalCollection.ts changeInterval — submitSerializedOperation
			// runs unconditionally when props is supplied. The port must not
			// filter no-op writes out of the wire stream.
			(SharedString local, FakeFluidDataObjectSender sender) = CreateAckedSharedString("client-a", "abcdefghij");
			IntervalCollection collection = local.GetIntervalCollection("comments");

			collection.Add(2, 5, intervalId: "i1", properties: new PropertySet { ["color"] = "red" });
			ProcessLocalAck(local, Assert.Single(sender.Sent), refSeq: 1, seq: 2);
			sender.Sent.Clear();

			// Re-writing the same value with the same key should still
			// emit a wire op — TS does the same.
			collection.ChangeProperties("i1", new PropertySet { ["color"] = "red" });

			Assert.Single(sender.Sent);
			using JsonDocument document = JsonDocument.Parse(sender.Sent[0].OpJson);
			JsonElement value = document.RootElement.GetProperty("value");
			Assert.Equal("change", value.GetProperty("opName").GetString());
		}

		[Fact]
		public void ChangeProperties_DeleteAbsentKey_StillSubmitsWireOp()
		{
			// V2-W04 regression. TS ref: packages/dds/sequence/src/
			// intervalCollection.ts changeInterval — deleting a key that was
			// never present still submits the op. The receiver-side
			// property-manager handles the no-op semantics.
			(SharedString local, FakeFluidDataObjectSender sender) = CreateAckedSharedString("client-a", "abcdefghij");
			IntervalCollection collection = local.GetIntervalCollection("comments");

			collection.Add(2, 5, intervalId: "i1");
			ProcessLocalAck(local, Assert.Single(sender.Sent), refSeq: 1, seq: 2);
			sender.Sent.Clear();

			collection.ChangeProperties("i1", new PropertySet { ["missing"] = null });

			Assert.Single(sender.Sent);
		}

		[Fact]
		public void ChangeProperties_EqualValue_DoesNotSuppressLaterRemoteEndpointChange()
		{
			// V2-W04 regression. Before the fix, ChangeProperties() with an
			// equal value would allocate _pendingChanges[id] but never emit a
			// wire op, so no ACK could clear it; a subsequent remote endpoint
			// change would then be routed into UpdatePendingConsensusNoLock
			// and quietly reconciled into a phantom pending record rather than
			// applied to the live interval.
			(SharedString local, FakeFluidDataObjectSender sender) = CreateAckedSharedString("client-a", "abcdefghij");
			IntervalCollection collection = local.GetIntervalCollection("comments");

			SequenceInterval interval = collection.Add(2, 5, intervalId: "i1", properties: new PropertySet { ["color"] = "red" });
			ProcessLocalAck(local, Assert.Single(sender.Sent), refSeq: 1, seq: 2);
			sender.Sent.Clear();

			// Local no-op property write.
			collection.ChangeProperties("i1", new PropertySet { ["color"] = "red" });
			var sentPropertyOp = Assert.Single(sender.Sent);
			sender.Sent.Clear();

			// ACK the property op so pending state clears.
			ProcessLocalAck(local, sentPropertyOp, refSeq: 2, seq: 3);

			// Remote endpoint change now must reach the live interval, not
			// be swallowed by a leaked pending record.
			ProcessRemoteIntervalChange(local, "comments", "i1", 3, 7, refSeq: 3, seq: 4);

			SequenceInterval? current = collection.GetIntervalById("i1");
			Assert.NotNull(current);
			Assert.Equal(3, current!.StartPosition);
			Assert.Equal(7, current.EndPosition);
		}

		[Fact]
		public void Change_NoEndpointNoProps_DoesNotAllocatePendingState()
		{
			// V2-W04 regression. A Change() call with nothing to do must not
			// leak a pending record. The observable effect is that a later
			// remote endpoint change on the same interval reaches the live
			// interval.
			(SharedString local, FakeFluidDataObjectSender sender) = CreateAckedSharedString("client-a", "abcdefghij");
			IntervalCollection collection = local.GetIntervalCollection("comments");

			collection.Add(2, 5, intervalId: "i1");
			ProcessLocalAck(local, Assert.Single(sender.Sent), refSeq: 1, seq: 2);
			sender.Sent.Clear();

			// Change with everything null: no endpoint change, no props.
			collection.Change("i1", newStart: null, newEnd: null);
			Assert.Empty(sender.Sent);

			// Remote endpoint change must reach the live interval.
			ProcessRemoteIntervalChange(local, "comments", "i1", 3, 7, refSeq: 2, seq: 3);

			SequenceInterval? current = collection.GetIntervalById("i1");
			Assert.NotNull(current);
			Assert.Equal(3, current!.StartPosition);
			Assert.Equal(7, current.EndPosition);
		}

		private static SharedString CreateSharedStringWithText(string text)
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, text);
			return sharedString;
		}

		private static (SharedString SharedString, FakeFluidDataObjectSender Sender) CreateAckedSharedString(string id, string text)
		{
			FakeFluidDataObjectSender sender = new();
			SharedString sharedString = new(id, sender);
			sharedString.InsertText(0, text);
			ProcessLocalAck(sharedString, Assert.Single(sender.Sent), refSeq: 0, seq: 1);
			sender.Sent.Clear();
			return (sharedString, sender);
		}

		private static void ProcessRemoteIntervalAdd(
			SharedString sharedString,
			string collectionName,
			string intervalId,
			int start,
			int end,
			long refSeq,
			long seq)
		{
			string opJson = SharedStringOpSerializer.Serialize(new IntervalAddOpMsg()
			{
				CollectionName = collectionName,
				IntervalId = intervalId,
				Start = start,
				End = end,
				IntervalType = IntervalType.SlideOnRemove,
				Stickiness = IntervalStickiness.End,
				StartSide = IntervalSide.Before,
				EndSide = IntervalSide.Before,
			});
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq, seq, "remote-client"), opJson);
		}

		private static void ProcessRemoteIntervalChange(
			SharedString sharedString,
			string collectionName,
			string intervalId,
			int start,
			int end,
			long refSeq,
			long seq)
		{
			string opJson = SharedStringOpSerializer.Serialize(new IntervalChangeOpMsg()
			{
				CollectionName = collectionName,
				IntervalId = intervalId,
				Start = start,
				End = end,
				Stickiness = IntervalStickiness.End,
				StartSide = IntervalSide.Before,
				EndSide = IntervalSide.Before,
			});
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq, seq, "remote-client"), opJson);
		}

		private static void ProcessRemoteIntervalDelete(
			SharedString sharedString,
			string collectionName,
			string intervalId,
			long refSeq,
			long seq)
		{
			string opJson = SharedStringOpSerializer.Serialize(new IntervalDeleteOpMsg()
			{
				CollectionName = collectionName,
				IntervalId = intervalId,
			});
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq, seq, "remote-client"), opJson);
		}

		private static void ProcessRemoteIntervalPropertyChanged(
			SharedString sharedString,
			string collectionName,
			string intervalId,
			PropertySet props,
			long refSeq,
			long seq)
		{
			string opJson = SharedStringOpSerializer.Serialize(new IntervalPropertyChangedOpMsg()
			{
				CollectionName = collectionName,
				IntervalId = intervalId,
				Props = props,
			});
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq, seq, "remote-client"), opJson);
		}

		private static void ProcessRemoteInsert(
			SharedString sharedString,
			int position,
			string text,
			long refSeq,
			long seq,
			string clientId)
		{
			string opJson = SharedStringOpSerializer.Serialize(new MergeTreeInsertMsg()
			{
				Pos1 = position,
				Seg = text,
			});
			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq, seq, clientId), opJson);
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
				OpOrigin.Local,
				"client-a");
		}

		private static void AssertIntervalPositions(SequenceInterval interval, int start, int end)
		{
			Assert.Equal((int?)start, interval.StartPosition);
			Assert.Equal((int?)end, interval.EndPosition);
		}

		private static void AssertIds(IEnumerable<SequenceInterval> intervals, params string[] expectedIds)
		{
			Assert.Equal(expectedIds, intervals.Select(interval => interval.Id).ToArray());
		}
	}
}
