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
			// TS ref: packages/dds/sequence/src/intervalCollection.ts changeProperties —
			// TS throws UsageError when a caller tries to overwrite
			// reservedRangeLabelsKey (the collection label). The port previously
			// updated the local props bag while the wire path canonicalized to
			// the collection name, leaving peers divergent.
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
			// TS ref: packages/dds/sequence/src/intervalCollection.ts changeInterval —
			// TS emits ONE op carrying both endpoint delta and property delta.
			// The port previously emitted two ops (change + propertyChanged),
			// and receivers of the TS combined form dropped user properties.
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
