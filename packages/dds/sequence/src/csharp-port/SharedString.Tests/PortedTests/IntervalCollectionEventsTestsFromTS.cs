#nullable enable

using System.Collections.Generic;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	// Ported from packages/dds/sequence/src/test/intervalCollection.events.spec.ts
	//
	// Focused subset targeting the event-contract fixes from Waves V2-A11
	// through V2-A15: propertyChanged with previous values, changeInterval
	// with previous endpoints, and event firing on remote-op ack.
	//
	// Skipped:
	// - Handle-related event tests (covered by SharedStringHandleTests
	//   V2-A12 regression).
	// - Op.contents.type == "act" wire-shape assertions (covered by
	//   IntervalOpSerializerTests).
	public sealed class IntervalCollectionEventsTestsFromTS
	{
		// Ported from intervalCollection.events.spec.ts —
		// "addInterval / is emitted on initial local add"
		[Fact]
		public void AddInterval_LocalAdd_EmitsOnceWithLocalTrue()
		{
			PortedTestUtilities.TwoClientHarness harness = new();
			harness.LoadInitialText("hello world");
			IntervalCollection collection = harness.ClientA.GetIntervalCollection("test");

			List<(int start, int end, bool local)> log = new();
			collection.OnAddInterval += (_, args) =>
			{
				int? start = args.Interval.StartPosition;
				int? end = args.Interval.EndPosition;
				log.Add((start ?? -1, end ?? -1, args.Local));
			};

			collection.Add(0, 1, intervalId: "i1");

			(int start, int end, bool local) e = Assert.Single(log);
			Assert.Equal(0, e.start);
			Assert.Equal(1, e.end);
			Assert.True(e.local);

			// ACK of the outbound op must NOT fire the event a second time.
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			Assert.Single(log);
		}

		// Ported from intervalCollection.events.spec.ts —
		// "addInterval / is emitted on ack of a remote add"
		[Fact]
		public void AddInterval_RemoteAdd_EmitsOnceWithLocalFalse()
		{
			PortedTestUtilities.TwoClientHarness harness = new();
			harness.LoadInitialText("hello world");
			IntervalCollection collectionA = harness.ClientA.GetIntervalCollection("test");

			List<(int start, int end, bool local)> log = new();
			collectionA.OnAddInterval += (_, args) =>
			{
				int? start = args.Interval.StartPosition;
				int? end = args.Interval.EndPosition;
				log.Add((start ?? -1, end ?? -1, args.Local));
			};

			IntervalCollection collectionB = harness.ClientB.GetIntervalCollection("test");
			collectionB.Add(0, 1, intervalId: "remote");
			Assert.Empty(log);

			harness.DeliverBtoA(Assert.Single(harness.SenderB.Sent), refSeq: harness.CurrentServerSeq);
			(int start, int end, bool local) e = Assert.Single(log);
			Assert.Equal(0, e.start);
			Assert.Equal(1, e.end);
			Assert.False(e.local);
		}

		// Ported from intervalCollection.events.spec.ts —
		// "deleteInterval / is emitted on initial local delete"
		[Fact]
		public void DeleteInterval_LocalDelete_EmitsOnceWithLocalTrue()
		{
			PortedTestUtilities.TwoClientHarness harness = new();
			harness.LoadInitialText("hello world");
			IntervalCollection collection = harness.ClientA.GetIntervalCollection("test");
			collection.Add(0, 1, intervalId: "i1");
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderA.Sent.Clear();

			List<(int? start, int? end, bool local)> log = new();
			collection.OnDeleteInterval += (_, args) =>
			{
				log.Add((args.PreviousStart, args.PreviousEnd, args.Local));
			};

			collection.RemoveIntervalById("i1");

			(int? start, int? end, bool local) e = Assert.Single(log);
			Assert.Equal(0, e.start);
			Assert.Equal(1, e.end);
			Assert.True(e.local);

			// ACK must not double-fire.
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			Assert.Single(log);
		}

		// Ported from intervalCollection.events.spec.ts —
		// "deleteInterval / is emitted on ack of a remote delete"
		[Fact]
		public void DeleteInterval_RemoteDelete_EmitsOnceWithLocalFalse()
		{
			PortedTestUtilities.TwoClientHarness harness = new();
			harness.LoadInitialText("hello world");
			IntervalCollection collectionA = harness.ClientA.GetIntervalCollection("test");
			collectionA.Add(0, 1, intervalId: "i1");
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderA.Sent.Clear();

			List<(int? start, int? end, bool local)> log = new();
			collectionA.OnDeleteInterval += (_, args) =>
			{
				log.Add((args.PreviousStart, args.PreviousEnd, args.Local));
			};

			IntervalCollection collectionB = harness.ClientB.GetIntervalCollection("test");
			collectionB.RemoveIntervalById("i1");
			Assert.Empty(log);

			harness.DeliverBtoA(Assert.Single(harness.SenderB.Sent), refSeq: harness.CurrentServerSeq);
			(int? start, int? end, bool local) e = Assert.Single(log);
			Assert.Equal(0, e.start);
			Assert.Equal(1, e.end);
			Assert.False(e.local);
		}

		// Ported from intervalCollection.events.spec.ts —
		// "changeInterval / is emitted on initial local change"
		[Fact]
		public void ChangeInterval_LocalChange_EmitsWithPreviousEndpoints()
		{
			PortedTestUtilities.TwoClientHarness harness = new();
			harness.LoadInitialText("hello world");
			IntervalCollection collection = harness.ClientA.GetIntervalCollection("test");
			collection.Add(0, 1, intervalId: "i1");
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderA.Sent.Clear();

			List<(int? previousStart, int? previousEnd, bool local, bool slide)> log = new();
			collection.OnChangeInterval += (_, args) =>
			{
				log.Add((args.PreviousStart, args.PreviousEnd, args.Local, args.Slide));
			};

			collection.Change("i1", newStart: 2, newEnd: 3);

			(int? previousStart, int? previousEnd, bool local, bool slide) e = Assert.Single(log);
			Assert.Equal(0, e.previousStart);
			Assert.Equal(1, e.previousEnd);
			Assert.True(e.local);
			Assert.False(e.slide);
		}

		// Ported from intervalCollection.events.spec.ts —
		// "changeInterval / is emitted on a remote change"
		[Fact]
		public void ChangeInterval_RemoteChange_EmitsWithPreviousEndpointsAndLocalFalse()
		{
			PortedTestUtilities.TwoClientHarness harness = new();
			harness.LoadInitialText("hello world");
			IntervalCollection collectionA = harness.ClientA.GetIntervalCollection("test");
			collectionA.Add(0, 1, intervalId: "i1");
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderA.Sent.Clear();

			List<(int? previousStart, int? previousEnd, bool local, bool slide)> log = new();
			collectionA.OnChangeInterval += (_, args) =>
			{
				log.Add((args.PreviousStart, args.PreviousEnd, args.Local, args.Slide));
			};

			IntervalCollection collectionB = harness.ClientB.GetIntervalCollection("test");
			collectionB.Change("i1", newStart: 2, newEnd: 3);
			harness.DeliverBtoA(Assert.Single(harness.SenderB.Sent), refSeq: harness.CurrentServerSeq);

			(int? previousStart, int? previousEnd, bool local, bool slide) e = Assert.Single(log);
			Assert.Equal(0, e.previousStart);
			Assert.Equal(1, e.previousEnd);
			Assert.False(e.local);
			Assert.False(e.slide);
		}

		// Ported from intervalCollection.events.spec.ts —
		// "propertyChanged / includes prior property values"
		[Fact]
		public void PropertyChanged_LocalChange_ExposesPreviousValuesInDeltas()
		{
			// V2-A13 regression at the event contract. When a property is
			// mutated, propertyChanged's delta bag reports the PREVIOUS
			// value at each key — set to null for keys that had no prior
			// value.
			PortedTestUtilities.TwoClientHarness harness = new();
			harness.LoadInitialText("hello world");
			IntervalCollection collection = harness.ClientA.GetIntervalCollection("test");
			collection.Add(0, 1, intervalId: "i1", properties: new PropertySet() { ["color"] = "red" });
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderA.Sent.Clear();

			List<IntervalPropertyChangedEventArgs> log = new();
			collection.OnPropertyChanged += (_, args) => log.Add(args);

			collection.ChangeProperties("i1", new PropertySet()
			{
				["color"] = "blue",
				["added"] = "new",
			});

			IntervalPropertyChangedEventArgs e = Assert.Single(log);
			Assert.True(e.Local);
			Assert.Equal("blue", e.ChangedProps["color"]);
			Assert.Equal("new", e.ChangedProps["added"]);
			// PropertyDeltas carries the PREVIOUS value at each key.
			Assert.Equal("red", e.PropertyDeltas["color"]);
			Assert.Null(e.PropertyDeltas["added"]);
		}

		// Ported from intervalCollection.events.spec.ts —
		// "propertyChanged / null removes and reports previous"
		[Fact]
		public void PropertyChanged_NullValue_RemovesPropertyAndReportsPrevious()
		{
			PortedTestUtilities.TwoClientHarness harness = new();
			harness.LoadInitialText("hello world");
			IntervalCollection collection = harness.ClientA.GetIntervalCollection("test");
			SequenceInterval interval = collection.Add(0, 1, intervalId: "i1", properties: new PropertySet() { ["color"] = "red" });
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderA.Sent.Clear();

			List<IntervalPropertyChangedEventArgs> log = new();
			collection.OnPropertyChanged += (_, args) => log.Add(args);

			collection.ChangeProperties("i1", new PropertySet() { ["color"] = null });

			IntervalPropertyChangedEventArgs e = Assert.Single(log);
			Assert.Null(e.ChangedProps["color"]);
			Assert.Equal("red", e.PropertyDeltas["color"]);
			// Interval no longer has the property.
			Assert.False(interval.Properties?.ContainsKey("color") ?? false);
		}

		// Ported from intervalCollection.events.spec.ts (combined test) —
		// remote combined change fires changeInterval with previous
		// endpoints AND propertyChanged with previous values.
		[Fact]
		public void RemoteCombinedChange_FiresChangeAndPropertyChangedWithPreviousValues()
		{
			PortedTestUtilities.TwoClientHarness harness = new();
			harness.LoadInitialText("hello world");
			IntervalCollection collectionA = harness.ClientA.GetIntervalCollection("test");
			collectionA.Add(0, 1, intervalId: "i1", properties: new PropertySet() { ["color"] = "red" });
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderA.Sent.Clear();

			List<(int? previousStart, int? previousEnd)> changeLog = new();
			List<IntervalPropertyChangedEventArgs> propertyLog = new();
			collectionA.OnChangeInterval += (_, args) => changeLog.Add((args.PreviousStart, args.PreviousEnd));
			collectionA.OnPropertyChanged += (_, args) => propertyLog.Add(args);

			IntervalCollection collectionB = harness.ClientB.GetIntervalCollection("test");
			collectionB.Change("i1", newStart: 3, newEnd: 5, props: new PropertySet() { ["color"] = "blue" });
			harness.DeliverBtoA(Assert.Single(harness.SenderB.Sent), refSeq: harness.CurrentServerSeq);

			(int? previousStart, int? previousEnd) change = Assert.Single(changeLog);
			Assert.Equal(0, change.previousStart);
			Assert.Equal(1, change.previousEnd);

			IntervalPropertyChangedEventArgs property = Assert.Single(propertyLog);
			Assert.Equal("blue", property.ChangedProps["color"]);
			Assert.Equal("red", property.PropertyDeltas["color"]);
		}
	}
}
