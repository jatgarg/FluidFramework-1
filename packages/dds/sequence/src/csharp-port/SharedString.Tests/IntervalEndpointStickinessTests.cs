// -----------------------------------------------------------------------------
// TS-parity interval endpoint stickiness and side semantics.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Text.Json;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using IntervalSide = Microsoft.Office.Web.Fluid.Intervals.Side;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class IntervalEndpointStickinessTests
	{
		[Fact]
		public void Add_DefaultStickiness_MatchesTS()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdefghijklmnop");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");

			SequenceInterval interval = collection.Add(5, 10, intervalId: "i1");

			// TS-parity: default stickiness is End (Finding IC2/IC3).
			Assert.Equal(IntervalStickiness.End, interval.Stickiness);
			Assert.Equal(IntervalSide.Before, interval.StartSide);
			Assert.Equal(IntervalSide.Before, interval.EndSide);
			Assert.Equal(SlidingPreference.Forward, interval.Start.SlidingPreference);
			Assert.Equal(SlidingPreference.Forward, interval.End.SlidingPreference);
			Assert.Equal(0x10 | 0x40, (int)interval.Start.RefType);
			Assert.Equal(0x20 | 0x40, (int)interval.End.RefType);
		}

		[Fact]
		public void Add_Stickiness_End_InsertAtStart_EndpointStaysRight()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdefghijklmnop");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval interval = collection.Add(5, 10, intervalId: "i1", stickiness: IntervalStickiness.End);

			sharedString.InsertText(5, "XX");

			AssertIntervalPositions(interval, 7, 12);
		}

		[Fact]
		public void Add_Stickiness_Start_InsertAtStart_EndpointExtendsLeft()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdefghijklmnop");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval interval = collection.Add(5, 10, intervalId: "i1", stickiness: IntervalStickiness.Start);

			sharedString.InsertText(5, "XX");

			AssertIntervalPositions(interval, 5, 12);
		}

		[Fact]
		public void Add_Stickiness_Full_InsertAtBothEnds()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdefghijklmnop");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval interval = collection.Add(5, 10, intervalId: "i1", stickiness: IntervalStickiness.Full);

			sharedString.InsertText(5, "XX");
			sharedString.InsertText(interval.EndPosition!.Value, "YY");

			AssertIntervalPositions(interval, 5, 14);
		}

		[Fact]
		public void Add_Stickiness_None_InsertAtEndpoints_ExcludedBoth()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdefghijklmnop");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval interval = collection.Add(5, 10, intervalId: "i1", stickiness: IntervalStickiness.None);

			sharedString.InsertText(5, "XX");
			sharedString.InsertText(interval.EndPosition!.Value, "YY");

			AssertIntervalPositions(interval, 7, 12);
		}

		[Fact]
		public void EndpointSide_Before_InsertAtPos_ExcludedFromRange()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdefghijklmnop");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval interval = collection.Add(5, IntervalSide.Before, 10, IntervalSide.Before, intervalId: "i1");

			sharedString.InsertText(5, "XX");

			Assert.Equal(IntervalSide.Before, interval.StartSide);
			AssertIntervalPositions(interval, 7, 12);
		}

		[Fact]
		public void EndpointSide_After_InsertAtPos_IncludedInRange()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdefghijklmnop");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval interval = collection.Add(5, IntervalSide.After, 10, IntervalSide.Before, intervalId: "i1");

			sharedString.InsertText(5, "XX");

			Assert.Equal(IntervalSide.After, interval.StartSide);
			AssertIntervalPositions(interval, 5, 12);
		}

		[Fact]
		public void Interval_WireRoundTrip_PreservesStickinessAndSides()
		{
			foreach (IntervalStickiness stickiness in Enum.GetValues<IntervalStickiness>())
			{
				foreach (IntervalSide startSide in Enum.GetValues<IntervalSide>())
				{
					foreach (IntervalSide endSide in Enum.GetValues<IntervalSide>())
					{
						IntervalAddOpMsg add = new()
						{
							CollectionName = "comments",
							IntervalId = $"{(int)stickiness}-{(int)startSide}-{(int)endSide}",
							Start = 5,
							End = 10,
							IntervalType = IntervalType.SlideOnRemove,
							Stickiness = stickiness,
							StartSide = startSide,
							EndSide = endSide,
						};

						string json = SharedStringOpSerializer.Serialize(add);
						// TS wire: op is wrapped in IntervalCollectionMap "act" envelope;
						// stickiness/startSide/endSide live at value.value.* per
						// intervalCollection.ts's serialized interval header.
						using JsonDocument document = JsonDocument.Parse(json);
						JsonElement payload = document.RootElement.GetProperty("value").GetProperty("value");
						Assert.Equal((int)stickiness, payload.GetProperty("stickiness").GetInt32());
						Assert.Equal((int)startSide, payload.GetProperty("startSide").GetInt32());
						Assert.Equal((int)endSide, payload.GetProperty("endSide").GetInt32());

						IntervalAddOpMsg roundTripped = Assert.IsType<IntervalAddOpMsg>(SharedStringOpSerializer.Deserialize(json));
						Assert.Equal(stickiness, roundTripped.Stickiness);
						Assert.Equal(startSide, roundTripped.StartSide);
						Assert.Equal(endSide, roundTripped.EndSide);
					}
				}
			}

			IntervalChangeOpMsg change = new()
			{
				CollectionName = "comments",
				IntervalId = "change",
				Start = 1,
				End = 3,
				Stickiness = IntervalStickiness.Full,
				StartSide = IntervalSide.After,
				EndSide = IntervalSide.Before,
			};

			IntervalChangeOpMsg changeRoundTripped = Assert.IsType<IntervalChangeOpMsg>(
				SharedStringOpSerializer.Deserialize(SharedStringOpSerializer.Serialize(change)));
			Assert.Equal(IntervalStickiness.Full, changeRoundTripped.Stickiness);
			Assert.Equal(IntervalSide.After, changeRoundTripped.StartSide);
			Assert.Equal(IntervalSide.Before, changeRoundTripped.EndSide);
		}

		private static SharedString CreateSharedStringWithText(string text)
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, text);
			return sharedString;
		}

		private static void AssertIntervalPositions(SequenceInterval interval, int start, int end)
		{
			Assert.Equal((int?)start, interval.StartPosition);
			Assert.Equal((int?)end, interval.EndPosition);
		}
	}
}
