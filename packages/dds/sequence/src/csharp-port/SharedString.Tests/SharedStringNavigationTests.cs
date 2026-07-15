// -----------------------------------------------------------------------------
// Wave 13 navigation API tests for SharedString POC.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.IO;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SharedStringNavigationTests
	{
		[Fact]
		public void GetContainingSegment_ValidPosition_ReturnsSegmentAndOffset()
		{
			var sharedString = CreateSharedStringWithText("hello");

			(ISegment segment, int offsetInSegment) result = GetRequiredContainingSegment(sharedString, 2);

			AssertTextSegment(result.segment, "hello");
			Assert.Equal(2, result.offsetInSegment);
		}

		[Fact]
		public void GetContainingSegment_AfterMultipleInserts_ReturnsCorrectSegment()
		{
			var sharedString = CreateSharedStringWithText("hello");
			sharedString.InsertText(5, " world");

			(ISegment segment, int offsetInSegment) first = GetRequiredContainingSegment(sharedString, 1);
			(ISegment segment, int offsetInSegment) second = GetRequiredContainingSegment(sharedString, 6);

			AssertTextSegment(first.segment, "hello");
			Assert.Equal(1, first.offsetInSegment);
			AssertTextSegment(second.segment, " world");
			Assert.Equal(1, second.offsetInSegment);
			Assert.NotSame(first.segment, second.segment);
		}

		[Fact]
		public void GetContainingSegment_OutOfRange_ReturnsNull()
		{
			var sharedString = CreateSharedStringWithText("hello");

			Assert.Null(sharedString.GetContainingSegment(5));
			Assert.Null(sharedString.GetContainingSegment(-1));
		}

		[Fact]
		public void GetPosition_OfExistingSegment_ReturnsCurrentPosition()
		{
			var sharedString = CreateSharedStringWithText("hello");
			ISegment segment = GetRequiredContainingSegment(sharedString, 0).segment;

			Assert.Equal(0, sharedString.GetPosition(segment));
		}

		[Fact]
		public void GetPosition_AfterInsertBefore_ReturnsShiftedPosition()
		{
			var sharedString = CreateSharedStringWithText("hello");
			ISegment segment = GetRequiredContainingSegment(sharedString, 0).segment;

			sharedString.InsertText(0, "XY");

			Assert.Equal(2, sharedString.GetPosition(segment));
		}

		[Fact]
		public void GetPosition_OfRemovedSegment_ReturnsNull()
		{
			var sharedString = CreateSharedStringWithText("hello");
			ISegment segment = GetRequiredContainingSegment(sharedString, 0).segment;

			sharedString.DeleteText(0, 5);

			Assert.Null(sharedString.GetPosition(segment));
		}

		[Fact]
		public void GetRangeExtentsOfPosition_MidSegment_ReturnsFullSegmentBounds()
		{
			var sharedString = CreateSharedStringWithText("hello world");

			AssertRange(sharedString.GetRangeExtentsOfPosition(5), 0, 11);
		}

		[Fact]
		public void GetRangeExtentsOfPosition_AfterAnnotateSplit_ReturnsSplitBounds()
		{
			var sharedString = CreateSharedStringWithText("hello world");

			sharedString.AnnotateRange(0, 5, new PropertySet()
			{
				["color"] = "red",
			});

			AssertRange(sharedString.GetRangeExtentsOfPosition(3), 0, 5);
			AssertRange(sharedString.GetRangeExtentsOfPosition(7), 5, 11);
		}

		[Fact]
		public void GetRangeExtentsOfPosition_OutOfRange_ReturnsNull()
		{
			var sharedString = CreateSharedStringWithText("hello");

			Assert.Null(sharedString.GetRangeExtentsOfPosition(5));
			Assert.Null(sharedString.GetRangeExtentsOfPosition(-1));
		}

		[Fact]
		public void LocalReferencePositionToPosition_AtCreation_ReturnsOriginalPosition()
		{
			var sharedString = CreateSharedStringWithText("hello");
			SequenceInterval interval = CreateReferenceInterval(sharedString, 3);

			Assert.Equal(3, sharedString.LocalReferencePositionToPosition(interval.Start));
		}

		[Fact]
		public void LocalReferencePositionToPosition_AfterInsertBefore_ReturnsShiftedPosition()
		{
			var sharedString = CreateSharedStringWithText("hello");
			SequenceInterval interval = CreateReferenceInterval(sharedString, 3);

			sharedString.InsertText(0, "XY");

			Assert.Equal(5, sharedString.LocalReferencePositionToPosition(interval.Start));
		}

		[Fact]
		public void LocalReferencePositionToPosition_Detached_ReturnsMinusOne()
		{
			var sharedString = CreateSharedStringWithText("hello");
			SequenceInterval interval = CreateReferenceInterval(sharedString, 3);

			sharedString.DeleteText(0, sharedString.GetLength());

			// TS-parity: detached reference positions use -1 rather than null (Finding L7).
			Assert.Equal(MergeTree.MergeTree.DetachedReferencePosition, sharedString.LocalReferencePositionToPosition(interval.Start));
		}

		[Fact]
		public void AllApis_AgreeOnPosition_AfterEdits()
		{
			var sharedString = CreateSharedStringWithText("hello world");
			ISegment originalSegment = GetRequiredContainingSegment(sharedString, 3).segment;

			sharedString.InsertText(0, "XX");

			(ISegment segment, int offsetInSegment) shifted = GetRequiredContainingSegment(sharedString, 5);
			Assert.Same(originalSegment, shifted.segment);
			Assert.Equal(3, shifted.offsetInSegment);
			Assert.Equal(2, sharedString.GetPosition(originalSegment));
			AssertRange(sharedString.GetRangeExtentsOfPosition(5), 2, 13);
		}

		[Fact]
		public void AfterSnapshotLoad_NavigationApisWork()
		{
			var sharedString = new SharedString();
			sharedString.LoadFromSnapshot(LoadSimpleHelloSnapshot());

			(ISegment segment, int offsetInSegment) result = GetRequiredContainingSegment(sharedString, 3);

			AssertTextSegment(result.segment, "Hello, world!");
			Assert.Equal(3, result.offsetInSegment);
			AssertRange(sharedString.GetRangeExtentsOfPosition(3), 0, 13);
		}

		private static SharedString CreateSharedStringWithText(string text)
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, text);
			return sharedString;
		}

		private static SequenceInterval CreateReferenceInterval(SharedString sharedString, int position)
		{
			IntervalCollection collection = sharedString.GetIntervalCollection("navigation");
			return collection.Add(position, position);
		}

		private static (ISegment segment, int offsetInSegment) GetRequiredContainingSegment(SharedString sharedString, int position)
		{
			(ISegment segment, int offsetInSegment)? result = sharedString.GetContainingSegment(position);
			Assert.NotNull(result);
			return result.Value;
		}

		private static void AssertTextSegment(ISegment segment, string expectedText)
		{
			TextSegment textSegment = Assert.IsType<TextSegment>(segment);
			Assert.Equal(expectedText, textSegment.Text);
		}

		private static void AssertRange((int start, int end)? actual, int expectedStart, int expectedEnd)
		{
			Assert.NotNull(actual);
			Assert.Equal(expectedStart, actual.Value.start);
			Assert.Equal(expectedEnd, actual.Value.end);
		}

		private static SharedStringSnapshotDto LoadSimpleHelloSnapshot()
		{
			string fixturePath = Path.Combine(
				AppContext.BaseDirectory,
				"Fixtures",
				"simple-hello.snapshot.json.txt");
			return SharedStringSnapshotLoader.Parse(File.ReadAllText(fixturePath));
		}
	}
}
