#nullable enable

using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using static Microsoft.Office.Web.Fluid.Tests.PortedTestUtilities;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class ClientSearchForMarkerTestsFromTS
	{
		// Ported from packages/dds/merge-tree/src/test/client.searchForMarker.spec.ts — "Should return marker at the search position in either direction"
		[Fact]
		public void SearchAtMarkerPosition_ReturnsMarkerInEitherDirection()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdefg");
			sharedString.InsertMarker(4, ReferenceType.Tile, MarkerProps("marker", "Eop"));

			Marker? forward = sharedString.SearchForMarker(4, "Eop", forwards: true);
			Marker? backward = sharedString.SearchForMarker(4, "Eop", forwards: false);

			Assert.NotNull(forward);
			Assert.Same(forward, backward);
			Assert.Equal(4, sharedString.GetPositionOfMarker(forward!));
		}

		// Ported from packages/dds/merge-tree/src/test/client.searchForMarker.spec.ts — "Should return the marker at the search position in either direction from multiple blocks"
		[Fact]
		public void SearchAtMarkerPositionAcrossMultipleSegments_ReturnsMarkerInEitherDirection()
		{
			SharedString sharedString = CreateSharedStringWithText("abcd");
			sharedString.InsertMarker(4, ReferenceType.Tile, MarkerProps("marker", "Eop"));
			sharedString.InsertText(5, "efg");

			Marker? forward = sharedString.SearchForMarker(4, "Eop", forwards: true);
			Marker? backward = sharedString.SearchForMarker(4, "Eop", forwards: false);

			Assert.NotNull(forward);
			Assert.Same(forward, backward);
			Assert.Equal(4, sharedString.GetPositionOfMarker(forward!));
		}

		// Ported from packages/dds/merge-tree/src/test/client.searchForMarker.spec.ts — "Should be able to find forward marker position based on label"
		[Fact]
		public void ForwardSearch_UsesTileLabel()
		{
			SharedString sharedString = new();
			sharedString.InsertMarker(0, ReferenceType.Tile, MarkerProps("m1", "EOP"));
			sharedString.InsertText(0, "abc");

			Marker? marker = sharedString.SearchForMarker(0, "EOP", forwards: true);

			Assert.NotNull(marker);
			Assert.Equal("m1", marker!.GetId());
			Assert.Equal(3, sharedString.GetPositionOfMarker(marker));
		}

		// Ported from packages/dds/merge-tree/src/test/client.searchForMarker.spec.ts — "Should be able to find backward marker position based on label from client with multiple marker"
		[Fact]
		public void BackwardSearch_UsesNearestMatchingTileLabel()
		{
			SharedString sharedString = CreateSharedStringWithText("abcdef");
			sharedString.InsertMarker(1, ReferenceType.Tile, MarkerProps("first", "EOP"));
			sharedString.InsertMarker(5, ReferenceType.Tile, MarkerProps("second", "EOP"));

			Marker? marker = sharedString.SearchForMarker(sharedString.GetLength() - 1, "EOP", forwards: false);

			Assert.NotNull(marker);
			Assert.Equal("second", marker!.GetId());
			Assert.Equal(5, sharedString.GetPositionOfMarker(marker));
		}

		// Ported from packages/dds/merge-tree/src/test/client.searchForMarker.spec.ts — "Should be able to find forward marker position with multiple segments and markers"
		[Fact]
		public void ForwardSearch_WithManySegmentsAndMarkers_FindsNearestMarkerAtOrAfterStart()
		{
			SharedString sharedString = new();
			for (int i = 0; i < 24; i++)
			{
				sharedString.InsertText(sharedString.GetLength(), i.ToString("00"));
			}

			for (int position = 0, id = 0; position <= sharedString.GetLength(); position += 5, id++)
			{
				sharedString.InsertMarker(position, ReferenceType.Tile, MarkerProps($"m{id}", "EOP"));
			}

			for (int position = 0; position < sharedString.GetLength(); position++)
			{
				Marker? marker = sharedString.SearchForMarker(position, "EOP", forwards: true);

				Assert.NotNull(marker);
				int markerPosition = sharedString.GetPositionOfMarker(marker!)!.Value;
				Assert.True(markerPosition >= position, $"Expected marker at or after {position}, got {markerPosition}.");
			}
		}

		// Ported from packages/dds/merge-tree/src/test/client.searchForMarker.spec.ts — "Should be able to find backward marker position with multiple segments and markers"
		[Fact]
		public void BackwardSearch_WithManySegmentsAndMarkers_FindsNearestMarkerAtOrBeforeStart()
		{
			SharedString sharedString = new();
			for (int i = 0; i < 24; i++)
			{
				sharedString.InsertText(sharedString.GetLength(), i.ToString("00"));
			}

			for (int position = 0, id = 0; position <= sharedString.GetLength(); position += 5, id++)
			{
				sharedString.InsertMarker(position, ReferenceType.Tile, MarkerProps($"m{id}", "EOP"));
			}

			for (int position = sharedString.GetLength() - 1; position >= 0; position--)
			{
				Marker? marker = sharedString.SearchForMarker(position, "EOP", forwards: false);

				Assert.NotNull(marker);
				int markerPosition = sharedString.GetPositionOfMarker(marker!)!.Value;
				Assert.True(markerPosition <= position, $"Expected marker at or before {position}, got {markerPosition}.");
			}
		}

		// Ported from packages/dds/merge-tree/src/test/client.searchForMarker.spec.ts — "Should be able to find a marker at 0 searching at 0 in both directions"
		[Fact]
		public void SearchAtZero_FindsMarkerAtZeroBothDirections()
		{
			SharedString sharedString = CreateSharedStringWithText("abc");
			sharedString.InsertMarker(0, ReferenceType.Tile, MarkerProps("marker", "Eop"));

			Marker? forward = sharedString.SearchForMarker(0, "Eop", forwards: true);
			Marker? backward = sharedString.SearchForMarker(0, "Eop", forwards: false);

			Assert.NotNull(forward);
			Assert.Same(forward, backward);
			Assert.Equal(0, sharedString.GetPositionOfMarker(forward!));
		}

		// Ported from packages/dds/merge-tree/src/test/client.searchForMarker.spec.ts — "Should be able to find a marker at length-1 searching at length-1 in both directions"
		[Fact]
		public void SearchAtLastPosition_FindsMarkerAtLastPositionBothDirections()
		{
			SharedString sharedString = new();
			sharedString.InsertMarker(0, ReferenceType.Tile, MarkerProps("marker", "Eop"));
			sharedString.InsertText(0, "abc");
			int lastPosition = sharedString.GetLength() - 1;

			Marker? forward = sharedString.SearchForMarker(lastPosition, "Eop", forwards: true);
			Marker? backward = sharedString.SearchForMarker(lastPosition, "Eop", forwards: false);

			Assert.NotNull(forward);
			Assert.Same(forward, backward);
			Assert.Equal(lastPosition, sharedString.GetPositionOfMarker(forward!));
		}

		// Ported from packages/dds/merge-tree/src/test/client.searchForMarker.spec.ts — "Should return undefined when searching past the end of a string length 1"
		[Fact]
		public void SearchPastEndOfSingleMarkerString_ReturnsNull()
		{
			SharedString sharedString = new();
			sharedString.InsertMarker(0, ReferenceType.Tile, MarkerProps("marker", "Eop"));

			Assert.Null(sharedString.SearchForMarker(sharedString.GetLength(), "Eop", forwards: true));
		}

		// Ported from packages/dds/merge-tree/src/test/client.searchForMarker.spec.ts — "Should return undefined when searching before the start of a string length 1"
		[Fact]
		public void SearchBeforeStartOfSingleMarkerString_ReturnsNull()
		{
			SharedString sharedString = new();
			sharedString.InsertMarker(0, ReferenceType.Tile, MarkerProps("marker", "Eop"));

			Assert.Null(sharedString.SearchForMarker(-1, "Eop", forwards: false));
		}

		// Ported from packages/dds/merge-tree/src/test/client.searchForMarker.spec.ts — "Should return undefined when trying to find marker from text without the specified marker"
		[Fact]
		public void SearchTextWithoutMatchingMarker_ReturnsNullBothDirections()
		{
			SharedString sharedString = CreateSharedStringWithText("abc");

			Assert.Null(sharedString.SearchForMarker(1, "EOP", forwards: true));
			Assert.Null(sharedString.SearchForMarker(1, "EOP", forwards: false));
		}

		// Ported from packages/dds/merge-tree/src/test/client.searchForMarker.spec.ts — "Should return undefined when trying to find a removed marker"
		[Fact]
		public void SearchRemovedMarker_ReturnsNull()
		{
			SharedString sharedString = CreateSharedStringWithText("abc");
			sharedString.InsertMarker(1, ReferenceType.Tile, MarkerProps("marker", "Eop"));

			sharedString.DeleteText(1, 2);

			Assert.Null(sharedString.SearchForMarker(0, "Eop", forwards: true));
			Assert.Null(sharedString.GetMarkerFromId("marker"));
		}

		// Ported from packages/dds/merge-tree/src/test/client.searchForMarker.spec.ts — "Should be able to find remotely inserted marker"
		[Fact]
		public void SearchRemotelyInsertedMarker_FindsMarker()
		{
			SharedString sharedString = CreateAckedSharedString("client-a", "abc").SharedString;

			ProcessRemoteMarkerInsert(sharedString, 0, "marker", "Eop", refSeq: 1, seq: 2, clientId: "client-b");

			Marker? marker = sharedString.SearchForMarker(0, "Eop", forwards: true);
			Assert.NotNull(marker);
			Assert.Equal("marker", marker!.GetId());
			Assert.Equal(0, sharedString.GetPositionOfMarker(marker));
		}

		// Ported from packages/dds/merge-tree/src/test/client.searchForMarker.spec.ts — "Should not be able to find remotely removed marker"
		[Fact]
		public void SearchRemotelyRemovedMarker_ReturnsNull()
		{
			SharedString sharedString = CreateAckedSharedString("client-a", "abc").SharedString;
			ProcessRemoteMarkerInsert(sharedString, 0, "marker", "Eop", refSeq: 1, seq: 2, clientId: "client-b");

			ProcessRemoteRemove(sharedString, 0, 1, refSeq: 2, seq: 3, clientId: "client-b");

			Assert.Null(sharedString.SearchForMarker(0, "Eop", forwards: true));
			Assert.Null(sharedString.GetMarkerFromId("marker"));
		}
	}
}
