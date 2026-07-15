// -----------------------------------------------------------------------------
// Marker tests for TS parity: markerId is optional.
// -----------------------------------------------------------------------------

#nullable enable

using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SharedStringMarkerWithoutIdTests
	{
		[Fact]
		public void InsertMarker_NoProps_DoesNotThrow()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "abc");

			sharedString.InsertMarker(1, ReferenceType.Simple);

			Assert.Equal("abc", sharedString.GetText());
			Assert.Equal(4, sharedString.GetLength());
			Marker marker = AssertMarkerAt(sharedString, 1);
			Assert.Equal(ReferenceType.Simple, marker.RefType);
			Assert.Null(marker.Properties);
		}

		[Fact]
		public void InsertMarker_PropsWithoutMarkerId_DoesNotThrow()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "abc");
			PropertySet props = new()
			{
				["nodeType"] = "Paragraph",
			};

			sharedString.InsertMarker(2, ReferenceType.Tile, props);

			Assert.Equal("abc", sharedString.GetText());
			Assert.Equal(4, sharedString.GetLength());
			Marker marker = AssertMarkerAt(sharedString, 2);
			Assert.NotNull(marker.Properties);
			Assert.Equal("Paragraph", marker.Properties!["nodeType"]);
			Assert.False(marker.Properties.ContainsKey(Marker.reservedMarkerIdKey));
		}

		[Fact]
		public void InsertMarker_NoMarkerId_NotFindableById()
		{
			var sharedString = new SharedString();

			sharedString.InsertMarker(0, ReferenceType.Simple);

			Assert.Null(sharedString.GetMarkerFromId("anything"));
		}

		[Fact]
		public void InsertMarker_MixedWithAndWithoutId_BothWork()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "ab");

			sharedString.InsertMarker(
				1,
				ReferenceType.Tile,
				new PropertySet()
				{
					[Marker.reservedTileLabelsKey] = new[] { "anonymous" },
				});
			sharedString.InsertMarker(
				2,
				ReferenceType.Tile,
				new PropertySet()
				{
					[Marker.reservedMarkerIdKey] = "named",
					[Marker.reservedTileLabelsKey] = new[] { "namedLabel" },
				});

			Marker anonymousMarker = AssertMarkerAt(sharedString, 1);
			Marker? namedMarker = sharedString.GetMarkerFromId("named");
			Assert.NotNull(namedMarker);
			Assert.Null(sharedString.GetMarkerFromId("anonymous"));
			Assert.Equal("ab", sharedString.GetText());
			Assert.Equal(4, sharedString.GetLength());
			Assert.True(anonymousMarker.HasTileLabel("anonymous"));
			Assert.Equal(2, sharedString.GetPositionOfMarker(namedMarker!));
		}

		[Fact]
		public void RemoveMarker_WithoutId_NoIndexCrash()
		{
			var removeString = new SharedString();
			removeString.InsertText(0, "abc");
			removeString.InsertMarker(1, ReferenceType.Simple);

			removeString.DeleteText(1, 2);

			Assert.Equal("abc", removeString.GetText());
			Assert.Equal(3, removeString.GetLength());
			Assert.Null(removeString.GetMarkerFromId("anything"));

			var obliterateString = new SharedString();
			obliterateString.InsertText(0, "abc");
			obliterateString.InsertMarker(1, ReferenceType.Simple);

			obliterateString.ObliterateRange(1, 2);

			Assert.Equal("abc", obliterateString.GetText());
			Assert.Equal(3, obliterateString.GetLength());
			Assert.Null(obliterateString.GetMarkerFromId("anything"));
		}

		private static Marker AssertMarkerAt(SharedString sharedString, int position)
		{
			var segmentInfo = sharedString.GetContainingSegment(position);
			Assert.NotNull(segmentInfo);
			Marker marker = Assert.IsType<Marker>(segmentInfo!.Value.segment);
			Assert.Equal(0, segmentInfo.Value.offsetInSegment);
			return marker;
		}
	}
}
