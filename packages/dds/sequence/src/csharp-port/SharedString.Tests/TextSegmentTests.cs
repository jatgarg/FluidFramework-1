// -----------------------------------------------------------------------------
// Unit tests for TextSegment behaviors that are hard to exercise through the
// SharedString public API — specifically metadata preservation on Clone.
// -----------------------------------------------------------------------------

using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class TextSegmentTests
	{
		[Fact]
		public void Clone_PreservesSequenceMetadata()
		{
			TextSegment original = TextSegment.Make("hello world");
			original.Seq = 42L;
			original.ClientId = 7;
			original.LocalSeq = 3L;

			TextSegment clone = original.Clone();

			Assert.Equal("hello world", clone.Text);
			Assert.Equal(42L, clone.Seq);
			Assert.Equal(7, clone.ClientId);
			Assert.Equal(3L, clone.LocalSeq);
		}

		[Fact]
		public void Clone_WithSlice_PreservesSequenceMetadata()
		{
			TextSegment original = TextSegment.Make("hello world");
			original.Seq = 100L;
			original.ClientId = 12;

			TextSegment sliced = original.Clone(0, 5);

			Assert.Equal("hello", sliced.Text);
			Assert.Equal(100L, sliced.Seq);
			Assert.Equal(12, sliced.ClientId);
		}

		[Fact]
		public void Clone_PreservesRemovalMetadata()
		{
			TextSegment original = TextSegment.Make("hello");
			original.Seq = 5L;
			original.ClientId = 3;
			original.RemovedSeq = 20L;
			original.RemovedClientId = 8;
			original.RemovedLocalSeq = 15L;

			TextSegment clone = original.Clone();

			Assert.Equal(20L, clone.RemovedSeq);
			Assert.Equal(8, clone.RemovedClientId);
			Assert.Equal(15L, clone.RemovedLocalSeq);
		}

		[Fact]
		public void Clone_PreservesProperties()
		{
			PropertySet props = new() { ["color"] = "red", ["bold"] = true };
			TextSegment original = TextSegment.Make("hello", props);

			TextSegment clone = original.Clone();

			Assert.NotNull(clone.Properties);
			Assert.Equal("red", clone.Properties!["color"]);
			Assert.Equal(true, clone.Properties["bold"]);
		}
	}
}
