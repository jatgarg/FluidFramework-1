// -----------------------------------------------------------------------------
// Regression tests for text segment split property isolation.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SegmentSplitPropertyIsolationTests
	{
		[Fact]
		public void SplitTextSegment_AnnotateLeft_DoesNotAffectRight()
		{
			TextSegment left = TextSegment.Make("hello world", new PropertySet()
			{
				["bold"] = true,
			});

			TextSegment right = Assert.IsType<TextSegment>(left.SplitAt(5));
			left.Properties!["color"] = "blue";

			Assert.Equal("hello", left.Text);
			Assert.Equal(" world", right.Text);
			Assert.NotSame(left.Properties, right.Properties);
			Assert.True(Assert.IsType<bool>(left.Properties!["bold"]));
			Assert.True(Assert.IsType<bool>(right.Properties!["bold"]));
			Assert.Equal("blue", Assert.IsType<string>(left.Properties["color"]));
			Assert.False(right.Properties!.ContainsKey("color"));
		}

		[Fact]
		public void SplitTextSegment_ModifyRightProp_DoesNotAffectLeft()
		{
			TextSegment left = TextSegment.Make("hello world", new PropertySet()
			{
				["bold"] = true,
			});

			TextSegment right = Assert.IsType<TextSegment>(left.SplitAt(5));
			right.Properties!["italic"] = true;

			Assert.NotSame(left.Properties, right.Properties);
			Assert.True(Assert.IsType<bool>(left.Properties!["bold"]));
			Assert.True(Assert.IsType<bool>(right.Properties!["bold"]));
			Assert.True(Assert.IsType<bool>(right.Properties["italic"]));
			Assert.False(left.Properties.ContainsKey("italic"));
		}

		[Fact]
		public void SplitTextSegment_CopyMetadataTo_DeepCopiesProps()
		{
			var source = new CopyMetadataSegment(new PropertySet()
			{
				["bold"] = true,
			});
			var destination = new CopyMetadataSegment();

			source.CopyMetadataToSegment(destination);
			destination.Properties!["italic"] = true;

			Assert.NotSame(source.Properties, destination.Properties);
			Assert.True(Assert.IsType<bool>(source.Properties!["bold"]));
			Assert.True(Assert.IsType<bool>(destination.Properties!["bold"]));
			Assert.True(Assert.IsType<bool>(destination.Properties["italic"]));
			Assert.False(source.Properties.ContainsKey("italic"));
		}

		[Fact]
		public void SplitTextSegment_NoProps_NoException()
		{
			TextSegment left = TextSegment.Make("hello");

			TextSegment right = Assert.IsType<TextSegment>(left.SplitAt(2));

			Assert.Equal("he", left.Text);
			Assert.Equal("llo", right.Text);
			Assert.Null(left.Properties);
			Assert.Null(right.Properties);
		}

		[Fact]
		public void SplitTextSegment_NestedPropsShallowCloned_ToMatchTS()
		{
			List<string> tags = new() { "a", "b" };
			TextSegment left = TextSegment.Make("hello", new PropertySet()
			{
				["tags"] = tags,
			});

			TextSegment right = Assert.IsType<TextSegment>(left.SplitAt(2));
			Assert.NotSame(left.Properties, right.Properties);
			Assert.Same(left.Properties!["tags"], right.Properties!["tags"]);

			Assert.Same(tags, Assert.IsType<List<string>>(right.Properties["tags"]));
			Assert.Equal(new[] { "a", "b" }, Assert.IsType<List<string>>(left.Properties["tags"]));
		}

		[Fact]
		public void InsertMidSegment_TS_ParityCheck()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "Blocker");
			sharedString.AnnotateRange(0, 7, new PropertySet()
			{
				["bold"] = true,
			});

			sharedString.InsertText(3, "!");

			Assert.Equal("Blo!cker", sharedString.GetText());
			TextSegment left = GetTextSegmentAt(sharedString, 0);
			TextSegment inserted = GetTextSegmentAt(sharedString, 3);
			TextSegment right = GetTextSegmentAt(sharedString, 4);
			Assert.Equal("Blo", left.Text);
			Assert.Equal("!", inserted.Text);
			Assert.Equal("cker", right.Text);
			Assert.Null(inserted.Properties);
			Assert.NotSame(left.Properties, right.Properties);
			Assert.True(Assert.IsType<bool>(left.Properties!["bold"]));
			Assert.True(Assert.IsType<bool>(right.Properties!["bold"]));

			sharedString.AnnotateRange(4, 8, new PropertySet()
			{
				["italic"] = true,
			});

			Assert.True(Assert.IsType<bool>(left.Properties["bold"]));
			Assert.False(left.Properties.ContainsKey("italic"));
			PropertySet rightProperties = GetRequiredProperties(sharedString, 4);
			Assert.True(Assert.IsType<bool>(rightProperties["bold"]));
			Assert.True(Assert.IsType<bool>(rightProperties["italic"]));
		}

		private static TextSegment GetTextSegmentAt(SharedString sharedString, int position)
		{
			var containingSegment = sharedString.GetContainingSegment(position);
			Assert.NotNull(containingSegment);
			return Assert.IsType<TextSegment>(containingSegment!.Value.segment);
		}

		private static PropertySet GetRequiredProperties(SharedString sharedString, int position)
		{
			PropertySet? properties = sharedString.GetPropertiesAtPosition(position);
			Assert.NotNull(properties);
			return properties!;
		}

		private sealed class CopyMetadataSegment : BaseSegment
		{
			public CopyMetadataSegment(PropertySet? properties = null)
				: base(properties)
			{
				CachedLength = 1;
			}

			public override string Type => "CopyMetadataSegment";

			public void CopyMetadataToSegment(ISegment segment)
			{
				CopyMetadataTo(segment);
			}

			public override ISegment Clone()
			{
				var clone = new CopyMetadataSegment();
				CopyMetadataTo(clone);
				return clone;
			}

			public override object ToJSONObject()
			{
				return Type;
			}

			protected override BaseSegment? CreateSplitSegmentAt(int pos)
			{
				return pos > 0 ? new CopyMetadataSegment() : null;
			}
		}
	}
}
