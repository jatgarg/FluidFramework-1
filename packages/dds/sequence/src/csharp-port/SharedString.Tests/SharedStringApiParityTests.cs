// -----------------------------------------------------------------------------
// SharedString public API and event-surface parity regressions.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;
using System.Text.Json;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SharedStringApiParityTests
	{
		[Fact]
		public void InsertText_WithProperties_AppliesPropertiesAndEmitsSegmentRange()
		{
			SharedString sharedString = new();
			List<SequenceDeltaEventArgs> events = CaptureEvents(sharedString);
			PropertySet props = new()
			{
				["style"] = "bold",
			};

			sharedString.InsertText(0, "hello", props);

			Assert.Equal("hello", sharedString.GetText());
			Assert.Equal("bold", Assert.IsType<string>(sharedString.GetPropertiesAtPosition(0)!["style"]));
			SequenceDeltaEventArgs captured = Assert.Single(events);
			Assert.True(captured.IsLocal);
			Assert.Equal("insert", captured.OpType);
			SequenceDeltaRange range = Assert.Single(captured.Ranges);
			Assert.Same(range, captured.First);
			Assert.Same(range, captured.Last);
			Assert.Equal(0, range.Position);
			Assert.Equal(5, range.Length);
			TextSegment segment = Assert.IsType<TextSegment>(range.Segment);
			Assert.Equal("hello", segment.Text);
			Assert.Equal("bold", Assert.IsType<string>(segment.Properties!["style"]));
		}

		[Fact]
		public void InsertText_WithProperties_SerializesTextSegmentProps()
		{
			FakeFluidDataObjectSender sender = new();
			SharedString sharedString = new("shared-string", sender);

			sharedString.InsertText(0, "hi", new PropertySet()
			{
				["lang"] = "en",
			});

			var sent = Assert.Single(sender.Sent);
			using JsonDocument document = JsonDocument.Parse(sent.OpJson);
			JsonElement segment = document.RootElement.GetProperty("seg");
			Assert.Equal("hi", segment.GetProperty("text").GetString());
			Assert.Equal("en", segment.GetProperty("props").GetProperty("lang").GetString());
		}

		[Fact]
		public void RemoveText_Alias_RemovesRangeUsingTsName()
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, "hello world");

			sharedString.RemoveText(5, 11);

			Assert.Equal("hello", sharedString.GetText());
		}

		[Fact]
		public void ReplaceText_ReplacesRangeByInsertingThenRemoving()
		{
			FakeFluidDataObjectSender sender = new();
			SharedString sharedString = new("doc", sender);
			sharedString.InsertText(0, "hello world");
			sender.Sent.Clear();
			List<SequenceDeltaEventArgs> events = CaptureEvents(sharedString);

			sharedString.ReplaceText(6, 11, "there!");

			Assert.Equal("hello there!", sharedString.GetText());
			Assert.Collection(
				sender.Sent,
				insert => Assert.Equal("insert", insert.OpTypeName),
				remove => Assert.Equal("remove", remove.OpTypeName));
			Assert.Collection(
				events,
				insert =>
				{
					Assert.Equal("insert", insert.OpType);
					Assert.Equal(11, insert.Position);
					Assert.Equal("there!", insert.Text);
				},
				remove =>
				{
					Assert.Equal("remove", remove.OpType);
					Assert.Equal(6, remove.Position);
					Assert.Equal(5, remove.Length);
				});
		}

		[Fact]
		public void ReplaceText_ZeroAndReverseRanges_InsertAtMaxEndWithoutRemoving()
		{
			SharedString zeroRange = new();
			zeroRange.InsertText(0, "123");
			zeroRange.ReplaceText(1, 1, "\u00E4\u00C4");

			SharedString reverseRange = new();
			reverseRange.InsertText(0, "123");
			reverseRange.ReplaceText(2, 1, "aaa");

			Assert.Equal("1\u00E4\u00C423", zeroRange.GetText());
			Assert.Equal("12aaa3", reverseRange.GetText());
		}

		[Fact]
		public void InsertText_EmptyText_EmitsNoOpAndNoEvent()
		{
			// TS ref: packages/dds/merge-tree/src/client.ts insertSegmentLocal —
			// empty segments early-return undefined; TS emits neither a wire op
			// nor a sequence-delta event.
			var sender = new FakeFluidDataObjectSender();
			SharedString sharedString = new("shared-string", sender);
			int eventCount = 0;
			sharedString.OnSequenceDelta += (_, _) => eventCount++;

			sharedString.InsertText(0, string.Empty);

			Assert.Equal(0, sharedString.GetLength());
			Assert.Empty(sender.Sent);
			Assert.Equal(0, eventCount);
		}

		[Fact]
		public void ReplaceText_EmptyText_IsNoOp_DoesNotDelete()
		{
			// TS ref: packages/dds/sequence/src/sequence.ts replaceRange — the
			// remove is guarded by the insert op's truthiness. Empty text
			// produces no insert op, so no remove is issued. The whole
			// replaceText call is a no-op instead of a delete.
			var sender = new FakeFluidDataObjectSender();
			SharedString sharedString = new("shared-string", sender);
			sharedString.InsertText(0, "hello");
			sender.Sent.Clear();

			sharedString.ReplaceText(1, 4, string.Empty);

			Assert.Equal("hello", sharedString.GetText());
			Assert.Empty(sender.Sent);
		}

		[Fact]
		public void ReplaceText_WithProperties_AppliesReplacementProperties()
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, "hello world");

			sharedString.ReplaceText(0, 5, "hi", new PropertySet()
			{
				["style"] = "replacement",
			});

			Assert.Equal("hi world", sharedString.GetText());
			Assert.Equal("replacement", Assert.IsType<string>(sharedString.GetPropertiesAtPosition(0)!["style"]));
			Assert.Null(sharedString.GetPropertiesAtPosition(3));
		}

		[Fact]
		public void InsertTextRelative_AfterMarker_InsertsAfterMarker()
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, "ab");
			sharedString.InsertMarker(1, ReferenceType.Tile, MarkerProps("m1", "anchor"));

			sharedString.InsertTextRelative(new RelativePosition() { Id = "m1" }, "X");

			Assert.Equal("aXb", sharedString.GetText());
			Assert.Equal(1, sharedString.GetPositionOfMarker(sharedString.GetMarkerFromId("m1")!));
		}

		[Fact]
		public void InsertMarkerRelative_BeforeMarker_InsertsBeforeExistingMarker()
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, "ab");
			sharedString.InsertMarker(1, ReferenceType.Tile, MarkerProps("m1", "anchor"));
			Dictionary<string, object?> relativePosition = new()
			{
				["id"] = "m1",
				["before"] = true,
			};

			sharedString.InsertMarkerRelative(relativePosition, ReferenceType.Tile, MarkerProps("m0", "before"));

			Marker? before = sharedString.SearchForMarker(0, "before");
			Marker? anchor = sharedString.SearchForMarker(0, "anchor");
			Assert.NotNull(before);
			Assert.NotNull(anchor);
			Assert.Equal(1, sharedString.GetPositionOfMarker(before!));
			Assert.Equal(2, sharedString.GetPositionOfMarker(anchor!));
			Assert.Equal("a  b", sharedString.GetTextWithPlaceholders());
		}

		[Fact]
		public void SearchForMarker_TsSignatureSearchesByRequiredLabel()
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, "ab");
			sharedString.InsertMarker(1, ReferenceType.Tile, MarkerProps("other", "other"));
			sharedString.InsertMarker(2, ReferenceType.Tile, MarkerProps("target", "target"));

			Marker? marker = sharedString.SearchForMarker(0, "target");

			Assert.NotNull(marker);
			Assert.Equal("target", marker!.GetId());
		}

		[Fact]
		public void SearchForMarker_OmittedLegacyLabel_DoesNotMatchAnyTile()
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, "ab");
			sharedString.InsertMarker(1, ReferenceType.Tile, MarkerProps("target", "target"));

			Assert.Null(sharedString.SearchForMarker(0));
		}

		[Fact]
		public void GetTextWithPlaceholders_IncludesMarkerSpaces()
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, "ab");
			sharedString.InsertMarker(1, ReferenceType.Tile, MarkerProps("m1", "anchor"));

			Assert.Equal("ab", sharedString.GetText());
			Assert.Equal("a b", sharedString.GetTextWithPlaceholders());
			Assert.Equal(" b", sharedString.GetTextWithPlaceholders(1, sharedString.GetLength()));
		}

		[Fact]
		public void GetTextRangeWithMarkers_IncludesMarkerString()
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, "ab");
			sharedString.InsertMarker(1, ReferenceType.Tile, MarkerProps("m1", "anchor"));

			Assert.Equal("a\n\u200Eb", sharedString.GetTextRangeWithMarkers(0, sharedString.GetLength()));
		}

		[Fact]
		public void LocalRemoveSequenceDelta_RangesExposeRemovedSlice()
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, "hello");
			List<SequenceDeltaEventArgs> events = CaptureEvents(sharedString);

			sharedString.RemoveText(1, 4);

			SequenceDeltaEventArgs captured = Assert.Single(events);
			Assert.Equal("remove", captured.OpType);
			SequenceDeltaRange range = Assert.Single(captured.Ranges);
			Assert.Equal(1, range.Position);
			Assert.Equal(3, range.Length);
			Assert.Equal("ell", Assert.IsType<TextSegment>(range.Segment).Text);
		}

		[Fact]
		public void LocalAnnotateSequenceDelta_RangesExposeUpdatedSegmentAndPropertyDeltas()
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, "abcde", new PropertySet()
			{
				["color"] = "blue",
			});
			List<SequenceDeltaEventArgs> events = CaptureEvents(sharedString);

			sharedString.AnnotateRange(1, 4, new PropertySet()
			{
				["color"] = "red",
				["bold"] = true,
			});

			SequenceDeltaEventArgs captured = Assert.Single(events);
			Assert.Equal("annotate", captured.OpType);
			SequenceDeltaRange range = Assert.Single(captured.Ranges);
			TextSegment segment = Assert.IsType<TextSegment>(range.Segment);
			Assert.Equal("bcd", segment.Text);
			Assert.Equal("red", Assert.IsType<string>(segment.Properties!["color"]));
			Assert.True(Assert.IsType<bool>(segment.Properties["bold"]));
			Assert.Equal("blue", Assert.IsType<string>(range.PropertyDeltas!["color"]));
			Assert.Null(range.PropertyDeltas["bold"]);
		}

		[Fact]
		public void LocalMutation_DuringLocalSequenceDelta_ThrowsReentrancyDetected()
		{
			FakeFluidDataObjectSender sender = new();
			SharedString sharedString = new("doc", sender);
			sharedString.InsertText(0, "abcX");
			sender.Sent.Clear();
			sharedString.OnSequenceDelta += (target, args) =>
			{
				if (args.IsLocal && args.OpType == "insert" && args.Text == "e")
				{
					((SharedString)target).RemoveText(3, 4);
				}
			};

			InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
				() => sharedString.InsertText(4, "e"));

			Assert.Contains("Reentrancy detected", exception.Message);
			Assert.Equal("abcXe", sharedString.GetText());
			Assert.Single(sender.Sent);
		}

		private static List<SequenceDeltaEventArgs> CaptureEvents(SharedString sharedString)
		{
			List<SequenceDeltaEventArgs> events = new();
			sharedString.OnSequenceDelta += (sender, args) => events.Add(args);
			return events;
		}

		private static PropertySet MarkerProps(string markerId, params string[] tileLabels)
		{
			return new PropertySet()
			{
				[Marker.reservedMarkerIdKey] = markerId,
				[Marker.reservedTileLabelsKey] = tileLabels,
			};
		}
	}
}
