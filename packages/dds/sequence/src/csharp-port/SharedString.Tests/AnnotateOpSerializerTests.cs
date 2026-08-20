// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------

using System.Collections.Generic;
using System.Text.Json;
using Microsoft.Office.Web.Fluid;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class AnnotateOpSerializerTests
	{
		[Fact]
		public void AnnotateOp_JsonRoundTrip_PreservesFields()
		{
			MergeTreeAnnotateMsg op = new MergeTreeAnnotateMsg()
			{
				Pos1 = 0,
				Pos2 = 5,
				Props = new PropertySet()
				{
					["color"] = "red",
				},
			};

			MergeTreeAnnotateMsg roundTripped = RoundTrip(op);

			AssertNullableInt(0, roundTripped.Pos1);
			AssertNullableInt(5, roundTripped.Pos2);
			PropertySet? props = roundTripped.Props;
			Assert.NotNull(props);
			Assert.Single(props!);
			Assert.Equal("red", Assert.IsType<string>(props!["color"]));
		}

		[Fact]
		public void AnnotateOp_MultipleProps_RoundTrip()
		{
			MergeTreeAnnotateMsg op = new MergeTreeAnnotateMsg()
			{
				Pos1 = 2,
				Pos2 = 9,
				Props = new PropertySet()
				{
					["color"] = "red",
					["weight"] = 700,
					["bold"] = true,
					["font"] = "arial",
				},
			};

			MergeTreeAnnotateMsg roundTripped = RoundTrip(op);

			AssertNullableInt(2, roundTripped.Pos1);
			AssertNullableInt(9, roundTripped.Pos2);
			PropertySet? props = roundTripped.Props;
			Assert.NotNull(props);
			Assert.Equal(4, props!.Count);
			Assert.Equal("red", Assert.IsType<string>(props["color"]));
			Assert.Equal(700, Assert.IsType<int>(props["weight"]));
			Assert.True(Assert.IsType<bool>(props["bold"]));
			Assert.Equal("arial", Assert.IsType<string>(props["font"]));
		}

		[Fact]
		public void AnnotateOp_NullPropValue_RoundTrips()
		{
			MergeTreeAnnotateMsg op = new MergeTreeAnnotateMsg()
			{
				Pos1 = 1,
				Pos2 = 4,
				Props = new PropertySet()
				{
					["color"] = null,
				},
			};

			MergeTreeAnnotateMsg roundTripped = RoundTrip(op);

			PropertySet? props = roundTripped.Props;
			Assert.NotNull(props);
			Assert.True(props!.ContainsKey("color"));
			Assert.Null(props["color"]);
		}

		[Fact]
		public void AnnotateOp_EmptyProps_RoundTrips()
		{
			MergeTreeAnnotateMsg op = new MergeTreeAnnotateMsg()
			{
				Pos1 = 3,
				Pos2 = 3,
				Props = new PropertySet(),
			};

			MergeTreeAnnotateMsg roundTripped = RoundTrip(op);

			AssertNullableInt(3, roundTripped.Pos1);
			AssertNullableInt(3, roundTripped.Pos2);
			PropertySet? props = roundTripped.Props;
			Assert.NotNull(props);
			Assert.Empty(props!);
		}

		[Fact]
		public void AnnotateOp_TypeDiscriminator_IsTwo()
		{
			string json = SharedStringOpSerializer.Serialize(new MergeTreeAnnotateMsg()
			{
				Pos1 = 0,
				Pos2 = 1,
				Props = new PropertySet(),
			});

			Assert.Contains("\"type\":2", json);
			using JsonDocument document = JsonDocument.Parse(json);
			Assert.Equal(2, document.RootElement.GetProperty("type").GetInt32());
		}

		[Fact]
		public void AnnotateOp_Deserialize_FromKnownJson()
		{
			MergeTreeAnnotateMsg op = Assert.IsType<MergeTreeAnnotateMsg>(
				SharedStringOpSerializer.Deserialize("{\"type\":2,\"pos1\":2,\"pos2\":6,\"props\":{\"color\":\"blue\",\"weight\":700,\"italic\":true}}"));

			AssertNullableInt(2, op.Pos1);
			AssertNullableInt(6, op.Pos2);
			PropertySet? props = op.Props;
			Assert.NotNull(props);
			Assert.Equal(3, props!.Count);
			Assert.Equal("blue", Assert.IsType<string>(props["color"]));
			Assert.Equal(700, Assert.IsType<int>(props["weight"]));
			Assert.True(Assert.IsType<bool>(props["italic"]));
		}

		[Fact]
		public void AnnotateOp_ComplexProps_RoundTrip()
		{
			MergeTreeAnnotateMsg op = new MergeTreeAnnotateMsg()
			{
				Pos1 = 4,
				Pos2 = 8,
				Props = new PropertySet()
				{
					["tags"] = new object?[] { "a", 3, false },
					["style"] = new PropertySet()
					{
						["name"] = "heading",
						["level"] = 2,
					},
				},
			};

			MergeTreeAnnotateMsg roundTripped = RoundTrip(op);

			PropertySet? props = roundTripped.Props;
			Assert.NotNull(props);
			List<object?> tags = Assert.IsType<List<object?>>(props!["tags"]);
			Assert.Equal("a", Assert.IsType<string>(tags[0]));
			Assert.Equal(3, Assert.IsType<int>(tags[1]));
			Assert.False(Assert.IsType<bool>(tags[2]));
			PropertySet style = Assert.IsType<PropertySet>(props["style"]);
			Assert.Equal("heading", Assert.IsType<string>(style["name"]));
			Assert.Equal(2, Assert.IsType<int>(style["level"]));
		}

		private static MergeTreeAnnotateMsg RoundTrip(MergeTreeAnnotateMsg op)
		{
			return Assert.IsType<MergeTreeAnnotateMsg>(
				SharedStringOpSerializer.Deserialize(SharedStringOpSerializer.Serialize(op)));
		}

		private static void AssertNullableInt(int expected, int? actual)
		{
			Assert.True(actual.HasValue);
			Assert.Equal(expected, actual.Value);
		}
	}
}
