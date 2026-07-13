// -----------------------------------------------------------------------------
// Wave 2c tests for DirectorySnapshotLoader. Structure-inspired by TS
// packages/dds/map/src/test/mocha/directory.snapshot.spec.ts; test bodies target
// the C# port's Option B pending-change model.
// -----------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class DirectorySnapshotLoaderTests
	{
		[Fact]
		public void Parse_EmptySnapshot_ReturnsEmptyDto()
		{
			DirectorySnapshotDto dto = DirectorySnapshotLoader.Parse("{}");

			Assert.Empty(dto.Storage);
			Assert.Empty(dto.Subdirectories);
		}

		[Fact]
		public void Parse_SimpleStorageSnapshot_PopulatesStorage()
		{
			const string json = "{\"storage\":{\"alpha\":{\"type\":\"Plain\",\"value\":\"one\"},\"beta\":{\"type\":\"Plain\",\"value\":\"2\"}}}";

			DirectorySnapshotDto dto = DirectorySnapshotLoader.Parse(json);

			Assert.Equal(2, dto.Storage.Count);
			Assert.Equal("Plain", dto.Storage["alpha"].Type);
			Assert.Equal("one", dto.Storage["alpha"].Value);
			Assert.Equal("2", dto.Storage["beta"].Value);
			Assert.Empty(dto.Subdirectories);
		}

		[Fact]
		public void Parse_NestedSnapshot_PopulatesRecursiveTree()
		{
			const string json = "{\"storage\":{\"root\":{\"type\":\"Plain\",\"value\":\"r\"}},\"subdirectories\":{\"foo\":{\"storage\":{\"child\":{\"type\":\"Plain\",\"value\":\"c\"}},\"subdirectories\":{\"bar\":{\"storage\":{\"leaf\":{\"type\":\"Plain\",\"value\":\"l\"}}}}}}}}";

			DirectorySnapshotDto dto = DirectorySnapshotLoader.Parse(json);

			Assert.Equal("r", dto.Storage["root"].Value);
			DirectorySnapshotDto foo = dto.Subdirectories["foo"];
			Assert.Equal("c", foo.Storage["child"].Value);
			DirectorySnapshotDto bar = foo.Subdirectories["bar"];
			Assert.Equal("l", bar.Storage["leaf"].Value);
		}

		[Fact]
		public void Parse_UnknownTopLevelField_IgnoresField()
		{
			const string json = "{\"storage\":{},\"ci\":{\"csn\":42}}";

			DirectorySnapshotDto dto = DirectorySnapshotLoader.Parse(json);

			Assert.Empty(dto.Storage);
			Assert.Empty(dto.Subdirectories);
		}

		[Fact]
		public void Parse_MalformedStorageEntry_ThrowsInvalidOperation()
		{
			const string json = "{\"storage\":{\"broken\":123}}";

			OcsException exception = Assert.Throws<OcsException>(() => DirectorySnapshotLoader.Parse(json));

			Assert.Equal(OcsGateErrorCode.InvalidOperation, exception.ErrorCode);
		}

		[Fact]
		public void Load_BlobSplitFormat_NullResolver_Throws()
		{
			const string json = "{\"blobs\":[\"blob0\"],\"content\":{}}";

			OcsException exception = Assert.Throws<OcsException>(() => DirectorySnapshotLoader.Load(json));

			Assert.Equal(OcsGateErrorCode.InvalidOperation, exception.ErrorCode);
			Assert.Contains("blob resolver", exception.Message);
		}

		[Fact]
		public void Load_DetectsBlobSplitFormat_ReadsMainContent()
		{
			const string json = "{\"blobs\":[],\"content\":{\"storage\":{\"a\":{\"type\":\"Plain\",\"value\":\"1\"}}}}";

			SharedDirectory directory = LoadDirectory(json, UnexpectedBlobResolver);

			Assert.Equal("1", directory.Get("a"));
			Assert.Equal(new[] { "a" }, directory.Keys.ToArray());
			Assert.Equal(0, directory.CountSubDirectory());
		}

		[Fact]
		public void Load_MergesSingleBlob_IntoMainTree()
		{
			const string json = "{\"blobs\":[\"blob0\"],\"content\":{\"storage\":{\"a\":{\"type\":\"Plain\",\"value\":1}}}}";
			var blobs = new Dictionary<string, string>()
			{
				["blob0"] = "{\"storage\":{\"b\":{\"type\":\"Plain\",\"value\":2}}}",
			};

			SharedDirectory directory = LoadDirectory(json, BlobResolver(blobs));

			AssertJsonNumber(1, directory.Get("a"));
			AssertJsonNumber(2, directory.Get("b"));
			Assert.Equal(new[] { "a", "b" }, directory.Keys.ToArray());
		}

		[Fact]
		public void Load_MergesMultipleBlobs_InOrder()
		{
			const string json = "{\"blobs\":[\"blob0\",\"blob1\"],\"content\":{\"storage\":{\"a\":{\"type\":\"Plain\",\"value\":\"1\"}},\"subdirectories\":{\"shared\":{\"storage\":{\"main\":{\"type\":\"Plain\",\"value\":\"m\"}}}}}}";
			var blobs = new Dictionary<string, string>()
			{
				["blob0"] = "{\"storage\":{\"b\":{\"type\":\"Plain\",\"value\":\"2\"}},\"subdirectories\":{\"shared\":{\"storage\":{\"first\":{\"type\":\"Plain\",\"value\":\"blob0\"}}}}}",
				["blob1"] = "{\"subdirectories\":{\"shared\":{\"storage\":{\"second\":{\"type\":\"Plain\",\"value\":\"blob1\"}},\"subdirectories\":{\"leaf\":{\"storage\":{\"z\":{\"type\":\"Plain\",\"value\":\"deep\"}}}}}}}",
			};

			SharedDirectory directory = LoadDirectory(json, BlobResolver(blobs));

			Assert.Equal("1", directory.Get("a"));
			Assert.Equal("2", directory.Get("b"));
			IDirectory shared = Assert.IsAssignableFrom<IDirectory>(directory.GetSubDirectory("shared"));
			Assert.Equal("m", shared.Get("main"));
			Assert.Equal("blob0", shared.Get("first"));
			Assert.Equal("blob1", shared.Get("second"));
			IDirectory leaf = Assert.IsAssignableFrom<IDirectory>(shared.GetSubDirectory("leaf"));
			Assert.Equal("deep", leaf.Get("z"));
		}

		[Fact]
		public void Load_OverlappingKeys_LastBlobWins()
		{
			const string json = "{\"blobs\":[\"blob0\",\"blob1\"],\"content\":{\"storage\":{\"k\":{\"type\":\"Plain\",\"value\":\"main\"}}}}";
			var blobs = new Dictionary<string, string>()
			{
				["blob0"] = "{\"storage\":{\"k\":{\"type\":\"Plain\",\"value\":\"first\"}}}",
				["blob1"] = "{\"storage\":{\"k\":{\"type\":\"Plain\",\"value\":\"blob\"}}}",
			};

			SharedDirectory directory = LoadDirectory(json, BlobResolver(blobs));

			Assert.Equal("blob", directory.Get("k"));
			Assert.Equal(1, directory.Count);
		}

		[Fact]
		public void Load_NestedSubdirInBlob_MergesCorrectly()
		{
			const string json = "{\"blobs\":[\"blob0\"],\"content\":{}}";
			var blobs = new Dictionary<string, string>()
			{
				["blob0"] = "{\"subdirectories\":{\"deep\":{\"subdirectories\":{\"nested\":{\"storage\":{\"path\":{\"type\":\"Plain\",\"value\":\"value\"}}}}}}}",
			};
			var directory = new SharedDirectory();

			directory.LoadFromSnapshot(json, BlobResolver(blobs));

			IDirectory nested = Assert.IsAssignableFrom<IDirectory>(directory.GetWorkingDirectory("/deep/nested"));
			Assert.Equal("value", nested.Get("path"));
		}

		[Fact]
		public void LoadFromSnapshot_EmptyTree_HydratesDirectory()
		{
			const string json = "{\"storage\":{\"root\":{\"type\":\"Plain\",\"value\":\"r\"}},\"subdirectories\":{\"foo\":{\"storage\":{\"child\":{\"type\":\"Plain\",\"value\":\"c\"}},\"subdirectories\":{\"bar\":{\"storage\":{\"leaf\":{\"type\":\"Plain\",\"value\":\"l\"}}}}}}}}";
			DirectorySnapshotDto dto = DirectorySnapshotLoader.Parse(json);
			var directory = new SharedDirectory();

			directory.LoadFromSnapshot(dto);

			Assert.Equal("r", directory.Get("root"));
			Assert.Contains("root", directory.Keys);
			Assert.Equal(1, directory.CountSubDirectory());
			IDirectory foo = Assert.IsAssignableFrom<IDirectory>(directory.GetSubDirectory("foo"));
			Assert.Equal("c", foo.Get("child"));
			IDirectory bar = Assert.IsAssignableFrom<IDirectory>(foo.GetSubDirectory("bar"));
			Assert.Equal("l", bar.Get("leaf"));
			Assert.Equal(new[] { "foo" }, directory.SubDirectories().Select(entry => entry.Key).ToArray());
		}

		private static SharedDirectory LoadDirectory(string json, Func<string, string>? blobResolver = null)
		{
			DirectorySnapshotDto dto = DirectorySnapshotLoader.Load(json, blobResolver);
			var directory = new SharedDirectory();
			directory.LoadFromSnapshot(dto);
			return directory;
		}

		private static Func<string, string> BlobResolver(IReadOnlyDictionary<string, string> blobs)
		{
			return blobName => blobs[blobName];
		}

		private static string UnexpectedBlobResolver(string blobName)
		{
			throw new InvalidOperationException($"Unexpected blob resolver call for '{blobName}'.");
		}

		private static void AssertJsonNumber(int expected, object? actual)
		{
			JsonElement element = Assert.IsType<JsonElement>(actual);
			Assert.Equal(JsonValueKind.Number, element.ValueKind);
			Assert.Equal(expected, element.GetInt32());
		}
	}
}
