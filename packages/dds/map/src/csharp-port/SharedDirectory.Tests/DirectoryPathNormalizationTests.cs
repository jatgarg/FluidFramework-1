// -----------------------------------------------------------------------------
// Path-normalization regression tests for SharedDirectory.
//
// Covers findings SD-W05, SD-W06, and SD-T02 from the independent audit
// (INDEPENDENT-AUDIT.md):
//   - MakeChildAbsolutePath must match posix.join semantics for '.', '..',
//     and other special names (not just literal concatenation).
//   - Incoming op paths must be normalized via posix.resolve('/', path)
//     before walking, so non-canonical wire paths still target the right
//     directory.
//   - GetWorkingDirectory must handle '.', '..', absolute reset, repeated
//     slashes, and attempts to walk above root the same way as TS
//     posix.resolve.
// TS ref: packages/dds/map/src/directory.ts
// -----------------------------------------------------------------------------

using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class DirectoryPathNormalizationTests
	{
		// -------------------------------------------------------------
		// PosixPath primitive tests (Join + ResolveAbsolute + Normalize)
		// -------------------------------------------------------------

		[Theory]
		// Trivial cases
		[InlineData("/", "a", "/a")]
		[InlineData("/", "", "/")]
		[InlineData("/a", "b", "/a/b")]
		[InlineData("/a/b", "c", "/a/b/c")]
		// Dot segments
		[InlineData("/", ".", "/")]
		[InlineData("/a", ".", "/a")]
		[InlineData("/a", "./b", "/a/b")]
		// Parent-directory segments (up-navigation)
		[InlineData("/a", "..", "/")]
		[InlineData("/a/b", "..", "/a")]
		[InlineData("/a/b", "../c", "/a/c")]
		[InlineData("/a", "../..", "/")] // walking above root stays at root
		// Consecutive slashes collapse
		[InlineData("/a", "//b", "/a/b")]
		[InlineData("/", "//", "/")]
		public void PosixPath_Join_MatchesNodePosixJoin(string parent, string child, string expected)
		{
			Assert.Equal(expected, PosixPath.Join(parent, child));
		}

		[Theory]
		[InlineData("", "/")]
		[InlineData("/", "/")]
		[InlineData("a", "/a")]
		[InlineData("/a", "/a")]
		[InlineData("a/b", "/a/b")]
		[InlineData("/a/b", "/a/b")]
		// Dot / dot-dot resolution
		[InlineData("/a/./b", "/a/b")]
		[InlineData("/a/../b", "/b")]
		[InlineData("/../a", "/a")] // extra '..' above root drop
		[InlineData("/a/b/..", "/a")]
		[InlineData("/./a", "/a")]
		// Consecutive slashes
		[InlineData("//a//b//", "/a/b")]
		public void PosixPath_ResolveAbsolute_MatchesNodePosixResolveRoot(string input, string expected)
		{
			Assert.Equal(expected, PosixPath.ResolveAbsolute(input));
		}

		// -------------------------------------------------------------
		// SD-W05: CreateSubDirectory with special names emits canonical path
		// -------------------------------------------------------------

		[Fact]
		public void CreateSubDirectory_DotName_CanonicalizesToRoot()
		{
			// TS ref: posix.join("/", ".") == "/". Literal concatenation would produce "/."
			// and put a wire op with path:"/." on the outbound channel.
			var dir = new SharedDirectory();
			IDirectory child = dir.CreateSubDirectory(".");

			// posix.join("/", ".") == "/" — dotting into root returns the root itself.
			Assert.Equal("/", child.AbsolutePath);
		}

		[Fact]
		public void CreateSubDirectory_DotDotName_CanonicalizesToParent()
		{
			// TS ref: posix.join("/a", "..") == "/". So '..' at root collapses to root.
			var dir = new SharedDirectory();
			IDirectory a = dir.CreateSubDirectory("a");
			IDirectory upFromA = a.CreateSubDirectory("..");

			Assert.Equal("/", upFromA.AbsolutePath);
		}

		[Fact]
		public void CreateSubDirectory_ChainedDotSegments_Canonicalize()
		{
			var dir = new SharedDirectory();
			IDirectory a = dir.CreateSubDirectory("a");
			IDirectory b = a.CreateSubDirectory("b");
			Assert.Equal("/a/b", b.AbsolutePath);
		}

		// -------------------------------------------------------------
		// SD-W06: incoming op paths are normalized before walking
		// -------------------------------------------------------------

		[Fact]
		public void RemoteSet_NonCanonicalPathWithDotDot_TargetsResolvedDirectory()
		{
			// TS ref: posix.resolve('/', '/a/../b') == '/b'. A remote op with a
			// non-canonical path must resolve to the canonical directory before
			// applying, otherwise the walk looks for a child literally named '..'.
			var dir = new SharedDirectory();
			dir.CreateSubDirectory("a");
			dir.CreateSubDirectory("b");
			string op = "{\"type\":\"set\",\"path\":\"/a/../b\",\"key\":\"k\",\"value\":{\"type\":\"Plain\",\"value\":\"v\"}}";

			dir.ProcessDataObjectOp(RemoteMessage(refSeq: 0, seq: 1), op);

			IDirectory? bDir = dir.GetSubDirectory("b");
			Assert.NotNull(bDir);
			Assert.Equal("v", bDir!.Get("k"));

			// Value should NOT have landed in "a" (the naive-walk destination).
			IDirectory? aDir = dir.GetSubDirectory("a");
			Assert.NotNull(aDir);
			Assert.Null(aDir!.Get("k"));
		}

		[Fact]
		public void RemoteSet_PathWithDotSegments_TargetsResolvedDirectory()
		{
			var dir = new SharedDirectory();
			dir.CreateSubDirectory("a");
			string op = "{\"type\":\"set\",\"path\":\"/./a\",\"key\":\"k\",\"value\":{\"type\":\"Plain\",\"value\":\"v\"}}";

			dir.ProcessDataObjectOp(RemoteMessage(refSeq: 0, seq: 1), op);

			IDirectory? aDir = dir.GetSubDirectory("a");
			Assert.NotNull(aDir);
			Assert.Equal("v", aDir!.Get("k"));
		}

		[Fact]
		public void RemoteSet_PathWithConsecutiveSlashes_TargetsResolvedDirectory()
		{
			var dir = new SharedDirectory();
			IDirectory a = dir.CreateSubDirectory("a");
			a.CreateSubDirectory("b");
			string op = "{\"type\":\"set\",\"path\":\"/a//b\",\"key\":\"k\",\"value\":{\"type\":\"Plain\",\"value\":\"v\"}}";

			dir.ProcessDataObjectOp(RemoteMessage(refSeq: 0, seq: 1), op);

			IDirectory? bDir = dir.GetWorkingDirectory("/a/b");
			Assert.NotNull(bDir);
			Assert.Equal("v", bDir!.Get("k"));
		}

		// -------------------------------------------------------------
		// SD-T02: GetWorkingDirectory posix.resolve coverage
		// -------------------------------------------------------------

		[Fact]
		public void GetWorkingDirectory_DotSegments_Normalize()
		{
			var dir = new SharedDirectory();
			IDirectory a = dir.CreateSubDirectory("a");
			IDirectory b = a.CreateSubDirectory("b");

			Assert.Same(a, dir.GetWorkingDirectory("/a/."));
			Assert.Same(b, dir.GetWorkingDirectory("/a/./b"));
			Assert.Same(a, dir.GetWorkingDirectory("/./a"));
		}

		[Fact]
		public void GetWorkingDirectory_DotDotSegments_WalkUp()
		{
			var dir = new SharedDirectory();
			IDirectory a = dir.CreateSubDirectory("a");
			IDirectory b = a.CreateSubDirectory("b");
			IDirectory c = a.CreateSubDirectory("c");

			// From b, "../c" should land at c (sibling of b).
			Assert.Same(c, b.GetWorkingDirectory("../c"));
			// From b, ".." should land at a.
			Assert.Same(a, b.GetWorkingDirectory(".."));
		}

		[Fact]
		public void GetWorkingDirectory_AbsolutePath_ResetsFromCurrent()
		{
			// TS ref: posix.resolve(cwd, "/absolute") drops cwd entirely.
			var dir = new SharedDirectory();
			IDirectory a = dir.CreateSubDirectory("a");
			IDirectory b = dir.CreateSubDirectory("b");
			IDirectory nested = a.CreateSubDirectory("nested");

			// From a nested directory, an absolute path should reset to root.
			Assert.Same(b, nested.GetWorkingDirectory("/b"));
		}

		[Fact]
		public void GetWorkingDirectory_RepeatedSlashes_Collapse()
		{
			var dir = new SharedDirectory();
			IDirectory a = dir.CreateSubDirectory("a");
			IDirectory b = a.CreateSubDirectory("b");

			Assert.Same(b, dir.GetWorkingDirectory("//a//b"));
			Assert.Same(a, dir.GetWorkingDirectory("//a"));
		}

		[Fact]
		public void GetWorkingDirectory_WalkAboveRoot_StaysAtRoot()
		{
			// TS ref: posix.resolve('/', '/../..') == '/'. Excess '..' segments above
			// root are silently dropped rather than producing a null or error.
			var dir = new SharedDirectory();
			IDirectory? result = dir.GetWorkingDirectory("/../..");
			Assert.NotNull(result);
			Assert.Equal("/", result!.AbsolutePath);
		}

		[Fact]
		public void GetWorkingDirectory_NonExistent_ReturnsNull()
		{
			var dir = new SharedDirectory();
			Assert.Null(dir.GetWorkingDirectory("/does/not/exist"));
			Assert.Null(dir.GetWorkingDirectory("/./does/../not/exist"));
		}

		// -------------------------------------------------------------
		// Helpers
		// -------------------------------------------------------------

		private static SequencedDocumentMessageDescriptor RemoteMessage(long refSeq, long seq)
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: 0, refSeq: refSeq, seq: seq),
				OpOrigin.Remote,
				clientId: "remote-client");
		}
	}
}
