// -----------------------------------------------------------------------------
// Feasibility-demo tests for SharedDirectory.
//
// Covers the in-memory API surface delivered by Wave 1:
//   - set/get/has/delete/keys/count/clear
//   - createSubDirectory/getSubDirectory/hasSubDirectory/deleteSubDirectory
//   - countSubDirectory/subDirectories creation-order iteration
//   - getWorkingDirectory posix-style path navigation
//   - OnValueChanged / OnSubDirectoryCreated / OnSubDirectoryDeleted events
//
// Op processing, snapshot loading, and pending-change tracking are
// intentionally out of scope for the demo and not tested here.
// -----------------------------------------------------------------------------

using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class SharedDirectoryBasicsTests
	{
		[Fact]
		public void Ctor_AssignsId_And_RootAbsolutePathIsSlash()
		{
			var dir = new SharedDirectory();
			Assert.False(string.IsNullOrEmpty(dir.Id));
			Assert.Equal("/", dir.AbsolutePath);
		}

		[Fact]
		public void Ctor_AcceptsExplicitId()
		{
			var dir = new SharedDirectory("test-id");
			Assert.Equal("test-id", dir.Id);
		}

		// -------------------------------------------------------------
		// Key/value APIs on the root
		// -------------------------------------------------------------

		[Fact]
		public void Set_And_Get_RoundTripsAValue()
		{
			var dir = new SharedDirectory();
			dir.Set("k", "hello");
			Assert.Equal("hello", dir.Get("k"));
		}

		[Fact]
		public void Set_ReturnsIDirectoryForChaining()
		{
			var dir = new SharedDirectory();
			IDirectory returned = dir.Set("k", 1).Set("k2", 2);
			Assert.Equal(1, returned.Get("k"));
			Assert.Equal(2, returned.Get("k2"));
		}

		[Fact]
		public void Has_ReturnsTrueForSetKey_And_FalseForMissing()
		{
			var dir = new SharedDirectory();
			dir.Set("present", 42);
			Assert.True(dir.Has("present"));
			Assert.False(dir.Has("absent"));
		}

		[Fact]
		public void Delete_RemovesKey_AndReturnsTrue_WhenPresent()
		{
			var dir = new SharedDirectory();
			dir.Set("k", "v");
			bool removed = dir.Delete("k");
			Assert.True(removed);
			Assert.False(dir.Has("k"));
			Assert.Null(dir.Get("k"));
		}

		[Fact]
		public void Delete_ReturnsFalse_WhenKeyMissing()
		{
			var dir = new SharedDirectory();
			Assert.False(dir.Delete("nope"));
		}

		[Fact]
		public void Clear_RemovesAllKeys()
		{
			var dir = new SharedDirectory();
			dir.Set("a", 1);
			dir.Set("b", 2);
			dir.Clear();
			Assert.False(dir.Has("a"));
			Assert.False(dir.Has("b"));
			Assert.Equal(0, dir.Count);
		}

		[Fact]
		public void Keys_ContainsSetKeys()
		{
			var dir = new SharedDirectory();
			dir.Set("a", 1);
			dir.Set("b", 2);
			var keys = new HashSet<string>(dir.Keys);
			Assert.Contains("a", keys);
			Assert.Contains("b", keys);
		}

		[Fact]
		public void Values_ContainsSetValues()
		{
			var dir = new SharedDirectory();
			dir.Set("a", 1);
			dir.Set("b", "two");
			var values = new HashSet<object?>(dir.Values);
			Assert.Contains(1, values);
			Assert.Contains("two", values);
		}

		[Fact]
		public void Entries_ReturnsAllKeyValuePairs()
		{
			var dir = new SharedDirectory();
			dir.Set("a", 1);
			dir.Set("b", 2);
			var entries = dir.Entries().ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
			Assert.Equal(2, entries.Count);
			Assert.Equal(1, entries["a"]);
			Assert.Equal(2, entries["b"]);
		}

		[Fact]
		public void Foreach_IteratesAllEntries()
		{
			var dir = new SharedDirectory();
			dir.Set("a", 1);
			dir.Set("b", 2);
			var collected = new Dictionary<string, object?>();
			foreach (KeyValuePair<string, object?> kvp in dir)
			{
				collected[kvp.Key] = kvp.Value;
			}
			Assert.Equal(2, collected.Count);
			Assert.Equal(1, collected["a"]);
			Assert.Equal(2, collected["b"]);
		}

		[Fact]
		public void Values_Entries_ReflectDelete()
		{
			var dir = new SharedDirectory();
			dir.Set("a", 1);
			dir.Set("b", 2);
			dir.Delete("a");
			Assert.DoesNotContain(1, dir.Values);
			Assert.Single(dir.Entries());
		}

		[Fact]
		public void Count_ReflectsNumberOfKeys()
		{
			var dir = new SharedDirectory();
			Assert.Equal(0, dir.Count);
			dir.Set("a", 1);
			dir.Set("b", 2);
			Assert.Equal(2, dir.Count);
			dir.Delete("a");
			Assert.Equal(1, dir.Count);
		}

		// -------------------------------------------------------------
		// Sub-directory APIs
		// -------------------------------------------------------------

		[Fact]
		public void CreateSubDirectory_CreatesChild_WithComposedAbsolutePath()
		{
			var dir = new SharedDirectory();
			IDirectory child = dir.CreateSubDirectory("foo");
			Assert.Equal("/foo", child.AbsolutePath);
			Assert.True(dir.HasSubDirectory("foo"));
		}

		[Fact]
		public void CreateSubDirectory_IsIdempotent_ReturnsExistingChild()
		{
			var dir = new SharedDirectory();
			IDirectory first = dir.CreateSubDirectory("foo");
			first.Set("k", 1);
			IDirectory second = dir.CreateSubDirectory("foo");
			Assert.Same(first, second);
			Assert.Equal(1, second.Get("k"));
		}

		[Fact]
		public void GetSubDirectory_ReturnsNull_WhenMissing()
		{
			var dir = new SharedDirectory();
			Assert.Null(dir.GetSubDirectory("missing"));
		}

		[Fact]
		public void DeleteSubDirectory_RemovesChild()
		{
			var dir = new SharedDirectory();
			dir.CreateSubDirectory("foo");
			Assert.True(dir.HasSubDirectory("foo"));
			bool removed = dir.DeleteSubDirectory("foo");
			Assert.True(removed);
			Assert.False(dir.HasSubDirectory("foo"));
		}

		[Fact]
		public void DeleteSubDirectory_ReturnsFalse_WhenMissing()
		{
			var dir = new SharedDirectory();
			Assert.False(dir.DeleteSubDirectory("nope"));
		}

		[Fact]
		public void CountSubDirectory_ReflectsNumberOfChildren()
		{
			var dir = new SharedDirectory();
			Assert.Equal(0, dir.CountSubDirectory());
			dir.CreateSubDirectory("a");
			dir.CreateSubDirectory("b");
			Assert.Equal(2, dir.CountSubDirectory());
			dir.DeleteSubDirectory("a");
			Assert.Equal(1, dir.CountSubDirectory());
		}

		[Fact]
		public void SubDirectories_IteratesInCreationOrder()
		{
			var dir = new SharedDirectory();
			dir.CreateSubDirectory("first");
			dir.CreateSubDirectory("second");
			dir.CreateSubDirectory("third");
			// Delete-then-recreate should append at the end (new creation).
			dir.DeleteSubDirectory("first");
			dir.CreateSubDirectory("first");

			List<string> ordered = dir.SubDirectories().Select(kvp => kvp.Key).ToList();
			Assert.Equal(new[] { "second", "third", "first" }, ordered);
		}

		[Fact]
		public void SubDirectories_YieldsSameInstance_AsGetSubDirectory()
		{
			var dir = new SharedDirectory();
			IDirectory a = dir.CreateSubDirectory("a");
			KeyValuePair<string, IDirectory> entry = dir.SubDirectories().Single();
			Assert.Same(a, entry.Value);
			Assert.Same(a, dir.GetSubDirectory("a"));
		}

		// -------------------------------------------------------------
		// Working-directory navigation
		// -------------------------------------------------------------

		[Fact]
		public void GetWorkingDirectory_NavigatesRelativePaths()
		{
			var dir = new SharedDirectory();
			IDirectory foo = dir.CreateSubDirectory("foo");
			IDirectory bar = foo.CreateSubDirectory("bar");
			bar.Set("k", "v");

			IDirectory? resolved = dir.GetWorkingDirectory("foo/bar");
			Assert.NotNull(resolved);
			Assert.Same(bar, resolved);
			Assert.Equal("v", resolved!.Get("k"));
		}

		[Fact]
		public void GetWorkingDirectory_NavigatesAbsolutePaths_FromASubdirectory()
		{
			var dir = new SharedDirectory();
			IDirectory foo = dir.CreateSubDirectory("foo");
			IDirectory bar = dir.CreateSubDirectory("bar");
			bar.Set("k", "v");

			IDirectory? resolved = foo.GetWorkingDirectory("/bar");
			Assert.NotNull(resolved);
			Assert.Same(bar, resolved);
		}

		[Fact]
		public void GetWorkingDirectory_ReturnsNull_ForMissingSegment()
		{
			var dir = new SharedDirectory();
			dir.CreateSubDirectory("foo");
			Assert.Null(dir.GetWorkingDirectory("foo/nonexistent"));
		}

		// -------------------------------------------------------------
		// Events
		// -------------------------------------------------------------

		[Fact]
		public void OnValueChanged_FiresOnSet_WithPreviousValue()
		{
			var dir = new SharedDirectory();
			dir.Set("k", "old");

			ValueChangedEventArgs? captured = null;
			dir.OnValueChanged += (s, e) => captured = e;

			dir.Set("k", "new");

			Assert.NotNull(captured);
			Assert.Equal("k", captured!.Key);
			Assert.Equal("old", captured.PreviousValue);
			Assert.Equal("/", captured.Path);
		}

		[Fact]
		public void OnValueChanged_FiresOnDelete_FromNestedSubDirectory()
		{
			var dir = new SharedDirectory();
			IDirectory foo = dir.CreateSubDirectory("foo");
			foo.Set("k", "v");

			ValueChangedEventArgs? captured = null;
			dir.OnValueChanged += (s, e) => captured = e;

			foo.Delete("k");

			Assert.NotNull(captured);
			Assert.Equal("k", captured!.Key);
			Assert.Equal("v", captured.PreviousValue);
			Assert.Equal("/foo", captured.Path);
		}

		[Fact]
		public void OnSubDirectoryCreated_FiresOnCreate()
		{
			var dir = new SharedDirectory();
			SubDirectoryEventArgs? captured = null;
			dir.OnSubDirectoryCreated += (s, e) => captured = e;

			dir.CreateSubDirectory("foo");

			Assert.NotNull(captured);
			Assert.Equal("foo", captured!.SubdirName);
			Assert.Equal("/", captured.ParentPath);
			// TS ref: directory.ts emits (relativePath, local, target) where relativePath
			// is the subdir name at the raising SubDirectory. Path mirrors that argument.
			Assert.Equal("foo", captured.Path);
		}

		[Fact]
		public void OnSubDirectoryCreated_DoesNotFire_OnIdempotentReCreate()
		{
			var dir = new SharedDirectory();
			dir.CreateSubDirectory("foo");

			int fireCount = 0;
			dir.OnSubDirectoryCreated += (s, e) => fireCount++;

			// Second call returns the existing child; should not fire again.
			dir.CreateSubDirectory("foo");

			Assert.Equal(0, fireCount);
		}

		[Fact]
		public void OnSubDirectoryDeleted_FiresOnDelete()
		{
			var dir = new SharedDirectory();
			dir.CreateSubDirectory("foo");

			SubDirectoryEventArgs? captured = null;
			dir.OnSubDirectoryDeleted += (s, e) => captured = e;

			dir.DeleteSubDirectory("foo");

			Assert.NotNull(captured);
			Assert.Equal("foo", captured!.SubdirName);
			Assert.Equal("/", captured.ParentPath);
			Assert.Equal("foo", captured.Path);
		}

		// TODO(w2c): op processing coverage lives in DirectoryOpProcessingTests.cs.
	}
}
