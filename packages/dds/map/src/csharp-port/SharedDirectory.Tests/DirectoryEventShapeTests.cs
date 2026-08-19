// -----------------------------------------------------------------------------
// Event-shape regression tests for SharedDirectory. TS refs are inline.
//
// Areas covered:
//   - Nested subdirectory events bubble a joined relative path (matching TS
//     posix.join(subDirName, relativePath) in directory.ts).
//   - Remote delete of an absent key still emits valueChanged with
//     previousValue null (matching TS previousValue: undefined behavior).
//   - OnValueChanged on a SubDirectory fires only on the direct container
//     (matching TS containedValueChanged semantics), not through ancestors.
//     SharedDirectory-level OnValueChanged continues to see events from
//     anywhere.
// -----------------------------------------------------------------------------

using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class DirectoryEventShapeTests
	{
		// -------------------------------------------------------------
		// Nested subdirectory events bubble the joined path
		// -------------------------------------------------------------

		[Fact]
		public void OnSubDirectoryCreated_BubblesJoinedPath()
		{
			// TS ref: directory.ts re-emits with posix.join(subDirName, relativePath).
			// So when grandchild is created under /parent/child, the parent listener sees
			// Path == "child/grandchild" (relative to itself), not just "grandchild".
			var dir = new SharedDirectory();
			IDirectory child = dir.CreateSubDirectory("child");

			SubDirectoryEventArgs? capturedOnChild = null;
			child.OnSubDirectoryCreated += (s, e) => capturedOnChild = e;

			SubDirectoryEventArgs? capturedOnRoot = null;
			dir.OnSubDirectoryCreated += (s, e) => capturedOnRoot = e;

			child.CreateSubDirectory("grandchild");

			Assert.NotNull(capturedOnChild);
			Assert.Equal("grandchild", capturedOnChild!.Path);
			Assert.Equal("grandchild", capturedOnChild.SubdirName);

			Assert.NotNull(capturedOnRoot);
			// SharedDirectory listener sees the fully-joined absolute-relative path.
			Assert.Equal("child/grandchild", capturedOnRoot!.Path);
			Assert.Equal("grandchild", capturedOnRoot.SubdirName);
		}

		[Fact]
		public void OnSubDirectoryCreated_ThreeLevelsDeep_BubblesFullJoinedPath()
		{
			// TS ref: successive posix.join calls chain the path segments.
			var dir = new SharedDirectory();
			IDirectory a = dir.CreateSubDirectory("a");
			IDirectory b = a.CreateSubDirectory("b");

			SubDirectoryEventArgs? capturedOnA = null;
			a.OnSubDirectoryCreated += (s, e) => capturedOnA = e;

			SubDirectoryEventArgs? capturedOnRoot = null;
			dir.OnSubDirectoryCreated += (s, e) => capturedOnRoot = e;

			b.CreateSubDirectory("c");

			// a sees Path == "b/c" (relative to a).
			Assert.NotNull(capturedOnA);
			Assert.Equal("b/c", capturedOnA!.Path);

			// Root sees Path == "a/b/c".
			Assert.NotNull(capturedOnRoot);
			Assert.Equal("a/b/c", capturedOnRoot!.Path);
		}

		[Theory]
		[InlineData(".")]
		[InlineData("..")]
		[InlineData("")]
		public void OnSubDirectoryCreated_BubblesLiteralLocalName_WhenAncestorNameNormalizesAway(string ancestorName)
		{
			// DirectoryPath.Join normalizes ".", "..", and "" away, so an ancestor
			// subdirectory created under one of those names would have an
			// AbsolutePath equal to its parent's. The bubble path must use the
			// literal local name (the parent-map key) so listeners still see
			// the caller-supplied segment.
			var dir = new SharedDirectory();
			IDirectory ancestor = dir.CreateSubDirectory(ancestorName);

			SubDirectoryEventArgs? capturedOnRoot = null;
			dir.OnSubDirectoryCreated += (s, e) => capturedOnRoot = e;

			ancestor.CreateSubDirectory("leaf");

			Assert.NotNull(capturedOnRoot);
			Assert.Equal(DirectoryPath.Join(ancestorName, "leaf"), capturedOnRoot!.Path);
		}

		[Fact]
		public void OnSubDirectoryDeleted_BubblesJoinedPath()
		{
			// TS ref: directory.ts mirrors subDirectoryCreated with posix.join.
			var dir = new SharedDirectory();
			IDirectory child = dir.CreateSubDirectory("child");
			child.CreateSubDirectory("grandchild");

			SubDirectoryEventArgs? capturedOnChild = null;
			child.OnSubDirectoryDeleted += (s, e) => capturedOnChild = e;

			SubDirectoryEventArgs? capturedOnRoot = null;
			dir.OnSubDirectoryDeleted += (s, e) => capturedOnRoot = e;

			child.DeleteSubDirectory("grandchild");

			Assert.NotNull(capturedOnChild);
			Assert.Equal("grandchild", capturedOnChild!.Path);

			Assert.NotNull(capturedOnRoot);
			Assert.Equal("child/grandchild", capturedOnRoot!.Path);
		}

		// -------------------------------------------------------------
		// Remote delete of absent key emits valueChanged event
		// -------------------------------------------------------------

		[Fact]
		public void RemoteDelete_AbsentKey_EmitsValueChangedEvent()
		{
			// TS ref: directory.ts. Even when the key isn't present locally, TS
			// emits valueChanged with previousValue: undefined (subject to
			// pending-op suppression).
			var dir = new SharedDirectory();

			ValueChangedEventArgs? captured = null;
			dir.OnValueChanged += (s, e) => captured = e;

			// Send a remote delete for a key that was never set.
			string op = "{\"type\":\"delete\",\"path\":\"/\",\"key\":\"never-existed\"}";
			dir.ProcessDataObjectOp(RemoteMessage(refSeq: 0, seq: 1), op);

			Assert.NotNull(captured);
			Assert.Equal("never-existed", captured!.Key);
			Assert.Null(captured.PreviousValue);
			Assert.Equal("/", captured.Path);
			Assert.False(captured.Local);
		}

		[Fact]
		public void RemoteDelete_PresentKey_EmitsValueChangedWithPreviousValue()
		{
			// Sanity: the present-key path is unchanged.
			var dir = new SharedDirectory();
			dir.Set("k", "v");

			ValueChangedEventArgs? captured = null;
			dir.OnValueChanged += (s, e) => captured = e;

			string op = "{\"type\":\"delete\",\"path\":\"/\",\"key\":\"k\"}";
			dir.ProcessDataObjectOp(RemoteMessage(refSeq: 0, seq: 1), op);

			Assert.NotNull(captured);
			Assert.Equal("k", captured!.Key);
			Assert.Equal("v", captured.PreviousValue);
		}

		// -------------------------------------------------------------
		// OnValueChanged does NOT bubble through ancestor subdirs
		// -------------------------------------------------------------

		[Fact]
		public void OnValueChanged_OnSubDirectory_FiresOnlyOnDirectContainer()
		{
			// TS ref: directory.ts. containedValueChanged fires only on the
			// SubDirectory that directly contains the key. valueChanged fires on
			// the root SharedDirectory. It does NOT bubble through intermediate
			// ancestor subdirectories.
			var dir = new SharedDirectory();
			IDirectory a = dir.CreateSubDirectory("a");
			IDirectory b = a.CreateSubDirectory("b");

			int rootHits = 0;
			int aHits = 0;
			int bHits = 0;
			dir.OnValueChanged += (s, e) => rootHits++;
			a.OnValueChanged += (s, e) => aHits++;
			b.OnValueChanged += (s, e) => bHits++;

			// Set a key on `b`. Only b's OnValueChanged should fire (direct container),
			// plus root's OnValueChanged (SharedDirectory sees all).
			b.Set("k", "v");

			Assert.Equal(1, bHits);
			Assert.Equal(0, aHits);   // Ancestor subdir must NOT receive the event.
			Assert.Equal(1, rootHits);
		}

		[Fact]
		public void OnValueChanged_RemoteSet_FiresOnlyOnDirectContainerAndRoot()
		{
			var dir = new SharedDirectory();
			IDirectory a = dir.CreateSubDirectory("a");
			IDirectory b = a.CreateSubDirectory("b");

			int rootHits = 0;
			int aHits = 0;
			int bHits = 0;
			dir.OnValueChanged += (s, e) => rootHits++;
			a.OnValueChanged += (s, e) => aHits++;
			b.OnValueChanged += (s, e) => bHits++;

			string op = "{\"type\":\"set\",\"path\":\"/a/b\",\"key\":\"k\",\"value\":{\"type\":\"Plain\",\"value\":\"v\"}}";
			dir.ProcessDataObjectOp(RemoteMessage(refSeq: 0, seq: 1), op);

			Assert.Equal(1, bHits);
			Assert.Equal(0, aHits);
			Assert.Equal(1, rootHits);
		}

		[Fact]
		public void OnValueChanged_OnRootDirectly_FiresOnceForRootKeys()
		{
			// Sanity: setting at root still hits both root SubDirectory's OnValueChanged
			// (as the direct container is root itself) and SharedDirectory's OnValueChanged.
			// These are two separate event sources so both should fire.
			var dir = new SharedDirectory();

			int rootHits = 0;
			dir.OnValueChanged += (s, e) => rootHits++;

			dir.Set("k", "v");

			// Root fires once because dir (the SharedDirectory) IS the propagation target,
			// AND it exposes OnValueChanged for its root SubDirectory. Set at root produces
			// one event on the SharedDirectory-level.
			Assert.Equal(1, rootHits);
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
